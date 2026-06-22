import type { GraphClient, ODataOptions } from "./graph-client.js";
import type {
  Message,
  MailListResponse,
  DeltaResponse,
  ListMailParams,
  GetMailParams,
  ListBodyMode,
  BodyFormat,
  SendMailParams,
  ReplyParams,
  DraftParams,
  SweepCondition,
  MassMoveParams,
  MassMoveResult,
} from "./types.js";
import { GRAPH_BASE_URL } from "./config.js";
import { renderBody } from "./body.js";
import { buildSearchQuery, isoDaysAgo, isOlderThan } from "./sweep-query.js";

export interface MoveBatchResult {
  moved: number;
  failed: Array<{ id: string; status: number }>;
}

const SUMMARY_FIELDS =
  "id,subject,from,toRecipients,receivedDateTime,sentDateTime,isRead,isDraft,conversationId";

function defaultSelect(mode: ListBodyMode): string {
  if (mode === "full") return `${SUMMARY_FIELDS},body`;
  if (mode === "preview") return `${SUMMARY_FIELDS},bodyPreview`;
  return SUMMARY_FIELDS;
}

function applyListBody(
  resp: MailListResponse,
  mode: ListBodyMode,
  fmt: BodyFormat
): MailListResponse {
  resp.value = resp.value.map((m) => {
    const next: Message = { ...m };
    if (mode === "none") {
      delete next.body;
      delete next.bodyPreview;
    } else if (mode === "preview") {
      delete next.body;
    } else if (mode === "full" && next.body) {
      next.body = {
        contentType: fmt,
        content: renderBody(next.body.content, next.body.contentType, fmt),
      };
    }
    return next;
  });
  return resp;
}

export class MailClient {
  constructor(private readonly graph: GraphClient) {}

  async list(params: Partial<ListMailParams> = {}): Promise<MailListResponse> {
    const mode = params.body ?? "preview";
    const fmt = params.bodyFormat ?? "text";

    if (params.cursor) {
      const result = (await this.graph.list<Message>(params.cursor, {})) as MailListResponse;
      // body mode governs post-processing on every page; callers should thread the same
      // body/bodyFormat through pagination so cursor pages shape consistently with page 1.
      return applyListBody(result, mode, fmt);
    }

    const folder = params.folder ?? "inbox";
    const opts = {
      $top: params.limit ?? 25,
      $select: params.select ?? defaultSelect(mode),
      ...(params.filter && { $filter: params.filter }),
      ...(params.orderby && { $orderby: params.orderby }),
    };
    const result = (await this.graph.list<Message>(
      `/me/mailFolders/${encodeURIComponent(folder)}/messages`,
      opts
    )) as MailListResponse;
    return applyListBody(result, mode, fmt);
  }

  async get(id: string, opts: Partial<GetMailParams> = {}): Promise<Message> {
    const fmt = opts.bodyFormat ?? "text";
    const msg = await this.graph.get<Message>(`/me/messages/${encodeURIComponent(id)}`);
    if (msg.body) {
      msg.body = {
        contentType: fmt,
        content: renderBody(msg.body.content, msg.body.contentType, fmt),
      };
    }
    return msg;
  }

  async send(params: SendMailParams): Promise<void> {
    await this.graph.post("/me/sendMail", {
      message: {
        subject: params.subject,
        body: { contentType: params.contentType ?? "HTML", content: params.body },
        toRecipients: [{ emailAddress: { address: params.to } }],
      },
      saveToSentItems: true,
    });
  }

  async reply(id: string, params: ReplyParams): Promise<void> {
    await this.graph.post(`/me/messages/${encodeURIComponent(id)}/reply`, {
      message: {
        body: { contentType: params.contentType ?? "HTML", content: params.body },
      },
    });
  }

  async createDraft(params: DraftParams): Promise<Message> {
    return this.graph.post<object, Message>("/me/messages", {
      subject: params.subject,
      body: { contentType: params.contentType ?? "HTML", content: params.body },
      toRecipients: [{ emailAddress: { address: params.to } }],
      isDraft: true,
    });
  }

  async sync(deltaLink?: string): Promise<DeltaResponse> {
    const path = deltaLink
      ? deltaLink.replace(GRAPH_BASE_URL, "")
      : "/me/mailFolders/inbox/messages/delta";
    const result = await this.graph.list<Message>(path);
    return result as DeltaResponse;
  }

  async move(id: string, destinationId: string): Promise<void> {
    await this.graph.post(`/me/messages/${encodeURIComponent(id)}/move`, {
      destinationId,
    });
  }

  async moveBatch(ids: string[], destinationId: string): Promise<MoveBatchResult> {
    const result: MoveBatchResult = { moved: 0, failed: [] };
    for (let i = 0; i < ids.length; i += 20) {
      const chunk = ids.slice(i, i + 20);
      const requests = chunk.map((id, j) => ({
        id: String(j),
        method: "POST",
        url: `/me/messages/${encodeURIComponent(id)}/move`,
        headers: { "Content-Type": "application/json" },
        body: { destinationId },
      }));
      const res = await this.graph.post<
        { requests: typeof requests },
        { responses: Array<{ id: string; status: number }> }
      >("/$batch", { requests });
      for (const r of res.responses) {
        const idx = Number(r.id);
        const originalId = Number.isInteger(idx) ? chunk[idx] : undefined;
        if (originalId === undefined) continue; // malformed/unknown batch id — skip
        if (r.status >= 200 && r.status < 300) result.moved++;
        else result.failed.push({ id: originalId, status: r.status });
      }
    }
    return result;
  }

  /**
   * Find messages matching any of `conditions` (OR across conditions), deduped by id,
   * bounded by `max`. Conditions are evaluated in order against a shared `max` budget,
   * so a high-volume earlier condition can exhaust the budget before later ones run.
   * Pagination is capped by a 50-page guard per condition; when combined with
   * `olderThanDays` + keyword search, heavily-filtered pages may hit that guard and
   * return fewer than `max` even when more old matches exist deeper in the result set.
   */
  async findMatches(
    conditions: SweepCondition[],
    folder: string,
    max: number
  ): Promise<Message[]> {
    const seen = new Map<string, Message>();
    const path = `/me/mailFolders/${encodeURIComponent(folder)}/messages`;

    for (const c of conditions) {
      if (seen.size >= max) break;
      const query = buildSearchQuery(c);
      const opts: ODataOptions = {
        $select: "id,subject,from,receivedDateTime",
        $top: Math.min(100, max),
      };
      if (query) {
        opts.$search = `"${query}"`;
      } else if (c.olderThanDays) {
        opts.$filter = `receivedDateTime lt ${isoDaysAgo(c.olderThanDays)}`;
      }

      let cursor: string | undefined;
      let guard = 0;
      while (seen.size < max && guard < 50) {
        guard++;
        const page = cursor
          ? ((await this.graph.list<Message>(
              cursor.replace(GRAPH_BASE_URL, ""),
              {}
            )) as MailListResponse)
          : ((await this.graph.list<Message>(path, opts)) as MailListResponse);

        for (const m of page.value) {
          // olderThanDays + search terms: search can't also date-filter, so filter here.
          if (query && c.olderThanDays && !isOlderThan(m.receivedDateTime, c.olderThanDays)) {
            continue;
          }
          if (!seen.has(m.id)) seen.set(m.id, m);
          if (seen.size >= max) break;
        }
        cursor = page["@odata.nextLink"];
        if (!cursor) break;
      }
    }

    return [...seen.values()].slice(0, max);
  }

  async massMove(params: MassMoveParams): Promise<MassMoveResult> {
    const folder = params.folder ?? "inbox";
    const max = params.max ?? 200;
    const matches = await this.findMatches(params.conditions, folder, max);
    const messages = matches.map((m) => ({
      id: m.id,
      subject: m.subject ?? null,
      from: m.from?.emailAddress?.address,
      receivedDateTime: m.receivedDateTime,
    }));
    const capped = matches.length >= max;

    if (params.dryRun) {
      return {
        destination: params.destination,
        dryRun: true,
        matched: matches.length,
        moved: 0,
        failed: [],
        capped,
        messages,
      };
    }

    const { moved, failed } = await this.moveBatch(
      matches.map((m) => m.id),
      params.destination
    );
    return {
      destination: params.destination,
      dryRun: false,
      matched: matches.length,
      moved,
      failed,
      capped,
      messages,
    };
  }
}
