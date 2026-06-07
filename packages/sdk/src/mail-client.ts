import type { GraphClient } from "./graph-client.js";
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
} from "./types.js";
import { GRAPH_BASE_URL } from "./config.js";
import { renderBody } from "./body.js";

const SUMMARY_FIELDS =
  "id,subject,from,toRecipients,receivedDateTime,sentDateTime,isRead,isDraft,conversationId";

function defaultSelect(mode: ListBodyMode): string {
  if (mode === "full") return `${SUMMARY_FIELDS},body`;
  if (mode === "preview") return `bodyPreview,${SUMMARY_FIELDS}`;
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
}
