import type { GraphClient } from "./graph-client.js";
import type { ODataListResponse } from "./graph-client.js";
import type {
  Message,
  MailListResponse,
  DeltaResponse,
  ListMailParams,
  SendMailParams,
  ReplyParams,
  DraftParams,
} from "./types.js";
import { GRAPH_BASE_URL } from "./config.js";

export class MailClient {
  constructor(private readonly graph: GraphClient) {}

  async list(params: Partial<ListMailParams> = {}): Promise<MailListResponse> {
    const folder = params.folder ?? "inbox";
    const opts = {
      $top: params.limit ?? 25,
      ...(params.filter && { $filter: params.filter }),
      ...(params.select && { $select: params.select }),
      ...(params.orderby && { $orderby: params.orderby }),
    };
    const result = await this.graph.list<Message>(
      `/me/mailFolders/${folder}/messages`,
      opts
    );
    return result as MailListResponse;
  }

  async get(id: string): Promise<Message> {
    return this.graph.get<Message>(`/me/messages/${id}`);
  }

  async send(params: SendMailParams): Promise<void> {
    await this.graph.post("/me/sendMail", {
      message: {
        subject: params.subject,
        body: {
          contentType: params.contentType ?? "HTML",
          content: params.body,
        },
        toRecipients: [{ emailAddress: { address: params.to } }],
      },
      saveToSentItems: true,
    });
  }

  async reply(id: string, params: ReplyParams): Promise<void> {
    await this.graph.post(`/me/messages/${id}/reply`, {
      message: {
        body: {
          contentType: params.contentType ?? "HTML",
          content: params.body,
        },
      },
    });
  }

  async createDraft(params: DraftParams): Promise<Message> {
    return this.graph.post<object, Message>("/me/messages", {
      subject: params.subject,
      body: {
        contentType: params.contentType ?? "HTML",
        content: params.body,
      },
      toRecipients: [{ emailAddress: { address: params.to } }],
      isDraft: true,
    });
  }

  async sync(deltaLink?: string): Promise<DeltaResponse> {
    // If a deltaLink is provided, use it as the path directly (it's a full URL)
    const path = deltaLink
      ? deltaLink.replace(GRAPH_BASE_URL, "")
      : "/me/mailFolders/inbox/messages/delta";
    const result = await this.graph.list<Message>(path);
    return result as DeltaResponse;
  }
}
