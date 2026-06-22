import type { GraphClient, ODataListResponse } from "./graph-client.js";
import { CreateMessageRuleParamsSchema } from "./types.js";
import type {
  MessageRule,
  CreateMessageRuleParams,
  UpdateMessageRuleParams,
} from "./types.js";

const RULES_PATH = "/me/mailFolders/inbox/messageRules";

export class RulesClient {
  constructor(private readonly graph: GraphClient) {}

  async list(): Promise<MessageRule[]> {
    const res = await this.graph.list<MessageRule>(RULES_PATH);
    return res.value;
  }

  async get(id: string): Promise<MessageRule> {
    return this.graph.get<MessageRule>(`${RULES_PATH}/${encodeURIComponent(id)}`);
  }

  async create(params: CreateMessageRuleParams): Promise<MessageRule> {
    const body = CreateMessageRuleParamsSchema.parse(params);
    return this.graph.post<typeof body, MessageRule>(RULES_PATH, body);
  }

  async update(id: string, patch: UpdateMessageRuleParams): Promise<void> {
    await this.graph.patch(`${RULES_PATH}/${encodeURIComponent(id)}`, patch);
  }

  async delete(id: string): Promise<void> {
    await this.graph.delete(`${RULES_PATH}/${encodeURIComponent(id)}`);
  }
}

export type { ODataListResponse };
