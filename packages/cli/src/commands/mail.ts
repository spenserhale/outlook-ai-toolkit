import { buildCommand, buildRouteMap } from "@stricli/core";
import {
  OutlookAuth,
  TokenStore,
  GraphClient,
  MailClient,
  RulesClient,
  renderOutput,
  BodyFormatSchema,
  ListBodyModeSchema,
  type BodyFormat,
  type ListBodyMode,
  type CreateMessageRuleParams,
  type UpdateMessageRuleParams,
  type MessageRulePredicates,
  type MessageRuleActions,
} from "@outlook-toolkit/sdk";
import { resolveCliConfig } from "../context.js";

function parseBodyFormat(s: string): BodyFormat {
  const r = BodyFormatSchema.safeParse(s);
  if (r.success) return r.data;
  throw new Error(`--bodyFormat must be one of: ${BodyFormatSchema.options.join(", ")} (got: "${s}")`);
}

function parseListBody(s: string): ListBodyMode {
  const r = ListBodyModeSchema.safeParse(s);
  if (r.success) return r.data;
  throw new Error(`--body must be one of: ${ListBodyModeSchema.options.join(", ")} (got: "${s}")`);
}

async function getMailClient(profile?: string): Promise<MailClient> {
  const config = await resolveCliConfig(profile);
  const store = new TokenStore(config.clientId);
  const auth = new OutlookAuth(config, store);
  let token: string;
  try {
    token = await auth.acquireToken();
  } catch {
    console.error("error: not authenticated (exit code 5). Run: outlook auth login");
    process.exit(5);
  }
  return new MailClient(new GraphClient(token));
}

async function getRulesClient(profile?: string): Promise<RulesClient> {
  const config = await resolveCliConfig(profile);
  const store = new TokenStore(config.clientId);
  const auth = new OutlookAuth(config, store);
  let token: string;
  try {
    token = await auth.acquireToken();
  } catch {
    console.error("error: not authenticated (exit code 5). Run: outlook auth login");
    process.exit(5);
  }
  return new RulesClient(new GraphClient(token));
}

function csvToList(s: string | undefined): string[] | undefined {
  if (!s) return undefined;
  const items = s.split(",").map((x) => x.trim()).filter(Boolean);
  return items.length ? items : undefined;
}

const listCommand = buildCommand({
  docs: { brief: "List messages in a mail folder" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      folder: { kind: "parsed", brief: "Folder (default: inbox)", parse: String, optional: true },
      limit: { kind: "parsed", brief: "Max messages (default: 25)", parse: Number, optional: true },
      cursor: { kind: "parsed", brief: "Pagination cursor ($skipToken URL)", parse: String, optional: true },
      body: { kind: "parsed", brief: "Body: none|preview|full (default: preview)", parse: parseListBody, optional: true },
      bodyFormat: { kind: "parsed", brief: "Body format: text|markdown|html (default: text)", parse: parseBodyFormat, optional: true },
      toon: { kind: "boolean", brief: "Output as TOON (default)", default: true },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
      csv: { kind: "boolean", brief: "Output as CSV", default: false },
    },
  },
  async func(
    this: void,
    flags: {
      profile?: string;
      folder?: string;
      limit?: number;
      cursor?: string;
      body?: ListBodyMode;
      bodyFormat?: BodyFormat;
      toon: boolean;
      json: boolean;
      csv: boolean;
    }
  ) {
    const mail = await getMailClient(flags.profile);
    const result = await mail.list({
      folder: flags.folder ?? "inbox",
      limit: flags.limit ?? 25,
      cursor: flags.cursor,
      body: flags.body ?? "preview",
      bodyFormat: flags.bodyFormat ?? "text",
    });

    if (flags.csv) {
      console.log("id,subject,from,receivedDateTime,isRead");
      for (const m of result.value) {
        const from = m.from?.emailAddress?.address ?? "";
        console.log(`${m.id},${JSON.stringify(m.subject ?? "")},${from},${m.receivedDateTime ?? ""},${m.isRead ?? ""}`);
      }
    } else {
      console.log(renderOutput(result, flags.json ? "json" : "toon"));
    }

    if (result["@odata.nextLink"]) {
      process.stderr.write(`\nMore results available. Use --cursor=<nextLink> to continue.\n`);
    }
  },
});

const getCommand = buildCommand({
  docs: { brief: "Get a single message by ID" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      bodyFormat: { kind: "parsed", brief: "Body format: text|markdown|html (default: text)", parse: parseBodyFormat, optional: true },
      toon: { kind: "boolean", brief: "Output as TOON (default)", default: true },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Message ID", parse: String }],
    },
  },
  async func(
    this: void,
    flags: { profile?: string; bodyFormat?: BodyFormat; toon: boolean; json: boolean },
    id: string
  ) {
    const mail = await getMailClient(flags.profile);
    const message = await mail.get(id, { bodyFormat: flags.bodyFormat ?? "text" });
    console.log(renderOutput(message, flags.json ? "json" : "toon"));
  },
});

const sendCommand = buildCommand({
  docs: { brief: "Send an email" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      to: { kind: "parsed", brief: "Recipient email address", parse: String },
      subject: { kind: "parsed", brief: "Email subject", parse: String },
      body: { kind: "parsed", brief: "Email body (HTML)", parse: String },
      dryRun: { kind: "boolean", brief: "Validate without sending", default: false },
      json: { kind: "boolean", brief: "Output result as JSON", default: false },
    },
  },
  async func(this: void, flags: { profile?: string; to: string; subject: string; body: string; dryRun: boolean; json: boolean }) {
    if (flags.dryRun) {
      const out = { status: "dry_run", to: flags.to, subject: flags.subject };
      console.log(renderOutput(out, flags.json ? "json" : "toon"));
      return;
    }
    const mail = await getMailClient(flags.profile);
    await mail.send({ to: flags.to, subject: flags.subject, body: flags.body, contentType: "HTML" });
    const out = { status: "sent", to: flags.to };
    flags.json ? console.log(JSON.stringify(out, null, 2)) : console.log("Sent.");
  },
});

const replyCommand = buildCommand({
  docs: { brief: "Reply to a message" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      body: { kind: "parsed", brief: "Reply body (HTML)", parse: String },
      dryRun: { kind: "boolean", brief: "Validate without sending", default: false },
      json: { kind: "boolean", brief: "Output result as JSON", default: false },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Message ID to reply to", parse: String }],
    },
  },
  async func(this: void, flags: { profile?: string; body: string; dryRun: boolean; json: boolean }, id: string) {
    if (flags.dryRun) {
      const out = { status: "dry_run", replyTo: id };
      console.log(renderOutput(out, flags.json ? "json" : "toon"));
      return;
    }
    const mail = await getMailClient(flags.profile);
    await mail.reply(id, { body: flags.body, contentType: "HTML" });
    const out = { status: "replied", messageId: id };
    flags.json ? console.log(JSON.stringify(out, null, 2)) : console.log("Reply sent.");
  },
});

const draftCommand = buildCommand({
  docs: { brief: "Create a draft email" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      to: { kind: "parsed", brief: "Recipient email address", parse: String },
      subject: { kind: "parsed", brief: "Email subject", parse: String },
      body: { kind: "parsed", brief: "Email body (HTML)", parse: String },
      dryRun: { kind: "boolean", brief: "Validate without creating", default: false },
      json: { kind: "boolean", brief: "Output draft as JSON", default: false },
    },
  },
  async func(this: void, flags: { profile?: string; to: string; subject: string; body: string; dryRun: boolean; json: boolean }) {
    if (flags.dryRun) {
      const out = { status: "dry_run", to: flags.to, subject: flags.subject };
      console.log(renderOutput(out, flags.json ? "json" : "toon"));
      return;
    }
    const mail = await getMailClient(flags.profile);
    const draft = await mail.createDraft({ to: flags.to, subject: flags.subject, body: flags.body, contentType: "HTML" });
    console.log(renderOutput(draft, flags.json ? "json" : "toon"));
  },
});

const syncCommand = buildCommand({
  docs: { brief: "Delta sync inbox (only changed messages since last sync)" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      deltaLink: { kind: "parsed", brief: "Delta link from previous sync", parse: String, optional: true },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
  },
  async func(this: void, flags: { profile?: string; deltaLink?: string; json: boolean }) {
    const mail = await getMailClient(flags.profile);
    const result = await mail.sync(flags.deltaLink);
    console.log(renderOutput(result, flags.json ? "json" : "toon"));
    if (result["@odata.deltaLink"]) {
      process.stderr.write(`\ndeltaLink: ${result["@odata.deltaLink"]}\n`);
    }
  },
});

const rulesListCommand = buildCommand({
  docs: { brief: "List inbox rules (messageRules)" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
  },
  async func(this: void, flags: { profile?: string; json: boolean }) {
    const rules = await getRulesClient(flags.profile);
    const result = await rules.list();
    console.log(renderOutput(result, flags.json ? "json" : "toon"));
  },
});

const rulesGetCommand = buildCommand({
  docs: { brief: "Get an inbox rule by ID" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
    positional: { kind: "tuple", parameters: [{ brief: "Rule ID", parse: String }] },
  },
  async func(this: void, flags: { profile?: string; json: boolean }, id: string) {
    const rules = await getRulesClient(flags.profile);
    const rule = await rules.get(id);
    console.log(renderOutput(rule, flags.json ? "json" : "toon"));
  },
});

const rulesCreateCommand = buildCommand({
  docs: { brief: "Create an inbox rule (e.g. from X -> move to Deleted Items)" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      name: { kind: "parsed", brief: "Rule display name", parse: String },
      fromContains: { kind: "parsed", brief: "Sender contains (comma-separated)", parse: String, optional: true },
      subjectContains: { kind: "parsed", brief: "Subject contains (comma-separated)", parse: String, optional: true },
      moveTo: { kind: "parsed", brief: "Destination folder id/well-known name", parse: String, optional: true },
      delete: { kind: "boolean", brief: "Action: move to Deleted Items", default: false },
      markRead: { kind: "boolean", brief: "Action: mark as read", default: false },
      stopProcessing: { kind: "boolean", brief: "Action: stop processing further rules", default: false },
      sequence: { kind: "parsed", brief: "Rule order (default: 1)", parse: Number, optional: true },
      disabled: { kind: "boolean", brief: "Create the rule disabled", default: false },
      conditions: { kind: "parsed", brief: "Raw conditions JSON (overrides convenience flags)", parse: String, optional: true },
      actions: { kind: "parsed", brief: "Raw actions JSON (overrides convenience flags)", parse: String, optional: true },
      dryRun: { kind: "boolean", brief: "Validate without creating", default: false },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
  },
  async func(
    this: void,
    flags: {
      profile?: string; name: string;
      fromContains?: string; subjectContains?: string;
      moveTo?: string; delete: boolean; markRead: boolean; stopProcessing: boolean;
      sequence?: number; disabled: boolean;
      conditions?: string; actions?: string;
      dryRun: boolean; json: boolean;
    }
  ) {
    let conditions: MessageRulePredicates | undefined;
    let actions: MessageRuleActions;
    try {
      conditions = flags.conditions
        ? (JSON.parse(flags.conditions) as MessageRulePredicates)
        : {
            ...(csvToList(flags.fromContains) && { senderContains: csvToList(flags.fromContains) }),
            ...(csvToList(flags.subjectContains) && { subjectContains: csvToList(flags.subjectContains) }),
          };
      if (conditions && Object.keys(conditions).length === 0) conditions = undefined;

      actions = flags.actions
        ? (JSON.parse(flags.actions) as MessageRuleActions)
        : {
            ...(flags.moveTo && { moveToFolder: flags.moveTo }),
            ...(flags.delete && { delete: true }),
            ...(flags.markRead && { markAsRead: true }),
            ...(flags.stopProcessing && { stopProcessingRules: true }),
          };
    } catch (err) {
      console.error(`error: invalid JSON for --conditions/--actions (exit code 2): ${(err as Error).message}`);
      process.exit(2);
    }

    if (!actions || Object.keys(actions).length === 0) {
      console.error("error: a rule needs at least one action (--moveTo, --delete, --markRead, or --actions) (exit code 2)");
      process.exit(2);
    }

    const params: CreateMessageRuleParams = {
      displayName: flags.name,
      sequence: flags.sequence ?? 1,
      isEnabled: !flags.disabled,
      ...(conditions && { conditions }),
      actions,
    };

    if (flags.dryRun) {
      console.log(renderOutput({ status: "dry_run", rule: params }, flags.json ? "json" : "toon"));
      return;
    }

    const rules = await getRulesClient(flags.profile);
    const created = await rules.create(params);
    console.log(renderOutput(created, flags.json ? "json" : "toon"));
  },
});

const rulesUpdateCommand = buildCommand({
  docs: { brief: "Update an inbox rule (enable/disable, reorder, replace conditions/actions)" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      name: { kind: "parsed", brief: "New display name", parse: String, optional: true },
      sequence: { kind: "parsed", brief: "New order", parse: Number, optional: true },
      enable: { kind: "boolean", brief: "Enable the rule", default: false },
      disable: { kind: "boolean", brief: "Disable the rule", default: false },
      conditions: { kind: "parsed", brief: "Raw conditions JSON", parse: String, optional: true },
      actions: { kind: "parsed", brief: "Raw actions JSON", parse: String, optional: true },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
    positional: { kind: "tuple", parameters: [{ brief: "Rule ID", parse: String }] },
  },
  async func(
    this: void,
    flags: {
      profile?: string; name?: string; sequence?: number;
      enable: boolean; disable: boolean; conditions?: string; actions?: string; json: boolean;
    },
    id: string
  ) {
    const patch: UpdateMessageRuleParams = {};
    try {
      if (flags.name !== undefined) patch.displayName = flags.name;
      if (flags.sequence !== undefined) patch.sequence = flags.sequence;
      if (flags.enable) patch.isEnabled = true;
      if (flags.disable) patch.isEnabled = false;
      if (flags.conditions) patch.conditions = JSON.parse(flags.conditions) as MessageRulePredicates;
      if (flags.actions) patch.actions = JSON.parse(flags.actions) as MessageRuleActions;
    } catch (err) {
      console.error(`error: invalid JSON for --conditions/--actions (exit code 2): ${(err as Error).message}`);
      process.exit(2);
    }
    if (Object.keys(patch).length === 0) {
      console.error("error: nothing to update (exit code 2)");
      process.exit(2);
    }
    const rules = await getRulesClient(flags.profile);
    await rules.update(id, patch);
    console.log(renderOutput({ status: "updated", ruleId: id }, flags.json ? "json" : "toon"));
  },
});

const rulesDeleteCommand = buildCommand({
  docs: { brief: "Delete an inbox rule by ID" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      dryRun: { kind: "boolean", brief: "Validate without deleting", default: false },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
    positional: { kind: "tuple", parameters: [{ brief: "Rule ID", parse: String }] },
  },
  async func(this: void, flags: { profile?: string; dryRun: boolean; json: boolean }, id: string) {
    if (flags.dryRun) {
      console.log(renderOutput({ status: "dry_run", ruleId: id }, flags.json ? "json" : "toon"));
      return;
    }
    const rules = await getRulesClient(flags.profile);
    await rules.delete(id);
    console.log(renderOutput({ status: "deleted", ruleId: id }, flags.json ? "json" : "toon"));
  },
});

const rulesRoutes = buildRouteMap({
  routes: {
    list: rulesListCommand,
    get: rulesGetCommand,
    create: rulesCreateCommand,
    update: rulesUpdateCommand,
    delete: rulesDeleteCommand,
  },
  docs: { brief: "Manage inbox rules (messageRules)" },
});

export const mailRoutes = buildRouteMap({
  routes: { list: listCommand, get: getCommand, send: sendCommand, reply: replyCommand, draft: draftCommand, sync: syncCommand, rules: rulesRoutes },
  docs: { brief: "Read and send Outlook mail" },
});
