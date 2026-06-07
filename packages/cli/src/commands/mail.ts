import { buildCommand, buildRouteMap } from "@stricli/core";
import {
  OutlookAuth,
  TokenStore,
  GraphClient,
  MailClient,
  renderOutput,
  type BodyFormat,
  type ListBodyMode,
} from "@outlook-toolkit/sdk";
import { resolveCliConfig } from "../context.js";

function parseBodyFormat(s: string): BodyFormat {
  if (s === "text" || s === "markdown" || s === "html") return s;
  throw new Error(`--bodyFormat must be one of: text, markdown, html (got: "${s}")`);
}

function parseListBody(s: string): ListBodyMode {
  if (s === "none" || s === "preview" || s === "full") return s;
  throw new Error(`--body must be one of: none, preview, full (got: "${s}")`);
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

export const mailRoutes = buildRouteMap({
  routes: { list: listCommand, get: getCommand, send: sendCommand, reply: replyCommand, draft: draftCommand, sync: syncCommand },
  docs: { brief: "Read and send Outlook mail" },
});
