import { buildCommand, buildRouteMap } from "@stricli/core";
import { encode } from "@toon-format/toon";
import {
  OutlookAuth,
  TokenStore,
  GraphClient,
  MailClient,
} from "@outlook-toolkit/sdk";
import { resolveCliConfig } from "../context.js";

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
      json: { kind: "boolean", brief: "Output as JSON", default: false },
      csv: { kind: "boolean", brief: "Output as CSV", default: false },
    },
  },
  async func(this: void, flags: { profile?: string; folder?: string; limit?: number; cursor?: string; json: boolean; csv: boolean }) {
    const mail = await getMailClient(flags.profile);
    const result = await mail.list({
      folder: flags.folder ?? "inbox",
      limit: flags.limit ?? 25,
      cursor: flags.cursor,
    });

    if (flags.json) {
      console.log(JSON.stringify(result, null, 2));
    } else if (flags.csv) {
      console.log("id,subject,from,receivedDateTime,isRead");
      for (const m of result.value) {
        const from = m.from?.emailAddress?.address ?? "";
        console.log(`${m.id},${JSON.stringify(m.subject ?? "")},${from},${m.receivedDateTime ?? ""},${m.isRead ?? ""}`);
      }
    } else {
      console.log(encode(result, { keyFolding: "safe" }));
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
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Message ID", parse: String }],
    },
  },
  async func(this: void, flags: { profile?: string; json: boolean }, id: string) {
    const mail = await getMailClient(flags.profile);
    const message = await mail.get(id);
    flags.json
      ? console.log(JSON.stringify(message, null, 2))
      : console.log(encode(message, { keyFolding: "safe" }));
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
      flags.json ? console.log(JSON.stringify(out, null, 2)) : console.log(encode(out, { keyFolding: "safe" }));
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
      flags.json ? console.log(JSON.stringify(out, null, 2)) : console.log(encode(out, { keyFolding: "safe" }));
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
      flags.json ? console.log(JSON.stringify(out, null, 2)) : console.log(encode(out, { keyFolding: "safe" }));
      return;
    }
    const mail = await getMailClient(flags.profile);
    const draft = await mail.createDraft({ to: flags.to, subject: flags.subject, body: flags.body, contentType: "HTML" });
    flags.json
      ? console.log(JSON.stringify(draft, null, 2))
      : console.log(encode(draft, { keyFolding: "safe" }));
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
    flags.json
      ? console.log(JSON.stringify(result, null, 2))
      : console.log(encode(result, { keyFolding: "safe" }));
    if (result["@odata.deltaLink"]) {
      process.stderr.write(`\ndeltaLink: ${result["@odata.deltaLink"]}\n`);
    }
  },
});

export const mailRoutes = buildRouteMap({
  routes: { list: listCommand, get: getCommand, send: sendCommand, reply: replyCommand, draft: draftCommand, sync: syncCommand },
  docs: { brief: "Read and send Outlook mail" },
});
