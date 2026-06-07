import type { FastMCP } from "fastmcp";
import { z } from "zod";
import {
  resolveConfig,
  OutlookAuth,
  TokenStore,
  GraphClient,
  MailClient,
  OutlookAuthError,
  OutlookConfigError,
  renderOutput,
  BodyFormatSchema,
  ListBodyModeSchema,
  OutputFormatSchema,
} from "@outlook-toolkit/sdk";

async function getMailClient(): Promise<MailClient> {
  let config;
  try {
    config = resolveConfig();
  } catch (err) {
    if (err instanceof OutlookConfigError) {
      throw new Error(
        "Outlook not configured. Set OUTLOOK_CLIENT_ID and OUTLOOK_TENANT_ID environment variables."
      );
    }
    throw err;
  }
  const store = new TokenStore(config.clientId);
  const auth = new OutlookAuth(config, store);
  try {
    const token = await auth.acquireToken();
    return new MailClient(new GraphClient(token));
  } catch (err) {
    if (err instanceof OutlookAuthError) {
      throw new Error(
        `Not authenticated. Run \`outlook auth login\` in your terminal with OUTLOOK_CLIENT_ID=${config.clientId} set.`
      );
    }
    throw err;
  }
}

export function registerMailTools(server: FastMCP) {
  server.addTool({
    name: "outlook_mail_list",
    description:
      "List messages in an Outlook mail folder (default: inbox). Bodies are omitted by default (body=preview keeps a short snippet, body=full returns the converted body). Returns a nextLink cursor for pagination.",
    parameters: z.object({
      folder: z.string().default("inbox").describe("Folder name (inbox, sentitems, drafts, deleteditems)"),
      limit: z.number().int().positive().max(999).default(25).describe("Max messages to return"),
      cursor: z.string().optional().describe("Pagination cursor from a previous call's nextLink"),
      filter: z.string().optional().describe("OData $filter expression"),
      select: z.string().optional().describe("Comma-separated fields to return (overrides body shaping)"),
      orderby: z.string().optional().describe("OData orderby expression (e.g. \"receivedDateTime desc\")"),
      body: ListBodyModeSchema.default("preview").describe("How much body each row carries"),
      bodyFormat: BodyFormatSchema.default("text").describe("Body format when body=full"),
      format: OutputFormatSchema.default("toon").describe("Output encoding (toon is cheaper for LLMs)"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      const result = await mail.list({
        folder: args.folder,
        limit: args.limit,
        cursor: args.cursor,
        filter: args.filter,
        select: args.select,
        orderby: args.orderby,
        body: args.body,
        bodyFormat: args.bodyFormat,
      });
      return renderOutput(result, args.format);
    },
  });

  server.addTool({
    name: "outlook_mail_get",
    description: "Get a single Outlook message by ID, including the full body (converted to text by default).",
    parameters: z.object({
      id: z.string().describe("Message ID"),
      bodyFormat: BodyFormatSchema.default("text").describe("Body format: text (default), markdown, or raw html"),
      format: OutputFormatSchema.default("toon").describe("Output encoding (toon is cheaper for LLMs)"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      const message = await mail.get(args.id, { bodyFormat: args.bodyFormat });
      return renderOutput(message, args.format);
    },
  });

  server.addTool({
    name: "outlook_mail_send",
    description: "Send an email from the authenticated Outlook account.",
    parameters: z.object({
      to: z.string().email().describe("Recipient email address"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body (HTML supported)"),
      contentType: z.enum(["text", "HTML"]).default("HTML").describe("Body content type"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      await mail.send({ to: args.to, subject: args.subject, body: args.body, contentType: args.contentType });
      return renderOutput({ status: "sent", to: args.to }, "toon");
    },
  });

  server.addTool({
    name: "outlook_mail_reply",
    description: "Reply to an existing Outlook message thread.",
    parameters: z.object({
      id: z.string().describe("Message ID to reply to"),
      body: z.string().describe("Reply body (HTML supported)"),
      contentType: z.enum(["text", "HTML"]).default("HTML").describe("Body content type"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      await mail.reply(args.id, { body: args.body, contentType: args.contentType });
      return renderOutput({ status: "replied", messageId: args.id }, "toon");
    },
  });

  server.addTool({
    name: "outlook_mail_create_draft",
    description: "Create a draft email without sending it. Returns the draft message including its ID.",
    parameters: z.object({
      to: z.string().email().describe("Recipient email address"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body (HTML supported)"),
      contentType: z.enum(["text", "HTML"]).default("HTML").describe("Body content type"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      const draft = await mail.createDraft({ to: args.to, subject: args.subject, body: args.body, contentType: args.contentType });
      return renderOutput(draft, "toon");
    },
  });

  server.addTool({
    name: "outlook_mail_sync",
    description:
      "Delta sync inbox — returns only messages that changed since the last sync. On first call, omit deltaLink to get the full initial sync. Save the returned deltaLink and pass it on subsequent calls to get only changes.",
    parameters: z.object({
      deltaLink: z.string().optional().describe("Delta link from a previous sync call. Omit for initial full sync."),
      format: OutputFormatSchema.default("toon").describe("Output encoding (toon is cheaper for LLMs)"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      const result = await mail.sync(args.deltaLink);
      return renderOutput(result, args.format);
    },
  });
}
