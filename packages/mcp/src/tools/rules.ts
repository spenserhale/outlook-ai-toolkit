import type { FastMCP } from "fastmcp";
import { z } from "zod";
import {
  resolveConfig,
  OutlookAuth,
  TokenStore,
  GraphClient,
  MailClient,
  RulesClient,
  OutlookAuthError,
  OutlookConfigError,
  renderOutput,
  MessageRulePredicatesSchema,
  MessageRuleActionsSchema,
  SweepConditionSchema,
  OutputFormatSchema,
} from "@outlook-toolkit/sdk";

async function getToken(): Promise<string> {
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
    return await auth.acquireToken();
  } catch (err) {
    if (err instanceof OutlookAuthError) {
      throw new Error(
        `Not authenticated. Run \`outlook auth login\` in your terminal with OUTLOOK_CLIENT_ID=${config.clientId} set.`
      );
    }
    throw err;
  }
}

async function getRulesClient(): Promise<RulesClient> {
  return new RulesClient(new GraphClient(await getToken()));
}

async function getMailClient(): Promise<MailClient> {
  return new MailClient(new GraphClient(await getToken()));
}

export function registerRulesTools(server: FastMCP) {
  server.addTool({
    name: "outlook_mail_rules_list",
    description:
      "List the authenticated user's Outlook inbox rules (messageRules). Rules act on incoming mail only — to act on existing mail use outlook_mail_mass_archive / outlook_mail_mass_delete.",
    parameters: z.object({
      format: OutputFormatSchema.default("toon").describe("Output encoding (toon is cheaper for LLMs)"),
    }),
    execute: async (args) => {
      const rules = await getRulesClient();
      return renderOutput(await rules.list(), args.format);
    },
  });

  server.addTool({
    name: "outlook_mail_rules_get",
    description: "Get a single Outlook inbox rule by ID.",
    parameters: z.object({
      id: z.string().describe("Rule ID"),
      format: OutputFormatSchema.default("toon"),
    }),
    execute: async (args) => {
      const rules = await getRulesClient();
      return renderOutput(await rules.get(args.id), args.format);
    },
  });

  server.addTool({
    name: "outlook_mail_rules_create",
    description:
      "Create an Outlook inbox rule for FUTURE incoming mail (e.g. from X -> move to Deleted Items). Provide at least one action. Does not affect existing mail.",
    parameters: z.object({
      displayName: z.string().describe("Rule display name"),
      conditions: MessageRulePredicatesSchema.optional().describe(
        "Match conditions, e.g. { senderContains: ['alice@x.com'] }"
      ),
      actions: MessageRuleActionsSchema.describe(
        "Actions, e.g. { moveToFolder: 'deleteditems', stopProcessingRules: true } or { delete: true }"
      ),
      sequence: z.number().int().default(1).describe("Rule order"),
      isEnabled: z.boolean().default(true),
      format: OutputFormatSchema.default("toon"),
    }),
    execute: async (args) => {
      const rules = await getRulesClient();
      const created = await rules.create({
        displayName: args.displayName,
        sequence: args.sequence,
        isEnabled: args.isEnabled,
        conditions: args.conditions,
        actions: args.actions,
      });
      return renderOutput(created, args.format);
    },
  });

  server.addTool({
    name: "outlook_mail_rules_update",
    description: "Update an existing Outlook inbox rule (enable/disable, reorder, or replace conditions/actions).",
    parameters: z.object({
      id: z.string().describe("Rule ID"),
      displayName: z.string().optional(),
      sequence: z.number().int().optional(),
      isEnabled: z.boolean().optional(),
      conditions: MessageRulePredicatesSchema.optional(),
      actions: MessageRuleActionsSchema.optional(),
      format: OutputFormatSchema.default("toon"),
    }),
    execute: async (args) => {
      const rules = await getRulesClient();
      const { id, format, ...patch } = args;
      if (Object.keys(patch).length === 0) {
        throw new Error(
          "nothing to update: provide at least one of displayName, sequence, isEnabled, conditions, actions"
        );
      }
      await rules.update(id, patch);
      return renderOutput({ status: "updated", ruleId: id }, format);
    },
  });

  server.addTool({
    name: "outlook_mail_rules_delete",
    description: "Delete an Outlook inbox rule by ID.",
    parameters: z.object({
      id: z.string().describe("Rule ID"),
      format: OutputFormatSchema.default("toon"),
    }),
    execute: async (args) => {
      const rules = await getRulesClient();
      await rules.delete(args.id);
      return renderOutput({ status: "deleted", ruleId: args.id }, args.format);
    },
  });

  const massParams = z.object({
    conditions: z
      .array(SweepConditionSchema)
      .min(1)
      .describe(
        "OR-list of conditions. Each: { from?, subjectContains?, bodyContains?, olderThanDays? }. Within a condition, fields AND together."
      ),
    folder: z.string().default("inbox").describe("Source folder to sweep"),
    max: z.number().int().positive().max(1000).default(200).describe("Max messages to move"),
    dryRun: z.boolean().default(false).describe("Preview matches without moving"),
    format: OutputFormatSchema.default("toon"),
  });

  server.addTool({
    name: "outlook_mail_mass_archive",
    description:
      "Move EXISTING mail matching the given conditions from a folder (default inbox) to Archive. This is the retroactive counterpart to an inbox rule. Use dryRun first to preview.",
    parameters: massParams,
    execute: async (args) => {
      const mail = await getMailClient();
      const result = await mail.massMove({
        conditions: args.conditions,
        destination: "archive",
        folder: args.folder,
        max: args.max,
        dryRun: args.dryRun,
      });
      return renderOutput(result, args.format);
    },
  });

  server.addTool({
    name: "outlook_mail_mass_delete",
    description:
      "Move EXISTING mail matching the given conditions from a folder (default inbox) to Deleted Items (recoverable). Retroactive counterpart to an inbox rule. Use dryRun first to preview.",
    parameters: massParams,
    execute: async (args) => {
      const mail = await getMailClient();
      const result = await mail.massMove({
        conditions: args.conditions,
        destination: "deleteditems",
        folder: args.folder,
        max: args.max,
        dryRun: args.dryRun,
      });
      return renderOutput(result, args.format);
    },
  });
}
