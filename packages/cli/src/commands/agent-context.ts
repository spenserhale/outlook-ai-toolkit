import { buildCommand } from "@stricli/core";

const AGENT_CONTEXT = {
  schema_version: "1",
  cli: "outlook",
  description: "Outlook toolkit — read and write Outlook mail via Microsoft Graph",
  config: {
    env_vars: {
      OUTLOOK_CLIENT_ID: { required: false, description: "Azure app registration client ID (required for non-interactive use if no profile is saved)" },
      OUTLOOK_TENANT_ID: { required: false, description: "Tenant ID: consumers | common | <guid> (required alongside OUTLOOK_CLIENT_ID)" },
      OUTLOOK_PROFILE: { required: false, description: "Named profile to activate" },
    },
    precedence: ["explicit-flag", "env-var", "default-profile", "guided-setup"],
  },
  first_run: "Run `outlook login` to sign in. Non-interactive callers must set OUTLOOK_CLIENT_ID and OUTLOOK_TENANT_ID (or `outlook profile save`) first.",
  commands: {
    login: { brief: "Sign in (guided setup on first run)", flags: { "--profile": "string (optional)" } },
    logout: { brief: "Sign out and clear saved tokens", flags: { "--profile": "string (optional)" } },
    status: { brief: "Show sign-in status", flags: { "--profile": "string (optional)", "--json": "boolean" } },
    profile: {
      add: { brief: "Add a named profile and sign in (prompts for client ID if interactive)", args: ["name"], flags: { "--clientId": "string (optional)", "--tenantId": "string (default: consumers)" } },
      create: { brief: "Alias of `add`", args: ["name"], flags: { "--clientId": "string (optional)", "--tenantId": "string (default: consumers)" } },
      save: { brief: "Save a profile without signing in (scripting)", args: ["name"], flags: { "--clientId": "string", "--tenantId": "string" } },
      list: { brief: "List profiles", flags: { "--json": "boolean" } },
      delete: { brief: "Delete a profile", args: ["name"], flags: { "--force": "boolean" } },
    },
    mail: {
      list: { brief: "List messages", flags: { "--folder": "string (default: inbox)", "--limit": "number (default: 25)", "--cursor": "string", "--body": "none|preview|full (default: preview)", "--bodyFormat": "text|markdown|html (default: text)", "--toon": "boolean (default: true)", "--json": "boolean", "--csv": "boolean", "--profile": "string (optional)" } },
      get: { brief: "Get message by ID", args: ["id"], flags: { "--bodyFormat": "text|markdown|html (default: text)", "--toon": "boolean (default: true)", "--json": "boolean", "--profile": "string (optional)" } },
      send: { brief: "Send an email", flags: { "--to": "string (email)", "--subject": "string", "--body": "string (HTML)", "--dryRun": "boolean", "--json": "boolean", "--profile": "string (optional)" } },
      reply: { brief: "Reply to a message", args: ["id"], flags: { "--body": "string (HTML)", "--dryRun": "boolean", "--json": "boolean", "--profile": "string (optional)" } },
      draft: { brief: "Create a draft email", flags: { "--to": "string (email)", "--subject": "string", "--body": "string (HTML)", "--dryRun": "boolean", "--json": "boolean", "--profile": "string (optional)" } },
      sync: { brief: "Delta sync inbox", flags: { "--deltaLink": "string (optional)", "--json": "boolean", "--profile": "string (optional)" } },
      rules: {
        list: { brief: "List inbox rules", flags: { "--json": "boolean", "--profile": "string (optional)" } },
        get: { brief: "Get an inbox rule by ID", args: ["id"], flags: { "--json": "boolean", "--profile": "string (optional)" } },
        create: { brief: "Create an inbox rule (from X -> move/delete)", flags: { "--name": "string", "--fromContains": "string (comma-separated)", "--subjectContains": "string (comma-separated)", "--moveTo": "string (folder id)", "--delete": "boolean", "--markRead": "boolean", "--stopProcessing": "boolean", "--sequence": "number (default: 1)", "--disabled": "boolean", "--conditions": "json (overrides flags)", "--actions": "json (overrides flags)", "--dryRun": "boolean", "--json": "boolean", "--profile": "string (optional)" } },
        update: { brief: "Update an inbox rule", args: ["id"], flags: { "--name": "string", "--sequence": "number", "--enable": "boolean", "--disable": "boolean", "--conditions": "json", "--actions": "json", "--json": "boolean", "--profile": "string (optional)" } },
        delete: { brief: "Delete an inbox rule by ID", args: ["id"], flags: { "--dryRun": "boolean", "--json": "boolean", "--profile": "string (optional)" } },
      },
      "mass-archive": { brief: "Move existing inbox mail matching conditions to Archive", flags: { "--from": "string", "--subjectContains": "string", "--bodyContains": "string", "--olderThanDays": "number", "--conditions": "json array (overrides single-condition flags)", "--to": "string (default: archive)", "--folder": "string (default: inbox)", "--max": "number (default: 200)", "--dryRun": "boolean", "--json": "boolean", "--profile": "string (optional)" } },
      "mass-delete": { brief: "Move existing inbox mail matching conditions to Deleted Items", flags: { "--from": "string", "--subjectContains": "string", "--bodyContains": "string", "--olderThanDays": "number", "--conditions": "json array (overrides single-condition flags)", "--to": "string (default: deleteditems)", "--folder": "string (default: inbox)", "--max": "number (default: 200)", "--dryRun": "boolean", "--json": "boolean", "--profile": "string (optional)" } },
    },
    upgrade: { brief: "Upgrade the CLI to the latest release (standalone binary only)", flags: { "--check": "boolean", "--force": "boolean", "--version": "string (optional, e.g. 0.1.1)" } },
  },
  exit_codes: {
    0: "success",
    1: "network error",
    2: "validation error",
    3: "config error (missing OUTLOOK_CLIENT_ID or OUTLOOK_TENANT_ID)",
    4: "not found",
    5: "auth error (run: outlook login)",
    6: "rate limited",
  },
};

export const agentContextCommand = buildCommand({
  docs: { brief: "Output machine-readable CLI schema for agent use" },
  parameters: {
    flags: {},
  },
  async func(this: void) {
    console.log(JSON.stringify(AGENT_CONTEXT, null, 2));
  },
});
