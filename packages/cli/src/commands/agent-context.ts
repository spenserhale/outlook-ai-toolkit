import { buildCommand } from "@stricli/core";

const AGENT_CONTEXT = {
  schema_version: "1",
  cli: "outlook",
  description: "Outlook toolkit — read and write Outlook mail via Microsoft Graph",
  config: {
    env_vars: {
      OUTLOOK_CLIENT_ID: { required: true, description: "Azure app registration client ID" },
      OUTLOOK_TENANT_ID: { required: true, description: "Tenant ID: consumers | common | <guid>" },
      OUTLOOK_PROFILE: { required: false, description: "Named profile to activate" },
    },
    precedence: ["explicit-flag", "env-var", "named-profile", "error"],
  },
  commands: {
    auth: {
      login: { brief: "Interactive browser PKCE sign-in", flags: { "--profile": "string (optional)" } },
      logout: { brief: "Clear saved tokens", flags: { "--profile": "string (optional)" } },
      status: { brief: "Show auth status", flags: { "--profile": "string (optional)", "--json": "boolean" } },
    },
    profile: {
      save: { brief: "Save a named profile", args: ["name"], flags: { "--client-id": "string", "--tenant-id": "string" } },
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
    },
  },
  exit_codes: {
    0: "success",
    1: "network error",
    2: "validation error",
    3: "config error (missing OUTLOOK_CLIENT_ID or OUTLOOK_TENANT_ID)",
    4: "not found",
    5: "auth error (run: outlook auth login)",
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
