import type { FastMCP } from "fastmcp";
import { z } from "zod";
import { resolveConfig, OutlookAuth, TokenStore, OutlookConfigError } from "@outlook-toolkit/sdk";

function getAuth() {
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
  return { auth: new OutlookAuth(config, store), config };
}

export function registerAuthTools(server: FastMCP) {
  server.addTool({
    name: "outlook_auth_status",
    description: "Check Outlook authentication status for the configured account.",
    parameters: z.object({}),
    execute: async () => {
      const { auth, config } = getAuth();
      const status = await auth.status();
      if (!status) {
        return JSON.stringify({
          authenticated: false,
          message: `Not authenticated. Run \`outlook auth login\` in your terminal with OUTLOOK_CLIENT_ID=${config.clientId} set.`,
        }, null, 2);
      }
      return JSON.stringify({
        authenticated: true,
        userEmail: status.userEmail,
        accessTokenExpiry: status.accessTokenExpiry
          ? new Date(status.accessTokenExpiry).toISOString()
          : null,
      }, null, 2);
    },
  });

  server.addTool({
    name: "outlook_auth_logout",
    description: "Sign out and clear saved tokens for the configured Outlook account.",
    parameters: z.object({}),
    execute: async () => {
      const { auth } = getAuth();
      await auth.logout();
      return JSON.stringify({ status: "signed_out", message: "Signed out. Run `outlook auth login` to re-authenticate." }, null, 2);
    },
  });
}
