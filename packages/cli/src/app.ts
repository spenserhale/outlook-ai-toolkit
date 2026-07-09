import { buildApplication, buildRouteMap } from "@stricli/core";
import { authRoutes } from "./commands/auth.js";
import { profileRoutes } from "./commands/profile.js";
import { mailRoutes } from "./commands/mail.js";
import { agentContextCommand } from "./commands/agent-context.js";
import { upgradeCommand } from "./commands/upgrade.js";

const routes = buildRouteMap({
  routes: {
    auth: authRoutes,
    profile: profileRoutes,
    mail: mailRoutes,
    "agent-context": agentContextCommand,
    upgrade: upgradeCommand,
  },
  docs: {
    brief: "SDK, CLI, and MCP server for Microsoft Outlook via Microsoft Graph",
  },
});

export const app = buildApplication(routes, {
  name: "outlook",
  versionInfo: {
    currentVersion: "0.1.0",
  },
});
