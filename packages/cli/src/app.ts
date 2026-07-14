import { buildApplication, buildRouteMap } from "@stricli/core";
import { loginCommand, logoutCommand, statusCommand } from "./commands/session.js";
import { profileRoutes } from "./commands/profile.js";
import { mailRoutes } from "./commands/mail.js";
import { agentContextCommand } from "./commands/agent-context.js";
import { upgradeCommand } from "./commands/upgrade.js";

const routes = buildRouteMap({
  routes: {
    login: loginCommand,
    logout: logoutCommand,
    status: statusCommand,
    profile: profileRoutes,
    mail: mailRoutes,
    "agent-context": agentContextCommand,
    upgrade: upgradeCommand,
  },
  docs: {
    // Shown as the top-level description in `outlook --help`.
    brief: "Send and read Microsoft Outlook mail from the terminal. New here? Run `outlook login` first.",
  },
});

export const app = buildApplication(routes, {
  name: "outlook",
  versionInfo: {
    currentVersion: "0.1.1",
  },
});
