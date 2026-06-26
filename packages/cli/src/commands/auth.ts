import { buildCommand, buildRouteMap } from "@stricli/core";
import { OutlookAuth, TokenStore } from "@outlook-toolkit/sdk";
import { resolveCliConfig } from "../context.js";
import { openBrowser } from "../browser.js";

const loginCommand = buildCommand({
  docs: { brief: "Sign in to an Outlook account (interactive browser flow)" },
  parameters: {
    flags: {
      profile: {
        kind: "parsed",
        brief: "Named profile to use",
        parse: String,
        optional: true,
      },
    },
  },
  async func(this: void, flags: { profile?: string }) {
    const config = await resolveCliConfig(flags.profile);
    const effectiveProfile = flags.profile ?? process.env.OUTLOOK_PROFILE;
    const tokenKey = effectiveProfile ?? config.clientId;
    const store = new TokenStore(tokenKey);
    const auth = new OutlookAuth(config, store);
    try {
      const { userEmail } = await auth.login(openBrowser);
      console.log(`Authenticated as: ${userEmail}`);
    } catch (err) {
      console.error(`Login failed: ${err instanceof Error ? err.message : err}`);
      process.exit(5);
    }
  },
});

const logoutCommand = buildCommand({
  docs: { brief: "Sign out and clear saved tokens" },
  parameters: {
    flags: {
      profile: {
        kind: "parsed",
        brief: "Named profile to use",
        parse: String,
        optional: true,
      },
    },
  },
  async func(this: void, flags: { profile?: string }) {
    const config = await resolveCliConfig(flags.profile);
    const effectiveProfile = flags.profile ?? process.env.OUTLOOK_PROFILE;
    const tokenKey = effectiveProfile ?? config.clientId;
    const store = new TokenStore(tokenKey);
    const auth = new OutlookAuth(config, store);
    await auth.logout();
    console.log("Signed out.");
  },
});

const statusCommand = buildCommand({
  docs: { brief: "Show current authentication status" },
  parameters: {
    flags: {
      profile: {
        kind: "parsed",
        brief: "Named profile to use",
        parse: String,
        optional: true,
      },
      json: {
        kind: "boolean",
        brief: "Output as JSON",
        default: false,
      },
    },
  },
  async func(this: void, flags: { profile?: string; json: boolean }) {
    const config = await resolveCliConfig(flags.profile);
    const effectiveProfile = flags.profile ?? process.env.OUTLOOK_PROFILE;
    const tokenKey = effectiveProfile ?? config.clientId;
    const store = new TokenStore(tokenKey);
    const auth = new OutlookAuth(config, store);
    const status = await auth.status();

    if (!status) {
      const out = { authenticated: false };
      flags.json
        ? console.log(JSON.stringify(out, null, 2))
        : console.log("Not authenticated. Run: outlook auth login");
      return;
    }

    if (flags.json) {
      console.log(JSON.stringify(status, null, 2));
    } else {
      console.log(`Authenticated: ${status.userEmail ?? "unknown"}`);
      if (status.accessTokenExpiry) {
        console.log(`Access token expires: ${new Date(status.accessTokenExpiry).toISOString()}`);
      }
    }
  },
});

export const authRoutes = buildRouteMap({
  routes: { login: loginCommand, logout: logoutCommand, status: statusCommand },
  docs: { brief: "Manage Outlook authentication" },
});
