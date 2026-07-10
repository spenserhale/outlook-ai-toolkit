import { buildCommand } from "@stricli/core";
import { OutlookAuth, TokenStore, ProfileStore } from "@outlook-toolkit/sdk";
import { loadContext, DEFAULT_PROFILE, type CliContext } from "../context.js";
import { isInteractive, promptLine } from "../prompt.js";
import { openBrowser } from "../browser.js";

/** Run the interactive browser sign-in for a resolved account. */
export async function performLogin(ctx: CliContext): Promise<void> {
  const store = new TokenStore(ctx.tokenKey);
  const auth = new OutlookAuth(ctx.config, store);
  try {
    const { userEmail } = await auth.login(openBrowser);
    const where = ctx.profileName ? ` (profile "${ctx.profileName}")` : "";
    console.log(`Authenticated as: ${userEmail}${where}`);
  } catch (err) {
    console.error(`Login failed: ${err instanceof Error ? err.message : err}`);
    process.exit(5);
  }
}

/**
 * First-run setup: prompt for an Azure app Client ID, save it as the given
 * profile (defaulting the tenant to personal `consumers`), and return the
 * resulting context. Non-interactive shells get clear instructions instead.
 */
async function guidedSetup(profileName: string): Promise<CliContext> {
  if (!isInteractive()) {
    console.error("error: no account configured (exit code 3).");
    console.error("Set OUTLOOK_CLIENT_ID (and OUTLOOK_TENANT_ID), or run `outlook login` in an");
    console.error("interactive terminal to set one up.");
    console.error("You need a free Azure app Client ID — see the README 'Azure App Registration'.");
    process.exit(3);
  }

  console.log("No account configured. Let's set one up.");
  console.log("");
  console.log("You need a free Azure app Client ID (one-time, ~2 min).");
  console.log("See the README 'Azure App Registration' section for the exact steps.");
  console.log("");

  const clientId = await promptLine("Paste your Azure app Client ID: ");
  if (!clientId) {
    console.error("error: no Client ID provided (exit code 2).");
    process.exit(2);
  }

  // Q: personal vs work/school. Personal accounts use the "consumers" tenant;
  // power users with a work/school tenant can pass it via `profile add --tenantId`.
  const tenantId = "consumers";

  await new ProfileStore().save(profileName, { clientId, tenantId });
  console.log(`Saved profile "${profileName}".`);
  console.log("");

  return { config: { clientId, tenantId }, tokenKey: profileName, profileName };
}

const loginCommand = buildCommand({
  docs: { brief: "Sign in to Outlook (guides you through setup on first run)" },
  parameters: {
    flags: {
      profile: {
        kind: "parsed",
        brief: "Named profile to sign in (default: the default profile)",
        parse: String,
        optional: true,
      },
    },
  },
  async func(this: void, flags: { profile?: string }) {
    let ctx = await loadContext(flags.profile);
    if (!ctx) {
      ctx = await guidedSetup(flags.profile ?? DEFAULT_PROFILE);
    }
    await performLogin(ctx);
  },
});

const logoutCommand = buildCommand({
  docs: { brief: "Sign out and clear saved tokens" },
  parameters: {
    flags: {
      profile: {
        kind: "parsed",
        brief: "Named profile to sign out (default: the default profile)",
        parse: String,
        optional: true,
      },
    },
  },
  async func(this: void, flags: { profile?: string }) {
    const ctx = await loadContext(flags.profile);
    if (!ctx) {
      console.log("No account configured; nothing to sign out.");
      return;
    }
    const store = new TokenStore(ctx.tokenKey);
    const auth = new OutlookAuth(ctx.config, store);
    await auth.logout();
    console.log("Signed out.");
  },
});

const statusCommand = buildCommand({
  docs: { brief: "Show current sign-in status" },
  parameters: {
    flags: {
      profile: {
        kind: "parsed",
        brief: "Named profile to check (default: the default profile)",
        parse: String,
        optional: true,
      },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
  },
  async func(this: void, flags: { profile?: string; json: boolean }) {
    const ctx = await loadContext(flags.profile);
    if (!ctx) {
      const out = { configured: false, authenticated: false };
      flags.json
        ? console.log(JSON.stringify(out, null, 2))
        : console.log("No account configured. Run: outlook login");
      return;
    }

    const store = new TokenStore(ctx.tokenKey);
    const auth = new OutlookAuth(ctx.config, store);
    const status = await auth.status();

    if (!status) {
      const out = { configured: true, authenticated: false, profile: ctx.profileName };
      flags.json
        ? console.log(JSON.stringify(out, null, 2))
        : console.log("Not signed in. Run: outlook login");
      return;
    }

    if (flags.json) {
      console.log(JSON.stringify({ configured: true, ...status, profile: ctx.profileName }, null, 2));
    } else {
      console.log(`Signed in: ${status.userEmail ?? "unknown"}${ctx.profileName ? ` (profile "${ctx.profileName}")` : ""}`);
      if (status.accessTokenExpiry) {
        console.log(`Access token expires: ${new Date(status.accessTokenExpiry).toISOString()}`);
      }
    }
  },
});

export { loginCommand, logoutCommand, statusCommand };
