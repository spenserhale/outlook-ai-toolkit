import {
  resolveConfig,
  ProfileStore,
  type OutlookConfig,
} from "@outlook-toolkit/sdk";

/** Profile name used for the single-account "just log me in" flow. */
export const DEFAULT_PROFILE = "default";

export interface CliContext {
  config: OutlookConfig;
  /** Key the token cache is stored under (profile name, or client id for env config). */
  tokenKey: string;
  /** Set when the config came from a saved profile. */
  profileName?: string;
}

/**
 * Resolve which account to act as, in priority order:
 *   1. an explicit `--profile` / `OUTLOOK_PROFILE`
 *   2. `OUTLOOK_CLIENT_ID` / `OUTLOOK_TENANT_ID` from the environment or a .env file
 *   3. the saved "default" profile (created by `outlook login`)
 *
 * Returns `null` when nothing is configured yet (so callers can offer setup).
 * Exits (code 3) when a profile is explicitly named but doesn't exist.
 */
export async function loadContext(
  profile?: string,
  store: ProfileStore = new ProfileStore()
): Promise<CliContext | null> {
  const profileName = profile ?? process.env.OUTLOOK_PROFILE;

  if (profileName) {
    const p = await store.get(profileName);
    if (!p) {
      console.error(
        `error: profile "${profileName}" not found (exit code 3). Create it with: outlook profile add ${profileName}`
      );
      process.exit(3);
    }
    return { config: resolveConfig(p), tokenKey: profileName, profileName };
  }

  try {
    const config = resolveConfig();
    return { config, tokenKey: config.clientId };
  } catch {
    // No env/.env config — fall back to the saved default profile below.
  }

  const def = await store.get(DEFAULT_PROFILE);
  if (def) {
    return { config: resolveConfig(def), tokenKey: DEFAULT_PROFILE, profileName: DEFAULT_PROFILE };
  }

  return null;
}

/**
 * Like {@link loadContext} but for commands that cannot proceed without an
 * account. Exits (code 3) with a setup hint when nothing is configured.
 */
export async function requireContext(
  profile?: string,
  store?: ProfileStore
): Promise<CliContext> {
  const ctx = await loadContext(profile, store);
  if (ctx) return ctx;
  console.error("error: no account configured (exit code 3). Run: outlook login");
  process.exit(3);
}
