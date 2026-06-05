import {
  resolveConfig,
  ProfileStore,
  type OutlookConfig,
} from "@outlook-toolkit/sdk";

export async function resolveCliConfig(profile?: string): Promise<OutlookConfig> {
  const profileName = profile ?? process.env.OUTLOOK_PROFILE;
  if (profileName) {
    const store = new ProfileStore();
    const p = await store.get(profileName);
    if (!p) {
      console.error(
        `error: profile "${profileName}" not found (exit code 3). Run: outlook profile list`
      );
      process.exit(3);
    }
    try {
      return resolveConfig(p);
    } catch (err) {
      console.error(err instanceof Error ? err.message : String(err));
      process.exit(3);
    }
  }
  try {
    return resolveConfig();
  } catch (err) {
    console.error(err instanceof Error ? err.message : String(err));
    process.exit(3);
  }
}
