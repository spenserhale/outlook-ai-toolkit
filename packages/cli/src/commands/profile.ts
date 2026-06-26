import { buildCommand, buildRouteMap } from "@stricli/core";
import { ProfileStore, OutlookAuth, TokenStore, resolveConfig } from "@outlook-toolkit/sdk";
import { openBrowser } from "../browser.js";

const saveCommand = buildCommand({
  docs: { brief: "Save a named profile" },
  parameters: {
    flags: {
      clientId: {
        kind: "parsed",
        brief: "Azure app client ID",
        parse: String,
      },
      tenantId: {
        kind: "parsed",
        brief: "Tenant ID (consumers, common, or GUID)",
        parse: String,
      },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Profile name", parse: String }],
    },
  },
  async func(this: void, flags: { clientId: string; tenantId: string }, name: string) {
    const store = new ProfileStore();
    await store.save(name, { clientId: flags.clientId, tenantId: flags.tenantId });
    console.log(`Profile "${name}" saved.`);
  },
});

const listCommand = buildCommand({
  docs: { brief: "List all saved profiles" },
  parameters: {
    flags: {
      json: {
        kind: "boolean",
        brief: "Output as JSON",
        default: false,
      },
    },
  },
  async func(this: void, flags: { json: boolean }) {
    const store = new ProfileStore();
    const profiles = await store.list();
    const names = Object.keys(profiles);
    if (names.length === 0) {
      console.log("No profiles saved. Run: outlook profile save <name> --clientId=... --tenantId=...");
      return;
    }
    if (flags.json) {
      console.log(JSON.stringify(profiles, null, 2));
    } else {
      for (const [name, p] of Object.entries(profiles)) {
        console.log(`${name}  clientId=${p.clientId}  tenantId=${p.tenantId}`);
      }
    }
  },
});

const deleteCommand = buildCommand({
  docs: { brief: "Delete a saved profile" },
  parameters: {
    flags: {
      force: {
        kind: "boolean",
        brief: "Skip confirmation",
        default: false,
      },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Profile name", parse: String }],
    },
  },
  async func(this: void, _flags: { force: boolean }, name: string) {
    const store = new ProfileStore();
    const deleted = await store.delete(name);
    if (!deleted) {
      console.error(`error: profile "${name}" not found`);
      process.exit(4);
    }
    console.log(`Profile "${name}" deleted.`);
  },
});

const createCommand = buildCommand({
  docs: { brief: "Create a named profile and authenticate in one step" },
  parameters: {
    flags: {
      clientId: {
        kind: "parsed",
        brief: "Azure app client ID (reads OUTLOOK_CLIENT_ID if omitted)",
        parse: String,
        optional: true,
      },
      tenantId: {
        kind: "parsed",
        brief: "Tenant ID (consumers, common, or GUID)",
        parse: String,
        optional: true,
      },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Profile name", parse: String }],
    },
  },
  async func(this: void, flags: { clientId?: string; tenantId?: string }, name: string) {
    const clientId = (() => {
      if (flags.clientId) return flags.clientId;
      try {
        return resolveConfig({ tenantId: flags.tenantId ?? "consumers" }).clientId;
      } catch (err) {
        console.error(err instanceof Error ? err.message : String(err));
        return process.exit(5);
      }
    })();

    const tenantId = flags.tenantId ?? "consumers";
    const store = new TokenStore(name);
    const auth = new OutlookAuth({ clientId, tenantId }, store);

    try {
      const { userEmail } = await auth.login(openBrowser);
      await new ProfileStore().save(name, { clientId, tenantId });
      console.log(`Profile "${name}" created and authenticated as: ${userEmail}`);
    } catch (err) {
      console.error(err instanceof Error ? err.message : String(err));
      process.exit(5);
    }
  },
});

export const profileRoutes = buildRouteMap({
  routes: { save: saveCommand, list: listCommand, delete: deleteCommand, create: createCommand },
  docs: { brief: "Manage named account profiles" },
});
