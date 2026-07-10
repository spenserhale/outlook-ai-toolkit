import { buildCommand, buildRouteMap } from "@stricli/core";
import { ProfileStore } from "@outlook-toolkit/sdk";
import { isInteractive, promptLine } from "../prompt.js";
import { performLogin } from "./session.js";

const addCommand = buildCommand({
  docs: { brief: "Add a named account profile and sign in (for multiple accounts)" },
  parameters: {
    flags: {
      clientId: {
        kind: "parsed",
        brief: "Azure app client ID (prompted if omitted in an interactive shell)",
        parse: String,
        optional: true,
      },
      tenantId: {
        kind: "parsed",
        brief: "Tenant: consumers (personal), common, or a work/school GUID (default: consumers)",
        parse: String,
        optional: true,
      },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Profile name", placeholder: "name", parse: String, optional: true }],
    },
  },
  async func(
    this: void,
    flags: { clientId?: string; tenantId?: string },
    nameArg?: string
  ) {
    let name = nameArg;
    if (!name) {
      if (!isInteractive()) {
        console.error("error: profile name required (exit code 2). Usage: outlook profile add <name>");
        process.exit(2);
      }
      name = await promptLine("Profile name: ");
      if (!name) {
        console.error("error: no profile name provided (exit code 2).");
        process.exit(2);
      }
    }

    let clientId = flags.clientId;
    if (!clientId) {
      if (!isInteractive()) {
        console.error("error: --clientId is required (exit code 2). See the README 'Azure App Registration' to create one.");
        process.exit(2);
      }
      console.log("You need a free Azure app Client ID — see the README 'Azure App Registration'.");
      clientId = await promptLine("Paste your Azure app Client ID: ");
      if (!clientId) {
        console.error("error: no Client ID provided (exit code 2).");
        process.exit(2);
      }
    }

    const tenantId = flags.tenantId ?? "consumers";
    await new ProfileStore().save(name, { clientId, tenantId });
    console.log(`Profile "${name}" saved.`);

    // `add` saves and signs in; use `profile save` to store one without signing in.
    await performLogin({ config: { clientId, tenantId }, tokenKey: name, profileName: name });
  },
});

const saveCommand = buildCommand({
  docs: { brief: "Save a profile without signing in (for scripting)" },
  parameters: {
    flags: {
      clientId: { kind: "parsed", brief: "Azure app client ID", parse: String },
      tenantId: { kind: "parsed", brief: "Tenant ID (consumers, common, or GUID)", parse: String },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Profile name", placeholder: "name", parse: String }],
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
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
  },
  async func(this: void, flags: { json: boolean }) {
    const store = new ProfileStore();
    const profiles = await store.list();
    const names = Object.keys(profiles);
    if (names.length === 0) {
      console.log("No profiles saved. Run: outlook login  (or, for a named account: outlook profile add <name>)");
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
      force: { kind: "boolean", brief: "Skip confirmation", default: false },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Profile name", placeholder: "name", parse: String }],
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

export const profileRoutes = buildRouteMap({
  routes: {
    add: addCommand,
    create: addCommand,
    save: saveCommand,
    list: listCommand,
    delete: deleteCommand,
  },
  docs: { brief: "Manage named account profiles (for multiple accounts)" },
});
