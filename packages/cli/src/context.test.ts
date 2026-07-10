import { describe, it, expect, beforeEach, afterEach } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { ProfileStore } from "@outlook-toolkit/sdk";
import { loadContext, DEFAULT_PROFILE } from "./context.js";

// NOTE: bun's os.homedir() ignores process.env.HOME, so we cannot isolate the
// profile store by overriding HOME — doing so writes to the real
// ~/.outlook-toolkit. Instead we inject a ProfileStore pointed at a temp dir.
const ENV_KEYS = ["OUTLOOK_CLIENT_ID", "OUTLOOK_TENANT_ID", "OUTLOOK_PROFILE"] as const;

describe("loadContext", () => {
  let dir: string;
  let store: ProfileStore;
  const saved: Record<string, string | undefined> = {};

  beforeEach(() => {
    for (const k of ENV_KEYS) saved[k] = process.env[k];
    delete process.env.OUTLOOK_CLIENT_ID;
    delete process.env.OUTLOOK_TENANT_ID;
    delete process.env.OUTLOOK_PROFILE;
    dir = mkdtempSync(join(tmpdir(), "outlook-ctx-"));
    store = new ProfileStore(dir);
  });

  afterEach(() => {
    for (const k of ENV_KEYS) {
      if (saved[k] === undefined) delete process.env[k];
      else process.env[k] = saved[k];
    }
    rmSync(dir, { recursive: true, force: true });
  });

  it("returns null when nothing is configured", async () => {
    expect(await loadContext(undefined, store)).toBeNull();
  });

  it("resolves the default profile when no env config is set", async () => {
    await store.save(DEFAULT_PROFILE, { clientId: "cid-default", tenantId: "consumers" });
    const ctx = await loadContext(undefined, store);
    expect(ctx).not.toBeNull();
    expect(ctx?.config.clientId).toBe("cid-default");
    expect(ctx?.tokenKey).toBe(DEFAULT_PROFILE);
    expect(ctx?.profileName).toBe(DEFAULT_PROFILE);
  });

  it("prefers environment config over the default profile", async () => {
    await store.save(DEFAULT_PROFILE, { clientId: "cid-default", tenantId: "consumers" });
    process.env.OUTLOOK_CLIENT_ID = "cid-env";
    process.env.OUTLOOK_TENANT_ID = "common";
    const ctx = await loadContext(undefined, store);
    expect(ctx?.config.clientId).toBe("cid-env");
    expect(ctx?.tokenKey).toBe("cid-env"); // env config keys the token cache by client id
    expect(ctx?.profileName).toBeUndefined();
  });

  it("resolves an explicitly named profile", async () => {
    await store.save("work", { clientId: "cid-work", tenantId: "guid-123" });
    const ctx = await loadContext("work", store);
    expect(ctx?.config.clientId).toBe("cid-work");
    expect(ctx?.config.tenantId).toBe("guid-123");
    expect(ctx?.tokenKey).toBe("work");
    expect(ctx?.profileName).toBe("work");
  });

  it("honors OUTLOOK_PROFILE when no explicit profile is passed", async () => {
    await store.save("personal", { clientId: "cid-personal", tenantId: "consumers" });
    process.env.OUTLOOK_PROFILE = "personal";
    const ctx = await loadContext(undefined, store);
    expect(ctx?.tokenKey).toBe("personal");
    expect(ctx?.profileName).toBe("personal");
  });
});
