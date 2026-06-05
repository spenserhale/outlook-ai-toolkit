import { describe, expect, it, beforeAll, afterAll } from "bun:test";
import { join } from "node:path";
import { mkdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { ProfileStore } from "../src/profile-store.js";

const testDir = join(tmpdir(), `outlook-profile-test-${Date.now()}`);

beforeAll(() => mkdirSync(testDir, { recursive: true }));
afterAll(() => rmSync(testDir, { recursive: true, force: true }));

describe("ProfileStore", () => {
  it("get returns null for unknown profile", async () => {
    const store = new ProfileStore(testDir);
    expect(await store.get("nonexistent")).toBeNull();
  });

  it("save then get round-trips profile data", async () => {
    const store = new ProfileStore(testDir);
    await store.save("personal", { clientId: "abc", tenantId: "consumers" });
    const p = await store.get("personal");
    expect(p).not.toBeNull();
    expect(p!.clientId).toBe("abc");
    expect(p!.tenantId).toBe("consumers");
  });

  it("list returns all saved profiles", async () => {
    const store = new ProfileStore(testDir);
    await store.save("p1", { clientId: "c1", tenantId: "consumers" });
    await store.save("p2", { clientId: "c2", tenantId: "work-guid" });
    const all = await store.list();
    expect(all["p1"]?.clientId).toBe("c1");
    expect(all["p2"]?.clientId).toBe("c2");
  });

  it("delete removes a profile and returns true", async () => {
    const store = new ProfileStore(testDir);
    await store.save("temp", { clientId: "t", tenantId: "consumers" });
    expect(await store.delete("temp")).toBe(true);
    expect(await store.get("temp")).toBeNull();
  });

  it("delete returns false for unknown profile", async () => {
    const store = new ProfileStore(testDir);
    expect(await store.delete("ghost")).toBe(false);
  });
});
