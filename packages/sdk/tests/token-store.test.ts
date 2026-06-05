import { describe, expect, it, beforeAll, afterAll } from "bun:test";
import { join } from "node:path";
import { mkdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { TokenStore } from "../src/token-store.js";
import type { TokenData } from "../src/types.js";

const testDir = join(tmpdir(), `outlook-token-test-${Date.now()}`);

beforeAll(() => mkdirSync(testDir, { recursive: true }));
afterAll(() => rmSync(testDir, { recursive: true, force: true }));

const sampleData: TokenData = {
  accessToken: "access-abc",
  refreshToken: "refresh-xyz",
  accessTokenExpiry: Date.now() + 3600_000,
  refreshTokenExpiry: Date.now() + 90 * 86400_000,
  userEmail: "test@example.com",
};

describe("TokenStore", () => {
  it("load returns null when no file exists", async () => {
    const store = new TokenStore("no-client", testDir);
    expect(await store.load()).toBeNull();
  });

  it("save then load round-trips token data", async () => {
    const store = new TokenStore("client-1", testDir);
    await store.save(sampleData);
    const loaded = await store.load();
    expect(loaded).not.toBeNull();
    expect(loaded!.accessToken).toBe(sampleData.accessToken);
    expect(loaded!.refreshToken).toBe(sampleData.refreshToken);
    expect(loaded!.userEmail).toBe(sampleData.userEmail);
  });

  it("clear removes the file — subsequent load returns null", async () => {
    const store = new TokenStore("client-2", testDir);
    await store.save(sampleData);
    await store.clear();
    expect(await store.load()).toBeNull();
  });

  it("different clientIds use separate files", async () => {
    const store1 = new TokenStore("client-a", testDir);
    const store2 = new TokenStore("client-b", testDir);
    await store1.save({ ...sampleData, userEmail: "a@example.com" });
    await store2.save({ ...sampleData, userEmail: "b@example.com" });
    expect((await store1.load())!.userEmail).toBe("a@example.com");
    expect((await store2.load())!.userEmail).toBe("b@example.com");
  });
});
