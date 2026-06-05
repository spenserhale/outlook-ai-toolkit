import { describe, expect, it, mock, beforeEach } from "bun:test";
import {
  generateCodeVerifier,
  generateCodeChallenge,
  generateState,
  buildAuthorizationUrl,
  OutlookAuth,
} from "../src/auth.js";
import { OutlookAuthError } from "../src/errors.js";
import type { TokenData } from "../src/types.js";
import type { OutlookConfig } from "../src/types.js";

// ── Pure PKCE helpers ──────────────────────────────────────────────────────

describe("generateCodeVerifier", () => {
  it("returns a base64url string of at least 43 chars", () => {
    const v = generateCodeVerifier();
    expect(v.length).toBeGreaterThanOrEqual(43);
    expect(v).toMatch(/^[A-Za-z0-9\-_]+$/);
  });

  it("returns a different value each call", () => {
    expect(generateCodeVerifier()).not.toBe(generateCodeVerifier());
  });
});

describe("generateCodeChallenge", () => {
  it("produces a deterministic base64url SHA-256 hash", async () => {
    const { createHash } = await import("node:crypto");
    const verifier = "dGVzdC12ZXJpZmllcg";
    const expected = createHash("sha256").update(verifier).digest("base64url");
    expect(generateCodeChallenge(verifier)).toBe(expected);
  });
});

describe("generateState", () => {
  it("returns a 32-char hex string", () => {
    const s = generateState();
    expect(s).toMatch(/^[0-9a-f]{32}$/);
  });
});

describe("buildAuthorizationUrl", () => {
  it("includes required OAuth params", () => {
    const url = buildAuthorizationUrl(
      "my-client",
      "consumers",
      "http://localhost:12345/callback",
      "challenge-abc",
      "state-xyz"
    );
    const parsed = new URL(url);
    expect(parsed.searchParams.get("client_id")).toBe("my-client");
    expect(parsed.searchParams.get("response_type")).toBe("code");
    expect(parsed.searchParams.get("code_challenge")).toBe("challenge-abc");
    expect(parsed.searchParams.get("code_challenge_method")).toBe("S256");
    expect(parsed.searchParams.get("state")).toBe("state-xyz");
  });

  it("uses consumers authority for tenantId='consumers'", () => {
    const url = buildAuthorizationUrl("c", "consumers", "http://localhost/cb", "ch", "st");
    expect(url).toContain("login.microsoftonline.com/consumers");
  });

  it("uses tenant authority for a GUID tenantId", () => {
    const guid = "11111111-2222-3333-4444-555555555555";
    const url = buildAuthorizationUrl("c", guid, "http://localhost/cb", "ch", "st");
    expect(url).toContain(`login.microsoftonline.com/${guid}`);
  });
});

// ── OutlookAuth class ──────────────────────────────────────────────────────

const config: OutlookConfig = { clientId: "test-client", tenantId: "consumers" };

function makeMockStore(data: TokenData | null) {
  return {
    save: mock((_data: TokenData) => Promise.resolve()),
    load: mock(() => Promise.resolve(data)),
    clear: mock(() => Promise.resolve()),
  };
}

describe("OutlookAuth.acquireToken", () => {
  it("returns access token when not near expiry", async () => {
    const store = makeMockStore({
      accessToken: "valid-token",
      refreshToken: "refresh",
      accessTokenExpiry: Date.now() + 2 * 60 * 60 * 1000, // 2 hours
      refreshTokenExpiry: Date.now() + 90 * 86400_000,
    });
    const auth = new OutlookAuth(config, store);
    expect(await auth.acquireToken()).toBe("valid-token");
  });

  it("throws OutlookAuthError when no tokens exist", async () => {
    const store = makeMockStore(null);
    const auth = new OutlookAuth(config, store);
    await expect(auth.acquireToken()).rejects.toBeInstanceOf(OutlookAuthError);
  });

  it("throws OutlookAuthError when refresh token is expired", async () => {
    const store = makeMockStore({
      accessToken: "old-token",
      refreshToken: "old-refresh",
      accessTokenExpiry: Date.now() - 10_000,
      refreshTokenExpiry: Date.now() - 10_000, // expired
    });
    const auth = new OutlookAuth(config, store);
    await expect(auth.acquireToken()).rejects.toBeInstanceOf(OutlookAuthError);
    expect(store.clear).toHaveBeenCalled();
  });
});

describe("OutlookAuth.logout", () => {
  it("calls store.clear()", async () => {
    const store = makeMockStore(null);
    const auth = new OutlookAuth(config, store);
    await auth.logout();
    expect(store.clear).toHaveBeenCalledTimes(1);
  });
});

describe("OutlookAuth.status", () => {
  it("returns null when no tokens exist", async () => {
    const store = makeMockStore(null);
    const auth = new OutlookAuth(config, store);
    expect(await auth.status()).toBeNull();
  });

  it("returns status object when tokens exist", async () => {
    const expiry = Date.now() + 3600_000;
    const store = makeMockStore({
      accessToken: "tok",
      refreshToken: "ref",
      accessTokenExpiry: expiry,
      refreshTokenExpiry: expiry + 90 * 86400_000,
      userEmail: "user@example.com",
    });
    const auth = new OutlookAuth(config, store);
    const status = await auth.status();
    expect(status).not.toBeNull();
    expect(status!.authenticated).toBe(true);
    expect(status!.userEmail).toBe("user@example.com");
  });
});
