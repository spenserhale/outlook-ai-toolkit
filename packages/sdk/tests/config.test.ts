import { describe, expect, it, beforeEach, afterEach } from "bun:test";
import { resolveConfig, resolveAuthority } from "../src/config.js";
import { OutlookConfigError } from "../src/errors.js";

describe("resolveAuthority", () => {
  it("maps 'consumers' to consumers authority", () => {
    expect(resolveAuthority("consumers")).toBe(
      "https://login.microsoftonline.com/consumers"
    );
  });

  it("maps 'common' to consumers authority (MSA-capable)", () => {
    expect(resolveAuthority("common")).toBe(
      "https://login.microsoftonline.com/consumers"
    );
  });

  it("maps a GUID tenant to tenant-specific authority", () => {
    const guid = "11111111-2222-3333-4444-555555555555";
    expect(resolveAuthority(guid)).toBe(
      `https://login.microsoftonline.com/${guid}`
    );
  });
});

describe("resolveConfig", () => {
  const origClientId = process.env.OUTLOOK_CLIENT_ID;
  const origTenantId = process.env.OUTLOOK_TENANT_ID;

  beforeEach(() => {
    delete process.env.OUTLOOK_CLIENT_ID;
    delete process.env.OUTLOOK_TENANT_ID;
  });

  afterEach(() => {
    if (origClientId !== undefined) process.env.OUTLOOK_CLIENT_ID = origClientId;
    else delete process.env.OUTLOOK_CLIENT_ID;
    if (origTenantId !== undefined) process.env.OUTLOOK_TENANT_ID = origTenantId;
    else delete process.env.OUTLOOK_TENANT_ID;
  });

  it("throws OutlookConfigError when clientId is missing", () => {
    process.env.OUTLOOK_TENANT_ID = "consumers";
    expect(() => resolveConfig()).toThrow(OutlookConfigError);
  });

  it("throws OutlookConfigError when tenantId is missing", () => {
    process.env.OUTLOOK_CLIENT_ID = "abc123";
    expect(() => resolveConfig()).toThrow(OutlookConfigError);
  });

  it("reads from env vars", () => {
    process.env.OUTLOOK_CLIENT_ID = "env-client";
    process.env.OUTLOOK_TENANT_ID = "consumers";
    const config = resolveConfig();
    expect(config.clientId).toBe("env-client");
    expect(config.tenantId).toBe("consumers");
  });

  it("overrides take precedence over env vars", () => {
    process.env.OUTLOOK_CLIENT_ID = "env-client";
    process.env.OUTLOOK_TENANT_ID = "consumers";
    const config = resolveConfig({ clientId: "override-client" });
    expect(config.clientId).toBe("override-client");
    expect(config.tenantId).toBe("consumers");
  });
});
