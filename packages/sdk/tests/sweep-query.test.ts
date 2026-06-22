import { describe, expect, it } from "bun:test";
import { buildSearchQuery, isoDaysAgo, isOlderThan } from "../src/sweep-query.js";

describe("buildSearchQuery", () => {
  it("builds a KQL AND query from from/subject/body", () => {
    expect(
      buildSearchQuery({ from: "alice@x.com", subjectContains: "invoice", bodyContains: "due" })
    ).toBe('from:alice@x.com AND subject:invoice AND body:due');
  });

  it("returns null when only olderThanDays is set", () => {
    expect(buildSearchQuery({ olderThanDays: 10 })).toBeNull();
  });

  it("emits only the provided fields", () => {
    expect(buildSearchQuery({ from: "alice@x.com" })).toBe("from:alice@x.com");
  });
});

describe("isoDaysAgo", () => {
  it("subtracts N days from the given now and returns ISO", () => {
    const now = Date.parse("2026-06-21T00:00:00.000Z");
    expect(isoDaysAgo(10, now)).toBe("2026-06-11T00:00:00.000Z");
  });
});

describe("isOlderThan", () => {
  const now = Date.parse("2026-06-21T00:00:00.000Z");
  it("true when received before the cutoff", () => {
    expect(isOlderThan("2026-06-01T00:00:00.000Z", 10, now)).toBe(true);
  });
  it("false when received after the cutoff", () => {
    expect(isOlderThan("2026-06-20T00:00:00.000Z", 10, now)).toBe(false);
  });
  it("false when receivedDateTime is missing", () => {
    expect(isOlderThan(undefined, 10, now)).toBe(false);
  });
});
