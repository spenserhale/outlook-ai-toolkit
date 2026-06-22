import { describe, expect, it } from "bun:test";
import { SweepConditionSchema, MassMoveParamsSchema } from "../src/types.js";

describe("SweepConditionSchema", () => {
  it("accepts a condition with only `from`", () => {
    expect(SweepConditionSchema.parse({ from: "alice@x.com" }).from).toBe("alice@x.com");
  });

  it("rejects an empty condition", () => {
    expect(() => SweepConditionSchema.parse({})).toThrow();
  });
});

describe("MassMoveParamsSchema", () => {
  it("defaults folder, max, dryRun", () => {
    const p = MassMoveParamsSchema.parse({
      conditions: [{ from: "alice@x.com" }],
      destination: "archive",
    });
    expect(p.folder).toBe("inbox");
    expect(p.max).toBe(200);
    expect(p.dryRun).toBe(false);
  });

  it("requires at least one condition", () => {
    expect(() =>
      MassMoveParamsSchema.parse({ conditions: [], destination: "archive" })
    ).toThrow();
  });
});
