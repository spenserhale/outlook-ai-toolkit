import { describe, expect, it } from "bun:test";
import * as sdk from "../src/index.js";

describe("public SDK exports", () => {
  it("exports the rules + sweep clients and schemas", () => {
    expect(typeof sdk.RulesClient).toBe("function");
    expect(typeof sdk.MessageRuleSchema?.parse).toBe("function");
    expect(typeof sdk.CreateMessageRuleParamsSchema?.parse).toBe("function");
    expect(typeof sdk.SweepConditionSchema?.parse).toBe("function");
    expect(typeof sdk.MassMoveParamsSchema?.parse).toBe("function");
  });
});
