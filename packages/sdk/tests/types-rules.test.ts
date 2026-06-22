import { describe, expect, it } from "bun:test";
import {
  MessageRuleSchema,
  CreateMessageRuleParamsSchema,
} from "../src/types.js";

describe("MessageRuleSchema", () => {
  it("parses a Graph rule and keeps unknown fields (passthrough)", () => {
    const rule = MessageRuleSchema.parse({
      id: "AQAAAJ5dZqA=",
      displayName: "From partner",
      sequence: 2,
      isEnabled: true,
      conditions: { senderContains: ["adele"] },
      actions: { moveToFolder: "deleteditems", stopProcessingRules: true },
      hasError: false,
    });
    expect(rule.displayName).toBe("From partner");
    expect(rule.actions?.moveToFolder).toBe("deleteditems");
  });
});

describe("CreateMessageRuleParamsSchema", () => {
  it("defaults sequence and isEnabled", () => {
    const p = CreateMessageRuleParamsSchema.parse({
      displayName: "x",
      actions: { delete: true },
    });
    expect(p.sequence).toBe(1);
    expect(p.isEnabled).toBe(true);
  });

  it("requires actions", () => {
    expect(() =>
      CreateMessageRuleParamsSchema.parse({ displayName: "x" })
    ).toThrow();
  });
});
