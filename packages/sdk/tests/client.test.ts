import { describe, expect, it } from "bun:test";
import { OutlookClient } from "../src/client.js";

describe("OutlookClient", () => {
  it("should require an API key", () => {
    expect(() => new OutlookClient({ apiKey: "" })).toThrow();
  });

  it("should accept a valid config", () => {
    const client = new OutlookClient({
      apiKey: "test-key",
      baseUrl: "https://api.example.com",
    });
    expect(client).toBeDefined();
  });
});
