import { describe, expect, it } from "bun:test";
import { encode } from "@toon-format/toon";
import { renderOutput } from "../src/format.js";

const DATA = { value: [{ id: "1", subject: "Hi" }] };

describe("renderOutput", () => {
  it("json matches JSON.stringify with 2-space indent", () => {
    expect(renderOutput(DATA, "json")).toBe(JSON.stringify(DATA, null, 2));
  });

  it("toon matches encode with safe key folding", () => {
    expect(renderOutput(DATA, "toon")).toBe(encode(DATA, { keyFolding: "safe" }));
  });

  it("toon output is shorter than json for list shapes", () => {
    expect(renderOutput(DATA, "toon").length).toBeLessThan(
      renderOutput(DATA, "json").length
    );
  });
});
