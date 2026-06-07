import { describe, expect, it } from "bun:test";
import { renderBody } from "../src/body.js";

const HTML = '<h1>Title</h1><p>Hello <strong>world</strong> see <a href="https://x.co">link</a></p><ul><li>one</li><li>two</li></ul>';

describe("renderBody", () => {
  it("html target returns content unchanged", () => {
    expect(renderBody(HTML, "html", "html")).toBe(HTML);
  });

  it("text target strips tags but keeps readable content", () => {
    const out = renderBody(HTML, "html", "text");
    expect(out).toContain("Title");
    expect(out).toContain("Hello world");
    expect(out).not.toContain("<p>");
    expect(out).not.toContain("<strong>");
  });

  it("markdown target preserves links, bold, and lists", () => {
    const out = renderBody(HTML, "html", "markdown");
    expect(out).toContain("**world**");
    expect(out).toContain("[link](https://x.co)");
    expect(out).toContain("# Title");
    expect(out).toMatch(/[-*]\s+one/);
  });

  it("already-plain text source passes through for every target", () => {
    const plain = "Just plain text.";
    expect(renderBody(plain, "text", "text")).toBe(plain);
    expect(renderBody(plain, "text", "markdown")).toBe(plain);
    expect(renderBody(plain, "text", "html")).toBe(plain);
  });

  it("is case-insensitive about the source content type", () => {
    const out = renderBody("<p>hi</p>", "HTML", "text");
    expect(out).toContain("hi");
    expect(out).not.toContain("<p>");
  });

  it("falls back to raw content if conversion throws", () => {
    const weird = "<<<not really html>>>";
    expect(() => renderBody(weird, "html", "markdown")).not.toThrow();
  });
});
