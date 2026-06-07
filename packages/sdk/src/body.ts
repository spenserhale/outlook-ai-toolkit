// html-to-text v9 ships no type declarations and there is no
// @types/html-to-text package; suppress the implicit-any import error.
// @ts-expect-error -- no bundled types for html-to-text
import { convert } from "html-to-text";
import TurndownService from "turndown";
import type { BodyFormat } from "./types.js";

const turndown = new TurndownService({
  headingStyle: "atx",
  codeBlockStyle: "fenced",
});

// Turndown's default list rule pads each marker to a fixed width
// (e.g. "*   item"). Override it to emit a single space ("* item")
// for cleaner, more conventional Markdown output.
turndown.addRule("listItem", {
  filter: "li",
  replacement(content, node, options) {
    const body = content
      .replace(/^\n+/, "")
      .replace(/\n+$/, "\n")
      .replace(/\n/gm, "\n    ");
    const parent = node.parentNode as Element | null;
    let prefix = `${options.bulletListMarker} `;
    if (parent && parent.nodeName === "OL") {
      const start = parent.getAttribute("start");
      const index = Array.prototype.indexOf.call(parent.children, node);
      prefix = `${start ? Number(start) + index : index + 1}. `;
    }
    const trailing =
      node.nextSibling && !/\n$/.test(body) ? "\n" : "";
    return prefix + body + trailing;
  },
});

/**
 * Render an email body into the requested format.
 *
 * - `html`   → return content unchanged
 * - `text`   → strip HTML to clean readable text (html-to-text)
 * - `markdown` → convert HTML to Markdown (turndown)
 *
 * Plain-text sources pass through unchanged (no structure to extract).
 * Best-effort: any converter failure falls back to the raw content.
 */
export function renderBody(
  content: string,
  sourceContentType: string,
  target: BodyFormat
): string {
  if (target === "html") return content;
  if (sourceContentType.toLowerCase() === "text") return content;
  try {
    if (target === "markdown") return turndown.turndown(content);
    return convert(content, {
      wordwrap: false,
      selectors: [
        { selector: "img", format: "skip" },
        { selector: "a", options: { hideLinkHrefIfSameAsText: true } },
        // Preserve original heading casing instead of upper-casing it.
        { selector: "h1", options: { uppercase: false } },
        { selector: "h2", options: { uppercase: false } },
        { selector: "h3", options: { uppercase: false } },
        { selector: "h4", options: { uppercase: false } },
        { selector: "h5", options: { uppercase: false } },
        { selector: "h6", options: { uppercase: false } },
      ],
    });
  } catch {
    return content;
  }
}
