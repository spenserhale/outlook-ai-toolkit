import { convert } from "html-to-text";
import TurndownService from "turndown";
import type { BodyFormat } from "./types.js";

const turndown = new TurndownService({
  headingStyle: "atx",
  codeBlockStyle: "fenced",
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
