import { encode } from "@toon-format/toon";
import type { OutputFormat } from "./types.js";

/**
 * Encode a data envelope for output. TOON is the default everywhere;
 * JSON is available for tooling compatibility.
 */
export function renderOutput(data: unknown, format: OutputFormat): string {
  if (format === "json") return JSON.stringify(data, null, 2);
  return encode(data, { keyFolding: "safe" });
}
