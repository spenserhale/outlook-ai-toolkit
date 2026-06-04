import type { OutlookConfig } from "./types.js";

/**
 * Resolve configuration from environment variables.
 * Useful for both CLI and MCP contexts.
 */
export function resolveConfig(
  overrides: Partial<OutlookConfig> = {}
): OutlookConfig {
  return {
    apiKey: overrides.apiKey ?? process.env.OUTLOOK_API_KEY ?? "",
    baseUrl:
      overrides.baseUrl ??
      process.env.OUTLOOK_BASE_URL ??
      "https://api.outlook.com",
  };
}
