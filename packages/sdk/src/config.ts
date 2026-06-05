import { existsSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { OutlookConfigSchema, type OutlookConfig } from "./types.js";
import { OutlookConfigError } from "./errors.js";

export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
export const AUTH_BASE_URL = "https://login.microsoftonline.com";

export const SCOPES = [
  "Mail.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "offline_access",
  "openid",
  "profile",
  "User.Read",
];

export function resolveAuthority(tenantId: string): string {
  const t = tenantId.toLowerCase();
  if (t === "common" || t === "consumers") {
    return `${AUTH_BASE_URL}/consumers`;
  }
  return `${AUTH_BASE_URL}/${tenantId}`;
}

function parseEnvLine(line: string): [string, string] | null {
  const trimmed = line.trim();
  if (!trimmed || trimmed.startsWith("#")) return null;
  const sep = trimmed.indexOf("=");
  if (sep <= 0) return null;
  const key = trimmed.slice(0, sep).trim();
  let value = trimmed.slice(sep + 1).trim();
  if (
    (value.startsWith('"') && value.endsWith('"')) ||
    (value.startsWith("'") && value.endsWith("'"))
  ) {
    value = value.slice(1, -1);
  }
  return [key, value];
}

function readDotenv(filePath: string, key: string): string | undefined {
  if (!existsSync(filePath)) return undefined;
  for (const line of readFileSync(filePath, "utf8").split(/\r?\n/)) {
    const parsed = parseEnvLine(line);
    if (parsed && parsed[0] === key) return parsed[1];
  }
  return undefined;
}

function getFromNearestDotenv(key: string): string | undefined {
  const moduleDir = dirname(fileURLToPath(import.meta.url));
  const visited = new Set<string>();
  for (const startDir of [process.cwd(), moduleDir]) {
    let dir = startDir;
    while (true) {
      if (!visited.has(dir)) {
        visited.add(dir);
        for (const name of [".env.local", ".env"]) {
          const val = readDotenv(join(dir, name), key);
          if (val !== undefined) return val;
        }
      }
      const parent = dirname(dir);
      if (parent === dir) break;
      dir = parent;
    }
  }
  return undefined;
}

export function resolveConfig(overrides: Partial<OutlookConfig> = {}): OutlookConfig {
  const clientId =
    overrides.clientId ??
    process.env.OUTLOOK_CLIENT_ID ??
    getFromNearestDotenv("OUTLOOK_CLIENT_ID") ??
    "";
  const tenantId =
    overrides.tenantId ??
    process.env.OUTLOOK_TENANT_ID ??
    getFromNearestDotenv("OUTLOOK_TENANT_ID") ??
    "";

  const result = OutlookConfigSchema.safeParse({ clientId, tenantId });
  if (!result.success) {
    const missing = result.error.errors.map((e) => e.message).join("; ");
    throw new OutlookConfigError(
      `${missing}. Set OUTLOOK_CLIENT_ID and OUTLOOK_TENANT_ID.`
    );
  }
  return result.data;
}
