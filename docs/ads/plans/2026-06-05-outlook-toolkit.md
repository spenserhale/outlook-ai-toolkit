# Outlook Toolkit — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use ads:subagent-driven-development (recommended) or ads:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement a Bun monorepo toolkit for programmatic Outlook/Microsoft 365 access via Microsoft Graph — SDK (auth + mail), CLI (Stricli), and MCP (FastMCP).

**Architecture:** Hand-rolled PKCE auth (no MSAL) with keytar+AES token storage, plain `fetch` Graph client, one active account per instance configured via env vars or named CLI profiles. SDK is the sole logic layer; CLI and MCP are thin wires.

**Tech Stack:** Bun, TypeScript (strict/ESM), Zod, Stricli, FastMCP, keytar, `@toon-format/toon`, `node:crypto`, `node:http`

---

## File Map

**Packages (existing scaffold — will be rewritten/replaced):**
```
packages/sdk/package.json            modify: add keytar dep
packages/sdk/src/types.ts            rewrite: Outlook Zod schemas
packages/sdk/src/errors.ts           rewrite: add OutlookConfigError, OutlookRateLimitError
packages/sdk/src/config.ts           rewrite: clientId/tenantId + resolveAuthority
packages/sdk/src/client.ts           delete: replaced by graph-client + mail-client
packages/sdk/src/index.ts            rewrite: updated exports
packages/sdk/tests/client.test.ts    delete: replaced by new tests

packages/cli/package.json            modify: add @toon-format/toon dep
packages/cli/src/app.ts              rewrite: auth/profile/mail route map
packages/cli/src/commands/create.ts  delete
packages/cli/src/commands/delete.ts  delete
packages/cli/src/commands/get.ts     delete
packages/cli/src/commands/list.ts    delete

packages/mcp/src/index.ts            rewrite: import from server.ts
packages/mcp/src/tools/resources.ts  delete: replaced by auth.ts + mail.ts
```

**New files to create:**
```
.env.example                               OUTLOOK_CLIENT_ID, OUTLOOK_TENANT_ID

packages/sdk/src/token-store.ts            TokenStore class (keytar + AES file fallback)
packages/sdk/src/auth.ts                   PKCE helpers + OutlookAuth class
packages/sdk/src/profile-store.ts          ProfileStore class
packages/sdk/src/graph-client.ts           GraphClient with retry/backoff
packages/sdk/src/mail-client.ts            MailClient (list/get/send/reply/draft/sync)

packages/sdk/tests/config.test.ts
packages/sdk/tests/token-store.test.ts
packages/sdk/tests/auth.test.ts
packages/sdk/tests/profile-store.test.ts
packages/sdk/tests/graph-client.test.ts
packages/sdk/tests/mail-client.test.ts

packages/cli/src/context.ts               resolveCliConfig() helper
packages/cli/src/commands/auth.ts         login/logout/status route map
packages/cli/src/commands/profile.ts      save/list/delete route map
packages/cli/src/commands/mail.ts         list/get/send/reply/draft/sync route map
packages/cli/src/commands/agent-context.ts machine-readable schema

packages/mcp/src/server.ts                FastMCP server instance
packages/mcp/src/tools/auth.ts            outlook_auth_status, outlook_auth_logout
packages/mcp/src/tools/mail.ts            outlook_mail_* tools
```

---

## Task 1: Add dependencies and update .env.example

**Files:**
- Modify: `packages/sdk/package.json`
- Modify: `packages/cli/package.json`
- Modify: `.env.example`

- [ ] **Step 1: Add keytar to SDK dependencies**

Edit `packages/sdk/package.json` — replace the `dependencies` block:
```json
"dependencies": {
  "zod": "^3.24.0",
  "keytar": "^7.9.0"
}
```

- [ ] **Step 2: Add toon to CLI dependencies**

Edit `packages/cli/package.json` — add to `dependencies`:
```json
"dependencies": {
  "@outlook-toolkit/sdk": "workspace:*",
  "@stricli/core": "^1.0.0",
  "@toon-format/toon": "^2.1.0"
}
```

- [ ] **Step 3: Update .env.example**

Replace `.env.example` entirely:
```
# Required for all operations
OUTLOOK_CLIENT_ID=your-azure-app-client-id
OUTLOOK_TENANT_ID=consumers

# consumers = personal MSA (outlook.com)
# common    = personal + work (multi-tenant)
# <guid>    = specific Entra ID tenant (work/school)

# Optional: select a saved profile by name (CLI and MCP)
# OUTLOOK_PROFILE=personal
```

- [ ] **Step 4: Install dependencies**

Run from workspace root:
```bash
bun install
```
Expected: no errors, `keytar` and `@toon-format/toon` appear in lockfile.

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/package.json packages/cli/package.json .env.example bun.lockb
git commit -m "chore: add keytar and toon dependencies"
```

---

## Task 2: Rewrite types.ts and errors.ts

**Files:**
- Modify: `packages/sdk/src/types.ts`
- Modify: `packages/sdk/src/errors.ts`
- Delete: `packages/sdk/src/client.ts`
- Delete: `packages/sdk/tests/client.test.ts`

- [ ] **Step 1: Rewrite types.ts**

Replace `packages/sdk/src/types.ts` entirely:
```typescript
import { z } from "zod";

// Config
export const OutlookConfigSchema = z.object({
  clientId: z.string().min(1, "OUTLOOK_CLIENT_ID is required"),
  tenantId: z.string().min(1, "OUTLOOK_TENANT_ID is required"),
});
export type OutlookConfig = z.infer<typeof OutlookConfigSchema>;

// Token storage
export const TokenDataSchema = z.object({
  accessToken: z.string(),
  refreshToken: z.string(),
  accessTokenExpiry: z.number(),
  refreshTokenExpiry: z.number(),
  userEmail: z.string().optional(),
});
export type TokenData = z.infer<typeof TokenDataSchema>;

// Auth status (non-sensitive, safe to display)
export const AuthStatusSchema = z.object({
  authenticated: z.boolean(),
  userEmail: z.string().optional(),
  accessTokenExpiry: z.number().optional(),
  refreshTokenExpiry: z.number().optional(),
});
export type AuthStatus = z.infer<typeof AuthStatusSchema>;

// Graph API — message types
export const MessageBodySchema = z.object({
  contentType: z.enum(["text", "HTML"]),
  content: z.string(),
});
export type MessageBody = z.infer<typeof MessageBodySchema>;

export const EmailAddressSchema = z.object({
  address: z.string(),
  name: z.string().optional(),
});
export type EmailAddress = z.infer<typeof EmailAddressSchema>;

export const RecipientSchema = z.object({
  emailAddress: EmailAddressSchema,
});
export type Recipient = z.infer<typeof RecipientSchema>;

export const MessageSchema = z
  .object({
    id: z.string(),
    subject: z.string().nullable().optional(),
    bodyPreview: z.string().optional(),
    body: MessageBodySchema.optional(),
    from: RecipientSchema.optional(),
    toRecipients: z.array(RecipientSchema).optional(),
    receivedDateTime: z.string().optional(),
    sentDateTime: z.string().optional(),
    isRead: z.boolean().optional(),
    isDraft: z.boolean().optional(),
    conversationId: z.string().optional(),
  })
  .passthrough();
export type Message = z.infer<typeof MessageSchema>;

export const MailListResponseSchema = z.object({
  value: z.array(MessageSchema),
  "@odata.nextLink": z.string().optional(),
});
export type MailListResponse = z.infer<typeof MailListResponseSchema>;

export const DeltaResponseSchema = z.object({
  value: z.array(MessageSchema),
  "@odata.nextLink": z.string().optional(),
  "@odata.deltaLink": z.string().optional(),
});
export type DeltaResponse = z.infer<typeof DeltaResponseSchema>;

// Mail operation params
export const ListMailParamsSchema = z.object({
  folder: z.string().default("inbox"),
  limit: z.number().int().positive().max(999).default(25),
  cursor: z.string().optional(),
  filter: z.string().optional(),
  select: z.string().optional(),
  orderby: z.string().optional(),
});
export type ListMailParams = z.infer<typeof ListMailParamsSchema>;

export const SendMailParamsSchema = z.object({
  to: z.string().email(),
  subject: z.string(),
  body: z.string(),
  contentType: z.enum(["text", "HTML"]).default("HTML"),
});
export type SendMailParams = z.infer<typeof SendMailParamsSchema>;

export const ReplyParamsSchema = z.object({
  body: z.string(),
  contentType: z.enum(["text", "HTML"]).default("HTML"),
});
export type ReplyParams = z.infer<typeof ReplyParamsSchema>;

export const DraftParamsSchema = z.object({
  to: z.string().email(),
  subject: z.string(),
  body: z.string(),
  contentType: z.enum(["text", "HTML"]).default("HTML"),
});
export type DraftParams = z.infer<typeof DraftParamsSchema>;

// Profile storage
export const ProfileSchema = z.object({
  clientId: z.string().min(1),
  tenantId: z.string().min(1),
});
export type Profile = z.infer<typeof ProfileSchema>;

export const ProfilesFileSchema = z.record(z.string(), ProfileSchema);
export type ProfilesFile = z.infer<typeof ProfilesFileSchema>;
```

- [ ] **Step 2: Rewrite errors.ts**

Replace `packages/sdk/src/errors.ts` entirely:
```typescript
export class OutlookError extends Error {
  constructor(
    message: string,
    public readonly code: string,
    public readonly statusCode?: number
  ) {
    super(message);
    this.name = "OutlookError";
  }
}

export class OutlookConfigError extends OutlookError {
  constructor(message: string) {
    super(message, "CONFIG_ERROR");
    this.name = "OutlookConfigError";
  }
}

export class OutlookAuthError extends OutlookError {
  constructor(message = "Not authenticated. Run `outlook auth login` first.") {
    super(message, "AUTH_ERROR", 401);
    this.name = "OutlookAuthError";
  }
}

export class OutlookNotFoundError extends OutlookError {
  constructor(resource: string, id: string) {
    super(`${resource} "${id}" not found`, "NOT_FOUND", 404);
    this.name = "OutlookNotFoundError";
  }
}

export class OutlookRateLimitError extends OutlookError {
  constructor(retryAfterSeconds?: number) {
    const msg = retryAfterSeconds
      ? `Rate limited by Microsoft Graph. Retry after ${retryAfterSeconds}s.`
      : "Rate limited by Microsoft Graph API.";
    super(msg, "RATE_LIMIT", 429);
    this.name = "OutlookRateLimitError";
  }
}
```

- [ ] **Step 3: Delete obsolete scaffold files**

```bash
rm packages/sdk/src/client.ts
rm packages/sdk/tests/client.test.ts
```

- [ ] **Step 4: Verify lint passes**

```bash
bun run --filter '@outlook-toolkit/sdk' lint
```
Expected: no errors.

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/src/types.ts packages/sdk/src/errors.ts
git rm packages/sdk/src/client.ts packages/sdk/tests/client.test.ts
git commit -m "feat(sdk): rewrite types and errors for Outlook Graph API"
```

---

## Task 3: Config module — resolveConfig + resolveAuthority

**Files:**
- Modify: `packages/sdk/src/config.ts`
- Create: `packages/sdk/tests/config.test.ts`

- [ ] **Step 1: Write failing tests**

Create `packages/sdk/tests/config.test.ts`:
```typescript
import { describe, expect, it, beforeEach, afterEach } from "bun:test";
import { resolveConfig, resolveAuthority } from "../src/config.js";
import { OutlookConfigError } from "../src/errors.js";

describe("resolveAuthority", () => {
  it("maps 'consumers' to consumers authority", () => {
    expect(resolveAuthority("consumers")).toBe(
      "https://login.microsoftonline.com/consumers"
    );
  });

  it("maps 'common' to consumers authority (MSA-capable)", () => {
    expect(resolveAuthority("common")).toBe(
      "https://login.microsoftonline.com/consumers"
    );
  });

  it("maps a GUID tenant to tenant-specific authority", () => {
    const guid = "11111111-2222-3333-4444-555555555555";
    expect(resolveAuthority(guid)).toBe(
      `https://login.microsoftonline.com/${guid}`
    );
  });
});

describe("resolveConfig", () => {
  const origClientId = process.env.OUTLOOK_CLIENT_ID;
  const origTenantId = process.env.OUTLOOK_TENANT_ID;

  beforeEach(() => {
    delete process.env.OUTLOOK_CLIENT_ID;
    delete process.env.OUTLOOK_TENANT_ID;
  });

  afterEach(() => {
    if (origClientId !== undefined) process.env.OUTLOOK_CLIENT_ID = origClientId;
    else delete process.env.OUTLOOK_CLIENT_ID;
    if (origTenantId !== undefined) process.env.OUTLOOK_TENANT_ID = origTenantId;
    else delete process.env.OUTLOOK_TENANT_ID;
  });

  it("throws OutlookConfigError when clientId is missing", () => {
    process.env.OUTLOOK_TENANT_ID = "consumers";
    expect(() => resolveConfig()).toThrow(OutlookConfigError);
  });

  it("throws OutlookConfigError when tenantId is missing", () => {
    process.env.OUTLOOK_CLIENT_ID = "abc123";
    expect(() => resolveConfig()).toThrow(OutlookConfigError);
  });

  it("reads from env vars", () => {
    process.env.OUTLOOK_CLIENT_ID = "env-client";
    process.env.OUTLOOK_TENANT_ID = "consumers";
    const config = resolveConfig();
    expect(config.clientId).toBe("env-client");
    expect(config.tenantId).toBe("consumers");
  });

  it("overrides take precedence over env vars", () => {
    process.env.OUTLOOK_CLIENT_ID = "env-client";
    process.env.OUTLOOK_TENANT_ID = "consumers";
    const config = resolveConfig({ clientId: "override-client" });
    expect(config.clientId).toBe("override-client");
    expect(config.tenantId).toBe("consumers");
  });
});
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
bun test packages/sdk/tests/config.test.ts
```
Expected: FAIL — `resolveAuthority` not found, `OutlookConfigError` not found.

- [ ] **Step 3: Rewrite config.ts**

Replace `packages/sdk/src/config.ts` entirely:
```typescript
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
```

- [ ] **Step 4: Run tests — confirm they pass**

```bash
bun test packages/sdk/tests/config.test.ts
```
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/src/config.ts packages/sdk/tests/config.test.ts
git commit -m "feat(sdk): rewrite config with resolveAuthority and typed config error"
```

---

## Task 4: Token store — keytar + AES-256-CBC file fallback

**Files:**
- Create: `packages/sdk/src/token-store.ts`
- Create: `packages/sdk/tests/token-store.test.ts`

- [ ] **Step 1: Write failing tests**

Create `packages/sdk/tests/token-store.test.ts`:
```typescript
import { describe, expect, it, beforeAll, afterAll } from "bun:test";
import { join } from "node:path";
import { mkdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { TokenStore } from "../src/token-store.js";
import type { TokenData } from "../src/types.js";

const testDir = join(tmpdir(), `outlook-token-test-${Date.now()}`);

beforeAll(() => mkdirSync(testDir, { recursive: true }));
afterAll(() => rmSync(testDir, { recursive: true, force: true }));

const sampleData: TokenData = {
  accessToken: "access-abc",
  refreshToken: "refresh-xyz",
  accessTokenExpiry: Date.now() + 3600_000,
  refreshTokenExpiry: Date.now() + 90 * 86400_000,
  userEmail: "test@example.com",
};

describe("TokenStore", () => {
  it("load returns null when no file exists", async () => {
    const store = new TokenStore("no-client", testDir);
    expect(await store.load()).toBeNull();
  });

  it("save then load round-trips token data", async () => {
    const store = new TokenStore("client-1", testDir);
    await store.save(sampleData);
    const loaded = await store.load();
    expect(loaded).not.toBeNull();
    expect(loaded!.accessToken).toBe(sampleData.accessToken);
    expect(loaded!.refreshToken).toBe(sampleData.refreshToken);
    expect(loaded!.userEmail).toBe(sampleData.userEmail);
  });

  it("clear removes the file — subsequent load returns null", async () => {
    const store = new TokenStore("client-2", testDir);
    await store.save(sampleData);
    await store.clear();
    expect(await store.load()).toBeNull();
  });

  it("different clientIds use separate files", async () => {
    const store1 = new TokenStore("client-a", testDir);
    const store2 = new TokenStore("client-b", testDir);
    await store1.save({ ...sampleData, userEmail: "a@example.com" });
    await store2.save({ ...sampleData, userEmail: "b@example.com" });
    expect((await store1.load())!.userEmail).toBe("a@example.com");
    expect((await store2.load())!.userEmail).toBe("b@example.com");
  });
});
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
bun test packages/sdk/tests/token-store.test.ts
```
Expected: FAIL — `TokenStore` not found.

- [ ] **Step 3: Create token-store.ts**

Create `packages/sdk/src/token-store.ts`:
```typescript
import { randomBytes, createCipheriv, createDecipheriv } from "node:crypto";
import { existsSync, mkdirSync, chmodSync, renameSync } from "node:fs";
import { readFile, writeFile, unlink } from "node:fs/promises";
import { join } from "node:path";
import { homedir } from "node:os";
import { statSync } from "node:fs";
import { TokenDataSchema, type TokenData } from "./types.js";

const SERVICE_NAME = "outlook-toolkit";

// Optional keytar (OS keychain); gracefully absent in headless environments
let _keytar: typeof import("keytar") | null | undefined;
async function getKeytar(): Promise<typeof import("keytar") | null> {
  if (_keytar === undefined) {
    try {
      _keytar = await import("keytar");
    } catch {
      _keytar = null;
    }
  }
  return _keytar;
}

function encrypt(plaintext: string, key: Buffer): string {
  const iv = randomBytes(16);
  const cipher = createCipheriv("aes-256-cbc", key, iv);
  let enc = cipher.update(plaintext, "utf8", "hex");
  enc += cipher.final("hex");
  return `${iv.toString("hex")}:${enc}`;
}

function decrypt(ciphertext: string, key: Buffer): string {
  const colonIdx = ciphertext.indexOf(":");
  const ivHex = ciphertext.slice(0, colonIdx);
  const enc = ciphertext.slice(colonIdx + 1);
  const iv = Buffer.from(ivHex, "hex");
  const decipher = createDecipheriv("aes-256-cbc", key, iv);
  let dec = decipher.update(enc, "hex", "utf8");
  dec += decipher.final("utf8");
  return dec;
}

async function getOrCreateKey(dir: string): Promise<Buffer> {
  const keyFile = join(dir, "encryption.key");

  const kt = await getKeytar();
  if (kt) {
    try {
      const stored = await kt.getPassword(SERVICE_NAME, "encryption-key");
      if (stored) return Buffer.from(stored, "base64");
      const newKey = randomBytes(32);
      await kt.setPassword(SERVICE_NAME, "encryption-key", newKey.toString("base64"));
      return newKey;
    } catch { /* fall through to file */ }
  }

  mkdirSync(dir, { recursive: true });
  if (existsSync(keyFile)) {
    const stat = statSync(keyFile);
    if ((stat.mode & 0o777) !== 0o600) chmodSync(keyFile, 0o600);
    const raw = await readFile(keyFile, "utf8");
    const key = Buffer.from(raw.trim(), "base64");
    if (key.length !== 32) throw new Error("Invalid encryption key length in " + keyFile);
    return key;
  }

  const newKey = randomBytes(32);
  try {
    await writeFile(keyFile, newKey.toString("base64"), { flag: "wx", mode: 0o600 });
  } catch (err: unknown) {
    if ((err as NodeJS.ErrnoException).code !== "EEXIST") throw err;
    const raw = await readFile(keyFile, "utf8");
    return Buffer.from(raw.trim(), "base64");
  }
  return newKey;
}

export class TokenStore {
  private readonly dir: string;
  private readonly filePath: string;

  constructor(clientId: string, baseDir?: string) {
    this.dir = baseDir ?? join(homedir(), ".outlook-toolkit");
    const safeName = clientId.replace(/[^a-zA-Z0-9-]/g, "_");
    this.filePath = join(this.dir, `tokens-${safeName}.enc`);
  }

  async save(data: TokenData): Promise<void> {
    mkdirSync(this.dir, { recursive: true });
    const key = await getOrCreateKey(this.dir);
    const encrypted = encrypt(JSON.stringify(data), key);
    const tmp = `${this.filePath}.tmp`;
    await writeFile(tmp, encrypted, { mode: 0o600 });
    renameSync(tmp, this.filePath);
  }

  async load(): Promise<TokenData | null> {
    if (!existsSync(this.filePath)) return null;
    try {
      const key = await getOrCreateKey(this.dir);
      const raw = await readFile(this.filePath, "utf8");
      const decrypted = decrypt(raw.trim(), key);
      const parsed = TokenDataSchema.safeParse(JSON.parse(decrypted));
      return parsed.success ? parsed.data : null;
    } catch {
      return null;
    }
  }

  async clear(): Promise<void> {
    if (existsSync(this.filePath)) {
      await unlink(this.filePath);
    }
  }
}
```

- [ ] **Step 4: Run tests — confirm they pass**

```bash
bun test packages/sdk/tests/token-store.test.ts
```
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/src/token-store.ts packages/sdk/tests/token-store.test.ts
git commit -m "feat(sdk): add TokenStore with AES-256-CBC file encryption and keytar fallback"
```

---

## Task 5: Auth module — PKCE helpers + OutlookAuth class

**Files:**
- Create: `packages/sdk/src/auth.ts`
- Create: `packages/sdk/tests/auth.test.ts`

- [ ] **Step 1: Write failing tests**

Create `packages/sdk/tests/auth.test.ts`:
```typescript
import { describe, expect, it, mock, beforeEach } from "bun:test";
import {
  generateCodeVerifier,
  generateCodeChallenge,
  generateState,
  buildAuthorizationUrl,
  OutlookAuth,
} from "../src/auth.js";
import { OutlookAuthError } from "../src/errors.js";
import type { TokenData } from "../src/types.js";
import type { OutlookConfig } from "../src/types.js";

// ── Pure PKCE helpers ──────────────────────────────────────────────────────

describe("generateCodeVerifier", () => {
  it("returns a base64url string of at least 43 chars", () => {
    const v = generateCodeVerifier();
    expect(v.length).toBeGreaterThanOrEqual(43);
    expect(v).toMatch(/^[A-Za-z0-9\-_]+$/);
  });

  it("returns a different value each call", () => {
    expect(generateCodeVerifier()).not.toBe(generateCodeVerifier());
  });
});

describe("generateCodeChallenge", () => {
  it("produces a deterministic base64url SHA-256 hash", () => {
    const { createHash } = await import("node:crypto");
    const verifier = "dGVzdC12ZXJpZmllcg";
    const expected = createHash("sha256").update(verifier).digest("base64url");
    expect(generateCodeChallenge(verifier)).toBe(expected);
  });
});

describe("generateState", () => {
  it("returns a 32-char hex string", () => {
    const s = generateState();
    expect(s).toMatch(/^[0-9a-f]{32}$/);
  });
});

describe("buildAuthorizationUrl", () => {
  it("includes required OAuth params", () => {
    const url = buildAuthorizationUrl(
      "my-client",
      "consumers",
      "http://localhost:12345/callback",
      "challenge-abc",
      "state-xyz"
    );
    const parsed = new URL(url);
    expect(parsed.searchParams.get("client_id")).toBe("my-client");
    expect(parsed.searchParams.get("response_type")).toBe("code");
    expect(parsed.searchParams.get("code_challenge")).toBe("challenge-abc");
    expect(parsed.searchParams.get("code_challenge_method")).toBe("S256");
    expect(parsed.searchParams.get("state")).toBe("state-xyz");
  });

  it("uses consumers authority for tenantId='consumers'", () => {
    const url = buildAuthorizationUrl("c", "consumers", "http://localhost/cb", "ch", "st");
    expect(url).toContain("login.microsoftonline.com/consumers");
  });

  it("uses tenant authority for a GUID tenantId", () => {
    const guid = "11111111-2222-3333-4444-555555555555";
    const url = buildAuthorizationUrl("c", guid, "http://localhost/cb", "ch", "st");
    expect(url).toContain(`login.microsoftonline.com/${guid}`);
  });
});

// ── OutlookAuth class ──────────────────────────────────────────────────────

const config: OutlookConfig = { clientId: "test-client", tenantId: "consumers" };

function makeMockStore(data: TokenData | null) {
  return {
    save: mock((_data: TokenData) => Promise.resolve()),
    load: mock(() => Promise.resolve(data)),
    clear: mock(() => Promise.resolve()),
  };
}

describe("OutlookAuth.acquireToken", () => {
  it("returns access token when not near expiry", async () => {
    const store = makeMockStore({
      accessToken: "valid-token",
      refreshToken: "refresh",
      accessTokenExpiry: Date.now() + 2 * 60 * 60 * 1000, // 2 hours
      refreshTokenExpiry: Date.now() + 90 * 86400_000,
    });
    const auth = new OutlookAuth(config, store);
    expect(await auth.acquireToken()).toBe("valid-token");
  });

  it("throws OutlookAuthError when no tokens exist", async () => {
    const store = makeMockStore(null);
    const auth = new OutlookAuth(config, store);
    await expect(auth.acquireToken()).rejects.toBeInstanceOf(OutlookAuthError);
  });

  it("throws OutlookAuthError when refresh token is expired", async () => {
    const store = makeMockStore({
      accessToken: "old-token",
      refreshToken: "old-refresh",
      accessTokenExpiry: Date.now() - 10_000,
      refreshTokenExpiry: Date.now() - 10_000, // expired
    });
    const auth = new OutlookAuth(config, store);
    await expect(auth.acquireToken()).rejects.toBeInstanceOf(OutlookAuthError);
    expect(store.clear).toHaveBeenCalled();
  });
});

describe("OutlookAuth.logout", () => {
  it("calls store.clear()", async () => {
    const store = makeMockStore(null);
    const auth = new OutlookAuth(config, store);
    await auth.logout();
    expect(store.clear).toHaveBeenCalledTimes(1);
  });
});

describe("OutlookAuth.status", () => {
  it("returns null when no tokens exist", async () => {
    const store = makeMockStore(null);
    const auth = new OutlookAuth(config, store);
    expect(await auth.status()).toBeNull();
  });

  it("returns status object when tokens exist", async () => {
    const expiry = Date.now() + 3600_000;
    const store = makeMockStore({
      accessToken: "tok",
      refreshToken: "ref",
      accessTokenExpiry: expiry,
      refreshTokenExpiry: expiry + 90 * 86400_000,
      userEmail: "user@example.com",
    });
    const auth = new OutlookAuth(config, store);
    const status = await auth.status();
    expect(status).not.toBeNull();
    expect(status!.authenticated).toBe(true);
    expect(status!.userEmail).toBe("user@example.com");
  });
});
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
bun test packages/sdk/tests/auth.test.ts
```
Expected: FAIL — `auth.js` module not found.

- [ ] **Step 3: Create auth.ts**

Create `packages/sdk/src/auth.ts`:
```typescript
import { createServer } from "node:http";
import { randomBytes, createHash } from "node:crypto";
import { exec } from "node:child_process";
import { SCOPES, GRAPH_BASE_URL, resolveAuthority } from "./config.js";
import { OutlookAuthError } from "./errors.js";
import type { OutlookConfig, TokenData, AuthStatus } from "./types.js";

const REFRESH_THRESHOLD_MS = 55 * 60 * 1000; // refresh 55 min before expiry
const REFRESH_TOKEN_LIFETIME_MS = 90 * 24 * 60 * 60 * 1000; // 90 day default

interface ITokenStore {
  save(data: TokenData): Promise<void>;
  load(): Promise<TokenData | null>;
  clear(): Promise<void>;
}

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
}

// ── Pure PKCE helpers (exported for testing) ──────────────────────────────

export function generateCodeVerifier(): string {
  return randomBytes(32).toString("base64url");
}

export function generateCodeChallenge(verifier: string): string {
  return createHash("sha256").update(verifier).digest("base64url");
}

export function generateState(): string {
  return randomBytes(16).toString("hex");
}

export function buildAuthorizationUrl(
  clientId: string,
  tenantId: string,
  redirectUri: string,
  codeChallenge: string,
  state: string
): string {
  const authority = resolveAuthority(tenantId);
  const params = new URLSearchParams({
    client_id: clientId,
    response_type: "code",
    redirect_uri: redirectUri,
    response_mode: "query",
    scope: SCOPES.join(" "),
    code_challenge: codeChallenge,
    code_challenge_method: "S256",
    state,
    prompt: "select_account",
  });
  return `${authority}/oauth2/v2.0/authorize?${params.toString()}`;
}

// ── Callback server ───────────────────────────────────────────────────────

async function startCallbackServer(
  expectedState: string,
  timeoutMs = 120_000
): Promise<{ promise: Promise<string>; port: number }> {
  let resolveCode: (code: string) => void;
  let rejectCode: (err: Error) => void;
  let settled = false;

  const promise = new Promise<string>((res, rej) => {
    resolveCode = (c) => { settled = true; res(c); };
    rejectCode = (e) => { settled = true; rej(e); };
  });

  const server = createServer((req, res) => {
    const url = new URL(req.url ?? "/", "http://localhost");
    if (url.pathname !== "/callback") { res.writeHead(404); res.end(); return; }

    const code = url.searchParams.get("code");
    const state = url.searchParams.get("state");
    const error = url.searchParams.get("error");

    if (error) {
      const desc = url.searchParams.get("error_description") ?? error;
      res.writeHead(400, { "Content-Type": "text/html" });
      res.end(`<h1>Auth failed</h1><p>${desc}</p>`);
      rejectCode(new OutlookAuthError(`OAuth error: ${desc}`));
      return;
    }
    if (state !== expectedState) {
      res.writeHead(400, { "Content-Type": "text/html" });
      res.end("<h1>State mismatch</h1>");
      rejectCode(new Error("State mismatch — possible CSRF"));
      return;
    }
    if (!code) {
      res.writeHead(400, { "Content-Type": "text/html" });
      res.end("<h1>No code received</h1>");
      rejectCode(new Error("No authorization code in callback"));
      return;
    }
    res.writeHead(200, { "Content-Type": "text/html" });
    res.end("<h1>Authenticated!</h1><p>You can close this tab.</p>");
    resolveCode(code);
  });

  const port = await new Promise<number>((res) =>
    server.listen(0, "127.0.0.1", () => {
      const addr = server.address();
      res(typeof addr === "object" && addr ? addr.port : 0);
    })
  );

  const timer = setTimeout(() => {
    if (!settled) rejectCode(new Error("Authentication timed out after 2 minutes."));
    server.close();
  }, timeoutMs);

  promise.then(
    () => { clearTimeout(timer); server.close(); },
    () => { clearTimeout(timer); server.close(); }
  );

  return { promise, port };
}

// ── Token exchange and refresh ───────────────────────────────────────────

async function exchangeCode(
  clientId: string,
  tenantId: string,
  code: string,
  redirectUri: string,
  codeVerifier: string
): Promise<TokenResponse> {
  const authority = resolveAuthority(tenantId);
  const body = new URLSearchParams({
    client_id: clientId,
    scope: SCOPES.join(" "),
    code,
    redirect_uri: redirectUri,
    grant_type: "authorization_code",
    code_verifier: codeVerifier,
  });
  const res = await fetch(`${authority}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new OutlookAuthError(`Token exchange failed (${res.status}): ${text}`);
  }
  return res.json() as Promise<TokenResponse>;
}

async function refreshTokenRequest(
  clientId: string,
  tenantId: string,
  refreshToken: string
): Promise<TokenResponse> {
  const authority = resolveAuthority(tenantId);
  const body = new URLSearchParams({
    client_id: clientId,
    scope: SCOPES.join(" "),
    refresh_token: refreshToken,
    grant_type: "refresh_token",
  });
  const res = await fetch(`${authority}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new OutlookAuthError(`Token refresh failed (${res.status}): ${text}`);
  }
  return res.json() as Promise<TokenResponse>;
}

async function fetchUserEmail(accessToken: string): Promise<string> {
  try {
    const res = await fetch(`${GRAPH_BASE_URL}/me?$select=mail,userPrincipalName`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (!res.ok) return "unknown";
    const data = (await res.json()) as { mail?: string; userPrincipalName?: string };
    return data.mail ?? data.userPrincipalName ?? "unknown";
  } catch {
    return "unknown";
  }
}

function openBrowserDefault(url: string): void {
  const cmd =
    process.platform === "darwin"
      ? `open "${url}"`
      : process.platform === "win32"
      ? `start "" "${url}"`
      : `xdg-open "${url}"`;
  exec(cmd, () => {});
}

// ── OutlookAuth class ────────────────────────────────────────────────────

export class OutlookAuth {
  private pendingRefresh: Promise<string> | null = null;

  constructor(
    private readonly config: OutlookConfig,
    private readonly store: ITokenStore
  ) {}

  async login(
    openBrowser: (url: string) => void | Promise<void> = openBrowserDefault
  ): Promise<{ userEmail: string }> {
    const verifier = generateCodeVerifier();
    const challenge = generateCodeChallenge(verifier);
    const state = generateState();
    const { promise: codePromise, port } = await startCallbackServer(state);
    const redirectUri = `http://localhost:${port}/callback`;
    const url = buildAuthorizationUrl(
      this.config.clientId,
      this.config.tenantId,
      redirectUri,
      challenge,
      state
    );

    await openBrowser(url);

    const code = await codePromise;
    const tokens = await exchangeCode(
      this.config.clientId,
      this.config.tenantId,
      code,
      redirectUri,
      verifier
    );
    const userEmail = await fetchUserEmail(tokens.access_token);

    await this.store.save({
      accessToken: tokens.access_token,
      refreshToken: tokens.refresh_token ?? "",
      accessTokenExpiry: Date.now() + tokens.expires_in * 1000,
      refreshTokenExpiry: Date.now() + REFRESH_TOKEN_LIFETIME_MS,
      userEmail,
    });

    return { userEmail };
  }

  async logout(): Promise<void> {
    await this.store.clear();
  }

  async acquireToken(): Promise<string> {
    if (this.pendingRefresh) return this.pendingRefresh;

    const data = await this.store.load();
    if (!data) throw new OutlookAuthError();

    const now = Date.now();

    if (now > data.refreshTokenExpiry) {
      await this.store.clear();
      throw new OutlookAuthError("Session expired. Run `outlook auth login` again.");
    }

    if (now < data.accessTokenExpiry - REFRESH_THRESHOLD_MS) {
      return data.accessToken;
    }

    this.pendingRefresh = this._doRefresh(data.refreshToken).finally(() => {
      this.pendingRefresh = null;
    });
    return this.pendingRefresh;
  }

  private async _doRefresh(refreshToken: string): Promise<string> {
    const tokens = await refreshTokenRequest(
      this.config.clientId,
      this.config.tenantId,
      refreshToken
    );
    const existing = await this.store.load();
    await this.store.save({
      accessToken: tokens.access_token,
      refreshToken: tokens.refresh_token ?? refreshToken,
      accessTokenExpiry: Date.now() + tokens.expires_in * 1000,
      refreshTokenExpiry: existing?.refreshTokenExpiry ?? Date.now() + REFRESH_TOKEN_LIFETIME_MS,
      userEmail: existing?.userEmail,
    });
    return tokens.access_token;
  }

  async status(): Promise<AuthStatus | null> {
    const data = await this.store.load();
    if (!data) return null;
    return {
      authenticated: true,
      userEmail: data.userEmail,
      accessTokenExpiry: data.accessTokenExpiry,
      refreshTokenExpiry: data.refreshTokenExpiry,
    };
  }
}
```

- [ ] **Step 4: Run tests — confirm they pass**

```bash
bun test packages/sdk/tests/auth.test.ts
```
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/src/auth.ts packages/sdk/tests/auth.test.ts
git commit -m "feat(sdk): add PKCE auth helpers and OutlookAuth class"
```

---

## Task 6: Profile store

**Files:**
- Create: `packages/sdk/src/profile-store.ts`
- Create: `packages/sdk/tests/profile-store.test.ts`

- [ ] **Step 1: Write failing tests**

Create `packages/sdk/tests/profile-store.test.ts`:
```typescript
import { describe, expect, it, beforeAll, afterAll } from "bun:test";
import { join } from "node:path";
import { mkdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { ProfileStore } from "../src/profile-store.js";

const testDir = join(tmpdir(), `outlook-profile-test-${Date.now()}`);

beforeAll(() => mkdirSync(testDir, { recursive: true }));
afterAll(() => rmSync(testDir, { recursive: true, force: true }));

describe("ProfileStore", () => {
  it("get returns null for unknown profile", async () => {
    const store = new ProfileStore(testDir);
    expect(await store.get("nonexistent")).toBeNull();
  });

  it("save then get round-trips profile data", async () => {
    const store = new ProfileStore(testDir);
    await store.save("personal", { clientId: "abc", tenantId: "consumers" });
    const p = await store.get("personal");
    expect(p).not.toBeNull();
    expect(p!.clientId).toBe("abc");
    expect(p!.tenantId).toBe("consumers");
  });

  it("list returns all saved profiles", async () => {
    const store = new ProfileStore(testDir);
    await store.save("p1", { clientId: "c1", tenantId: "consumers" });
    await store.save("p2", { clientId: "c2", tenantId: "work-guid" });
    const all = await store.list();
    expect(all["p1"]?.clientId).toBe("c1");
    expect(all["p2"]?.clientId).toBe("c2");
  });

  it("delete removes a profile and returns true", async () => {
    const store = new ProfileStore(testDir);
    await store.save("temp", { clientId: "t", tenantId: "consumers" });
    expect(await store.delete("temp")).toBe(true);
    expect(await store.get("temp")).toBeNull();
  });

  it("delete returns false for unknown profile", async () => {
    const store = new ProfileStore(testDir);
    expect(await store.delete("ghost")).toBe(false);
  });
});
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
bun test packages/sdk/tests/profile-store.test.ts
```
Expected: FAIL — `ProfileStore` not found.

- [ ] **Step 3: Create profile-store.ts**

Create `packages/sdk/src/profile-store.ts`:
```typescript
import { existsSync, mkdirSync } from "node:fs";
import { readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { homedir } from "node:os";
import { ProfilesFileSchema, ProfileSchema, type Profile, type ProfilesFile } from "./types.js";

export class ProfileStore {
  private readonly filePath: string;

  constructor(baseDir?: string) {
    const dir = baseDir ?? join(homedir(), ".outlook-toolkit");
    this.filePath = join(dir, "profiles.json");
  }

  private async read(): Promise<ProfilesFile> {
    if (!existsSync(this.filePath)) return {};
    try {
      const raw = await readFile(this.filePath, "utf8");
      const parsed = ProfilesFileSchema.safeParse(JSON.parse(raw));
      return parsed.success ? parsed.data : {};
    } catch {
      return {};
    }
  }

  private async write(data: ProfilesFile): Promise<void> {
    const dir = join(this.filePath, "..");
    mkdirSync(dir, { recursive: true });
    await writeFile(this.filePath, JSON.stringify(data, null, 2), "utf8");
  }

  async save(name: string, profile: Profile): Promise<void> {
    const validated = ProfileSchema.parse(profile);
    const all = await this.read();
    all[name] = validated;
    await this.write(all);
  }

  async get(name: string): Promise<Profile | null> {
    const all = await this.read();
    return all[name] ?? null;
  }

  async list(): Promise<ProfilesFile> {
    return this.read();
  }

  async delete(name: string): Promise<boolean> {
    const all = await this.read();
    if (!(name in all)) return false;
    delete all[name];
    await this.write(all);
    return true;
  }
}
```

- [ ] **Step 4: Run tests — confirm they pass**

```bash
bun test packages/sdk/tests/profile-store.test.ts
```
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/src/profile-store.ts packages/sdk/tests/profile-store.test.ts
git commit -m "feat(sdk): add ProfileStore for named account profiles"
```

---

## Task 7: Graph client — fetch wrapper with retry/backoff

**Files:**
- Create: `packages/sdk/src/graph-client.ts`
- Create: `packages/sdk/tests/graph-client.test.ts`

- [ ] **Step 1: Write failing tests**

Create `packages/sdk/tests/graph-client.test.ts`:
```typescript
import { describe, expect, it, mock, beforeEach, afterEach } from "bun:test";
import { GraphClient } from "../src/graph-client.js";
import { OutlookAuthError, OutlookNotFoundError, OutlookRateLimitError } from "../src/errors.js";

type MockResponse = {
  ok: boolean;
  status: number;
  json: () => Promise<unknown>;
  text: () => Promise<string>;
  headers: { get: (key: string) => string | null };
};

function makeFetch(response: MockResponse) {
  return mock(() => Promise.resolve(response as unknown as Response));
}

function makeOk(body: unknown): MockResponse {
  return {
    ok: true,
    status: 200,
    json: () => Promise.resolve(body),
    text: () => Promise.resolve(JSON.stringify(body)),
    headers: { get: () => null },
  };
}

function makeError(status: number, retryAfter?: string): MockResponse {
  return {
    ok: false,
    status,
    json: () => Promise.resolve({ error: { code: "err", message: "fail" } }),
    text: () => Promise.resolve("error"),
    headers: { get: (k) => (k === "retry-after" ? (retryAfter ?? null) : null) },
  };
}

describe("GraphClient", () => {
  let originalFetch: typeof globalThis.fetch;

  beforeEach(() => { originalFetch = globalThis.fetch; });
  afterEach(() => { globalThis.fetch = originalFetch; });

  it("get() calls the correct URL with Authorization header", async () => {
    const fetchMock = makeFetch(makeOk({ id: "123" }));
    globalThis.fetch = fetchMock as unknown as typeof fetch;

    const client = new GraphClient("test-token");
    await client.get("/me");

    expect(fetchMock).toHaveBeenCalledTimes(1);
    const [url, opts] = (fetchMock.mock.calls[0] ?? []) as [string, RequestInit];
    expect(url).toContain("/v1.0/me");
    expect((opts.headers as Record<string, string>)["Authorization"]).toBe("Bearer test-token");
  });

  it("get() throws OutlookAuthError on 401", async () => {
    globalThis.fetch = makeFetch(makeError(401)) as unknown as typeof fetch;
    const client = new GraphClient("bad-token");
    await expect(client.get("/me")).rejects.toBeInstanceOf(OutlookAuthError);
  });

  it("get() throws OutlookNotFoundError on 404", async () => {
    globalThis.fetch = makeFetch(makeError(404)) as unknown as typeof fetch;
    const client = new GraphClient("tok");
    await expect(client.get("/me/messages/nope")).rejects.toBeInstanceOf(OutlookNotFoundError);
  });

  it("get() throws OutlookRateLimitError after exhausting retries on 429", async () => {
    let callCount = 0;
    globalThis.fetch = mock(() => {
      callCount++;
      return Promise.resolve(makeError(429) as unknown as Response);
    }) as unknown as typeof fetch;

    const client = new GraphClient("tok", { maxAttempts: 2, initialDelayMs: 1, maxDelayMs: 1 });
    await expect(client.get("/me/messages")).rejects.toBeInstanceOf(OutlookRateLimitError);
    expect(callCount).toBe(3); // 1 initial + 2 retries
  });

  it("list() returns paginated response with nextLink", async () => {
    const body = { value: [{ id: "1" }], "@odata.nextLink": "https://..." };
    globalThis.fetch = makeFetch(makeOk(body)) as unknown as typeof fetch;
    const client = new GraphClient("tok");
    const result = await client.list<{ id: string }>("/me/messages");
    expect(result.value).toHaveLength(1);
    expect(result["@odata.nextLink"]).toBe("https://...");
  });

  it("post() sends JSON body", async () => {
    const fetchMock = makeFetch(makeOk(null));
    globalThis.fetch = fetchMock as unknown as typeof fetch;
    const client = new GraphClient("tok");
    await client.post("/me/sendMail", { message: { subject: "hi" } });
    const [, opts] = (fetchMock.mock.calls[0] ?? []) as [string, RequestInit];
    expect(opts.method).toBe("POST");
    expect(opts.body).toContain("hi");
  });
});
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
bun test packages/sdk/tests/graph-client.test.ts
```
Expected: FAIL — `GraphClient` not found.

- [ ] **Step 3: Create graph-client.ts**

Create `packages/sdk/src/graph-client.ts`:
```typescript
import { GRAPH_BASE_URL } from "./config.js";
import {
  OutlookAuthError,
  OutlookNotFoundError,
  OutlookRateLimitError,
  OutlookError,
} from "./errors.js";

export interface ODataOptions {
  $select?: string;
  $filter?: string;
  $top?: number;
  $orderby?: string;
  $search?: string;
  $expand?: string;
}

export interface ODataListResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.deltaLink"?: string;
}

interface RetryOptions {
  maxAttempts: number;
  initialDelayMs: number;
  maxDelayMs: number;
}

const DEFAULT_RETRY: RetryOptions = {
  maxAttempts: 2,
  initialDelayMs: 1000,
  maxDelayMs: 30_000,
};

const RETRYABLE = new Set([429, 503, 504]);

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

function calcDelay(err: unknown, attempt: number, opts: RetryOptions): number {
  // Respect Retry-After header if present
  const resp = err instanceof GraphResponseError ? err.response : null;
  const retryAfter = resp?.headers.get("retry-after");
  if (retryAfter) {
    const n = parseInt(retryAfter, 10);
    if (!isNaN(n)) return n * 1000;
  }
  const base = opts.initialDelayMs * Math.pow(2, attempt - 1);
  const capped = Math.min(base, opts.maxDelayMs);
  return Math.round(capped * (0.75 + Math.random() * 0.5)); // ±25% jitter
}

class GraphResponseError extends Error {
  constructor(public readonly response: Response) {
    super(`HTTP ${response.status}`);
  }
}

export class GraphClient {
  private readonly retry: RetryOptions;

  constructor(
    private readonly accessToken: string,
    retryOpts?: Partial<RetryOptions>
  ) {
    this.retry = { ...DEFAULT_RETRY, ...retryOpts };
  }

  private url(path: string, opts?: ODataOptions): string {
    const base = `${GRAPH_BASE_URL}${path}`;
    if (!opts) return base;
    const params = new URLSearchParams();
    if (opts.$select) params.set("$select", opts.$select);
    if (opts.$filter) params.set("$filter", opts.$filter);
    if (opts.$top !== undefined) params.set("$top", String(opts.$top));
    if (opts.$orderby) params.set("$orderby", opts.$orderby);
    if (opts.$search) params.set("$search", opts.$search);
    if (opts.$expand) params.set("$expand", opts.$expand);
    const qs = params.toString();
    return qs ? `${base}?${qs}` : base;
  }

  private headers(extra?: Record<string, string>): Record<string, string> {
    return {
      Authorization: `Bearer ${this.accessToken}`,
      "Content-Type": "application/json",
      ...extra,
    };
  }

  private async request<T>(
    method: string,
    path: string,
    body?: unknown,
    odataOpts?: ODataOptions
  ): Promise<T> {
    let attempt = 0;
    while (true) {
      const res = await fetch(this.url(path, odataOpts), {
        method,
        headers: this.headers(),
        body: body !== undefined ? JSON.stringify(body) : undefined,
      });

      if (res.ok) {
        if (res.status === 204) return undefined as T;
        return res.json() as Promise<T>;
      }

      attempt++;

      if (res.status === 401) throw new OutlookAuthError("Access token rejected. Re-authenticate.");
      if (res.status === 404) {
        throw new OutlookNotFoundError("resource", path);
      }

      if (RETRYABLE.has(res.status) && attempt <= this.retry.maxAttempts) {
        const err = new GraphResponseError(res);
        const delay = calcDelay(err, attempt, this.retry);
        await sleep(delay);
        continue;
      }

      if (res.status === 429) {
        const retryAfter = res.headers.get("retry-after");
        throw new OutlookRateLimitError(retryAfter ? parseInt(retryAfter, 10) : undefined);
      }

      const text = await res.text().catch(() => "");
      throw new OutlookError(`Graph API error ${res.status}: ${text}`, "GRAPH_ERROR", res.status);
    }
  }

  async get<T>(path: string, opts?: ODataOptions): Promise<T> {
    return this.request<T>("GET", path, undefined, opts);
  }

  async list<T>(path: string, opts?: ODataOptions): Promise<ODataListResponse<T>> {
    return this.request<ODataListResponse<T>>("GET", path, undefined, opts);
  }

  async post<TReq, TRes = void>(path: string, body: TReq): Promise<TRes> {
    return this.request<TRes>("POST", path, body);
  }

  async patch(path: string, body: unknown): Promise<void> {
    await this.request<void>("PATCH", path, body);
  }

  async delete(path: string): Promise<void> {
    await this.request<void>("DELETE", path);
  }
}
```

- [ ] **Step 4: Run tests — confirm they pass**

```bash
bun test packages/sdk/tests/graph-client.test.ts
```
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/src/graph-client.ts packages/sdk/tests/graph-client.test.ts
git commit -m "feat(sdk): add GraphClient with OData helpers and retry/backoff"
```

---

## Task 8: Mail client — list, get, send, reply, createDraft, sync

**Files:**
- Create: `packages/sdk/src/mail-client.ts`
- Create: `packages/sdk/tests/mail-client.test.ts`

- [ ] **Step 1: Write failing tests**

Create `packages/sdk/tests/mail-client.test.ts`:
```typescript
import { describe, expect, it, mock } from "bun:test";
import { MailClient } from "../src/mail-client.js";
import type { GraphClient } from "../src/graph-client.js";

function makeGraphClient(overrides: Partial<GraphClient> = {}): GraphClient {
  return {
    get: mock(() => Promise.resolve({})),
    list: mock(() => Promise.resolve({ value: [] })),
    post: mock(() => Promise.resolve(undefined)),
    patch: mock(() => Promise.resolve()),
    delete: mock(() => Promise.resolve()),
    ...overrides,
  } as unknown as GraphClient;
}

describe("MailClient.list", () => {
  it("calls inbox messages endpoint with defaults", async () => {
    const gc = makeGraphClient();
    const client = new MailClient(gc);
    await client.list();
    expect(gc.list).toHaveBeenCalledTimes(1);
    const [path, opts] = (gc.list as ReturnType<typeof mock>).mock.calls[0] as [string, unknown];
    expect(path).toBe("/me/mailFolders/inbox/messages");
    expect((opts as { $top?: number }).$top).toBe(25);
  });

  it("passes custom folder and limit", async () => {
    const gc = makeGraphClient();
    const client = new MailClient(gc);
    await client.list({ folder: "sentitems", limit: 10 });
    const [path, opts] = (gc.list as ReturnType<typeof mock>).mock.calls[0] as [string, unknown];
    expect(path).toBe("/me/mailFolders/sentitems/messages");
    expect((opts as { $top?: number }).$top).toBe(10);
  });
});

describe("MailClient.get", () => {
  it("calls /me/messages/{id}", async () => {
    const gc = makeGraphClient();
    const client = new MailClient(gc);
    await client.get("msg-abc");
    const [path] = (gc.get as ReturnType<typeof mock>).mock.calls[0] as [string];
    expect(path).toBe("/me/messages/msg-abc");
  });
});

describe("MailClient.send", () => {
  it("posts to /me/sendMail with correct shape", async () => {
    const gc = makeGraphClient();
    const client = new MailClient(gc);
    await client.send({ to: "user@example.com", subject: "Hi", body: "<p>Hello</p>" });
    const [path, body] = (gc.post as ReturnType<typeof mock>).mock.calls[0] as [string, unknown];
    expect(path).toBe("/me/sendMail");
    const b = body as { message: { toRecipients: { emailAddress: { address: string } }[] } };
    expect(b.message.toRecipients[0]!.emailAddress.address).toBe("user@example.com");
  });
});

describe("MailClient.reply", () => {
  it("posts to /me/messages/{id}/reply", async () => {
    const gc = makeGraphClient();
    const client = new MailClient(gc);
    await client.reply("msg-xyz", { body: "Reply text" });
    const [path] = (gc.post as ReturnType<typeof mock>).mock.calls[0] as [string];
    expect(path).toBe("/me/messages/msg-xyz/reply");
  });
});

describe("MailClient.createDraft", () => {
  it("posts to /me/messages", async () => {
    const gc = makeGraphClient();
    const client = new MailClient(gc);
    await client.createDraft({ to: "a@b.com", subject: "Draft", body: "Draft body" });
    const [path] = (gc.post as ReturnType<typeof mock>).mock.calls[0] as [string];
    expect(path).toBe("/me/messages");
  });
});

describe("MailClient.sync", () => {
  it("calls delta endpoint without deltaLink on initial sync", async () => {
    const gc = makeGraphClient();
    const client = new MailClient(gc);
    await client.sync();
    const [path] = (gc.list as ReturnType<typeof mock>).mock.calls[0] as [string];
    expect(path).toBe("/me/mailFolders/inbox/messages/delta");
  });

  it("fetches deltaLink URL directly when provided", async () => {
    const gc = makeGraphClient();
    const client = new MailClient(gc);
    const deltaLink = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$deltaToken=abc";
    await client.sync(deltaLink);
    const [path] = (gc.list as ReturnType<typeof mock>).mock.calls[0] as [string];
    expect(path).toContain("deltaToken=abc");
  });
});
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
bun test packages/sdk/tests/mail-client.test.ts
```
Expected: FAIL — `MailClient` not found.

- [ ] **Step 3: Create mail-client.ts**

Create `packages/sdk/src/mail-client.ts`:
```typescript
import type { GraphClient } from "./graph-client.js";
import type { ODataListResponse } from "./graph-client.js";
import type {
  Message,
  MailListResponse,
  DeltaResponse,
  ListMailParams,
  SendMailParams,
  ReplyParams,
  DraftParams,
} from "./types.js";
import { GRAPH_BASE_URL } from "./config.js";

export class MailClient {
  constructor(private readonly graph: GraphClient) {}

  async list(params: Partial<ListMailParams> = {}): Promise<MailListResponse> {
    const folder = params.folder ?? "inbox";
    const opts = {
      $top: params.limit ?? 25,
      ...(params.filter && { $filter: params.filter }),
      ...(params.select && { $select: params.select }),
      ...(params.orderby && { $orderby: params.orderby }),
    };
    const result = await this.graph.list<Message>(
      `/me/mailFolders/${folder}/messages`,
      opts
    );
    return result as MailListResponse;
  }

  async get(id: string): Promise<Message> {
    return this.graph.get<Message>(`/me/messages/${id}`);
  }

  async send(params: SendMailParams): Promise<void> {
    await this.graph.post("/me/sendMail", {
      message: {
        subject: params.subject,
        body: {
          contentType: params.contentType ?? "HTML",
          content: params.body,
        },
        toRecipients: [{ emailAddress: { address: params.to } }],
      },
      saveToSentItems: true,
    });
  }

  async reply(id: string, params: ReplyParams): Promise<void> {
    await this.graph.post(`/me/messages/${id}/reply`, {
      message: {
        body: {
          contentType: params.contentType ?? "HTML",
          content: params.body,
        },
      },
    });
  }

  async createDraft(params: DraftParams): Promise<Message> {
    return this.graph.post<object, Message>("/me/messages", {
      subject: params.subject,
      body: {
        contentType: params.contentType ?? "HTML",
        content: params.body,
      },
      toRecipients: [{ emailAddress: { address: params.to } }],
      isDraft: true,
    });
  }

  async sync(deltaLink?: string): Promise<DeltaResponse> {
    // If a deltaLink is provided, use it as the path directly (it's a full URL)
    const path = deltaLink
      ? deltaLink.replace(GRAPH_BASE_URL, "")
      : "/me/mailFolders/inbox/messages/delta";
    const result = await this.graph.list<Message>(path);
    return result as DeltaResponse;
  }
}
```

- [ ] **Step 4: Run tests — confirm they pass**

```bash
bun test packages/sdk/tests/mail-client.test.ts
```
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/src/mail-client.ts packages/sdk/tests/mail-client.test.ts
git commit -m "feat(sdk): add MailClient (list/get/send/reply/createDraft/sync)"
```

---

## Task 9: Update SDK index.ts — wire all exports

**Files:**
- Modify: `packages/sdk/src/index.ts`

- [ ] **Step 1: Rewrite index.ts**

Replace `packages/sdk/src/index.ts` entirely:
```typescript
// Config
export { resolveConfig, resolveAuthority, SCOPES, GRAPH_BASE_URL, AUTH_BASE_URL } from "./config.js";

// Types
export type {
  OutlookConfig,
  TokenData,
  AuthStatus,
  Message,
  MessageBody,
  EmailAddress,
  Recipient,
  MailListResponse,
  DeltaResponse,
  ListMailParams,
  SendMailParams,
  ReplyParams,
  DraftParams,
  Profile,
  ProfilesFile,
} from "./types.js";
export {
  OutlookConfigSchema,
  TokenDataSchema,
  AuthStatusSchema,
  MessageSchema,
  MailListResponseSchema,
  DeltaResponseSchema,
  ListMailParamsSchema,
  SendMailParamsSchema,
  ReplyParamsSchema,
  DraftParamsSchema,
  ProfileSchema,
  ProfilesFileSchema,
} from "./types.js";

// Errors
export {
  OutlookError,
  OutlookConfigError,
  OutlookAuthError,
  OutlookNotFoundError,
  OutlookRateLimitError,
} from "./errors.js";

// Auth
export {
  OutlookAuth,
  generateCodeVerifier,
  generateCodeChallenge,
  generateState,
  buildAuthorizationUrl,
} from "./auth.js";

// Token + Profile storage
export { TokenStore } from "./token-store.js";
export { ProfileStore } from "./profile-store.js";

// Graph + Mail
export { GraphClient } from "./graph-client.js";
export type { ODataOptions, ODataListResponse } from "./graph-client.js";
export { MailClient } from "./mail-client.js";
```

- [ ] **Step 2: Run full SDK test suite**

```bash
bun run --filter '@outlook-toolkit/sdk' test
```
Expected: all tests PASS, no import errors.

- [ ] **Step 3: Run lint**

```bash
bun run --filter '@outlook-toolkit/sdk' lint
```
Expected: no TypeScript errors.

- [ ] **Step 4: Commit**

```bash
git add packages/sdk/src/index.ts
git commit -m "feat(sdk): update index exports for full Outlook API surface"
```

---

## Task 10: CLI — context helper + auth + profile commands

**Files:**
- Create: `packages/cli/src/context.ts`
- Create: `packages/cli/src/commands/auth.ts`
- Create: `packages/cli/src/commands/profile.ts`

- [ ] **Step 1: Create context.ts — shared config resolver for CLI**

Create `packages/cli/src/context.ts`:
```typescript
import {
  resolveConfig,
  ProfileStore,
  type OutlookConfig,
  OutlookConfigError,
} from "@outlook-toolkit/sdk";

export async function resolveCliConfig(profile?: string): Promise<OutlookConfig> {
  const profileName = profile ?? process.env.OUTLOOK_PROFILE;
  if (profileName) {
    const store = new ProfileStore();
    const p = await store.get(profileName);
    if (!p) {
      console.error(
        `error: profile "${profileName}" not found (config: ${3}). Run: outlook profile list`
      );
      process.exit(3);
    }
    try {
      return resolveConfig(p);
    } catch (err) {
      console.error(err instanceof Error ? err.message : String(err));
      process.exit(3);
    }
  }
  try {
    return resolveConfig();
  } catch (err) {
    console.error(err instanceof Error ? err.message : String(err));
    process.exit(3);
  }
}
```

- [ ] **Step 2: Create commands/auth.ts**

Create `packages/cli/src/commands/auth.ts`:
```typescript
import { buildCommand, buildRouteMap } from "@stricli/core";
import { OutlookAuth, TokenStore } from "@outlook-toolkit/sdk";
import { resolveCliConfig } from "../context.js";
import { exec } from "node:child_process";

function openBrowser(url: string): void {
  const cmd =
    process.platform === "darwin"
      ? `open "${url}"`
      : process.platform === "win32"
      ? `start "" "${url}"`
      : `xdg-open "${url}"`;
  exec(cmd, () => {});
  process.stderr.write(`Opening browser for Microsoft sign-in...\n`);
  process.stderr.write(`If browser does not open, visit:\n${url}\n`);
}

const loginCommand = buildCommand({
  docs: { brief: "Sign in to an Outlook account (interactive browser flow)" },
  parameters: {
    flags: {
      profile: {
        kind: "parseable",
        brief: "Named profile to use",
        parse: String,
        optional: true,
      },
    },
  },
  async func(this: void, flags: { profile?: string }) {
    const config = await resolveCliConfig(flags.profile);
    const store = new TokenStore(config.clientId);
    const auth = new OutlookAuth(config, store);
    try {
      const { userEmail } = await auth.login(openBrowser);
      console.log(`Authenticated as: ${userEmail}`);
    } catch (err) {
      console.error(`Login failed: ${err instanceof Error ? err.message : err}`);
      process.exit(5);
    }
  },
});

const logoutCommand = buildCommand({
  docs: { brief: "Sign out and clear saved tokens" },
  parameters: {
    flags: {
      profile: {
        kind: "parseable",
        brief: "Named profile to use",
        parse: String,
        optional: true,
      },
    },
  },
  async func(this: void, flags: { profile?: string }) {
    const config = await resolveCliConfig(flags.profile);
    const store = new TokenStore(config.clientId);
    const auth = new OutlookAuth(config, store);
    await auth.logout();
    console.log("Signed out.");
  },
});

const statusCommand = buildCommand({
  docs: { brief: "Show current authentication status" },
  parameters: {
    flags: {
      profile: {
        kind: "parseable",
        brief: "Named profile to use",
        parse: String,
        optional: true,
      },
      json: {
        kind: "boolean",
        brief: "Output as JSON",
        default: false,
      },
    },
  },
  async func(this: void, flags: { profile?: string; json: boolean }) {
    const config = await resolveCliConfig(flags.profile);
    const store = new TokenStore(config.clientId);
    const auth = new OutlookAuth(config, store);
    const status = await auth.status();

    if (!status) {
      const out = { authenticated: false };
      flags.json
        ? console.log(JSON.stringify(out, null, 2))
        : console.log("Not authenticated. Run: outlook auth login");
      return;
    }

    if (flags.json) {
      console.log(JSON.stringify(status, null, 2));
    } else {
      console.log(`Authenticated: ${status.userEmail ?? "unknown"}`);
      if (status.accessTokenExpiry) {
        console.log(`Access token expires: ${new Date(status.accessTokenExpiry).toISOString()}`);
      }
    }
  },
});

export const authRoutes = buildRouteMap({
  routes: { login: loginCommand, logout: logoutCommand, status: statusCommand },
  docs: { brief: "Manage Outlook authentication" },
});
```

- [ ] **Step 3: Create commands/profile.ts**

Create `packages/cli/src/commands/profile.ts`:
```typescript
import { buildCommand, buildRouteMap } from "@stricli/core";
import { ProfileStore } from "@outlook-toolkit/sdk";

const saveCommand = buildCommand({
  docs: { brief: "Save a named profile" },
  parameters: {
    flags: {
      clientId: {
        kind: "parseable",
        brief: "Azure app client ID",
        parse: String,
      },
      tenantId: {
        kind: "parseable",
        brief: "Tenant ID (consumers, common, or GUID)",
        parse: String,
      },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Profile name", parse: String }],
    },
  },
  async func(this: void, flags: { clientId: string; tenantId: string }, name: string) {
    const store = new ProfileStore();
    await store.save(name, { clientId: flags.clientId, tenantId: flags.tenantId });
    console.log(`Profile "${name}" saved.`);
  },
});

const listCommand = buildCommand({
  docs: { brief: "List all saved profiles" },
  parameters: {
    flags: {
      json: {
        kind: "boolean",
        brief: "Output as JSON",
        default: false,
      },
    },
  },
  async func(this: void, flags: { json: boolean }) {
    const store = new ProfileStore();
    const profiles = await store.list();
    const names = Object.keys(profiles);
    if (names.length === 0) {
      console.log("No profiles saved. Run: outlook profile save <name> --client-id=... --tenant-id=...");
      return;
    }
    if (flags.json) {
      console.log(JSON.stringify(profiles, null, 2));
    } else {
      for (const [name, p] of Object.entries(profiles)) {
        console.log(`${name}  clientId=${p.clientId}  tenantId=${p.tenantId}`);
      }
    }
  },
});

const deleteCommand = buildCommand({
  docs: { brief: "Delete a saved profile" },
  parameters: {
    flags: {
      force: {
        kind: "boolean",
        brief: "Skip confirmation",
        default: false,
      },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Profile name", parse: String }],
    },
  },
  async func(this: void, _flags: { force: boolean }, name: string) {
    const store = new ProfileStore();
    const deleted = await store.delete(name);
    if (!deleted) {
      console.error(`error: profile "${name}" not found`);
      process.exit(4);
    }
    console.log(`Profile "${name}" deleted.`);
  },
});

export const profileRoutes = buildRouteMap({
  routes: { save: saveCommand, list: listCommand, delete: deleteCommand },
  docs: { brief: "Manage named account profiles" },
});
```

- [ ] **Step 4: Run lint**

```bash
bun run --filter '@outlook-toolkit/cli' lint
```
Expected: no TypeScript errors.

- [ ] **Step 5: Commit**

```bash
git add packages/cli/src/context.ts packages/cli/src/commands/auth.ts packages/cli/src/commands/profile.ts
git commit -m "feat(cli): add auth login/logout/status and profile save/list/delete commands"
```

---

## Task 11: CLI — mail commands + agent-context

**Files:**
- Create: `packages/cli/src/commands/mail.ts`
- Create: `packages/cli/src/commands/agent-context.ts`

- [ ] **Step 1: Create commands/mail.ts**

Create `packages/cli/src/commands/mail.ts`:
```typescript
import { buildCommand, buildRouteMap } from "@stricli/core";
import { encode } from "@toon-format/toon";
import {
  OutlookAuth,
  TokenStore,
  GraphClient,
  MailClient,
} from "@outlook-toolkit/sdk";
import { resolveCliConfig } from "../context.js";

async function getMailClient(profile?: string): Promise<MailClient> {
  const config = await resolveCliConfig(profile);
  const store = new TokenStore(config.clientId);
  const auth = new OutlookAuth(config, store);
  let token: string;
  try {
    token = await auth.acquireToken();
  } catch {
    console.error("error: not authenticated (exit code 5). Run: outlook auth login");
    process.exit(5);
  }
  return new MailClient(new GraphClient(token));
}

const listCommand = buildCommand({
  docs: { brief: "List messages in a mail folder" },
  parameters: {
    flags: {
      profile: { kind: "parseable", brief: "Profile name", parse: String, optional: true },
      folder: { kind: "parseable", brief: "Folder (default: inbox)", parse: String, optional: true },
      limit: { kind: "parseable", brief: "Max messages (default: 25)", parse: Number, optional: true },
      cursor: { kind: "parseable", brief: "Pagination cursor ($skipToken URL)", parse: String, optional: true },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
      csv: { kind: "boolean", brief: "Output as CSV", default: false },
    },
  },
  async func(this: void, flags: { profile?: string; folder?: string; limit?: number; cursor?: string; json: boolean; csv: boolean }) {
    const mail = await getMailClient(flags.profile);
    const result = await mail.list({
      folder: flags.folder ?? "inbox",
      limit: flags.limit ?? 25,
      cursor: flags.cursor,
    });

    if (flags.json) {
      console.log(JSON.stringify(result, null, 2));
    } else if (flags.csv) {
      console.log("id,subject,from,receivedDateTime,isRead");
      for (const m of result.value) {
        const from = m.from?.emailAddress?.address ?? "";
        console.log(`${m.id},${JSON.stringify(m.subject ?? "")},${from},${m.receivedDateTime ?? ""},${m.isRead ?? ""}`);
      }
    } else {
      console.log(encode(result, { keyFolding: "safe" }));
    }

    if (result["@odata.nextLink"]) {
      process.stderr.write(`\nMore results available. Use --cursor=<nextLink> to continue.\n`);
    }
  },
});

const getCommand = buildCommand({
  docs: { brief: "Get a single message by ID" },
  parameters: {
    flags: {
      profile: { kind: "parseable", brief: "Profile name", parse: String, optional: true },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Message ID", parse: String }],
    },
  },
  async func(this: void, flags: { profile?: string; json: boolean }, id: string) {
    const mail = await getMailClient(flags.profile);
    const message = await mail.get(id);
    flags.json
      ? console.log(JSON.stringify(message, null, 2))
      : console.log(encode(message, { keyFolding: "safe" }));
  },
});

const sendCommand = buildCommand({
  docs: { brief: "Send an email" },
  parameters: {
    flags: {
      profile: { kind: "parseable", brief: "Profile name", parse: String, optional: true },
      to: { kind: "parseable", brief: "Recipient email address", parse: String },
      subject: { kind: "parseable", brief: "Email subject", parse: String },
      body: { kind: "parseable", brief: "Email body (HTML)", parse: String },
      dryRun: { kind: "boolean", brief: "Validate without sending", default: false },
      json: { kind: "boolean", brief: "Output result as JSON", default: false },
    },
  },
  async func(this: void, flags: { profile?: string; to: string; subject: string; body: string; dryRun: boolean; json: boolean }) {
    if (flags.dryRun) {
      const out = { status: "dry_run", to: flags.to, subject: flags.subject };
      flags.json ? console.log(JSON.stringify(out, null, 2)) : console.log(encode(out, { keyFolding: "safe" }));
      return;
    }
    const mail = await getMailClient(flags.profile);
    await mail.send({ to: flags.to, subject: flags.subject, body: flags.body });
    const out = { status: "sent", to: flags.to };
    flags.json ? console.log(JSON.stringify(out, null, 2)) : console.log("Sent.");
  },
});

const replyCommand = buildCommand({
  docs: { brief: "Reply to a message" },
  parameters: {
    flags: {
      profile: { kind: "parseable", brief: "Profile name", parse: String, optional: true },
      body: { kind: "parseable", brief: "Reply body (HTML)", parse: String },
      dryRun: { kind: "boolean", brief: "Validate without sending", default: false },
      json: { kind: "boolean", brief: "Output result as JSON", default: false },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Message ID to reply to", parse: String }],
    },
  },
  async func(this: void, flags: { profile?: string; body: string; dryRun: boolean; json: boolean }, id: string) {
    if (flags.dryRun) {
      const out = { status: "dry_run", replyTo: id };
      flags.json ? console.log(JSON.stringify(out, null, 2)) : console.log(encode(out, { keyFolding: "safe" }));
      return;
    }
    const mail = await getMailClient(flags.profile);
    await mail.reply(id, { body: flags.body });
    const out = { status: "replied", messageId: id };
    flags.json ? console.log(JSON.stringify(out, null, 2)) : console.log("Reply sent.");
  },
});

const draftCommand = buildCommand({
  docs: { brief: "Create a draft email" },
  parameters: {
    flags: {
      profile: { kind: "parseable", brief: "Profile name", parse: String, optional: true },
      to: { kind: "parseable", brief: "Recipient email address", parse: String },
      subject: { kind: "parseable", brief: "Email subject", parse: String },
      body: { kind: "parseable", brief: "Email body (HTML)", parse: String },
      dryRun: { kind: "boolean", brief: "Validate without creating", default: false },
      json: { kind: "boolean", brief: "Output draft as JSON", default: false },
    },
  },
  async func(this: void, flags: { profile?: string; to: string; subject: string; body: string; dryRun: boolean; json: boolean }) {
    if (flags.dryRun) {
      const out = { status: "dry_run", to: flags.to, subject: flags.subject };
      flags.json ? console.log(JSON.stringify(out, null, 2)) : console.log(encode(out, { keyFolding: "safe" }));
      return;
    }
    const mail = await getMailClient(flags.profile);
    const draft = await mail.createDraft({ to: flags.to, subject: flags.subject, body: flags.body });
    flags.json
      ? console.log(JSON.stringify(draft, null, 2))
      : console.log(encode(draft, { keyFolding: "safe" }));
  },
});

const syncCommand = buildCommand({
  docs: { brief: "Delta sync inbox (only changed messages since last sync)" },
  parameters: {
    flags: {
      profile: { kind: "parseable", brief: "Profile name", parse: String, optional: true },
      deltaLink: { kind: "parseable", brief: "Delta link from previous sync", parse: String, optional: true },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
  },
  async func(this: void, flags: { profile?: string; deltaLink?: string; json: boolean }) {
    const mail = await getMailClient(flags.profile);
    const result = await mail.sync(flags.deltaLink);
    flags.json
      ? console.log(JSON.stringify(result, null, 2))
      : console.log(encode(result, { keyFolding: "safe" }));
    if (result["@odata.deltaLink"]) {
      process.stderr.write(`\ndeltaLink: ${result["@odata.deltaLink"]}\n`);
    }
  },
});

export const mailRoutes = buildRouteMap({
  routes: { list: listCommand, get: getCommand, send: sendCommand, reply: replyCommand, draft: draftCommand, sync: syncCommand },
  docs: { brief: "Read and send Outlook mail" },
});
```

- [ ] **Step 2: Create commands/agent-context.ts**

Create `packages/cli/src/commands/agent-context.ts`:
```typescript
import { buildCommand } from "@stricli/core";

const AGENT_CONTEXT = {
  schema_version: "1",
  cli: "outlook",
  description: "Outlook toolkit — read and write Outlook mail via Microsoft Graph",
  config: {
    env_vars: {
      OUTLOOK_CLIENT_ID: { required: true, description: "Azure app registration client ID" },
      OUTLOOK_TENANT_ID: { required: true, description: "Tenant ID: consumers | common | <guid>" },
      OUTLOOK_PROFILE: { required: false, description: "Named profile to activate" },
    },
    precedence: ["explicit-flag", "env-var", "named-profile", "error"],
  },
  commands: {
    auth: {
      login: { brief: "Interactive browser PKCE sign-in", flags: { "--profile": "string (optional)" } },
      logout: { brief: "Clear saved tokens", flags: { "--profile": "string (optional)" } },
      status: { brief: "Show auth status", flags: { "--profile": "string (optional)", "--json": "boolean" } },
    },
    profile: {
      save: { brief: "Save a named profile", args: ["name"], flags: { "--client-id": "string", "--tenant-id": "string" } },
      list: { brief: "List profiles", flags: { "--json": "boolean" } },
      delete: { brief: "Delete a profile", args: ["name"], flags: { "--force": "boolean" } },
    },
    mail: {
      list: { brief: "List messages", flags: { "--folder": "string (default: inbox)", "--limit": "number (default: 25)", "--cursor": "string", "--json": "boolean", "--csv": "boolean", "--profile": "string (optional)" } },
      get: { brief: "Get message by ID", args: ["id"], flags: { "--json": "boolean", "--profile": "string (optional)" } },
      send: { brief: "Send an email", flags: { "--to": "string (email)", "--subject": "string", "--body": "string (HTML)", "--dry-run": "boolean", "--json": "boolean", "--profile": "string (optional)" } },
      reply: { brief: "Reply to a message", args: ["id"], flags: { "--body": "string (HTML)", "--dry-run": "boolean", "--json": "boolean", "--profile": "string (optional)" } },
      draft: { brief: "Create a draft email", flags: { "--to": "string (email)", "--subject": "string", "--body": "string (HTML)", "--dry-run": "boolean", "--json": "boolean", "--profile": "string (optional)" } },
      sync: { brief: "Delta sync inbox", flags: { "--delta-link": "string (optional)", "--json": "boolean", "--profile": "string (optional)" } },
    },
  },
  exit_codes: {
    0: "success",
    1: "network error",
    2: "validation error",
    3: "config error (missing OUTLOOK_CLIENT_ID or OUTLOOK_TENANT_ID)",
    4: "not found",
    5: "auth error (run: outlook auth login)",
    6: "rate limited",
  },
};

export const agentContextCommand = buildCommand({
  docs: { brief: "Output machine-readable CLI schema for agent use" },
  parameters: {
    flags: {
      json: {
        kind: "boolean",
        brief: "Output as JSON (default: true)",
        default: true,
      },
    },
  },
  async func(this: void, _flags: { json: boolean }) {
    console.log(JSON.stringify(AGENT_CONTEXT, null, 2));
  },
});
```

- [ ] **Step 3: Run lint**

```bash
bun run --filter '@outlook-toolkit/cli' lint
```
Expected: no TypeScript errors.

- [ ] **Step 4: Commit**

```bash
git add packages/cli/src/commands/mail.ts packages/cli/src/commands/agent-context.ts
git commit -m "feat(cli): add mail list/get/send/reply/draft/sync commands and agent-context"
```

---

## Task 12: CLI — wire app.ts + delete scaffold commands

**Files:**
- Modify: `packages/cli/src/app.ts`
- Delete: `packages/cli/src/commands/create.ts`
- Delete: `packages/cli/src/commands/delete.ts`
- Delete: `packages/cli/src/commands/get.ts`
- Delete: `packages/cli/src/commands/list.ts`

- [ ] **Step 1: Rewrite app.ts**

Replace `packages/cli/src/app.ts` entirely:
```typescript
import { buildApplication, buildRouteMap } from "@stricli/core";
import { authRoutes } from "./commands/auth.js";
import { profileRoutes } from "./commands/profile.js";
import { mailRoutes } from "./commands/mail.js";
import { agentContextCommand } from "./commands/agent-context.js";

const routes = buildRouteMap({
  routes: {
    auth: authRoutes,
    profile: profileRoutes,
    mail: mailRoutes,
    "agent-context": agentContextCommand,
  },
  docs: {
    brief: "SDK, CLI, and MCP server for Microsoft Outlook via Microsoft Graph",
  },
});

export const app = buildApplication(routes, {
  name: "outlook",
  versionInfo: {
    currentVersion: "0.1.0",
  },
});
```

- [ ] **Step 2: Delete scaffold commands**

```bash
rm packages/cli/src/commands/create.ts
rm packages/cli/src/commands/delete.ts
rm packages/cli/src/commands/get.ts
rm packages/cli/src/commands/list.ts
```

- [ ] **Step 3: Run lint**

```bash
bun run --filter '@outlook-toolkit/cli' lint
```
Expected: no TypeScript errors.

- [ ] **Step 4: Smoke-test CLI help**

```bash
bun run dev:cli -- --help
```
Expected: shows `auth`, `profile`, `mail`, `agent-context` commands.

```bash
bun run dev:cli -- mail list --help
```
Expected: shows `--folder`, `--limit`, `--cursor`, `--json`, `--csv`, `--profile` flags.

- [ ] **Step 5: Commit**

```bash
git add packages/cli/src/app.ts
git rm packages/cli/src/commands/create.ts packages/cli/src/commands/delete.ts packages/cli/src/commands/get.ts packages/cli/src/commands/list.ts
git commit -m "feat(cli): wire auth/profile/mail routes, remove scaffold commands"
```

---

## Task 13: MCP — auth + mail tools + server wiring

**Files:**
- Create: `packages/mcp/src/server.ts`
- Create: `packages/mcp/src/tools/auth.ts`
- Create: `packages/mcp/src/tools/mail.ts`
- Modify: `packages/mcp/src/index.ts`
- Delete: `packages/mcp/src/tools/resources.ts`

- [ ] **Step 1: Create server.ts**

Create `packages/mcp/src/server.ts`:
```typescript
import { FastMCP } from "fastmcp";
import { registerAuthTools } from "./tools/auth.js";
import { registerMailTools } from "./tools/mail.js";

const server = new FastMCP({
  name: "outlook-toolkit",
  version: "0.1.0",
});

registerAuthTools(server);
registerMailTools(server);

export default server;
```

- [ ] **Step 2: Create tools/auth.ts**

Create `packages/mcp/src/tools/auth.ts`:
```typescript
import type { FastMCP } from "fastmcp";
import { z } from "zod";
import { resolveConfig, OutlookAuth, TokenStore } from "@outlook-toolkit/sdk";

function getAuth() {
  const config = resolveConfig();
  const store = new TokenStore(config.clientId);
  return { auth: new OutlookAuth(config, store), config };
}

export function registerAuthTools(server: FastMCP) {
  server.addTool({
    name: "outlook_auth_status",
    description: "Check Outlook authentication status for the configured account.",
    parameters: z.object({}),
    execute: async () => {
      const { auth, config } = getAuth();
      const status = await auth.status();
      if (!status) {
        return JSON.stringify({
          authenticated: false,
          message: `Not authenticated. Run \`outlook auth login\` in your terminal with OUTLOOK_CLIENT_ID=${config.clientId} set.`,
        }, null, 2);
      }
      return JSON.stringify({
        authenticated: true,
        userEmail: status.userEmail,
        accessTokenExpiry: status.accessTokenExpiry
          ? new Date(status.accessTokenExpiry).toISOString()
          : null,
      }, null, 2);
    },
  });

  server.addTool({
    name: "outlook_auth_logout",
    description: "Sign out and clear saved tokens for the configured Outlook account.",
    parameters: z.object({}),
    execute: async () => {
      const { auth } = getAuth();
      await auth.logout();
      return "Signed out. Run `outlook auth login` to re-authenticate.";
    },
  });
}
```

- [ ] **Step 3: Create tools/mail.ts**

Create `packages/mcp/src/tools/mail.ts`:
```typescript
import type { FastMCP } from "fastmcp";
import { z } from "zod";
import {
  resolveConfig,
  OutlookAuth,
  TokenStore,
  GraphClient,
  MailClient,
  OutlookAuthError,
} from "@outlook-toolkit/sdk";

async function getMailClient(): Promise<MailClient> {
  const config = resolveConfig();
  const store = new TokenStore(config.clientId);
  const auth = new OutlookAuth(config, store);
  try {
    const token = await auth.acquireToken();
    return new MailClient(new GraphClient(token));
  } catch (err) {
    if (err instanceof OutlookAuthError) {
      throw new Error(
        `Not authenticated. Run \`outlook auth login\` in your terminal with OUTLOOK_CLIENT_ID=${config.clientId} set.`
      );
    }
    throw err;
  }
}

export function registerMailTools(server: FastMCP) {
  server.addTool({
    name: "outlook_mail_list",
    description: "List messages in an Outlook mail folder (default: inbox). Returns messages and a nextLink cursor for pagination.",
    parameters: z.object({
      folder: z.string().default("inbox").describe("Folder name (inbox, sentitems, drafts, deleteditems)"),
      limit: z.number().int().positive().max(999).default(25).describe("Max messages to return"),
      cursor: z.string().optional().describe("Pagination cursor from a previous call's nextLink"),
      filter: z.string().optional().describe("OData $filter expression"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      const result = await mail.list({ folder: args.folder, limit: args.limit, cursor: args.cursor, filter: args.filter });
      return JSON.stringify(result, null, 2);
    },
  });

  server.addTool({
    name: "outlook_mail_get",
    description: "Get a single Outlook message by its ID, including the full body.",
    parameters: z.object({
      id: z.string().describe("Message ID"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      const message = await mail.get(args.id);
      return JSON.stringify(message, null, 2);
    },
  });

  server.addTool({
    name: "outlook_mail_send",
    description: "Send an email from the authenticated Outlook account.",
    parameters: z.object({
      to: z.string().email().describe("Recipient email address"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body (HTML supported)"),
      contentType: z.enum(["HTML", "text"]).default("HTML").describe("Body content type"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      await mail.send({ to: args.to, subject: args.subject, body: args.body, contentType: args.contentType });
      return JSON.stringify({ status: "sent", to: args.to }, null, 2);
    },
  });

  server.addTool({
    name: "outlook_mail_reply",
    description: "Reply to an existing Outlook message thread.",
    parameters: z.object({
      id: z.string().describe("Message ID to reply to"),
      body: z.string().describe("Reply body (HTML supported)"),
      contentType: z.enum(["HTML", "text"]).default("HTML").describe("Body content type"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      await mail.reply(args.id, { body: args.body, contentType: args.contentType });
      return JSON.stringify({ status: "replied", messageId: args.id }, null, 2);
    },
  });

  server.addTool({
    name: "outlook_mail_create_draft",
    description: "Create a draft email without sending it. Returns the draft message including its ID.",
    parameters: z.object({
      to: z.string().email().describe("Recipient email address"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body (HTML supported)"),
      contentType: z.enum(["HTML", "text"]).default("HTML").describe("Body content type"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      const draft = await mail.createDraft({ to: args.to, subject: args.subject, body: args.body, contentType: args.contentType });
      return JSON.stringify(draft, null, 2);
    },
  });

  server.addTool({
    name: "outlook_mail_sync",
    description: "Delta sync inbox — returns only messages that changed since the last sync. On first call, omit deltaLink to get the full initial sync. Save the returned deltaLink and pass it on subsequent calls to get only changes.",
    parameters: z.object({
      deltaLink: z.string().optional().describe("Delta link from a previous sync call. Omit for initial full sync."),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      const result = await mail.sync(args.deltaLink);
      return JSON.stringify(result, null, 2);
    },
  });
}
```

- [ ] **Step 4: Rewrite mcp/src/index.ts**

Replace `packages/mcp/src/index.ts` entirely:
```typescript
import server from "./server.js";
await server.start({ transportType: "stdio" });
```

- [ ] **Step 5: Delete scaffold resources.ts**

```bash
rm packages/mcp/src/tools/resources.ts
```

- [ ] **Step 6: Run lint**

```bash
bun run --filter '@outlook-toolkit/mcp' lint
```
Expected: no TypeScript errors.

- [ ] **Step 7: Smoke-test MCP inspector**

```bash
bun run dev:mcp &
sleep 2
npx fastmcp inspect packages/mcp/src/index.ts 2>&1 | head -30 || true
```
Expected: lists `outlook_auth_status`, `outlook_auth_logout`, and the six `outlook_mail_*` tools.

- [ ] **Step 8: Commit**

```bash
git add packages/mcp/src/server.ts packages/mcp/src/tools/auth.ts packages/mcp/src/tools/mail.ts packages/mcp/src/index.ts
git rm packages/mcp/src/tools/resources.ts
git commit -m "feat(mcp): add auth and mail tools, wire FastMCP server"
```

---

## Self-Review Checklist

### Spec coverage

| Spec requirement | Covered by |
|---|---|
| PKCE auth, no MSAL | Task 5 `auth.ts` |
| keytar + AES-256-CBC file fallback | Task 4 `token-store.ts` |
| Token keyed by clientId | Task 4 `TokenStore` constructor |
| resolveConfig + resolveAuthority | Task 3 `config.ts` |
| Env var + profile precedence | Task 3 + Task 10 `context.ts` |
| Named profiles (CLI) | Task 6 `ProfileStore` + Task 10 CLI |
| MCP configured via env vars | Task 13 `getMailClient()` + `getAuth()` |
| GraphClient retry/backoff | Task 7 |
| list, get, send, reply, createDraft, sync | Task 8 + Task 11 + Task 13 |
| `--toon`/`--json`/`--csv` on list | Task 11 mail list command |
| `--dry-run` on mutations | Task 11 send/reply/draft |
| `--profile` flag on all commands | Task 10 + Task 11 |
| agent-context introspection | Task 11 |
| OutlookConfigError (exit 3) | Task 2 + Task 10 |
| OutlookAuthError (exit 5) | Task 2 + Task 11 |
| OutlookRateLimitError (exit 6) | Task 2 + Task 7 |
| Stampede protection on refresh | Task 5 `pendingRefresh` |
| 0600 file permissions on token cache | Task 4 `save()` + `getOrCreateKey()` |
| Token values never logged | Review Task 10/11/13 (no `console.log(token)`) |

### No gaps found. All spec requirements have corresponding tasks.
