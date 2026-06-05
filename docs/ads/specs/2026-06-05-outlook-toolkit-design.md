# Outlook Toolkit — Design Spec

**Date:** 2026-06-05
**Scope:** Initial implementation — SDK (auth + mail), CLI (Stricli), MCP (FastMCP)

---

## Goal

A Bun monorepo toolkit for programmatic Outlook/Microsoft 365 access via Microsoft Graph. Follows the shared SDK → CLI → MCP three-layer architecture. Supports both personal (MSA/outlook.com) and work (Entra ID) accounts from a single codebase, one account active at a time.

---

## Auth Model

**Delegated OAuth 2.0 with PKCE** (`authorization_code` + `offline_access`) for both account types. No client secret; public client only. No MSAL — hand-rolled PKCE using plain `fetch` and `node:crypto`.

**Why not MSAL:** The best reference implementations skip it (see `refs/notes.md`). It adds a heavy dependency, uses an opaque cache format, and has unverified Bun compatibility. Hand-rolled PKCE is ~150 lines and gives full control over token storage and refresh behavior.

**Account type resolution:** derived from `tenantId`:
- `common` or `consumers` → personal/MSA (authority: `https://login.microsoftonline.com/consumers`)
- A GUID → work/Entra ID (authority: `https://login.microsoftonline.com/{tenantId}`)

**Interactive sign-in (once per account):**
1. Generate PKCE `code_verifier` + `code_challenge` (SHA-256, base64url)
2. Spin up `http://localhost:{port}/callback` HTTP server
3. Open browser to Microsoft authorization URL
4. Exchange returned code for access + refresh tokens
5. Store encrypted tokens; shut down callback server

**Silent token acquisition (every run after first):**
- Check token expiry; if within 55-minute window, refresh proactively
- `acquireToken()` returns a valid access token or throws `OutlookAuthError`
- In-flight refresh is a shared Promise (stampede protection)
- If refresh token is expired → throw `OutlookAuthError` prompting re-login

**Scopes:** `Mail.Read Mail.ReadWrite Mail.Send offline_access openid profile User.Read`

---

## Config Resolution

**Env vars (minimal required):**
```
OUTLOOK_CLIENT_ID=<app registration client id>
OUTLOOK_TENANT_ID=<consumers|common|tenant-guid>
```

**Precedence (strict):**
```
explicit function arg > env var > named profile > error
```

`resolveConfig(overrides?)` reads env vars, then falls back to the active named profile if `OUTLOOK_PROFILE` is set or a `--profile` flag is passed.

**Named profiles** (CLI only, stored at `~/.outlook-toolkit/profiles.json`):
```bash
outlook profile save personal --client-id=abc --tenant-id=consumers
outlook profile save work --client-id=xyz --tenant-id=<guid>
outlook mail list --profile=personal
```

**MCP instances** are configured entirely via env vars at boot — no profile flag needed. The same binary serves as both `outlook-personal` and `outlook-work` when launched with different env vars:
```json
"outlook-personal": { "env": { "OUTLOOK_CLIENT_ID": "abc", "OUTLOOK_TENANT_ID": "consumers" } }
"outlook-work":     { "env": { "OUTLOOK_CLIENT_ID": "xyz", "OUTLOOK_TENANT_ID": "<guid>" } }
```

---

## Token Storage

**Storage priority:**
1. OS keychain via `keytar` (macOS Keychain, Windows Credential Store, Linux Secret Service)
2. AES-256-CBC encrypted file fallback at `~/.outlook-toolkit/tokens-{clientId}.enc`

**Encryption key:** generated with `crypto.randomBytes(32)`, stored in keychain under service `outlook-toolkit`, key `encryption-key`. File fallback stores key at `~/.outlook-toolkit/encryption.key` (mode `0600`).

**Cache keyed by `clientId`** — two accounts (personal + work) with different client IDs never share a cache, whether configured via separate MCP instances or CLI `--profile`.

**Stored per account:**
- `accessToken` (encrypted)
- `refreshToken` (encrypted)
- `accessTokenExpiry` (timestamp)
- `refreshTokenExpiry` (timestamp, ~90 days)
- `userEmail` (plain, for `auth status` display)

---

## Graph Client

Plain `fetch` against `https://graph.microsoft.com/v1.0`. No `@microsoft/microsoft-graph-client`.

**Retry behavior** (from `refs/outlook-mcp/`):
- Retry on HTTP 429, 503, 504
- Exponential backoff with ±25% jitter
- Respects `Retry-After` header when present
- Max 2 retries, initial delay 1000ms, max delay 30 000ms

**Request helpers:** `get<T>()`, `list<T>()`, `post<Req, Res>()`, `patch()`, `delete()` — each accepts an OData options object (`$select`, `$filter`, `$top`, `$orderby`, `$search`).

---

## Mail Operations

All operations use the `/me` endpoint (delegated flow). Implemented in `MailClient` which wraps `GraphClient`.

| Operation | Graph endpoint | Notes |
|---|---|---|
| `list(opts)` | `GET /me/mailFolders/{folder}/messages` | Default folder: `inbox`. Supports `$top`, `$filter`, `$select`, `$orderby`. Returns messages + `@odata.nextLink` for pagination. |
| `get(id)` | `GET /me/messages/{id}` | Full message including body. |
| `send(params)` | `POST /me/sendMail` | To, subject, HTML body, `saveToSentItems: true`. |
| `reply(id, params)` | `POST /me/messages/{id}/reply` | Reply to existing message thread. |
| `createDraft(params)` | `POST /me/messages` | Creates draft without sending. |
| `sync(deltaLink?)` | `GET /me/mailFolders/inbox/messages/delta` | First call: no deltaLink → returns full sync + new deltaLink. Subsequent calls: pass deltaLink → returns only changed items. |

---

## Package Structure

```
packages/sdk/src/
  config.ts          resolveConfig(), OutlookConfig type, authority URL resolver
  auth.ts            PKCE flow (getAuthorizationCode, exchangeCode, refreshToken), acquireToken()
  token-store.ts     keytar + AES-256-CBC file fallback, keyed by clientId
  graph-client.ts    fetch wrapper with OData helpers + retry/backoff
  mail-client.ts     list, get, send, reply, createDraft, sync
  profile-store.ts   read/write ~/.outlook-toolkit/profiles.json
  types.ts           Zod schemas: OutlookConfig, Message, MailFolder, DeltaResponse, SendMailParams, ReplyParams, DraftParams
  errors.ts          OutlookError, OutlookAuthError, OutlookNotFoundError, OutlookRateLimitError
  index.ts           public re-exports

packages/cli/src/
  app.ts             Stricli route map
  bin.ts             entry point
  commands/
    auth.ts          login, logout, status
    profile.ts       save, list, delete
    mail.ts          list, get, send, reply, draft, sync
    agent-context.ts machine-readable CLI schema

packages/mcp/src/
  index.ts           FastMCP server bootstrap
  tools/
    auth.ts          outlook_auth_login, outlook_auth_logout, outlook_auth_status
    mail.ts          outlook_mail_list, outlook_mail_get, outlook_mail_send,
                     outlook_mail_reply, outlook_mail_create_draft, outlook_mail_sync
```

---

## CLI Command Surface

All commands follow agent-native conventions (Stricli):
- `--toon` (default) / `--json` / `--csv` on all data-returning commands
- `--dry-run` on all mutation commands
- `--profile <name>` as a root flag (overridden by env vars)
- Enumerated errors with valid values on bad input
- Paginated list with `--limit` and `--cursor`; truncation messages name the narrowing flags

```bash
# Auth
outlook auth login   [--profile=<name>]
outlook auth logout  [--profile=<name>]
outlook auth status  [--profile=<name>] [--json]

# Profiles
outlook profile save <name> --client-id=<id> --tenant-id=<id>
outlook profile list [--json]
outlook profile delete <name> [--force]

# Mail
outlook mail list   [--folder=inbox] [--limit=25] [--cursor=<token>] [--toon|--json|--csv]
outlook mail get    <id>             [--toon|--json]
outlook mail send   --to=<addr> --subject=<text> --body=<html>  [--dry-run] [--json]
outlook mail reply  <id> --body=<html>                          [--dry-run] [--json]
outlook mail draft  --to=<addr> --subject=<text> --body=<html>  [--dry-run] [--json]
outlook mail sync   [--delta-link=<token>] [--toon|--json]

# Introspection
outlook agent-context  [--json]
```

---

## MCP Tool Surface

Configured via env vars at boot. Each tool maps 1:1 to an SDK operation.

```
outlook_auth_login         triggers interactive PKCE (returns browser URL if headless)
outlook_auth_logout        clears token cache for the configured clientId
outlook_auth_status        returns email + expiry for active account

outlook_mail_list          list inbox/folder messages with optional filter/limit/cursor
outlook_mail_get           get single message by id
outlook_mail_send          send email (to, subject, body)
outlook_mail_reply         reply to message thread
outlook_mail_create_draft  create draft without sending
outlook_mail_sync          delta sync — pass deltaLink or omit for initial full sync
```

---

## Error Types & Exit Codes

| Error class | Exit code | Cause |
|---|---|---|
| network / fetch failure | 1 | Transient network error |
| validation | 2 | Bad flag value or missing required arg |
| config | 3 | Missing `OUTLOOK_CLIENT_ID` or `OUTLOOK_TENANT_ID` |
| not found | 4 | Message ID doesn't exist |
| auth | 5 | Not authenticated or token expired — run `outlook auth login` |
| rate limit | 6 | Graph 429 after retries exhausted |

---

## Security Constraints

- Public client only — no client secret in this toolkit. Confidential/app-only flow is out of scope.
- Token cache encrypted at rest (keytar or AES-256-CBC file).
- Cache files created with mode `0600`; toolkit validates permissions on load and corrects if wrong.
- Token values never logged — only metadata (expiry, email, status).
- `.env` and `~/.outlook-toolkit/` excluded from git.
- Request least-privilege scopes; `Mail.ReadWrite` is included to support draft creation and move operations, but callers can configure a read-only profile by registering an app with only `Mail.Read` scope (a `read_only` config flag is out of scope for this implementation).

---

## Build Order

1. **SDK foundation** — `config.ts`, `errors.ts`, `types.ts` (Zod schemas)
2. **Auth + token store** — `auth.ts`, `token-store.ts`; manual end-to-end test with personal account
3. **Graph client** — `graph-client.ts` with retry logic and unit tests
4. **Mail client** — `mail-client.ts` (list → get → send → reply → draft → sync)
5. **CLI** — Stricli wiring for auth, profile, mail commands; `agent-context`
6. **MCP** — FastMCP registration of auth + mail tools
7. **Work account** — add second app registration, validate with `OUTLOOK_TENANT_ID=<guid>`

---

## Open Questions (to resolve during implementation)

- **Bun + loopback server:** Verify `node:http` loopback callback server works cleanly under Bun. If not, the auth package may need a thin Node shim.
- **Corporate tenant:** Delegated mail scopes may require admin consent. Validate before investing in the work-account path.
- **Conditional Access:** Silent refresh from a laptop vs. a remote host may produce different CA outcomes. Validate corporate account on intended host early.
- **`keytar` on Bun:** `keytar` is a native Node addon. Test Bun compatibility; if broken, fall back to AES-file-only and document the limitation.
