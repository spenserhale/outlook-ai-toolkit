# Outlook Toolkit — Token-Efficient Output Design Spec

**Date:** 2026-06-07
**Scope:** Reduce LLM token usage when consuming the Outlook toolkit — shared TOON/JSON output rendering across CLI and MCP, and configurable email-body conversion (text / markdown / html).

---

## Goal

Cut the tokens an LLM spends reading toolkit output, without losing fidelity when it's needed.

Two levers:

1. **Output format.** TOON is materially cheaper than JSON for the record/list shapes this toolkit returns. It is already the CLI default; the MCP server still emits `JSON.stringify`. Make TOON the default everywhere with an easy switch to JSON for compatibility, using a single shared renderer.
2. **Email bodies.** Graph returns bodies as raw HTML (`{ contentType: "HTML", content: "<…huge…>" }`) — the single largest source of token bloat. Convert bodies to clean plain **text** by default, offer **markdown** when structure matters, and keep raw **html** as an opt-in escape hatch.

---

## Architecture (Approach A — shared SDK utilities)

All new logic lives in the SDK; CLI and MCP stay thin consumers that parse input and call shared functions. This matches the toolkit's "all API logic in the SDK, write-once" rule and guarantees identical behavior across both consumers.

```
packages/sdk/src/body.ts      renderBody()    — html→text | html→markdown | raw html
packages/sdk/src/format.ts    renderOutput()  — data → toon | json
packages/sdk/src/types.ts     BodyFormat, ListBodyMode, OutputFormat + param schemas
packages/sdk/src/mail-client.ts                — apply body conversion + $select shaping
        ^                                ^
        |                                |
packages/cli/src/commands/mail.ts   packages/mcp/src/tools/mail.ts
(flags → renderOutput/renderBody)   (params → renderOutput/renderBody)
```

Rejected alternatives:
- **B (convert in SDK, format per-consumer):** duplicates TOON/JSON rendering in CLI and MCP → drift risk.
- **C (everything in consumers):** duplicates logic, violates the architecture.

---

## Enums

```ts
BodyFormat   = "text" | "markdown" | "html"   // how a body's content is rendered
ListBodyMode = "none" | "preview" | "full"    // how much body a list row carries
OutputFormat = "toon" | "json"                // envelope encoding
```

Defaults: `BodyFormat = "text"`, `ListBodyMode = "preview"`, `OutputFormat = "toon"`.

All three are Zod enums so invalid input produces an enumerated error naming the valid set (agent-native principle 3).

---

## SDK — `body.ts`

```ts
renderBody(content: string, sourceContentType: "text" | "HTML", target: BodyFormat): string
```

Behavior:

| target | source = HTML | source = text |
|---|---|---|
| `html` | return content unchanged | return content unchanged |
| `text` | `html-to-text` → clean text | return content unchanged |
| `markdown` | `turndown` → markdown | return content unchanged |

- If the source is already plain `text`, return it as-is regardless of target (can't manufacture HTML/markdown structure from plain text).
- **Best-effort:** conversion is wrapped so a library failure falls back to the raw content rather than throwing. Body rendering must never crash a mail read.
- `html-to-text` configured to drop noise (skip images, sensible link formatting, no hard wrap that fights the consumer). `turndown` uses default options.

New SDK deps: `turndown`, `html-to-text`, `@toon-format/toon`, plus `@types/turndown` (html-to-text ships its own types).

---

## SDK — `format.ts`

```ts
renderOutput(data: unknown, format: OutputFormat): string
```

- `json` → `JSON.stringify(data, null, 2)`
- `toon` → `encode(data, { keyFolding: "safe" })`

This is the single source of truth for envelope encoding. CSV stays CLI-only (it is tabular and list-specific) and is not part of this function.

---

## SDK — MailClient changes

### `get(id, opts?: { bodyFormat?: BodyFormat })`
- Default `bodyFormat = "text"`.
- Fetch the full message, then replace `body` with `{ contentType: bodyFormat, content: renderBody(body.content, body.contentType, bodyFormat) }`.
- If the message has no body, leave it absent.

### `list(params)` — add `body` and `bodyFormat`
- Default `body = "preview"`, `bodyFormat = "text"`.
- `$select` shaping (only when the caller did **not** pass an explicit `select`):
  - `full` → select includes `body` (+ the standard summary fields); convert each message's body via `renderBody`.
  - `preview` → select includes `bodyPreview`, excludes `body`; strip any `body` that slips through.
  - `none` → select excludes both; strip `body` and `bodyPreview` from each row.
- An explicit user-provided `select` is respected verbatim (advanced escape hatch); body post-processing still applies to whatever `body` field is present.
- Standard summary fields for the default select: `id, subject, from, toRecipients, receivedDateTime, isRead, isDraft, conversationId` (+ `bodyPreview` for preview, + `body` for full).

`sync` is out of scope for body conversion in this iteration (delta payloads are change-tracking, not reading-oriented); it keeps current behavior but routes through `renderOutput` for format consistency.

---

## CLI — `commands/mail.ts`

- Replace every inline `flags.json ? JSON.stringify(...) : encode(...)` with `renderOutput(data, format)`.
- Add an explicit `--toon` flag (default true) to complete the canonical `--toon` / `--json` / `--csv` triplet. Precedence when multiple are passed: `--csv` > `--json` > `--toon`.
- `list`:
  - `--body <none|preview|full>` (default `preview`)
  - `--body-format <text|markdown|html>` (default `text`; applies when `--body=full`)
  - CSV columns unchanged: `id,subject,from,receivedDateTime,isRead`.
- `get`:
  - `--body-format <text|markdown|html>` (default `text`)
- Invalid enum value → stderr error enumerating valid values, exit code 2 (validation).

---

## MCP — `tools/mail.ts`

- Every data-returning tool (`outlook_mail_list`, `outlook_mail_get`, `outlook_mail_sync`) gets an optional `format: "toon" | "json"` param, default `"toon"`, and renders via `renderOutput` instead of `JSON.stringify`. Mutation acks (`send`, `reply`) also route through `renderOutput` for consistency.
- `outlook_mail_get`: add `bodyFormat: text|markdown|html` (default `text`).
- `outlook_mail_list`: add `body: none|preview|full` (default `preview`) and `bodyFormat: text|markdown|html` (default `text`).
- Tool descriptions updated to mention the new params and their defaults, kept terse (token budget).

---

## Error Handling

- Enum validation errors (CLI + MCP) name the valid set and echo the offending value.
- Body conversion failures fall back to raw content silently (logged to stderr in CLI dev only) — reading mail must not fail because one message has malformed HTML.
- No change to existing auth / network / Graph error taxonomy.

---

## Testing

New / updated SDK tests (TDD, red → green):

- `body.test.ts`
  - HTML → text strips tags, preserves readable content
  - HTML → markdown preserves links, bold, lists, headings
  - `html` target returns content unchanged
  - already-plain `text` source passes through for every target
  - malformed HTML falls back to raw content without throwing
- `format.test.ts`
  - `toon` output matches `encode(..., { keyFolding: "safe" })`
  - `json` output matches `JSON.stringify(..., 2)`
- `mail-client.test.ts` (extend, mocked GraphClient)
  - `get` converts body per `bodyFormat`
  - `list` `none` strips body + bodyPreview
  - `list` `preview` keeps bodyPreview, drops body
  - `list` `full` converts each body
  - explicit `select` is passed through untouched

Manual verification: run `dev:cli mail get <id>` across all three body formats and `dev:cli mail list --body=full` to confirm token reduction vs `--json` with raw HTML.

---

## Out of Scope

- Body conversion for `sync` (delta) payloads.
- CSV support for nested/body output.
- Changing the auth, profile, or Graph-client layers.
- Attachments handling.
