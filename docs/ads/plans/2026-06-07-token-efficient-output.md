# Token-Efficient Output Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use ads:subagent-driven-development (recommended) or ads:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reduce LLM token usage from the Outlook toolkit by making TOON the shared default output format (with easy JSON switch) and converting email bodies to clean text (default) / markdown / raw html.

**Architecture:** All new logic lives in the SDK as two pure utilities — `renderBody()` (html→text|markdown|raw) and `renderOutput()` (toon|json) — plus body-aware `MailClient.get/list`. CLI and MCP stay thin: they parse flags/params and call the shared functions. This follows the toolkit's "all logic in the SDK, write-once" rule.

**Tech Stack:** Bun, TypeScript (ESM, `.js` import extensions), Zod, Stricli (CLI), FastMCP (MCP), `@toon-format/toon`, `turndown` (HTML→Markdown), `html-to-text` (HTML→text).

---

## File Structure

| File | Responsibility | Action |
|---|---|---|
| `packages/sdk/package.json` | add `turndown`, `html-to-text`, `@toon-format/toon`, `@types/turndown` | Modify |
| `packages/sdk/src/types.ts` | add `BodyFormat`, `ListBodyMode`, `OutputFormat` enums; widen `MessageBody.contentType`; extend `ListMailParams`; add `GetMailParams` | Modify |
| `packages/sdk/src/body.ts` | `renderBody()` — body conversion | Create |
| `packages/sdk/src/format.ts` | `renderOutput()` — envelope encoding | Create |
| `packages/sdk/src/mail-client.ts` | body-aware `get`/`list` with `$select` shaping | Modify |
| `packages/sdk/src/index.ts` | export new utils, types, schemas | Modify |
| `packages/sdk/tests/body.test.ts` | `renderBody` tests | Create |
| `packages/sdk/tests/format.test.ts` | `renderOutput` tests | Create |
| `packages/sdk/tests/mail-client.test.ts` | body-mode + conversion tests | Modify |
| `packages/cli/src/commands/mail.ts` | new body flags, shared `renderOutput`, `--toon` flag | Modify |
| `packages/mcp/src/tools/mail.ts` | per-call `format` param, body params, `renderOutput` | Modify |

---

## Task 1: Add SDK dependencies

**Files:**
- Modify: `packages/sdk/package.json`

- [ ] **Step 1: Add the dependencies block**

Edit `packages/sdk/package.json` so the `dependencies` object reads exactly:

```json
  "dependencies": {
    "zod": "^3.24.0",
    "keytar": "^7.9.0",
    "turndown": "^7.2.0",
    "html-to-text": "^9.0.5",
    "@toon-format/toon": "^2.1.0"
  },
```

And the `devDependencies` object reads exactly:

```json
  "devDependencies": {
    "typescript": "^5.7.0",
    "@types/bun": "latest",
    "@types/turndown": "^5.0.5"
  }
```

(`html-to-text` ships its own type definitions, so it needs no `@types` package.)

- [ ] **Step 2: Install from the repo root**

Run: `bun install`
Expected: completes without error; `node_modules/turndown`, `node_modules/html-to-text`, and `node_modules/@toon-format/toon` exist.

- [ ] **Step 3: Commit**

```bash
git add packages/sdk/package.json bun.lock
git commit -m "build(sdk): add turndown, html-to-text, toon deps"
```

---

## Task 2: Add output-format and body enums to types

**Files:**
- Modify: `packages/sdk/src/types.ts`

- [ ] **Step 1: Add the three enums**

In `packages/sdk/src/types.ts`, immediately after the `import { z } from "zod";` line, add:

```ts
// Output + body format enums
export const BodyFormatSchema = z.enum(["text", "markdown", "html"]);
export type BodyFormat = z.infer<typeof BodyFormatSchema>;

export const ListBodyModeSchema = z.enum(["none", "preview", "full"]);
export type ListBodyMode = z.infer<typeof ListBodyModeSchema>;

export const OutputFormatSchema = z.enum(["toon", "json"]);
export type OutputFormat = z.infer<typeof OutputFormatSchema>;
```

- [ ] **Step 2: Widen `MessageBody.contentType`**

Graph returns `text`/`html`, and after rendering we set it to `text`/`markdown`/`html`. Replace the `MessageBodySchema` definition:

```ts
export const MessageBodySchema = z.object({
  contentType: z.string(),
  content: z.string(),
});
export type MessageBody = z.infer<typeof MessageBodySchema>;
```

- [ ] **Step 3: Extend `ListMailParams` and add `GetMailParams`**

Replace the `ListMailParamsSchema` definition with:

```ts
export const ListMailParamsSchema = z.object({
  folder: z.string().default("inbox"),
  limit: z.number().int().positive().max(999).default(25),
  cursor: z.string().optional(),
  filter: z.string().optional(),
  select: z.string().optional(),
  orderby: z.string().optional(),
  body: ListBodyModeSchema.default("preview"),
  bodyFormat: BodyFormatSchema.default("text"),
});
export type ListMailParams = z.infer<typeof ListMailParamsSchema>;

export const GetMailParamsSchema = z.object({
  bodyFormat: BodyFormatSchema.default("text"),
});
export type GetMailParams = z.infer<typeof GetMailParamsSchema>;
```

- [ ] **Step 4: Type-check**

Run: `bun run --filter '@outlook-toolkit/sdk' lint`
Expected: PASS (no `tsc` errors).

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/src/types.ts
git commit -m "feat(sdk): add body/output format enums and params"
```

---

## Task 3: Implement `renderBody`

**Files:**
- Create: `packages/sdk/src/body.ts`
- Test: `packages/sdk/tests/body.test.ts`

- [ ] **Step 1: Write the failing tests**

Create `packages/sdk/tests/body.test.ts`:

```ts
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
    expect(out).toMatch(/[-*] one/);
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
    // null content would throw inside a converter; renderBody must not propagate
    const weird = "<<<not really html>>>";
    expect(() => renderBody(weird, "html", "markdown")).not.toThrow();
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `bun test packages/sdk/tests/body.test.ts`
Expected: FAIL with "Cannot find module '../src/body.js'".

- [ ] **Step 3: Implement `renderBody`**

Create `packages/sdk/src/body.ts`:

```ts
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
      ],
    });
  } catch {
    return content;
  }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `bun test packages/sdk/tests/body.test.ts`
Expected: PASS (6 tests).

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/src/body.ts packages/sdk/tests/body.test.ts
git commit -m "feat(sdk): add renderBody html→text/markdown converter"
```

---

## Task 4: Implement `renderOutput`

**Files:**
- Create: `packages/sdk/src/format.ts`
- Test: `packages/sdk/tests/format.test.ts`

- [ ] **Step 1: Write the failing tests**

Create `packages/sdk/tests/format.test.ts`:

```ts
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `bun test packages/sdk/tests/format.test.ts`
Expected: FAIL with "Cannot find module '../src/format.js'".

- [ ] **Step 3: Implement `renderOutput`**

Create `packages/sdk/src/format.ts`:

```ts
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `bun test packages/sdk/tests/format.test.ts`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/src/format.ts packages/sdk/tests/format.test.ts
git commit -m "feat(sdk): add renderOutput toon/json encoder"
```

---

## Task 5: Make `MailClient` body-aware

**Files:**
- Modify: `packages/sdk/src/mail-client.ts`
- Test: `packages/sdk/tests/mail-client.test.ts`

- [ ] **Step 1: Write the failing tests**

Add these `describe` blocks to the end of `packages/sdk/tests/mail-client.test.ts` (keep the existing `makeGraphClient` helper and existing tests):

```ts
describe("MailClient.list body modes", () => {
  const sample = {
    value: [
      {
        id: "1",
        subject: "Hi",
        bodyPreview: "snippet",
        body: { contentType: "HTML", content: "<p>Hello <strong>world</strong></p>" },
      },
    ],
  };

  it("preview (default) drops body, keeps bodyPreview, selects bodyPreview not body", async () => {
    const gc = makeGraphClient({ list: mock(() => Promise.resolve(structuredClone(sample))) });
    const client = new MailClient(gc);
    const res = await client.list();
    const [, opts] = (gc.list as ReturnType<typeof mock>).mock.calls[0] as [string, { $select?: string }];
    expect(opts.$select).toContain("bodyPreview");
    expect(opts.$select).not.toContain(",body");
    expect(res.value[0]!.body).toBeUndefined();
    expect(res.value[0]!.bodyPreview).toBe("snippet");
  });

  it("none drops both body and bodyPreview", async () => {
    const gc = makeGraphClient({ list: mock(() => Promise.resolve(structuredClone(sample))) });
    const client = new MailClient(gc);
    const res = await client.list({ body: "none" });
    expect(res.value[0]!.body).toBeUndefined();
    expect(res.value[0]!.bodyPreview).toBeUndefined();
  });

  it("full converts each body to the requested format", async () => {
    const gc = makeGraphClient({ list: mock(() => Promise.resolve(structuredClone(sample))) });
    const client = new MailClient(gc);
    const res = await client.list({ body: "full", bodyFormat: "markdown" });
    const [, opts] = (gc.list as ReturnType<typeof mock>).mock.calls[0] as [string, { $select?: string }];
    expect(opts.$select).toContain("body");
    expect(res.value[0]!.body!.contentType).toBe("markdown");
    expect(res.value[0]!.body!.content).toContain("**world**");
  });

  it("respects an explicit select verbatim", async () => {
    const gc = makeGraphClient({ list: mock(() => Promise.resolve({ value: [] })) });
    const client = new MailClient(gc);
    await client.list({ select: "id,subject" });
    const [, opts] = (gc.list as ReturnType<typeof mock>).mock.calls[0] as [string, { $select?: string }];
    expect(opts.$select).toBe("id,subject");
  });
});

describe("MailClient.get body conversion", () => {
  it("converts the body to the requested format", async () => {
    const msg = { id: "1", body: { contentType: "HTML", content: "<p>Hi <em>there</em></p>" } };
    const gc = makeGraphClient({ get: mock(() => Promise.resolve(structuredClone(msg))) });
    const client = new MailClient(gc);
    const res = await client.get("1", { bodyFormat: "markdown" });
    expect(res.body!.contentType).toBe("markdown");
    expect(res.body!.content).toContain("_there_");
  });

  it("defaults to text and strips HTML", async () => {
    const msg = { id: "1", body: { contentType: "HTML", content: "<p>Hi there</p>" } };
    const gc = makeGraphClient({ get: mock(() => Promise.resolve(structuredClone(msg))) });
    const client = new MailClient(gc);
    const res = await client.get("1");
    expect(res.body!.contentType).toBe("text");
    expect(res.body!.content).toContain("Hi there");
    expect(res.body!.content).not.toContain("<p>");
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `bun test packages/sdk/tests/mail-client.test.ts`
Expected: FAIL (new tests fail — `list` ignores `body`, `get` takes no opts).

- [ ] **Step 3: Rewrite `mail-client.ts`**

Replace the full contents of `packages/sdk/src/mail-client.ts` with:

```ts
import type { GraphClient } from "./graph-client.js";
import type {
  Message,
  MailListResponse,
  DeltaResponse,
  ListMailParams,
  GetMailParams,
  ListBodyMode,
  BodyFormat,
  SendMailParams,
  ReplyParams,
  DraftParams,
} from "./types.js";
import { GRAPH_BASE_URL } from "./config.js";
import { renderBody } from "./body.js";

const SUMMARY_FIELDS =
  "id,subject,from,toRecipients,receivedDateTime,sentDateTime,isRead,isDraft,conversationId";

function defaultSelect(mode: ListBodyMode): string {
  if (mode === "full") return `${SUMMARY_FIELDS},body`;
  if (mode === "preview") return `${SUMMARY_FIELDS},bodyPreview`;
  return SUMMARY_FIELDS;
}

function applyListBody(
  resp: MailListResponse,
  mode: ListBodyMode,
  fmt: BodyFormat
): MailListResponse {
  resp.value = resp.value.map((m) => {
    const next: Message = { ...m };
    if (mode === "none") {
      delete next.body;
      delete next.bodyPreview;
    } else if (mode === "preview") {
      delete next.body;
    } else if (mode === "full" && next.body) {
      next.body = {
        contentType: fmt,
        content: renderBody(next.body.content, next.body.contentType, fmt),
      };
    }
    return next;
  });
  return resp;
}

export class MailClient {
  constructor(private readonly graph: GraphClient) {}

  async list(params: Partial<ListMailParams> = {}): Promise<MailListResponse> {
    const mode = params.body ?? "preview";
    const fmt = params.bodyFormat ?? "text";

    if (params.cursor) {
      const result = (await this.graph.list<Message>(params.cursor, {})) as MailListResponse;
      return applyListBody(result, mode, fmt);
    }

    const folder = params.folder ?? "inbox";
    const opts = {
      $top: params.limit ?? 25,
      $select: params.select ?? defaultSelect(mode),
      ...(params.filter && { $filter: params.filter }),
      ...(params.orderby && { $orderby: params.orderby }),
    };
    const result = (await this.graph.list<Message>(
      `/me/mailFolders/${encodeURIComponent(folder)}/messages`,
      opts
    )) as MailListResponse;
    return applyListBody(result, mode, fmt);
  }

  async get(id: string, opts: Partial<GetMailParams> = {}): Promise<Message> {
    const fmt = opts.bodyFormat ?? "text";
    const msg = await this.graph.get<Message>(`/me/messages/${encodeURIComponent(id)}`);
    if (msg.body) {
      msg.body = {
        contentType: fmt,
        content: renderBody(msg.body.content, msg.body.contentType, fmt),
      };
    }
    return msg;
  }

  async send(params: SendMailParams): Promise<void> {
    await this.graph.post("/me/sendMail", {
      message: {
        subject: params.subject,
        body: { contentType: params.contentType ?? "HTML", content: params.body },
        toRecipients: [{ emailAddress: { address: params.to } }],
      },
      saveToSentItems: true,
    });
  }

  async reply(id: string, params: ReplyParams): Promise<void> {
    await this.graph.post(`/me/messages/${encodeURIComponent(id)}/reply`, {
      message: {
        body: { contentType: params.contentType ?? "HTML", content: params.body },
      },
    });
  }

  async createDraft(params: DraftParams): Promise<Message> {
    return this.graph.post<object, Message>("/me/messages", {
      subject: params.subject,
      body: { contentType: params.contentType ?? "HTML", content: params.body },
      toRecipients: [{ emailAddress: { address: params.to } }],
      isDraft: true,
    });
  }

  async sync(deltaLink?: string): Promise<DeltaResponse> {
    const path = deltaLink
      ? deltaLink.replace(GRAPH_BASE_URL, "")
      : "/me/mailFolders/inbox/messages/delta";
    const result = await this.graph.list<Message>(path);
    return result as DeltaResponse;
  }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `bun test packages/sdk/tests/mail-client.test.ts`
Expected: PASS (existing tests + 6 new tests).

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/src/mail-client.ts packages/sdk/tests/mail-client.test.ts
git commit -m "feat(sdk): body-aware MailClient get/list with select shaping"
```

---

## Task 6: Export new utilities from the SDK

**Files:**
- Modify: `packages/sdk/src/index.ts`

- [ ] **Step 1: Add the new type exports**

In `packages/sdk/src/index.ts`, inside the `export type { ... } from "./types.js";` block, add these entries (e.g. after `ListMailParams,`):

```ts
  GetMailParams,
  BodyFormat,
  ListBodyMode,
  OutputFormat,
```

- [ ] **Step 2: Add the new schema exports**

In the same file, inside the `export { ... } from "./types.js";` block, add (e.g. after `ListMailParamsSchema,`):

```ts
  GetMailParamsSchema,
  BodyFormatSchema,
  ListBodyModeSchema,
  OutputFormatSchema,
```

- [ ] **Step 3: Export the new utilities**

At the end of `packages/sdk/src/index.ts`, after the `export { MailClient } from "./mail-client.js";` line, add:

```ts

// Output + body rendering
export { renderBody } from "./body.js";
export { renderOutput } from "./format.js";
```

- [ ] **Step 4: Type-check the SDK**

Run: `bun run --filter '@outlook-toolkit/sdk' lint`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/sdk/src/index.ts
git commit -m "feat(sdk): export renderBody, renderOutput, format enums"
```

---

## Task 7: Wire the CLI to shared rendering + body flags

**Files:**
- Modify: `packages/cli/src/commands/mail.ts`
- Modify: `packages/cli/package.json` (add SDK is already a dep; remove direct toon dep usage)

- [ ] **Step 1: Replace the imports and add parse helpers**

In `packages/cli/src/commands/mail.ts`, replace the top import block (lines 1–9, the `@stricli/core`, `@toon-format/toon`, SDK, and context imports) with:

```ts
import { buildCommand, buildRouteMap } from "@stricli/core";
import {
  OutlookAuth,
  TokenStore,
  GraphClient,
  MailClient,
  renderOutput,
  type BodyFormat,
  type ListBodyMode,
} from "@outlook-toolkit/sdk";
import { resolveCliConfig } from "../context.js";

function parseBodyFormat(s: string): BodyFormat {
  if (s === "text" || s === "markdown" || s === "html") return s;
  throw new Error(`--body-format must be one of: text, markdown, html (got: "${s}")`);
}

function parseListBody(s: string): ListBodyMode {
  if (s === "none" || s === "preview" || s === "full") return s;
  throw new Error(`--body must be one of: none, preview, full (got: "${s}")`);
}
```

(This removes the `import { encode } from "@toon-format/toon";` line — `renderOutput` replaces it.)

- [ ] **Step 2: Rewrite the `listCommand`**

Replace the entire `const listCommand = buildCommand({ ... });` block with:

```ts
const listCommand = buildCommand({
  docs: { brief: "List messages in a mail folder" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      folder: { kind: "parsed", brief: "Folder (default: inbox)", parse: String, optional: true },
      limit: { kind: "parsed", brief: "Max messages (default: 25)", parse: Number, optional: true },
      cursor: { kind: "parsed", brief: "Pagination cursor ($skipToken URL)", parse: String, optional: true },
      body: { kind: "parsed", brief: "Body: none|preview|full (default: preview)", parse: parseListBody, optional: true },
      bodyFormat: { kind: "parsed", brief: "Body format: text|markdown|html (default: text)", parse: parseBodyFormat, optional: true },
      toon: { kind: "boolean", brief: "Output as TOON (default)", default: true },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
      csv: { kind: "boolean", brief: "Output as CSV", default: false },
    },
  },
  async func(
    this: void,
    flags: {
      profile?: string;
      folder?: string;
      limit?: number;
      cursor?: string;
      body?: ListBodyMode;
      bodyFormat?: BodyFormat;
      toon: boolean;
      json: boolean;
      csv: boolean;
    }
  ) {
    const mail = await getMailClient(flags.profile);
    const result = await mail.list({
      folder: flags.folder ?? "inbox",
      limit: flags.limit ?? 25,
      cursor: flags.cursor,
      body: flags.body ?? "preview",
      bodyFormat: flags.bodyFormat ?? "text",
    });

    if (flags.csv) {
      console.log("id,subject,from,receivedDateTime,isRead");
      for (const m of result.value) {
        const from = m.from?.emailAddress?.address ?? "";
        console.log(`${m.id},${JSON.stringify(m.subject ?? "")},${from},${m.receivedDateTime ?? ""},${m.isRead ?? ""}`);
      }
    } else {
      console.log(renderOutput(result, flags.json ? "json" : "toon"));
    }

    if (result["@odata.nextLink"]) {
      process.stderr.write(`\nMore results available. Use --cursor=<nextLink> to continue.\n`);
    }
  },
});
```

- [ ] **Step 3: Rewrite the `getCommand`**

Replace the entire `const getCommand = buildCommand({ ... });` block with:

```ts
const getCommand = buildCommand({
  docs: { brief: "Get a single message by ID" },
  parameters: {
    flags: {
      profile: { kind: "parsed", brief: "Profile name", parse: String, optional: true },
      bodyFormat: { kind: "parsed", brief: "Body format: text|markdown|html (default: text)", parse: parseBodyFormat, optional: true },
      toon: { kind: "boolean", brief: "Output as TOON (default)", default: true },
      json: { kind: "boolean", brief: "Output as JSON", default: false },
    },
    positional: {
      kind: "tuple",
      parameters: [{ brief: "Message ID", parse: String }],
    },
  },
  async func(
    this: void,
    flags: { profile?: string; bodyFormat?: BodyFormat; toon: boolean; json: boolean },
    id: string
  ) {
    const mail = await getMailClient(flags.profile);
    const message = await mail.get(id, { bodyFormat: flags.bodyFormat ?? "text" });
    console.log(renderOutput(message, flags.json ? "json" : "toon"));
  },
});
```

- [ ] **Step 4: Swap `encode(...)` for `renderOutput(...)` in send/reply/draft/sync**

In `sendCommand`, `replyCommand`, `draftCommand`, and `syncCommand`, replace every occurrence of:

```ts
flags.json ? console.log(JSON.stringify(out, null, 2)) : console.log(encode(out, { keyFolding: "safe" }));
```

with:

```ts
console.log(renderOutput(out, flags.json ? "json" : "toon"));
```

And in `draftCommand` and `syncCommand`, replace:

```ts
    flags.json
      ? console.log(JSON.stringify(draft, null, 2))
      : console.log(encode(draft, { keyFolding: "safe" }));
```

(and the `result` variant in `syncCommand`) with the single-line form, e.g.:

```ts
    console.log(renderOutput(draft, flags.json ? "json" : "toon"));
```

and

```ts
    console.log(renderOutput(result, flags.json ? "json" : "toon"));
```

Leave the plain `console.log("Sent.")` / `console.log("Reply sent.")` success lines unchanged.

- [ ] **Step 5: Type-check the CLI**

Run: `bun run --filter '@outlook-toolkit/cli' lint`
Expected: PASS. (If `tsc` flags an unused `encode` import, ensure it was removed in Step 1.)

- [ ] **Step 6: Smoke-test help and a bad enum value**

Run: `bun run dev:cli -- mail list --help`
Expected: shows `--body`, `--body-format`, `--toon`, `--json`, `--csv` flags.

Run: `bun run dev:cli -- mail get xyz --body-format=foo`
Expected: stderr error `--body-format must be one of: text, markdown, html (got: "foo")`.

- [ ] **Step 7: Commit**

```bash
git add packages/cli/src/commands/mail.ts
git commit -m "feat(cli): body format flags + shared renderOutput"
```

---

## Task 8: Wire the MCP tools to format + body params

**Files:**
- Modify: `packages/mcp/src/tools/mail.ts`

- [ ] **Step 1: Import `renderOutput`**

In `packages/mcp/src/tools/mail.ts`, add `renderOutput` to the SDK import block:

```ts
import {
  resolveConfig,
  OutlookAuth,
  TokenStore,
  GraphClient,
  MailClient,
  OutlookAuthError,
  OutlookConfigError,
  renderOutput,
} from "@outlook-toolkit/sdk";
```

- [ ] **Step 2: Rewrite `outlook_mail_list`**

Replace the `outlook_mail_list` `server.addTool({ ... })` block with:

```ts
  server.addTool({
    name: "outlook_mail_list",
    description:
      "List messages in an Outlook mail folder (default: inbox). Bodies are omitted by default (body=preview keeps a short snippet, body=full returns the converted body). Returns a nextLink cursor for pagination.",
    parameters: z.object({
      folder: z.string().default("inbox").describe("Folder name (inbox, sentitems, drafts, deleteditems)"),
      limit: z.number().int().positive().max(999).default(25).describe("Max messages to return"),
      cursor: z.string().optional().describe("Pagination cursor from a previous call's nextLink"),
      filter: z.string().optional().describe("OData $filter expression"),
      select: z.string().optional().describe("Comma-separated fields to return (overrides body shaping)"),
      orderby: z.string().optional().describe("OData orderby expression (e.g. \"receivedDateTime desc\")"),
      body: z.enum(["none", "preview", "full"]).default("preview").describe("How much body each row carries"),
      bodyFormat: z.enum(["text", "markdown", "html"]).default("text").describe("Body format when body=full"),
      format: z.enum(["toon", "json"]).default("toon").describe("Output encoding (toon is cheaper for LLMs)"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      const result = await mail.list({
        folder: args.folder,
        limit: args.limit,
        cursor: args.cursor,
        filter: args.filter,
        select: args.select,
        orderby: args.orderby,
        body: args.body,
        bodyFormat: args.bodyFormat,
      });
      return renderOutput(result, args.format);
    },
  });
```

- [ ] **Step 3: Rewrite `outlook_mail_get`**

Replace the `outlook_mail_get` block with:

```ts
  server.addTool({
    name: "outlook_mail_get",
    description: "Get a single Outlook message by ID, including the full body (converted to text by default).",
    parameters: z.object({
      id: z.string().describe("Message ID"),
      bodyFormat: z.enum(["text", "markdown", "html"]).default("text").describe("Body format: text (default), markdown, or raw html"),
      format: z.enum(["toon", "json"]).default("toon").describe("Output encoding (toon is cheaper for LLMs)"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      const message = await mail.get(args.id, { bodyFormat: args.bodyFormat });
      return renderOutput(message, args.format);
    },
  });
```

- [ ] **Step 4: Add `format` to `outlook_mail_sync` and route mutation acks through `renderOutput`**

Replace the `outlook_mail_sync` block with:

```ts
  server.addTool({
    name: "outlook_mail_sync",
    description:
      "Delta sync inbox — returns only messages that changed since the last sync. On first call, omit deltaLink to get the full initial sync. Save the returned deltaLink and pass it on subsequent calls to get only changes.",
    parameters: z.object({
      deltaLink: z.string().optional().describe("Delta link from a previous sync call. Omit for initial full sync."),
      format: z.enum(["toon", "json"]).default("toon").describe("Output encoding (toon is cheaper for LLMs)"),
    }),
    execute: async (args) => {
      const mail = await getMailClient();
      const result = await mail.sync(args.deltaLink);
      return renderOutput(result, args.format);
    },
  });
```

Then, in the `outlook_mail_send`, `outlook_mail_reply`, and `outlook_mail_create_draft` `execute` functions, replace each `return JSON.stringify(<x>, null, 2);` with `return renderOutput(<x>, "toon");` (keeping the same `<x>` value — `{ status: "sent", to: args.to }`, `{ status: "replied", messageId: args.id }`, and `draft` respectively).

- [ ] **Step 5: Type-check the MCP package**

Run: `bun run --filter '@outlook-toolkit/mcp' lint`
Expected: PASS.

- [ ] **Step 6: Verify tool schemas load**

Run: `bun run --filter '@outlook-toolkit/mcp' inspect`
Expected: inspector lists `outlook_mail_list`/`outlook_mail_get` with `format`, `body`, `bodyFormat` params. (Ctrl-C to exit.)

- [ ] **Step 7: Commit**

```bash
git add packages/mcp/src/tools/mail.ts
git commit -m "feat(mcp): per-call format + body params via renderOutput"
```

---

## Task 9: Full verification

**Files:** none (verification only)

- [ ] **Step 1: Run the full test suite**

Run: `bun test packages`
Expected: PASS (all SDK tests including new body/format/mail-client tests).

- [ ] **Step 2: Lint every package**

Run: `bun run lint`
Expected: PASS for sdk, cli, mcp.

- [ ] **Step 3: Build every package**

Run: `bun run build`
Expected: all three packages build without error.

- [ ] **Step 4: Commit any incidental fixes**

```bash
git add -A
git commit -m "chore: verification fixes for token-efficient output" || echo "nothing to commit"
```

---

## Self-Review Notes

- **Spec coverage:** `renderBody` (Task 3), `renderOutput` (Task 4), enums + widened `MessageBody` + params (Task 2), body-aware `get`/`list` with `$select` shaping + explicit-select escape hatch (Task 5), SDK exports (Task 6), CLI flags + shared rendering + enumerated errors (Task 7), MCP per-call `format`/`body`/`bodyFormat` (Task 8), best-effort fallback (Task 3 implementation + test), TDD throughout. `sync` keeps current behavior but routes through `renderOutput` (Task 7/8) — matches spec "format routing yes, body conversion no."
- **Type consistency:** `BodyFormat`/`ListBodyMode`/`OutputFormat` names identical across types, body.ts, format.ts, mail-client.ts, CLI, MCP. `MessageBody.contentType` widened to `string` so rendered values (`markdown`/`html`) type-check. `get(id, opts?)` and `list({body, bodyFormat})` signatures consistent between SDK, CLI, and MCP call sites.
- **No placeholders:** every code step contains full code; every run step states expected output.
