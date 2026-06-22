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

  it("fetches cursor URL directly when cursor is provided", async () => {
    const gc = makeGraphClient();
    const client = new MailClient(gc);
    const cursor = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skiptoken=abc123";
    await client.list({ cursor });
    const [path] = (gc.list as ReturnType<typeof mock>).mock.calls[0] as [string];
    expect(path).toBe(cursor);
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
    const fields = (opts.$select ?? "").split(",");
    expect(fields).toContain("bodyPreview");
    expect(fields).not.toContain("body");
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
    expect((opts.$select ?? "").split(",")).toContain("body");
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

  it("applies body post-processing even when an explicit select is given", async () => {
    const sample = {
      value: [
        { id: "1", bodyPreview: "snip", body: { contentType: "HTML", content: "<p>x</p>" } },
      ],
    };
    const gc = makeGraphClient({ list: mock(() => Promise.resolve(structuredClone(sample))) });
    const client = new MailClient(gc);
    const res = await client.list({ select: "id,body,bodyPreview" }); // default body mode = preview
    const [, opts] = (gc.list as ReturnType<typeof mock>).mock.calls[0] as [string, { $select?: string }];
    expect(opts.$select).toBe("id,body,bodyPreview"); // honored verbatim
    expect(res.value[0]!.body).toBeUndefined();       // preview mode still strips body
    expect(res.value[0]!.bodyPreview).toBe("snip");
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

describe("MailClient.move / moveBatch", () => {
  it("move posts destinationId to the encoded message move endpoint", async () => {
    const post = mock(() => Promise.resolve(undefined));
    const mail = new MailClient(makeGraphClient({ post }));
    await mail.move("a/b", "archive");
    const [path, body] = post.mock.calls[0] as [string, Record<string, unknown>];
    expect(path).toBe("/me/messages/a%2Fb/move");
    expect(body).toEqual({ destinationId: "archive" });
  });

  it("moveBatch chunks into groups of 20 and tallies statuses", async () => {
    const ids = Array.from({ length: 25 }, (_, i) => `id${i}`);
    const post = mock((path: string, body: { requests: Array<{ id: string }> }) =>
      Promise.resolve({
        responses: body.requests.map((r) => ({ id: r.id, status: 200 })),
      })
    );
    const mail = new MailClient(makeGraphClient({ post }));
    const result = await mail.moveBatch(ids, "deleteditems");
    expect(post.mock.calls).toHaveLength(2); // 20 + 5
    expect(result.moved).toBe(25);
    expect(result.failed).toEqual([]);
    expect((post.mock.calls[0] as unknown[])[0]).toBe("/$batch");
  });

  it("moveBatch records failures by original id", async () => {
    const post = mock(() =>
      Promise.resolve({ responses: [{ id: "0", status: 404 }, { id: "1", status: 200 }] })
    );
    const mail = new MailClient(makeGraphClient({ post }));
    const result = await mail.moveBatch(["x", "y"], "archive");
    expect(result.moved).toBe(1);
    expect(result.failed).toEqual([{ id: "x", status: 404 }]);
  });

  it("moveBatch returns early for an empty id list", async () => {
    const post = mock(() => Promise.resolve({ responses: [] }));
    const mail = new MailClient(makeGraphClient({ post }));
    const result = await mail.moveBatch([], "archive");
    expect(post.mock.calls).toHaveLength(0);
    expect(result.moved).toBe(0);
  });
});

describe("MailClient.findMatches", () => {
  it("uses $search for keyword conditions and dedupes by id across conditions", async () => {
    const list = mock(() =>
      Promise.resolve({ value: [{ id: "m1" }, { id: "m2" }] })
    );
    const mail = new MailClient(makeGraphClient({ list }));
    const matches = await mail.findMatches(
      [{ from: "alice@x.com" }, { from: "alice@x.com" }],
      "inbox",
      200
    );
    const [path, opts] = list.mock.calls[0] as [string, Record<string, unknown>];
    expect(path).toBe("/me/mailFolders/inbox/messages");
    expect(opts.$search).toBe('"from:alice@x.com"');
    expect(matches.map((m) => m.id)).toEqual(["m1", "m2"]); // deduped
  });

  it("uses $filter on receivedDateTime for a date-only condition", async () => {
    const list = mock(() => Promise.resolve({ value: [{ id: "m1" }] }));
    const mail = new MailClient(makeGraphClient({ list }));
    await mail.findMatches([{ olderThanDays: 10 }], "inbox", 200);
    const [, opts] = list.mock.calls[0] as [string, Record<string, unknown>];
    expect(String(opts.$filter)).toContain("receivedDateTime lt ");
    expect(opts.$search).toBeUndefined();
  });

  it("applies olderThanDays client-side when combined with search terms", async () => {
    const recent = new Date().toISOString();
    const old = new Date(Date.now() - 40 * 86400_000).toISOString();
    const list = mock(() =>
      Promise.resolve({
        value: [
          { id: "old", receivedDateTime: old },
          { id: "recent", receivedDateTime: recent },
        ],
      })
    );
    const mail = new MailClient(makeGraphClient({ list }));
    const matches = await mail.findMatches(
      [{ from: "alice@x.com", olderThanDays: 10 }],
      "inbox",
      200
    );
    expect(matches.map((m) => m.id)).toEqual(["old"]);
  });

  it("stops at max", async () => {
    const list = mock(() =>
      Promise.resolve({ value: [{ id: "a" }, { id: "b" }, { id: "c" }] })
    );
    const mail = new MailClient(makeGraphClient({ list }));
    const matches = await mail.findMatches([{ from: "x@y.com" }], "inbox", 2);
    expect(matches).toHaveLength(2);
  });
});
