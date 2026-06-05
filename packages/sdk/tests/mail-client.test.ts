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
