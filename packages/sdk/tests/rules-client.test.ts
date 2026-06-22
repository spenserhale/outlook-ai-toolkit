import { describe, expect, it, mock } from "bun:test";
import { RulesClient } from "../src/rules-client.js";
import type { GraphClient } from "../src/graph-client.js";

function makeGraphClient(overrides: Partial<Record<string, unknown>> = {}): GraphClient {
  return {
    get: mock(() => Promise.resolve({})),
    list: mock(() => Promise.resolve({ value: [] })),
    post: mock(() => Promise.resolve({})),
    patch: mock(() => Promise.resolve()),
    delete: mock(() => Promise.resolve()),
    ...overrides,
  } as unknown as GraphClient;
}

describe("RulesClient", () => {
  it("list hits the inbox messageRules collection and returns value", async () => {
    const list = mock(() => Promise.resolve({ value: [{ id: "1", displayName: "r" }] }));
    const rules = new RulesClient(makeGraphClient({ list }));
    const result = await rules.list();
    expect((list.mock.calls[0] as unknown[])[0]).toBe("/me/mailFolders/inbox/messageRules");
    expect(result).toHaveLength(1);
    expect(result[0]!.id).toBe("1");
  });

  it("get encodes the id in the path", async () => {
    const get = mock(() => Promise.resolve({ id: "a/b", displayName: "r" }));
    const rules = new RulesClient(makeGraphClient({ get }));
    await rules.get("a/b");
    expect((get.mock.calls[0] as unknown[])[0]).toBe("/me/mailFolders/inbox/messageRules/a%2Fb");
  });

  it("create posts the parsed params with defaults", async () => {
    const post = mock(() => Promise.resolve({ id: "new", displayName: "x" }));
    const rules = new RulesClient(makeGraphClient({ post }));
    await rules.create({ displayName: "x", actions: { delete: true } });
    const [path, body] = post.mock.calls[0] as [string, Record<string, unknown>];
    expect(path).toBe("/me/mailFolders/inbox/messageRules");
    expect(body.sequence).toBe(1);
    expect(body.isEnabled).toBe(true);
    expect(body.actions).toEqual({ delete: true });
  });

  it("update patches the encoded id", async () => {
    const patch = mock(() => Promise.resolve());
    const rules = new RulesClient(makeGraphClient({ patch }));
    await rules.update("a/b", { isEnabled: false });
    const [path, body] = patch.mock.calls[0] as [string, Record<string, unknown>];
    expect(path).toBe("/me/mailFolders/inbox/messageRules/a%2Fb");
    expect(body).toEqual({ isEnabled: false });
  });

  it("delete removes the encoded id", async () => {
    const del = mock(() => Promise.resolve());
    const rules = new RulesClient(makeGraphClient({ delete: del }));
    await rules.delete("a/b");
    expect((del.mock.calls[0] as unknown[])[0]).toBe("/me/mailFolders/inbox/messageRules/a%2Fb");
  });
});
