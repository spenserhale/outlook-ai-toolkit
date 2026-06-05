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
