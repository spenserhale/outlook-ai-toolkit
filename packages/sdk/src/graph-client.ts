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

const RETRYABLE_ALL_METHODS = new Set([429]);
const RETRYABLE_SAFE_METHODS = new Set([503, 504]);
const SAFE_METHODS = new Set(["GET", "DELETE"]);

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

      const isRetryable =
        (RETRYABLE_ALL_METHODS.has(res.status) ||
          (RETRYABLE_SAFE_METHODS.has(res.status) && SAFE_METHODS.has(method))) &&
        attempt <= this.retry.maxAttempts;

      if (isRetryable) {
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
