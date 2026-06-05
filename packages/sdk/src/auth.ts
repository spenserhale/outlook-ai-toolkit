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
