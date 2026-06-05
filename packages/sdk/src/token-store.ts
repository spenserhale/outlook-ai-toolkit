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
