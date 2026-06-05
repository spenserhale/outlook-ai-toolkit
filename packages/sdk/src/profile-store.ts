import { existsSync, mkdirSync } from "node:fs";
import { readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { homedir } from "node:os";
import { ProfilesFileSchema, ProfileSchema, type Profile, type ProfilesFile } from "./types.js";

export class ProfileStore {
  private readonly filePath: string;

  constructor(baseDir?: string) {
    const dir = baseDir ?? join(homedir(), ".outlook-toolkit");
    this.filePath = join(dir, "profiles.json");
  }

  private async read(): Promise<ProfilesFile> {
    if (!existsSync(this.filePath)) return {};
    try {
      const raw = await readFile(this.filePath, "utf8");
      const parsed = ProfilesFileSchema.safeParse(JSON.parse(raw));
      return parsed.success ? parsed.data : {};
    } catch {
      return {};
    }
  }

  private async write(data: ProfilesFile): Promise<void> {
    const dir = join(this.filePath, "..");
    mkdirSync(dir, { recursive: true });
    await writeFile(this.filePath, JSON.stringify(data, null, 2), "utf8");
  }

  async save(name: string, profile: Profile): Promise<void> {
    const validated = ProfileSchema.parse(profile);
    const all = await this.read();
    all[name] = validated;
    await this.write(all);
  }

  async get(name: string): Promise<Profile | null> {
    const all = await this.read();
    return all[name] ?? null;
  }

  async list(): Promise<ProfilesFile> {
    return this.read();
  }

  async delete(name: string): Promise<boolean> {
    const all = await this.read();
    if (!(name in all)) return false;
    delete all[name];
    await this.write(all);
    return true;
  }
}
