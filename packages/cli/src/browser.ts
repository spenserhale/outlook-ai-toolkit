import { exec } from "node:child_process";

export function openBrowser(url: string): void {
  const cmd =
    process.platform === "darwin"
      ? `open "${url}"`
      : process.platform === "win32"
      ? `start "" "${url}"`
      : `xdg-open "${url}"`;
  exec(cmd, () => {});
  process.stderr.write(`Opening browser for Microsoft sign-in...\n`);
  process.stderr.write(`If browser does not open, visit:\n${url}\n`);
}
