import { createInterface } from "node:readline/promises";
import { stdin, stdout } from "node:process";

/** True when we can prompt the user interactively (a real terminal on stdin). */
export function isInteractive(): boolean {
  return Boolean(stdin.isTTY);
}

/** Ask a single question and return the trimmed answer. Caller must ensure isInteractive(). */
export async function promptLine(question: string): Promise<string> {
  const rl = createInterface({ input: stdin, output: stdout });
  try {
    const answer = await rl.question(question);
    return answer.trim();
  } finally {
    rl.close();
  }
}
