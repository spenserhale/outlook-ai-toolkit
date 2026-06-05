#!/usr/bin/env bun
import { run } from "@stricli/core";
import { app } from "./app.js";
import {
  OutlookConfigError,
  OutlookAuthError,
  OutlookNotFoundError,
  OutlookRateLimitError,
  OutlookError,
} from "@outlook-toolkit/sdk";

function exitCodeFor(err: unknown): number {
  if (err instanceof OutlookConfigError) return 3;
  if (err instanceof OutlookAuthError) return 5;
  if (err instanceof OutlookNotFoundError) return 4;
  if (err instanceof OutlookRateLimitError) return 6;
  return 1;
}

await run(app, process.argv.slice(2), { process }).catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  process.stderr.write(`Error: ${message}\n`);
  process.exit(exitCodeFor(err));
});
