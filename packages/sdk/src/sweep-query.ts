import type { SweepCondition } from "./types.js";

/**
 * Build a Graph $search (KQL) query for a condition's keyword fields.
 * Returns null when the condition has no keyword fields (date-only) — callers
 * fall back to a $filter on receivedDateTime in that case.
 */
// Embedded double-quotes would break the caller's `"<query>"` $search wrapper.
function sanitize(value: string): string {
  return value.replace(/"/g, "");
}

export function buildSearchQuery(c: SweepCondition): string | null {
  const parts: string[] = [];
  if (c.from) parts.push(`from:${sanitize(c.from)}`);
  if (c.subjectContains) parts.push(`subject:${sanitize(c.subjectContains)}`);
  if (c.bodyContains) parts.push(`body:${sanitize(c.bodyContains)}`);
  return parts.length ? parts.join(" AND ") : null;
}

/** ISO timestamp for `days` before `nowMs` (defaults to current time). */
export function isoDaysAgo(days: number, nowMs: number = Date.now()): string {
  return new Date(nowMs - days * 24 * 60 * 60 * 1000).toISOString();
}

/** True when `receivedDateTime` is strictly before the (now - days) cutoff. */
export function isOlderThan(
  receivedDateTime: string | undefined,
  days: number,
  nowMs: number = Date.now()
): boolean {
  if (!receivedDateTime) return false;
  const received = Date.parse(receivedDateTime);
  if (Number.isNaN(received)) return false;
  return received < nowMs - days * 24 * 60 * 60 * 1000;
}
