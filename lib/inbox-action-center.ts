import type { InboxIntelligenceRecord } from "./inbox-intelligence-models";

// Pure view-derivation over already-loaded, already-durable Inbox Intelligence records for
// the Home "Action Center" (docs/AI_HANDOFF.md Patch 1). No I/O, no AI, no fabricated dates —
// every value here comes from a field the record already stores. See
// docs/DASHBOARD_ACTION_CENTER_REPORT.md for the exact reasoning behind each rule.

const DISPLAY_LIMIT = 3;
const NEARBY_DAY_WINDOW = 6; // "Friday"-style weekday label through this many days out

/**
 * "Needs attention" uses the AI-set, human-reviewed `needsAttention` flag — never a blanket
 * status===\"open\" assumption. A resolved record is excluded even if it was once flagged:
 * "resolved" means nothing further currently requires attention (docs/INBOX_INTELLIGENCE_SHAREPOINT_REPORT.md
 * "Status model"), and combining two already-stored fields with a boolean AND is a deterministic
 * rule, not an inference or a scoring engine.
 */
export function filterNeedsAttention(records: InboxIntelligenceRecord[]): InboxIntelligenceRecord[] {
  return records.filter((record) => record.analysis.needsAttention && record.status !== "resolved");
}

export function filterWaiting(records: InboxIntelligenceRecord[]): InboxIntelligenceRecord[] {
  return records.filter((record) => record.status === "waiting");
}

const DATE_ONLY_PATTERN = /^\d{4}-\d{2}-\d{2}$/;

function parseDateOnly(value: string): Date | null {
  if (!DATE_ONLY_PATTERN.test(value)) return null;
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function startOfDay(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * Earliest already-stored, already-validated due date among a record's action items, or null.
 * Never infers or parses free text for a date — only reads `actionItems[].dueDate`, which the
 * AI pipeline already normalizes to a real YYYY-MM-DD value or null (lib/inbox-intelligence-models.ts
 * normalizeActionItemDueDate) before this record could ever have been saved.
 */
export function earliestDueDate(record: InboxIntelligenceRecord): Date | null {
  const dates = record.analysis.actionItems
    .map((item) => (item.dueDate ? parseDateOnly(item.dueDate) : null))
    .filter((date): date is Date => date !== null);
  if (!dates.length) return null;
  return dates.reduce((earliest, date) => (date < earliest ? date : earliest));
}

export type DueLabel = { text: string; overdue: boolean };

/** Deterministic formatting of an already-stored date compared to today. Returns null for no date — never fabricates one. */
export function formatDueLabel(due: Date | null, today: Date = new Date()): DueLabel | null {
  if (!due) return null;
  const diffDays = Math.round((startOfDay(due).getTime() - startOfDay(today).getTime()) / 86400000);
  if (diffDays < 0) return { text: "Overdue", overdue: true };
  if (diffDays === 0) return { text: "Due today", overdue: false };
  if (diffDays === 1) return { text: "Due tomorrow", overdue: false };
  if (diffDays <= NEARBY_DAY_WINDOW) return { text: new Intl.DateTimeFormat("en-US", { weekday: "long" }).format(due), overdue: false };
  return { text: new Intl.DateTimeFormat("en-US", { month: "short", day: "numeric" }).format(due), overdue: false };
}

/** The most actionable already-stored text for a record: a structured action item first, else the reviewed title, else the summary. */
export function primaryLabel(record: InboxIntelligenceRecord): string {
  return record.analysis.actionItems[0]?.action || record.analysis.suggestedWorkRecord.title || record.analysis.summary;
}

/**
 * Deterministic priority, no scoring engine: overdue first, then nearest real due date, then
 * items with no due date ordered by canonical `metadata.modifiedAt` (most recent first).
 */
export function sortNeedsAttention(records: InboxIntelligenceRecord[]): InboxIntelligenceRecord[] {
  return [...records].sort((a, b) => {
    const dueA = earliestDueDate(a);
    const dueB = earliestDueDate(b);
    if (dueA && dueB) return dueA.getTime() - dueB.getTime();
    if (dueA) return -1;
    if (dueB) return 1;
    return b.metadata.modifiedAt.localeCompare(a.metadata.modifiedAt);
  });
}

/** Ordered by canonical `metadata.modifiedAt` (most recently touched first) — the only timestamp the model actually has; see "waiting age" note below. */
export function sortWaiting(records: InboxIntelligenceRecord[]): InboxIntelligenceRecord[] {
  return [...records].sort((a, b) => b.metadata.modifiedAt.localeCompare(a.metadata.modifiedAt));
}

/**
 * The domain model has no dedicated "entered waiting at" timestamp — `metadata.modifiedAt` is a
 * general last-modified time that could, in principle, be bumped by something other than the
 * open→waiting transition (e.g. a Work Record link written while already waiting). Deliberately
 * NOT surfaced as "Waiting N days" for that reason — a plain label is preferable to false
 * precision, per instruction. If a future phase adds a real status-transition timestamp, this is
 * the place to compute an age label from it.
 */
export const WAITING_AGE_UNAVAILABLE_REASON =
  "No dedicated status-transition timestamp exists yet; metadata.modifiedAt is not reliably \"entered waiting.\"";

export function selectNeedsAttention(records: InboxIntelligenceRecord[], limit = DISPLAY_LIMIT): InboxIntelligenceRecord[] {
  return sortNeedsAttention(filterNeedsAttention(records)).slice(0, limit);
}

export function selectWaiting(records: InboxIntelligenceRecord[], limit = DISPLAY_LIMIT): InboxIntelligenceRecord[] {
  return sortWaiting(filterWaiting(records)).slice(0, limit);
}
