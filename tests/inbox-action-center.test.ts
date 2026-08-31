import { describe, expect, it } from "vitest";
import {
  buildInboxIntelligenceRecord,
  EmailAnalysisSchema,
  type EmailAnalysis,
  type InboxIntelligenceRecord,
} from "../lib/inbox-intelligence-models";
import { REFERENCE_DATA } from "../lib/reference-data";
import {
  earliestDueDate,
  filterNeedsAttention,
  filterWaiting,
  formatDueLabel,
  primaryLabel,
  selectNeedsAttention,
  selectWaiting,
  sortNeedsAttention,
} from "../lib/inbox-action-center";

function analysis(overrides: Partial<EmailAnalysis> = {}): EmailAnalysis {
  return EmailAnalysisSchema.parse({
    summary: "Test summary",
    priority: "medium",
    needsAttention: false,
    actionItems: [],
    followUp: "",
    people: [],
    organizations: [],
    districts: [],
    projects: [],
    tags: [],
    suggestedWorkType: null,
    suggestedWorkRecord: { title: "Suggested title", description: "Suggested description" },
    ...overrides,
  });
}

function record(overrides: Partial<InboxIntelligenceRecord> = {}): InboxIntelligenceRecord {
  const base = buildInboxIntelligenceRecord(analysis(), "excerpt", REFERENCE_DATA, "2026-08-29T12:00:00.000Z");
  return { ...base, appId: crypto.randomUUID(), metadata: { ...base.metadata, modifiedAt: "2026-08-29T12:00:00.000Z" }, ...overrides };
}

describe("filterNeedsAttention", () => {
  it("selects only records with the stored needsAttention flag set", () => {
    const flagged = record({ analysis: analysis({ needsAttention: true }) });
    const unflagged = record({ analysis: analysis({ needsAttention: false }) });
    expect(filterNeedsAttention([flagged, unflagged])).toEqual([flagged]);
  });

  it("excludes a resolved record even if it was flagged needsAttention", () => {
    const resolvedButFlagged = record({ analysis: analysis({ needsAttention: true }), status: "resolved", resolvedAt: "2026-08-29T12:00:00.000Z" });
    expect(filterNeedsAttention([resolvedButFlagged])).toEqual([]);
  });

  it("never infers attention from status alone — an unflagged open record is excluded", () => {
    const openButUnflagged = record({ analysis: analysis({ needsAttention: false }), status: "open" });
    expect(filterNeedsAttention([openButUnflagged])).toEqual([]);
  });
});

describe("filterWaiting", () => {
  it("selects only records whose status is exactly waiting", () => {
    const waiting = record({ status: "waiting" });
    const open = record({ status: "open" });
    const resolved = record({ status: "resolved", resolvedAt: "2026-08-29T12:00:00.000Z" });
    expect(filterWaiting([waiting, open, resolved])).toEqual([waiting]);
  });
});

describe("earliestDueDate / formatDueLabel — real dates only, never fabricated", () => {
  it("returns null when no action item has a due date", () => {
    const r = record({ analysis: analysis({ actionItems: [{ action: "a", dueDate: null, owner: "me" }] }) });
    expect(earliestDueDate(r)).toBeNull();
    expect(formatDueLabel(earliestDueDate(r))).toBeNull();
  });

  it("picks the earliest of several stored due dates", () => {
    const r = record({
      analysis: analysis({
        actionItems: [
          { action: "a", dueDate: "2026-09-10", owner: "me" },
          { action: "b", dueDate: "2026-09-01", owner: "me" },
        ],
      }),
    });
    expect(earliestDueDate(r)?.toISOString().slice(0, 10)).toBe("2026-09-01");
  });

  it("labels a past date Overdue, deterministically relative to 'today'", () => {
    const today = new Date(2026, 7, 29); // Aug 29, 2026
    const label = formatDueLabel(new Date(2026, 7, 28), today);
    expect(label).toEqual({ text: "Overdue", overdue: true });
  });

  it("labels today and tomorrow explicitly", () => {
    const today = new Date(2026, 7, 29);
    expect(formatDueLabel(new Date(2026, 7, 29), today)).toEqual({ text: "Due today", overdue: false });
    expect(formatDueLabel(new Date(2026, 7, 30), today)).toEqual({ text: "Due tomorrow", overdue: false });
  });

  it("labels a nearby date with its weekday name, and a far date with a short date", () => {
    const today = new Date(2026, 7, 24); // Monday, Aug 24 2026
    const nearby = formatDueLabel(new Date(2026, 7, 28), today); // Friday, 4 days out
    expect(nearby?.overdue).toBe(false);
    expect(nearby?.text).toBe("Friday");
    const far = formatDueLabel(new Date(2026, 8, 15), today); // more than a week out
    expect(far?.overdue).toBe(false);
    expect(far?.text).not.toBe("Friday");
  });
});

describe("primaryLabel — most actionable already-stored text, never invented", () => {
  it("prefers the first action item's action text", () => {
    const r = record({
      analysis: analysis({
        actionItems: [{ action: "Send the report", dueDate: null, owner: "me" }],
        suggestedWorkRecord: { title: "A title", description: "d" },
      }),
    });
    expect(primaryLabel(r)).toBe("Send the report");
  });

  it("falls back to the suggested work record title when there are no action items", () => {
    const r = record({ analysis: analysis({ actionItems: [], suggestedWorkRecord: { title: "Fallback title", description: "d" } }) });
    expect(primaryLabel(r)).toBe("Fallback title");
  });
});

describe("sortNeedsAttention — deterministic, no scoring engine", () => {
  it("orders overdue/nearest-due items before items with no due date", () => {
    const noDue = record({ analysis: analysis({ actionItems: [{ action: "no date", dueDate: null, owner: "me" }] }) });
    const dueSoon = record({ analysis: analysis({ actionItems: [{ action: "soon", dueDate: "2026-09-01", owner: "me" }] }) });
    const dueLater = record({ analysis: analysis({ actionItems: [{ action: "later", dueDate: "2026-09-10", owner: "me" }] }) });
    const sorted = sortNeedsAttention([noDue, dueLater, dueSoon]);
    expect(sorted.map((r) => primaryLabel(r))).toEqual(["soon", "later", "no date"]);
  });

  it("orders undated items by canonical modifiedAt, most recent first", () => {
    const older = record({ metadata: { providerId: "1", version: 1, createdAt: "x", modifiedAt: "2026-08-01T00:00:00.000Z", syncState: "saved" } });
    const newer = record({ metadata: { providerId: "2", version: 1, createdAt: "x", modifiedAt: "2026-08-15T00:00:00.000Z", syncState: "saved" } });
    expect(sortNeedsAttention([older, newer])).toEqual([newer, older]);
  });
});

describe("selectNeedsAttention / selectWaiting — respect the display limit", () => {
  it("caps needs-attention display to the given limit without dropping the underlying data elsewhere", () => {
    const flagged = Array.from({ length: 6 }, (_, i) =>
      record({ analysis: analysis({ needsAttention: true }), appId: `flagged-${i}` }),
    );
    expect(selectNeedsAttention(flagged, 3)).toHaveLength(3);
    expect(filterNeedsAttention(flagged)).toHaveLength(6); // full count still available for a badge/count display
  });

  it("caps waiting display to the given limit", () => {
    const waiting = Array.from({ length: 5 }, (_, i) => record({ status: "waiting", appId: `waiting-${i}` }));
    expect(selectWaiting(waiting, 3)).toHaveLength(3);
  });
});
