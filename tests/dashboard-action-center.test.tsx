// @vitest-environment jsdom
import { cleanup, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import IUWorkTracker from "../app/IUWorkTracker";
import { MemoryDataProvider } from "../lib/data-provider";
import {
  buildInboxIntelligenceRecord,
  EmailAnalysisSchema,
  type EmailAnalysis,
  type InboxIntelligenceRecord,
} from "../lib/inbox-intelligence-models";
import type { InboxIntelligenceProvider, InboxIntelligenceResult } from "../lib/inbox-intelligence-provider";
import { REFERENCE_DATA } from "../lib/reference-data";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
  vi.useRealTimers();
});

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
  return {
    ...base,
    appId: crypto.randomUUID(),
    metadata: { providerId: crypto.randomUUID(), version: 1, createdAt: "2026-08-29T12:00:00.000Z", modifiedAt: "2026-08-29T12:00:00.000Z", syncState: "saved" },
    ...overrides,
  };
}

class FakeInboxProvider implements InboxIntelligenceProvider {
  constructor(private records: InboxIntelligenceRecord[] = [], private failList = false) {}
  async list(): Promise<InboxIntelligenceResult<InboxIntelligenceRecord[]>> {
    if (this.failList) return { status: "network_error", message: "The inbox store could not be reached." };
    return { status: "success", value: this.records };
  }
  async create(r: InboxIntelligenceRecord): Promise<InboxIntelligenceResult<InboxIntelligenceRecord>> {
    return { status: "success", value: r };
  }
  async update(r: InboxIntelligenceRecord): Promise<InboxIntelligenceResult<InboxIntelligenceRecord>> {
    return { status: "success", value: r };
  }
}

async function renderHome(records: InboxIntelligenceRecord[] = [], options: { failList?: boolean } = {}) {
  render(
    <IUWorkTracker
      dataProvider={new MemoryDataProvider([])}
      inboxDataProvider={new FakeInboxProvider(records, options.failList ?? false)}
    />,
  );
  await screen.findByRole("heading", { name: "Action Center" });
}

describe("Dashboard Action Center — empty state", () => {
  it("renders an intentional empty state, not dead zero-cards, when there are no Inbox Intelligence records", async () => {
    await renderHome([]);
    expect(screen.getByText(/nothing needs your attention/i)).toBeTruthy();
    expect(screen.getByRole("button", { name: /open inbox intelligence/i })).toBeTruthy();
    // The summary metric strip (Open/Waiting/Resolved) only renders once there is at least one record.
    expect(screen.queryByText("open")).toBeNull();
  });
});

describe("Dashboard Action Center — status counts", () => {
  it("computes Open/Waiting/Resolved counts from durable records", async () => {
    const records = [
      ...Array.from({ length: 4 }, () => record({ status: "open" })),
      ...Array.from({ length: 2 }, () => record({ status: "waiting" })),
      ...Array.from({ length: 3 }, () => record({ status: "resolved", resolvedAt: "2026-08-29T12:00:00.000Z" })),
    ];
    await renderHome(records);
    const strip = await screen.findByText("open");
    const container = strip.closest(".metric-strip");
    expect(container).toBeTruthy();
    expect(container?.textContent).toContain("4");
    expect(container?.textContent).toContain("open");
    expect(container?.textContent).toContain("2");
    expect(container?.textContent).toContain("waiting");
    expect(container?.textContent).toContain("3");
    expect(container?.textContent).toContain("resolved");
  });
});

describe("Dashboard Action Center — Needs Attention", () => {
  it("shows only records with the stored needsAttention flag, never inferring from status alone", async () => {
    const flagged = record({
      status: "open",
      analysis: analysis({ needsAttention: true, actionItems: [{ action: "Send the STEELS agenda", dueDate: null, owner: "me" }] }),
    });
    const unflagged = record({ status: "open", analysis: analysis({ needsAttention: false, suggestedWorkRecord: { title: "Not flagged", description: "d" } }) });
    await renderHome([flagged, unflagged]);
    expect(await screen.findByText("Send the STEELS agenda")).toBeTruthy();
    expect(screen.queryByText("Not flagged")).toBeNull();
  });
});

describe("Dashboard Action Center — Waiting", () => {
  it("shows only records whose status is exactly waiting", async () => {
    const waiting = record({ status: "waiting", analysis: analysis({ suggestedWorkRecord: { title: "Equipment quote", description: "d" } }) });
    const open = record({ status: "open", analysis: analysis({ suggestedWorkRecord: { title: "Not waiting", description: "d" } }) });
    await renderHome([waiting, open]);
    expect(await screen.findByText("Equipment quote")).toBeTruthy();
    expect(screen.queryByText("Not waiting")).toBeNull();
  });
});

function isoDateOnly(date: Date): string {
  return new Intl.DateTimeFormat("en-CA", { year: "numeric", month: "2-digit", day: "2-digit" }).format(date);
}

describe("Dashboard Action Center — due date display", () => {
  it("renders a real stored due date deterministically", async () => {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const dueTomorrow = record({
      status: "open",
      analysis: analysis({
        needsAttention: true,
        actionItems: [{ action: "Follow up with district contact", dueDate: isoDateOnly(tomorrow), owner: "me" }],
      }),
    });
    await renderHome([dueTomorrow]);
    expect(await screen.findByText("Due tomorrow")).toBeTruthy();
  });

  it("never fabricates a date label when no action item has a stored due date", async () => {
    const noDueDate = record({
      status: "open",
      analysis: analysis({ needsAttention: true, actionItems: [{ action: "Review grant draft", dueDate: null, owner: "me" }] }),
    });
    await renderHome([noDueDate]);
    await screen.findByText("Review grant draft");
    expect(screen.queryByText(/due /i)).toBeNull();
    expect(screen.queryByText("Overdue")).toBeNull();
  });
});

describe("Dashboard Action Center — navigation", () => {
  it("reaches Inbox Intelligence when a Needs Attention row is clicked", async () => {
    const user = userEvent.setup();
    const flagged = record({
      status: "open",
      analysis: analysis({ needsAttention: true, actionItems: [{ action: "Send the agenda", dueDate: null, owner: "me" }] }),
    });
    await renderHome([flagged]);
    await user.click(await screen.findByText("Send the agenda"));
    await waitFor(() => expect(screen.getByText(/ai-assisted intake/i)).toBeTruthy());
  });

  it("View Intelligence reaches the same Inbox Intelligence screen", async () => {
    const user = userEvent.setup();
    const waiting = record({ status: "waiting" });
    await renderHome([waiting]);
    await user.click(await screen.findByRole("button", { name: /view intelligence/i }));
    await waitFor(() => expect(screen.getByText(/ai-assisted intake/i)).toBeTruthy());
  });
});

describe("Dashboard Action Center — zero AI requests from Home", () => {
  it("never calls the Anthropic analysis route while rendering/loading the Dashboard", async () => {
    const fetchSpy = vi.fn(async () => new Response(JSON.stringify({}), { status: 200 }));
    vi.stubGlobal("fetch", fetchSpy);
    const flagged = record({ status: "open", analysis: analysis({ needsAttention: true }) });
    await renderHome([flagged, record({ status: "waiting" }), record({ status: "resolved", resolvedAt: "2026-08-29T12:00:00.000Z" })]);
    // Give any stray microtask/effect a turn to fire before asserting.
    await new Promise((resolve) => setTimeout(resolve, 0));
    const calledInboxRoute = fetchSpy.mock.calls.some((call: unknown[]) => String(call[0]).includes("inbox-intelligence"));
    expect(calledInboxRoute).toBe(false);
    vi.unstubAllGlobals();
  });
});

describe("Dashboard Action Center — display limit", () => {
  it("caps Needs Attention rows without hiding the route into full Inbox Intelligence", async () => {
    const many = Array.from({ length: 6 }, (_, i) =>
      record({
        status: "open",
        analysis: analysis({ needsAttention: true, suggestedWorkRecord: { title: `Item ${i}`, description: "d" } }),
      }),
    );
    await renderHome(many);
    const rows = await screen.findAllByText(/^Item \d$/);
    expect(rows.length).toBeLessThanOrEqual(3);
    expect(screen.getByRole("button", { name: /view intelligence/i })).toBeTruthy();
  });
});

describe("Dashboard Action Center — graceful provider failure", () => {
  it("does not crash Home when the Inbox provider fails to load, and other Dashboard content remains", async () => {
    await renderHome([], { failList: true });
    expect(await screen.findByText(/inbox intelligence is temporarily unavailable/i)).toBeTruthy();
    // The rest of Home (Today panel) is still present — the Dashboard did not crash.
    expect(screen.getByRole("heading", { name: "Today" })).toBeTruthy();
  });
});
