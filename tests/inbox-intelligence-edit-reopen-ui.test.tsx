// @vitest-environment jsdom
//
// Patch 7.1 — Inbox Intelligence readability + Edit/Reopen. End-to-end through the real
// InboxIntelligence component and a real SessionInboxIntelligenceProvider (in-memory — zero
// SharePoint writes). Proves: Edit is available regardless of status, Cancel writes nothing,
// Save Changes makes exactly one update() using the current RecordVersion/AppId and never a
// create(), editing preserves status/content it doesn't expose, Reopen (which already existed)
// keeps working through the same update() pathway, and neither Edit nor Reopen ever calls
// Anthropic (no fetch) or creates a Work Record (openLog is never invoked).
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { readFileSync } from "node:fs";
import { useState } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import InboxIntelligence from "../app/InboxIntelligence";
import { buildInboxIntelligenceRecord, type EmailAnalysis } from "../lib/inbox-intelligence-models";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import type { InboxIntelligenceRecord } from "../lib/inbox-intelligence-models";
import { REFERENCE_DATA } from "../lib/reference-data";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

function baseWorkRecord(): WorkRecord {
  return {
    appId: "draft", title: "", activityDate: "2026-08-31", activityType: "", description: "", detailedNotes: "",
    durationMinutes: 60, status: "complete", engagementScope: "none", projectIds: [], organizationIds: [], contactIds: [],
    categoryIds: [], reach: { educatorsLeaders: 0, studentsFamilies: 0, workforceCommunity: 0, other: 0 }, evidenceSummary: "",
    evidenceReferenceIds: [], output: "", outcome: "", nextStep: "", followUpNeeded: false, followUpDate: null,
    orbit: { reportable: false, primaryDeliverable: null, supportingDeliverables: [], stemPocMinutes: 0, tacMinutes: 0, evidence: "" },
    schemaVersion: WORK_RECORD_SCHEMA_VERSION, metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" }, isSample: false,
  };
}

function analysisFixture(overrides: Partial<EmailAnalysis> = {}): EmailAnalysis {
  return {
    summary: "Original summary text describing the planning email.",
    priority: "medium",
    needsAttention: true,
    actionItems: [{ action: "Send the agenda", dueDate: "2026-02-01", owner: "me" }],
    followUp: "Follow up next week.",
    people: [],
    organizations: [],
    districts: [],
    projects: [],
    tags: [],
    suggestedWorkType: null,
    suggestedWorkRecord: { title: "Original Title", description: "Original description." },
    ...overrides,
  };
}

/** Seeds a real provider with one durable record BEFORE render — the provider's own update()/
 * create() lookups are keyed off its internal array, not whatever the test harness renders, so
 * the record must exist there first (the same pre-seed-before-render discipline used elsewhere
 * in this suite, e.g. tests/contact-import-ui.test.tsx). */
async function seededProvider(analysis: EmailAnalysis): Promise<{ provider: SessionInboxIntelligenceProvider; seeded: InboxIntelligenceRecord }> {
  const provider = new SessionInboxIntelligenceProvider();
  const draft = buildInboxIntelligenceRecord(analysis, "excerpt", REFERENCE_DATA, "2026-01-01T00:00:00.000Z");
  const result = await provider.create(draft);
  if (result.status !== "success") throw new Error("seed failed");
  return { provider, seeded: result.value };
}

function Harness({
  provider,
  initialRecords,
  openLogSpy,
}: {
  provider: SessionInboxIntelligenceProvider;
  initialRecords: InboxIntelligenceRecord[];
  openLogSpy?: (record?: WorkRecord, onSaved?: (saved: WorkRecord) => void) => void;
}) {
  const [records, setRecords] = useState<InboxIntelligenceRecord[]>(initialRecords);
  const saveRecord = async (record: InboxIntelligenceRecord) => {
    const result = await provider.create(record);
    if (result.status === "success") setRecords((current) => [result.value, ...current]);
    return result;
  };
  const updateRecord = async (record: InboxIntelligenceRecord, version: number) => {
    const result = await provider.update(record, version);
    if (result.status === "success") setRecords((current) => current.map((r) => (r.appId === result.value.appId ? result.value : r)));
    return result;
  };
  return (
    <InboxIntelligence
      references={REFERENCE_DATA}
      openLog={openLogSpy ?? vi.fn()}
      createDraftRecord={baseWorkRecord}
      records={records}
      saveRecord={saveRecord}
      updateRecord={updateRecord}
      saveContact={vi.fn()}
      updateContact={vi.fn()}
    />
  );
}

function rowFor(title: string): HTMLElement {
  return screen.getByText(title).closest(".inbox-row")!;
}

describe("Readability — Inbox-specific classes exist in app/globals.css", () => {
  it("defines .inbox-row-title, .inbox-row-summary, .inbox-row-meta, and .inbox-row-actions, scoped away from .record-row", () => {
    const css = readFileSync("app/globals.css", "utf-8");
    for (const selector of [".inbox-row-title", ".inbox-row-summary", ".inbox-row-meta", ".inbox-row-actions"]) {
      expect(css).toContain(selector);
    }
    // Scoped, not a change to the shared row class every other list (Work Records, etc.) still uses.
    expect(css).toContain(".record-row strong{font-size:11px}"); // unchanged from before this patch
  });
});

describe("Readability — Inbox-specific classes exist and are applied", () => {
  it("title, summary, metadata, and actions each render with their own scoped Inbox class", async () => {
    const { provider, seeded } = await seededProvider(analysisFixture());
    render(<Harness provider={provider} initialRecords={[seeded]} />);
    const row = rowFor("Original Title");
    expect(row.querySelector(".inbox-row-title")?.textContent).toBe("Original Title");
    expect(row.querySelector(".inbox-row-summary")?.textContent).toBe(seeded.analysis.summary);
    expect(row.querySelector(".inbox-row-meta")).toBeTruthy();
    expect(row.querySelector(".inbox-row-actions")).toBeTruthy();
    expect(within(row.querySelector(".inbox-row-actions")!).getByRole("button", { name: "Edit" })).toBeTruthy();
  });
});

describe("Edit — availability across every status", () => {
  it("an open record shows Edit", async () => {
    const { provider, seeded } = await seededProvider(analysisFixture());
    render(<Harness provider={provider} initialRecords={[seeded]} />);
    expect(within(rowFor("Original Title")).getByRole("button", { name: "Edit" })).toBeTruthy();
  });

  it("a waiting record shows Edit", async () => {
    const user = userEvent.setup();
    const { provider, seeded } = await seededProvider(analysisFixture());
    render(<Harness provider={provider} initialRecords={[seeded]} />);
    await user.click(within(rowFor("Original Title")).getByRole("button", { name: "Mark waiting" }));
    await waitFor(() => expect(screen.getByRole("heading", { name: "Waiting" }).closest("section")).toBeTruthy());
    const waitingSection = screen.getByRole("heading", { name: "Waiting" }).closest("section")!;
    expect(within(waitingSection).getByRole("button", { name: "Edit" })).toBeTruthy();
  });

  it("a resolved record shows Edit", async () => {
    const user = userEvent.setup();
    const { provider, seeded } = await seededProvider(analysisFixture());
    render(<Harness provider={provider} initialRecords={[seeded]} />);
    await user.click(within(rowFor("Original Title")).getByRole("button", { name: "Resolve" }));
    await waitFor(() => expect(screen.getByText("Recent / resolved").closest("section")).toBeTruthy());
    const resolvedSection = screen.getByText("Recent / resolved").closest("section")!;
    expect(within(resolvedSection).getByRole("button", { name: "Edit" })).toBeTruthy();
  });
});

describe("Edit — opening, Cancel, and Save Changes", () => {
  it("Edit opens prefilled with the record's current title/summary/priority/action item values", async () => {
    const user = userEvent.setup();
    const { provider, seeded } = await seededProvider(analysisFixture());
    render(<Harness provider={provider} initialRecords={[seeded]} />);
    await user.click(within(rowFor("Original Title")).getByRole("button", { name: "Edit" }));
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByDisplayValue("Original Title")).toBeTruthy();
    expect(within(dialog).getByDisplayValue(seeded.analysis.summary)).toBeTruthy();
    expect(within(dialog).getByDisplayValue("Send the agenda")).toBeTruthy();
    expect(within(dialog).getByDisplayValue("2026-02-01")).toBeTruthy();
    const prioritySelect = within(dialog).getByLabelText("Priority") as HTMLSelectElement;
    expect(prioritySelect.value).toBe("medium");
  });

  it("Cancel performs zero writes", async () => {
    const user = userEvent.setup();
    const { provider, seeded } = await seededProvider(analysisFixture());
    const updateSpy = vi.spyOn(provider, "update");
    const createSpy = vi.spyOn(provider, "create");
    render(<Harness provider={provider} initialRecords={[seeded]} />);
    await user.click(within(rowFor("Original Title")).getByRole("button", { name: "Edit" }));
    const dialog = await screen.findByRole("dialog");
    await user.clear(within(dialog).getByDisplayValue("Original Title"));
    await user.type(within(dialog).getByLabelText(/^Title/), "Should never be saved");
    await user.click(within(dialog).getByRole("button", { name: "Cancel" }));
    expect(screen.queryByRole("dialog")).toBeFalsy();
    expect(updateSpy).not.toHaveBeenCalled();
    expect(createSpy).not.toHaveBeenCalled();
    expect(screen.getByText("Original Title")).toBeTruthy();
  });

  it("Save Changes makes exactly one update() call and never a create() call", async () => {
    const user = userEvent.setup();
    const { provider, seeded } = await seededProvider(analysisFixture());
    const updateSpy = vi.spyOn(provider, "update");
    const createSpy = vi.spyOn(provider, "create");
    render(<Harness provider={provider} initialRecords={[seeded]} />);
    await user.click(within(rowFor("Original Title")).getByRole("button", { name: "Edit" }));
    const dialog = await screen.findByRole("dialog");
    await user.clear(within(dialog).getByLabelText(/^Title/));
    await user.type(within(dialog).getByLabelText(/^Title/), "Revised Title");
    await user.click(within(dialog).getByRole("button", { name: "Save Changes" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(updateSpy).toHaveBeenCalledTimes(1);
    expect(createSpy).not.toHaveBeenCalled();
    expect(screen.getByText("Revised Title")).toBeTruthy();
  });

  it("Save Changes does not call Anthropic and does not create a Work Record", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const openLogSpy = vi.fn();
    const { provider, seeded } = await seededProvider(analysisFixture());
    render(<Harness provider={provider} initialRecords={[seeded]} openLogSpy={openLogSpy} />);
    await user.click(within(rowFor("Original Title")).getByRole("button", { name: "Edit" }));
    const dialog = await screen.findByRole("dialog");
    await user.click(within(dialog).getByRole("button", { name: "Save Changes" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(fetchSpy).not.toHaveBeenCalled();
    expect(openLogSpy).not.toHaveBeenCalled();
  });

  it("editing a waiting record preserves the waiting status", async () => {
    const user = userEvent.setup();
    const { provider, seeded } = await seededProvider(analysisFixture());
    render(<Harness provider={provider} initialRecords={[seeded]} />);
    await user.click(within(rowFor("Original Title")).getByRole("button", { name: "Mark waiting" }));
    await waitFor(() => expect(screen.getByRole("heading", { name: "Waiting" }).closest("section")).toBeTruthy());
    await user.click(within(rowFor("Original Title")).getByRole("button", { name: "Edit" }));
    const dialog = await screen.findByRole("dialog");
    await user.clear(within(dialog).getByLabelText(/^Title/));
    await user.type(within(dialog).getByLabelText(/^Title/), "Still Waiting Title");
    await user.click(within(dialog).getByRole("button", { name: "Save Changes" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    const waitingSection = screen.getByRole("heading", { name: "Waiting" }).closest("section")!;
    expect(within(waitingSection).getByText("Still Waiting Title")).toBeTruthy();
    const resolvedSection = screen.getByText("Recent / resolved").closest("section")!;
    expect(within(resolvedSection).queryByText("Still Waiting Title")).toBeFalsy();
  });

  it("editing preserves the stable AppId and uses the record's current RecordVersion", async () => {
    const user = userEvent.setup();
    const { provider, seeded } = await seededProvider(analysisFixture());
    const updateSpy = vi.spyOn(provider, "update");
    render(<Harness provider={provider} initialRecords={[seeded]} />);
    await user.click(within(rowFor("Original Title")).getByRole("button", { name: "Edit" }));
    const dialog = await screen.findByRole("dialog");
    await user.click(within(dialog).getByRole("button", { name: "Save Changes" }));
    await waitFor(() => expect(updateSpy).toHaveBeenCalledTimes(1));
    const [sentRecord, sentVersion] = updateSpy.mock.calls[0];
    expect(sentRecord.appId).toBe(seeded.appId);
    expect(sentVersion).toBe(seeded.metadata.version);
  });
});

describe("Reopen — existing behavior, now with focused regression coverage", () => {
  it("a resolved record shows Reopen; an open record does not", async () => {
    const user = userEvent.setup();
    const { provider, seeded } = await seededProvider(analysisFixture());
    render(<Harness provider={provider} initialRecords={[seeded]} />);
    expect(within(rowFor("Original Title")).queryByRole("button", { name: "Reopen" })).toBeFalsy();
    await user.click(within(rowFor("Original Title")).getByRole("button", { name: "Resolve" }));
    await waitFor(() => expect(screen.getByText("Recent / resolved").closest("section")).toBeTruthy());
    const resolvedSection = screen.getByText("Recent / resolved").closest("section")!;
    expect(within(resolvedSection).getByRole("button", { name: "Reopen" })).toBeTruthy();
  });

  it("Reopen makes exactly one update() call, changes status back to open, and creates nothing", async () => {
    const user = userEvent.setup();
    const openLogSpy = vi.fn();
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const { provider, seeded } = await seededProvider(analysisFixture());
    render(<Harness provider={provider} initialRecords={[seeded]} openLogSpy={openLogSpy} />);
    await user.click(within(rowFor("Original Title")).getByRole("button", { name: "Resolve" }));
    await waitFor(() => expect(screen.getByText("Recent / resolved").closest("section")).toBeTruthy());

    const updateSpy = vi.spyOn(provider, "update");
    const createSpy = vi.spyOn(provider, "create");
    await user.click(within(screen.getByText("Recent / resolved").closest("section")!).getByRole("button", { name: "Reopen" }));
    await waitFor(() => expect(screen.getByText("Needs attention").closest("section")).toBeTruthy());

    expect(updateSpy).toHaveBeenCalledTimes(1);
    expect(createSpy).not.toHaveBeenCalled();
    const needsAttentionSection = screen.getByText("Needs attention").closest("section")!;
    expect(within(needsAttentionSection).getByText("Original Title")).toBeTruthy();
    expect(openLogSpy).not.toHaveBeenCalled();
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("Reopen preserves content and the stable AppId", async () => {
    const user = userEvent.setup();
    const { provider, seeded } = await seededProvider(analysisFixture());
    render(<Harness provider={provider} initialRecords={[seeded]} />);
    await user.click(within(rowFor("Original Title")).getByRole("button", { name: "Resolve" }));
    await waitFor(() => expect(screen.getByText("Recent / resolved").closest("section")).toBeTruthy());

    const updateSpy = vi.spyOn(provider, "update");
    await user.click(within(screen.getByText("Recent / resolved").closest("section")!).getByRole("button", { name: "Reopen" }));
    await waitFor(() => expect(updateSpy).toHaveBeenCalledTimes(1));
    const [sentRecord] = updateSpy.mock.calls[0];
    expect(sentRecord.appId).toBe(seeded.appId);
    expect(sentRecord.status).toBe("open");
    expect(sentRecord.analysis.suggestedWorkRecord.title).toBe("Original Title");
    expect(sentRecord.analysis.summary).toBe(seeded.analysis.summary);

    const needsAttentionSection = screen.getByText("Needs attention").closest("section")!;
    expect(within(needsAttentionSection).getByText("Original Title")).toBeTruthy();
  });
});
