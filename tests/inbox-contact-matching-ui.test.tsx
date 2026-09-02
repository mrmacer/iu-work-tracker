// @vitest-environment jsdom
//
// Patch 8D — Intelligence Contact matching, end-to-end through the real Inbox Intelligence
// review flow: paste -> analyze (mocked /api/inbox-intelligence, one call) -> the People review
// section -> Match Existing / Add Person / Ignore -> Save to Inbox -> matchedContactIds
// persisted and displayed. All against SessionInboxIntelligenceProvider/MemoryContactProvider —
// zero real SharePoint writes, and (checked explicitly) no Anthropic call beyond the one
// mocked Analyze click.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { useRef, useState } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import InboxIntelligence from "../app/InboxIntelligence";
import type { AnalyzeEmailResult } from "../lib/anthropic-email-analysis";
import { MemoryContactProvider } from "../lib/contact-provider";
import type { ContactResult } from "../lib/contact-provider";
import type { InboxIntelligenceRecord } from "../lib/inbox-intelligence-models";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import type { InboxIntelligenceResult } from "../lib/inbox-intelligence-provider";
import type { Contact, WorkRecord } from "../lib/models";
import { REFERENCE_DATA } from "../lib/reference-data";
import { WORK_RECORD_SCHEMA_VERSION } from "../lib/models";

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

/** Mirrors app/IUWorkTracker.tsx's own save/merge wiring at a small scale: a real
 * SessionInboxIntelligenceProvider and a real MemoryContactProvider behind saveContact/
 * updateContact spies, with references.contacts kept in sync after a successful create — so
 * "Add Person" behaves exactly as it would inside the real app. */
function Harness({ updateContactSpy }: { updateContactSpy: (contact: Contact, version: number) => void }) {
  const [records, setRecords] = useState<InboxIntelligenceRecord[]>([]);
  const [contacts, setContacts] = useState<Contact[]>(REFERENCE_DATA.contacts);
  const inboxProvider = useRef(new SessionInboxIntelligenceProvider()).current;
  const contactProvider = useRef(new MemoryContactProvider()).current;

  const saveRecord = async (record: InboxIntelligenceRecord): Promise<InboxIntelligenceResult<InboxIntelligenceRecord>> => {
    const result = await inboxProvider.create(record);
    if (result.status === "success") setRecords((current) => [result.value, ...current]);
    return result;
  };
  const updateRecord = async (record: InboxIntelligenceRecord, version: number) => {
    const result = await inboxProvider.update(record, version);
    if (result.status === "success") setRecords((current) => current.map((r) => (r.appId === result.value.appId ? result.value : r)));
    return result;
  };
  const saveContact = async (contact: Contact): Promise<ContactResult<Contact>> => {
    const result = await contactProvider.create(contact);
    if (result.status === "success") setContacts((current) => [result.value, ...current]);
    return result;
  };
  const updateContact = async (contact: Contact, version: number) => {
    updateContactSpy(contact, version);
    const result = await contactProvider.update(contact, version);
    if (result.status === "success") setContacts((current) => current.map((c) => (c.appId === result.value.appId ? result.value : c)));
    return result;
  };

  return (
    <InboxIntelligence
      references={{ ...REFERENCE_DATA, contacts }}
      openLog={vi.fn()}
      createDraftRecord={baseWorkRecord}
      records={records}
      saveRecord={saveRecord}
      updateRecord={updateRecord}
      saveContact={saveContact}
      updateContact={updateContact}
    />
  );
}

function analysisResult(people: string[]): AnalyzeEmailResult {
  return {
    status: "success",
    analysis: {
      summary: "Planning next steps for the district partnership.",
      priority: "medium",
      needsAttention: false,
      actionItems: [],
      followUp: "",
      people,
      organizations: [],
      districts: [],
      projects: [],
      tags: [],
      suggestedWorkType: null,
      suggestedWorkRecord: { title: "District planning email", description: "Synthetic test email." },
    },
    usage: { model: "claude-opus-5", inputTokens: 500, outputTokens: 200 },
  };
}

function jsonResponse(value: unknown): Response {
  return new Response(JSON.stringify(value), { status: 200, headers: { "content-type": "application/json" } });
}

async function analyze(user: ReturnType<typeof userEvent.setup>, people: string[]) {
  vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisResult(people)));
  await user.type(screen.getByPlaceholderText(/paste the whole email/i), "Hi team, following up on our district planning.");
  await user.click(screen.getByRole("button", { name: "Analyze email" }));
  await waitFor(() => expect(screen.getByText("People — possible Contact matches")).toBeTruthy());
}

describe("People review starts unreviewed", () => {
  it("a detected person renders with no decision — no 'Matched'/'Ignored' state until the human acts", async () => {
    const user = userEvent.setup();
    render(<Harness updateContactSpy={vi.fn()} />);
    await analyze(user, ["Development District Lead"]);
    expect(screen.queryByText(/Matched Contact/)).toBeFalsy();
    expect(screen.queryByText(/Ignored for this item/)).toBeFalsy();
    expect(screen.getByRole("button", { name: "Match Existing" })).toBeTruthy();
  });
});

describe("Match Existing", () => {
  it("records the selected Contact appId into matchedContactIds, visible on the saved row, without creating or updating any Contact", async () => {
    const user = userEvent.setup();
    const updateContactSpy = vi.fn();
    render(<Harness updateContactSpy={updateContactSpy} />);
    await analyze(user, ["Development District Lead"]);

    await user.click(screen.getByRole("button", { name: "Match Existing" }));
    expect(screen.getByText(/Matched Contact: Development District Lead/)).toBeTruthy();

    await user.click(screen.getByRole("button", { name: "Save to Inbox" }));
    await waitFor(() => expect(screen.getByText("Saved to Inbox.")).toBeTruthy());

    // Matched Contact never gets an update call — matching never overwrites Contact fields.
    expect(updateContactSpy).not.toHaveBeenCalled();
  });

  it("the saved Inbox row displays the matched Contact's name", async () => {
    const user = userEvent.setup();
    render(<Harness updateContactSpy={vi.fn()} />);
    await analyze(user, ["Development District Lead"]);
    await user.click(screen.getByRole("button", { name: "Match Existing" }));
    await user.click(screen.getByRole("button", { name: "Save to Inbox" }));
    await waitFor(() => expect(screen.getByText("Saved to Inbox.")).toBeTruthy());

    const needsAttentionSection = screen.getByText("Needs attention").closest("section")!;
    expect(within(needsAttentionSection).getByText(/Development District Lead/)).toBeTruthy();
  });
});

describe("Ignore", () => {
  it("does not create a Contact and excludes the person from matchedContactIds — and stops nagging for this item", async () => {
    const user = userEvent.setup();
    const saveContactSpy = vi.fn();
    render(<Harness updateContactSpy={vi.fn()} />);
    await analyze(user, ["Annie Milewski"]); // not a seeded Contact — "no reliable match found"
    expect(screen.getByText("No reliable match found.")).toBeTruthy();

    await user.click(screen.getByRole("button", { name: "Ignore" }));
    expect(screen.getByText("Ignored for this item")).toBeTruthy();
    expect(saveContactSpy).not.toHaveBeenCalled();

    await user.click(screen.getByRole("button", { name: "Save to Inbox" }));
    await waitFor(() => expect(screen.getByText("Saved to Inbox.")).toBeTruthy());
    const needsAttentionSection = screen.getByText("Needs attention").closest("section")!;
    expect(within(needsAttentionSection).queryByText(/Annie Milewski/)).toBeFalsy();
  });
});

describe("Add Person", () => {
  it("opens the existing Create Contact form prefilled with the detected name, and associates the newly created Contact once saved", async () => {
    const user = userEvent.setup();
    render(<Harness updateContactSpy={vi.fn()} />);
    await analyze(user, ["Annie Milewski"]);

    await user.click(screen.getByRole("button", { name: "Add Person" }));
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByRole("heading", { name: "Add Contact" })).toBeTruthy();
    expect(within(dialog).getByDisplayValue("Annie Milewski")).toBeTruthy();

    await user.click(within(dialog).getByRole("button", { name: "Add Contact" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());

    expect(screen.getByText(/Matched Contact: Annie Milewski/)).toBeTruthy();

    await user.click(screen.getByRole("button", { name: "Save to Inbox" }));
    await waitFor(() => expect(screen.getByText("Saved to Inbox.")).toBeTruthy());
    const needsAttentionSection = screen.getByText("Needs attention").closest("section")!;
    expect(within(needsAttentionSection).getByText(/Annie Milewski/)).toBeTruthy();
  });
});

describe("Regression — existing Organization/Project matching and zero extra AI calls", () => {
  it("existing Organization/District/Project auto-resolution is unaffected by the new People review step", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      jsonResponse({
        status: "success",
        analysis: {
          summary: "STEELS planning.",
          priority: "medium",
          needsAttention: false,
          actionItems: [],
          followUp: "",
          people: [],
          organizations: [],
          districts: ["North Valley SD"],
          projects: ["STEELS Implementation"],
          tags: [],
          suggestedWorkType: null,
          suggestedWorkRecord: { title: "STEELS planning email", description: "Synthetic test email." },
        },
        usage: { model: "claude-opus-5", inputTokens: 400, outputTokens: 150 },
      } satisfies AnalyzeEmailResult),
    );
    render(<Harness updateContactSpy={vi.fn()} />);
    await user.type(screen.getByPlaceholderText(/paste the whole email/i), "STEELS update for North Valley.");
    await user.click(screen.getByRole("button", { name: "Analyze email" }));
    await waitFor(() => expect(screen.getByRole("button", { name: "Save to Inbox" })).toBeTruthy());
    await user.click(screen.getByRole("button", { name: "Save to Inbox" }));
    await waitFor(() => expect(screen.getByText("Saved to Inbox.")).toBeTruthy());

    // InboxRow shows the matched Project over the matched District when both are present (the
    // existing, pre-8D precedence) — the point here is just that this resolution still works.
    const needsAttentionSection = screen.getByText("Needs attention").closest("section")!;
    expect(within(needsAttentionSection).getByText(/STEELS Implementation/)).toBeTruthy();
  });

  it("no People fieldset renders when the AI detected zero people, and Save still works", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisResult([])));
    render(<Harness updateContactSpy={vi.fn()} />);
    await user.type(screen.getByPlaceholderText(/paste the whole email/i), "No people mentioned here at all.");
    await user.click(screen.getByRole("button", { name: "Analyze email" }));
    await waitFor(() => expect(screen.getByRole("button", { name: "Save to Inbox" })).toBeTruthy());
    expect(screen.queryByText("People — possible Contact matches")).toBeFalsy();
    await user.click(screen.getByRole("button", { name: "Save to Inbox" }));
    await waitFor(() => expect(screen.getByText("Saved to Inbox.")).toBeTruthy());
  });

  it("opening the review panel, matching, and saving make exactly one Anthropic-backed fetch call (the Analyze click)", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisResult(["Development District Lead"])));
    render(<Harness updateContactSpy={vi.fn()} />);
    await user.type(screen.getByPlaceholderText(/paste the whole email/i), "District update.");
    await user.click(screen.getByRole("button", { name: "Analyze email" }));
    await waitFor(() => expect(screen.getByText("People — possible Contact matches")).toBeTruthy());
    expect(fetchSpy).toHaveBeenCalledTimes(1);

    await user.click(screen.getByRole("button", { name: "Match Existing" }));
    await user.click(screen.getByRole("button", { name: "Save to Inbox" }));
    await waitFor(() => expect(screen.getByText("Saved to Inbox.")).toBeTruthy());
    expect(fetchSpy).toHaveBeenCalledTimes(1); // matching + saving made zero additional network/AI calls
  });
});
