// @vitest-environment jsdom
//
// Patch 8D — Voice Intelligence's PERSON-candidate Contact matching. Voice has no durable
// persistence at all (see app/VoiceIntelligence.tsx's own doc comment), so this only proves the
// shared review UI/logic works correctly here — never that anything is saved beyond the
// Contact itself when "Add Person" is used (which reuses the real, durable Contact creation
// path). Zero AI calls beyond the one mocked Analyze click.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { useRef, useState } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import VoiceIntelligence from "../app/VoiceIntelligence";
import type { AnalyzeTranscriptResult } from "../lib/anthropic-voice-analysis";
import { MemoryContactProvider } from "../lib/contact-provider";
import { WORK_RECORD_SCHEMA_VERSION, type Contact, type WorkRecord } from "../lib/models";
import { REFERENCE_DATA } from "../lib/reference-data";

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

function jsonResponse(value: unknown): Response {
  return new Response(JSON.stringify(value), { status: 200, headers: { "content-type": "application/json" } });
}

function personAnalysis(title: string): AnalyzeTranscriptResult {
  return {
    status: "success",
    analysis: {
      candidates: [
        {
          type: "PERSON",
          title,
          detail: "Mentioned during the call.",
          sourceExcerpt: `I was talking with ${title} about the project`,
          durationText: null,
        },
      ],
    },
    usage: { model: "claude-opus-5", inputTokens: 300, outputTokens: 100 },
  };
}

function renderVoice() {
  return render(
    <VoiceIntelligence
      openLog={vi.fn()}
      createDraftRecord={baseWorkRecord}
      references={REFERENCE_DATA}
      saveContact={vi.fn()}
      updateContact={vi.fn()}
    />,
  );
}

async function analyze(user: ReturnType<typeof userEvent.setup>, title: string) {
  vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(personAnalysis(title)));
  await user.type(screen.getByPlaceholderText(/paste the transcript/i), "Talked with someone about the project today.");
  await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
  await waitFor(() => expect(screen.getByText("1 candidate")).toBeTruthy());
}

describe("Voice PERSON candidates get the shared Contact match panel", () => {
  it("a PERSON candidate renders the match panel; a non-PERSON candidate (e.g. ACTION) does not", async () => {
    const user = userEvent.setup();
    renderVoice();
    await analyze(user, "Development District Lead");
    expect(screen.getByRole("button", { name: "Match Existing" })).toBeTruthy();

    await user.selectOptions(screen.getByLabelText("Candidate type"), "ACTION");
    expect(screen.queryByRole("button", { name: "Match Existing" })).toBeFalsy();
  });

  it("Match Existing shows the matched state; it is never pre-selected", async () => {
    const user = userEvent.setup();
    renderVoice();
    await analyze(user, "Development District Lead");
    expect(screen.queryByText(/Matched Contact/)).toBeFalsy();
    await user.click(screen.getByRole("button", { name: "Match Existing" }));
    expect(screen.getByText(/Matched Contact: Development District Lead/)).toBeTruthy();
  });

  it("Ignore marks the candidate ignored without creating a Contact", async () => {
    const user = userEvent.setup();
    const saveContact = vi.fn();
    render(
      <VoiceIntelligence openLog={vi.fn()} createDraftRecord={baseWorkRecord} references={REFERENCE_DATA} saveContact={saveContact} updateContact={vi.fn()} />,
    );
    await analyze(user, "Someone Unmatched");
    expect(screen.getByText("No reliable match found.")).toBeTruthy();
    await user.click(screen.getByRole("button", { name: "Ignore" }));
    expect(screen.getByText("Ignored for this item")).toBeTruthy();
    expect(saveContact).not.toHaveBeenCalled();
  });

  it("Add Person opens the real Create Contact form prefilled with the candidate's title, and associates it once saved", async () => {
    const user = userEvent.setup();

    // Mirrors app/IUWorkTracker.tsx's own reactive merge: after a successful saveContact, the
    // newly created Contact must appear in `references.contacts` on the NEXT render — a real
    // requirement for ContactMatchPanel's matched-state lookup, not just this test's plumbing.
    function Harness() {
      const [contacts, setContacts] = useState<Contact[]>(REFERENCE_DATA.contacts);
      const contactProvider = useRef(new MemoryContactProvider()).current;
      return (
        <VoiceIntelligence
          openLog={vi.fn()}
          createDraftRecord={baseWorkRecord}
          references={{ ...REFERENCE_DATA, contacts }}
          saveContact={async (contact) => {
            const result = await contactProvider.create(contact);
            if (result.status === "success") setContacts((current) => [result.value, ...current]);
            return result;
          }}
          updateContact={(contact, version) => contactProvider.update(contact, version)}
        />
      );
    }
    render(<Harness />);
    await analyze(user, "Someone New");
    await user.click(screen.getByRole("button", { name: "Add Person" }));
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByRole("heading", { name: "Add Contact" })).toBeTruthy();
    expect(within(dialog).getByDisplayValue("Someone New")).toBeTruthy();
    await user.click(within(dialog).getByRole("button", { name: "Add Contact" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(screen.getByText(/Matched Contact: Someone New/)).toBeTruthy();
  });

  it("matching, ignoring, and reviewing a PERSON candidate make zero additional network/AI calls beyond the one Analyze click", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(personAnalysis("Development District Lead")));
    renderVoice();
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "Talked with the team today.");
    await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
    await waitFor(() => expect(screen.getByText("1 candidate")).toBeTruthy());
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    await user.click(screen.getByRole("button", { name: "Match Existing" }));
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });
});
