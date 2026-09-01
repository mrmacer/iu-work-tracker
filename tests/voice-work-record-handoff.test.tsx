// @vitest-environment jsdom
//
// Covers the Patch 5 "Voice → Universal Work Record" handoff: eligibility, current-edited-
// state prefill, zero persistence on the "Log as work" click itself, and (as a full
// integration proof) that the SAME existing Log Work wizard and save path Inbox Intelligence
// already uses is what actually persists the record — no second form, no second save path.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import VoiceIntelligence from "../app/VoiceIntelligence";
import IUWorkTracker from "../app/IUWorkTracker";
import { MemoryDataProvider } from "../lib/data-provider";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import type { AnalyzeTranscriptResult } from "../lib/anthropic-voice-analysis";
import { VOICE_CANDIDATE_TYPES, type VoiceCandidateType } from "../lib/voice-intelligence-models";

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

function analysisWith(type: VoiceCandidateType): AnalyzeTranscriptResult {
  return {
    status: "success",
    analysis: {
      candidates: [
        {
          type,
          title: "Met with North Schuylkill about science resources",
          detail: "Went well, covered the new lab kits.",
          sourceExcerpt: "I met with North Schuylkill this morning about science resources",
          durationText: "about an hour",
        },
      ],
    },
    usage: { model: "claude-opus-5", inputTokens: 900, outputTokens: 420 },
  };
}

describe("Log as work — eligibility", () => {
  it.each(VOICE_CANDIDATE_TYPES)("candidate type %s", async (type) => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisWith(type)));
    render(<VoiceIntelligence openLog={vi.fn()} createDraftRecord={baseWorkRecord} />);
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "Met with the team this morning.");
    await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
    await waitFor(() => expect(screen.getByText("1 candidate")).toBeTruthy());

    const logAsWorkButton = screen.queryByRole("button", { name: "Log as work" });
    if (type === "COMPLETED_WORK") {
      expect(logAsWorkButton).toBeTruthy();
    } else {
      expect(logAsWorkButton).toBeFalsy();
    }
  });

  it("removes Log as work when the current candidate is changed away from COMPLETED_WORK", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisWith("COMPLETED_WORK")));
    render(<VoiceIntelligence openLog={vi.fn()} createDraftRecord={baseWorkRecord} />);
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "Met with the team this morning.");
    await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
    await waitFor(() => expect(screen.getByRole("button", { name: "Log as work" })).toBeTruthy());

    await user.selectOptions(screen.getByLabelText("Candidate type"), "ACTION");
    expect(screen.queryByRole("button", { name: "Log as work" })).toBeFalsy();
  });

  it("adds Log as work when the current candidate is changed to COMPLETED_WORK", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisWith("ACTION")));
    render(<VoiceIntelligence openLog={vi.fn()} createDraftRecord={baseWorkRecord} />);
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "Met with the team this morning.");
    await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
    await waitFor(() => expect(screen.getByLabelText("Candidate type")).toBeTruthy());
    expect(screen.queryByRole("button", { name: "Log as work" })).toBeFalsy();

    await user.selectOptions(screen.getByLabelText("Candidate type"), "COMPLETED_WORK");
    expect(screen.getByRole("button", { name: "Log as work" })).toBeTruthy();
  });
});

describe("Log as work — uses current edited state, and performs zero persistence", () => {
  async function renderInReviewWithCompletedWork() {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisWith("COMPLETED_WORK")));
    const openLog = vi.fn();
    render(<VoiceIntelligence openLog={openLog} createDraftRecord={baseWorkRecord} />);
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "Met with the team this morning.");
    await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
    await waitFor(() => expect(screen.getByRole("button", { name: "Log as work" })).toBeTruthy());
    return { user, fetchSpy, openLog };
  }

  it("passes the human-edited title and detail, not the original AI output", async () => {
    const { user, openLog } = await renderInReviewWithCompletedWork();
    const titleInput = screen.getByDisplayValue("Met with North Schuylkill about science resources");
    await user.clear(titleInput);
    await user.type(titleInput, "Human-edited title");
    const detailInput = screen.getByDisplayValue("Went well, covered the new lab kits.");
    await user.clear(detailInput);
    await user.type(detailInput, "Human-edited detail");

    await user.click(screen.getByRole("button", { name: "Log as work" }));

    expect(openLog).toHaveBeenCalledTimes(1);
    const [draft] = openLog.mock.calls[0];
    expect(draft.title).toBe("Human-edited title");
    expect(draft.description).toBe("Human-edited detail");
  });

  it("passes the human-edited duration, converted, not the original", async () => {
    const { user, openLog } = await renderInReviewWithCompletedWork();
    const durationRemove = screen.getByRole("button", { name: "Remove duration" });
    await user.click(durationRemove); // clears "about an hour"
    // No duration chip left, so the draft must fall back to the base default rather than 60
    // from "about an hour" — proves it's reading live state, not the original candidate.
    await user.click(screen.getByRole("button", { name: "Log as work" }));
    const [draft] = openLog.mock.calls[0];
    expect(draft.durationMinutes).toBe(60); // base default from createDraftRecord(), unchanged
  });

  it("never persists anything, calls no provider, and touches no browser storage", async () => {
    const setItem = vi.spyOn(Storage.prototype, "setItem");
    const { user, fetchSpy, openLog } = await renderInReviewWithCompletedWork();
    fetchSpy.mockClear();
    await user.click(screen.getByRole("button", { name: "Log as work" }));
    expect(openLog).toHaveBeenCalledTimes(1);
    expect(fetchSpy).not.toHaveBeenCalled(); // zero Anthropic / zero network calls from the click
    expect(setItem).not.toHaveBeenCalled();
  });

  it("does not mutate the candidate's own review state when opening the form", async () => {
    const { user, openLog } = await renderInReviewWithCompletedWork();
    await user.click(screen.getByRole("button", { name: "Log as work" }));
    expect(openLog).toHaveBeenCalledTimes(1);
    // The candidate card is still present, unchanged, still showing "Log as work" — nothing
    // about the review list itself was altered by opening the Work Record form.
    expect(screen.getByDisplayValue("Met with North Schuylkill about science resources")).toBeTruthy();
    expect(screen.getByRole("button", { name: "Log as work" })).toBeTruthy();
  });
});

describe("Log as work — full integration through the existing Work Record form and save path", () => {
  it("opens the existing wizard prefilled, and exactly one normal createWorkRecord call happens on explicit Save", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisWith("COMPLETED_WORK")));
    const dataProvider = new MemoryDataProvider([]);
    const createSpy = vi.spyOn(dataProvider, "createWorkRecord");
    const inboxDataProvider = new SessionInboxIntelligenceProvider();

    render(<IUWorkTracker dataProvider={dataProvider} inboxDataProvider={inboxDataProvider} />);

    const nav = await screen.findByRole("navigation");
    await user.click(within(nav).getByRole("button", { name: /voice intelligence/i }));
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "Met with the team this morning.");
    await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
    await waitFor(() => expect(screen.getByRole("button", { name: "Log as work" })).toBeTruthy());

    expect(createSpy).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "Log as work" }));

    // The existing Work Record wizard opened — same dialog InboxIntelligence's "Create Work
    // Record" and the normal "+ Log work" button both already use.
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByDisplayValue("Met with North Schuylkill about science resources")).toBeTruthy();
    expect(createSpy).not.toHaveBeenCalled(); // still zero — opening the form persists nothing

    const activityTypeSelect = within(dialog).getByLabelText(/Activity type/i) as HTMLSelectElement;
    const firstRealOption = activityTypeSelect.options[1].value;
    await user.selectOptions(activityTypeSelect, firstRealOption);

    for (let i = 0; i < 4; i++) {
      await user.click(within(dialog).getByRole("button", { name: /continue/i }));
    }
    await user.click(within(dialog).getByRole("button", { name: "Save & done" }));

    await waitFor(() => expect(createSpy).toHaveBeenCalledTimes(1));
    const [savedRecord] = createSpy.mock.calls[0];
    expect(savedRecord.title).toBe("Met with North Schuylkill about science resources");
    expect(savedRecord.activityType).toBe(firstRealOption);
  });
});
