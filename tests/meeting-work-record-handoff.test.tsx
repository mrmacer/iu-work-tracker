// @vitest-environment jsdom
//
// Covers the Meeting Notes "Log as work" handoff: eligibility, current-edited-state prefill,
// zero persistence on the click itself, and (as a full integration proof) that the SAME
// existing Log Work wizard and save path Patch 5 (Voice Intelligence) already uses is what
// actually persists the record — no second form, no second save path, no new mapper risk to
// lib/voice-intelligence-work-record.ts's tested buildWorkRecordDraftFromVoiceCandidate.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import MeetingNotes from "../app/MeetingNotes";
import IUWorkTracker from "../app/IUWorkTracker";
import { MemoryDataProvider } from "../lib/data-provider";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import type { AnalyzeMeetingResult } from "../lib/anthropic-meeting-analysis";
import { MEETING_CANDIDATE_TYPES, type MeetingCandidateType } from "../lib/meeting-intelligence-models";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

function baseWorkRecord(): WorkRecord {
  return {
    appId: "draft", title: "", activityDate: "2026-09-01", activityType: "", description: "", detailedNotes: "",
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

function analysisWith(type: MeetingCandidateType): AnalyzeMeetingResult {
  return {
    status: "success",
    analysis: {
      candidates: [
        {
          type,
          title: "Reviewed the STEELS grant budget section",
          detail: "Walked through the budget together.",
          sourceExcerpt: "Spent about 30 minutes walking through the budget section together.",
          ownerText: null,
          dueText: null,
          durationText: "about 30 minutes",
        },
      ],
    },
    usage: { model: "claude-opus-5", inputTokens: 900, outputTokens: 420 },
  };
}

async function analyzeInto(user: ReturnType<typeof userEvent.setup>) {
  await user.type(screen.getByPlaceholderText(/take notes during the meeting/i), "Spent about 30 minutes on the budget.");
  await user.click(screen.getByRole("button", { name: "Analyze Meeting" }));
  await waitFor(() => expect(screen.getByText("1 candidate")).toBeTruthy());
}

describe("Meeting Notes — Log as work eligibility", () => {
  it.each(MEETING_CANDIDATE_TYPES)("candidate type %s", async (type) => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisWith(type)));
    render(<MeetingNotes openLog={vi.fn()} createDraftRecord={baseWorkRecord} />);
    await analyzeInto(user);

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
    render(<MeetingNotes openLog={vi.fn()} createDraftRecord={baseWorkRecord} />);
    await analyzeInto(user);
    expect(screen.getByRole("button", { name: "Log as work" })).toBeTruthy();

    await user.selectOptions(screen.getByLabelText("Candidate type"), "ACTION");
    expect(screen.queryByRole("button", { name: "Log as work" })).toBeFalsy();
  });

  it("adds Log as work when the current candidate is changed to COMPLETED_WORK", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisWith("ACTION")));
    render(<MeetingNotes openLog={vi.fn()} createDraftRecord={baseWorkRecord} />);
    await analyzeInto(user);
    expect(screen.queryByRole("button", { name: "Log as work" })).toBeFalsy();

    await user.selectOptions(screen.getByLabelText("Candidate type"), "COMPLETED_WORK");
    expect(screen.getByRole("button", { name: "Log as work" })).toBeTruthy();
  });

  it("DECISION candidates never show Log as work", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisWith("DECISION")));
    render(<MeetingNotes openLog={vi.fn()} createDraftRecord={baseWorkRecord} />);
    await analyzeInto(user);
    expect(screen.queryByRole("button", { name: "Log as work" })).toBeFalsy();
  });
});

describe("Meeting Notes — Log as work uses current edited state, performs zero persistence", () => {
  async function renderInReviewWithCompletedWork() {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisWith("COMPLETED_WORK")));
    const openLog = vi.fn();
    render(<MeetingNotes openLog={openLog} createDraftRecord={baseWorkRecord} />);
    await analyzeInto(user);
    return { user, fetchSpy, openLog };
  }

  it("passes the human-edited title and detail, not the original AI output", async () => {
    const { user, openLog } = await renderInReviewWithCompletedWork();
    const titleInput = screen.getByDisplayValue("Reviewed the STEELS grant budget section");
    await user.clear(titleInput);
    await user.type(titleInput, "Human-edited title");
    const detailInput = screen.getByDisplayValue("Walked through the budget together.");
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
    await user.click(screen.getByRole("button", { name: "Remove duration" })); // clears "about 30 minutes"
    await user.click(screen.getByRole("button", { name: "Log as work" }));
    const [draft] = openLog.mock.calls[0];
    expect(draft.durationMinutes).toBe(60); // base default from createDraftRecord(), unchanged
  });

  it("never persists anything, calls no provider, and makes no network/storage calls", async () => {
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
    expect(screen.getByDisplayValue("Reviewed the STEELS grant budget section")).toBeTruthy();
    expect(screen.getByRole("button", { name: "Log as work" })).toBeTruthy();
  });

  it("does not map category, project, district, or ORBIT", async () => {
    const { user, openLog } = await renderInReviewWithCompletedWork();
    await user.click(screen.getByRole("button", { name: "Log as work" }));
    const [draft] = openLog.mock.calls[0];
    expect(draft.activityType).toBe("");
    expect(draft.projectIds).toEqual([]);
    expect(draft.organizationIds).toEqual([]);
    expect(draft.orbit.reportable).toBe(false);
  });
});

describe("Meeting Notes — full integration through the existing Work Record form and save path", () => {
  it("opens the existing wizard prefilled, and exactly one normal createWorkRecord call happens on explicit Save", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(analysisWith("COMPLETED_WORK")));
    const dataProvider = new MemoryDataProvider([]);
    const createSpy = vi.spyOn(dataProvider, "createWorkRecord");
    const inboxDataProvider = new SessionInboxIntelligenceProvider();

    render(<IUWorkTracker dataProvider={dataProvider} inboxDataProvider={inboxDataProvider} />);

    const nav = await screen.findByRole("navigation");
    await user.click(within(nav).getByRole("button", { name: /meeting notes/i }));
    await analyzeInto(user);

    expect(createSpy).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "Log as work" }));

    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByDisplayValue("Reviewed the STEELS grant budget section")).toBeTruthy();
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
    expect(savedRecord.title).toBe("Reviewed the STEELS grant budget section");
    expect(savedRecord.activityType).toBe(firstRealOption);
  });
});
