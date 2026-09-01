// @vitest-environment jsdom
//
// Patch 6B durability coverage: Save Meeting create-vs-update discipline, zero-autosave, Reopen
// full-state reconstruction with zero AI call, "human review is authoritative" (edited/
// deselected/retyped candidate state survives save+reload, never the original AI output),
// pre-analysis save (agenda/notes only, no fabricated summary), and dirty-state tracking
// including the New Meeting / Re-analyze confirmation gates. All against a mocked
// saveRecord/updateRecord — zero real SharePoint writes, zero real Anthropic calls (fetch is
// only ever stubbed for the explicit Analyze click, and every test below asserts it is NOT
// called for any other action).
import { cleanup, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import MeetingNotes from "../app/MeetingNotes";
import type { AnalyzeMeetingResult } from "../lib/anthropic-meeting-analysis";
import {
  MEETING_RECORD_SCHEMA_VERSION,
  type MeetingRecord,
  type ReviewedMeetingCandidate,
} from "../lib/meeting-intelligence-models";
import type { MeetingRecordResult } from "../lib/meeting-record-provider";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
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

function candidate(overrides: Partial<ReviewedMeetingCandidate> = {}): ReviewedMeetingCandidate {
  return {
    type: "ACTION",
    title: "Call the district about the venue",
    detail: "Confirm venue details.",
    sourceExcerpt: "Annie will call the district about the venue by Friday.",
    ownerText: "Annie",
    dueText: "Friday",
    durationText: null,
    selected: true,
    ...overrides,
  };
}

function savedMeeting(overrides: Partial<MeetingRecord> = {}): MeetingRecord {
  return {
    appId: "saved-meeting-1",
    schemaVersion: MEETING_RECORD_SCHEMA_VERSION,
    title: "STEELS quarterly planning",
    meetingDate: "2026-08-15",
    meetingType: "District Meeting",
    attendeesText: "Greg, Annie, Kim",
    agendaText: "Review grant budget. Discuss fall meeting timing.",
    notesText: "Walked through the budget together. Annie will call the district.",
    reviewedCandidates: [
      candidate(),
      candidate({ type: "DECISION", title: "Move the fall network meeting to October", detail: "Agreed after discussion.", ownerText: null, dueText: null, sourceExcerpt: "" }),
      candidate({ type: "IDEA", title: "Consider a shared calendar", detail: "Might help scheduling.", ownerText: null, dueText: null, sourceExcerpt: "", selected: false }),
    ],
    minutesText: "STEELS quarterly planning\n2026-08-15 · District Meeting\n\nDecisions\n- Move the fall network meeting to October",
    analysisModel: "claude-opus-5",
    analyzedAt: "2026-08-15T18:00:00.000Z",
    metadata: { providerId: "sp-1", version: 1, createdAt: "2026-08-15T18:00:00Z", modifiedAt: "2026-08-15T18:00:00Z", syncState: "saved" },
    ...overrides,
  };
}

function successResult(saved: MeetingRecord): MeetingRecordResult<MeetingRecord> {
  return { status: "success", value: saved };
}

type SaveFn = (record: MeetingRecord) => Promise<MeetingRecordResult<MeetingRecord>>;
type UpdateFn = (record: MeetingRecord, expectedVersion: number) => Promise<MeetingRecordResult<MeetingRecord>>;

function renderMeeting(overrides: Partial<{
  records: MeetingRecord[];
  saveRecord: SaveFn;
  updateRecord: UpdateFn;
  loadFailed: boolean;
}> = {}) {
  const saveRecord = overrides.saveRecord ?? vi.fn<SaveFn>();
  const updateRecord = overrides.updateRecord ?? vi.fn<UpdateFn>();
  const utils = render(
    <MeetingNotes
      openLog={vi.fn()}
      createDraftRecord={baseWorkRecord}
      records={overrides.records ?? []}
      saveRecord={saveRecord}
      updateRecord={updateRecord}
      loadFailed={overrides.loadFailed ?? false}
      storageMode="sharepoint"
    />,
  );
  return { ...utils, saveRecord, updateRecord };
}

const SUCCESS_RESULT: AnalyzeMeetingResult = {
  status: "success",
  analysis: {
    candidates: [
      { type: "SUMMARY", title: "Grant on track", detail: "Reviewed the grant status.", sourceExcerpt: "", ownerText: null, dueText: null, durationText: null },
      { type: "ACTION", title: "Call the district about the venue", detail: "Confirm venue details.", sourceExcerpt: "Annie will call the district about the venue by Friday.", ownerText: "Annie", dueText: "Friday", durationText: null },
      { type: "DECISION", title: "Move the fall network meeting to October", detail: "Agreed after discussion.", sourceExcerpt: "Decided to move the fall network meeting to October.", ownerText: null, dueText: null, durationText: null },
    ],
  },
  usage: { model: "claude-opus-5", inputTokens: 1200, outputTokens: 480 },
};

describe("Save Meeting — create vs update, never autosaves", () => {
  it("is disabled until the meeting has an unsaved change", () => {
    renderMeeting();
    expect((screen.getByRole("button", { name: "Save Meeting" }) as HTMLButtonElement).disabled).toBe(true);
  });

  it("routes a brand-new meeting through saveRecord exactly once, never updateRecord", async () => {
    const user = userEvent.setup();
    const saveRecord = vi.fn().mockResolvedValue(successResult(savedMeeting()));
    const { updateRecord } = renderMeeting({ saveRecord });
    await user.type(screen.getByPlaceholderText(/steels quarterly planning/i), "Quick check-in");
    await user.click(screen.getByRole("button", { name: "Save Meeting" }));
    await waitFor(() => expect(saveRecord).toHaveBeenCalledTimes(1));
    expect(updateRecord).not.toHaveBeenCalled();
  });

  it("routes a second save of an already-saved meeting through updateRecord with the current version, never saveRecord again", async () => {
    const user = userEvent.setup();
    const created = savedMeeting({ appId: "brand-new", metadata: { providerId: "1", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" } });
    const saveRecord = vi.fn().mockResolvedValue(successResult(created));
    const updateRecord = vi.fn().mockResolvedValue(successResult({ ...created, metadata: { ...created.metadata, version: 2 } }));
    renderMeeting({ saveRecord, updateRecord });

    await user.type(screen.getByPlaceholderText(/steels quarterly planning/i), "Quick check-in");
    await user.click(screen.getByRole("button", { name: "Save Meeting" }));
    await waitFor(() => expect(saveRecord).toHaveBeenCalledTimes(1));

    await user.type(screen.getByPlaceholderText(/take notes during the meeting/i), "Some more notes.");
    await user.click(screen.getByRole("button", { name: "Save Meeting" }));
    await waitFor(() => expect(updateRecord).toHaveBeenCalledTimes(1));
    expect(updateRecord.mock.calls[0][1]).toBe(1); // expectedVersion from the create response
    expect(saveRecord).toHaveBeenCalledTimes(1); // still exactly once — never a second create
  });

  it("never calls saveRecord/updateRecord from typing, analyzing, editing a candidate, or copying minutes", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(SUCCESS_RESULT));
    const { saveRecord, updateRecord } = renderMeeting();

    await user.type(screen.getByPlaceholderText(/take notes during the meeting/i), "Reviewed the budget, decided to move the meeting.");
    await user.click(screen.getByRole("button", { name: "Analyze Meeting" }));
    await waitFor(() => expect(screen.getByText("3 candidates")).toBeTruthy());

    const titleInput = screen.getByDisplayValue("Call the district about the venue");
    await user.clear(titleInput);
    await user.type(titleInput, "Call the district today");
    await user.click(screen.getByRole("button", { name: "Copy Minutes" }));

    expect(saveRecord).not.toHaveBeenCalled();
    expect(updateRecord).not.toHaveBeenCalled();
  });

  it("Save Meeting itself makes zero Anthropic requests", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const saveRecord = vi.fn().mockResolvedValue(successResult(savedMeeting()));
    renderMeeting({ saveRecord });
    await user.type(screen.getByPlaceholderText(/steels quarterly planning/i), "Quick check-in");
    await user.click(screen.getByRole("button", { name: "Save Meeting" }));
    await waitFor(() => expect(saveRecord).toHaveBeenCalledTimes(1));
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});

describe("Reopen Meeting — full-state reconstruction, zero AI call", () => {
  it("restores every field from a saved meeting when its row is clicked", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const record = savedMeeting();
    renderMeeting({ records: [record] });

    await user.click(screen.getByRole("button", { name: /STEELS quarterly planning/ }));

    expect((screen.getByPlaceholderText(/steels quarterly planning/i) as HTMLInputElement).value).toBe(record.title);
    expect((screen.getByPlaceholderText(/greg, annie, kim/i) as HTMLInputElement).value).toBe(record.attendeesText);
    expect((screen.getByPlaceholderText(/paste or type agenda items/i) as HTMLTextAreaElement).value).toBe(record.agendaText);
    expect((screen.getByPlaceholderText(/take notes during the meeting/i) as HTMLTextAreaElement).value).toBe(record.notesText);
    expect((screen.getByText("Choose a meeting type…").closest("select") as HTMLSelectElement).value).toBe(record.meetingType);

    expect(screen.getByText("3 candidates")).toBeTruthy();
    expect(screen.getByText("2 selected")).toBeTruthy();
    expect(screen.getByText("1 ignored")).toBeTruthy();
    expect(screen.getByDisplayValue("Call the district about the venue")).toBeTruthy();
    expect(screen.getByDisplayValue("Move the fall network meeting to October")).toBeTruthy();
    expect(screen.getByDisplayValue("Consider a shared calendar")).toBeTruthy();
    expect(screen.getByText("Annie")).toBeTruthy();
    expect(screen.getByText("Friday")).toBeTruthy();

    const ideaCard = screen.getByDisplayValue("Consider a shared calendar").closest(".candidate-card");
    expect(ideaCard?.className).toContain("deselected");

    const minutes = document.querySelector(".draft-minutes")!;
    expect(minutes.textContent).toContain("Move the fall network meeting to October");

    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("does not show unsaved changes right after reopening", async () => {
    const user = userEvent.setup();
    renderMeeting({ records: [savedMeeting()] });
    await user.click(screen.getByRole("button", { name: /STEELS quarterly planning/ }));
    expect(screen.queryByText("Unsaved changes")).toBeFalsy();
    expect((screen.getByRole("button", { name: "Save Meeting" }) as HTMLButtonElement).disabled).toBe(true);
  });
});

describe("Human review is authoritative", () => {
  it("persists the human-edited candidate state on save, not the original AI output", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(SUCCESS_RESULT));
    const saveRecord = vi.fn().mockResolvedValue(successResult(savedMeeting()));
    renderMeeting({ saveRecord });

    await user.type(screen.getByPlaceholderText(/take notes during the meeting/i), "Reviewed the budget, decided to move the meeting.");
    await user.click(screen.getByRole("button", { name: "Analyze Meeting" }));
    await waitFor(() => expect(screen.getByText("3 candidates")).toBeTruthy());

    // Retype a title, remove an owner chip, and deselect a candidate — human edits.
    const titleInput = screen.getByDisplayValue("Call the district about the venue");
    await user.clear(titleInput);
    await user.type(titleInput, "Call the district ASAP");
    await user.click(screen.getByRole("button", { name: "Remove owner" }));
    const checkboxes = screen.getAllByRole("checkbox");
    await user.click(checkboxes[2]); // deselect the DECISION candidate

    await user.click(screen.getByRole("button", { name: "Save Meeting" }));
    await waitFor(() => expect(saveRecord).toHaveBeenCalledTimes(1));

    const persisted = saveRecord.mock.calls[0][0] as MeetingRecord;
    const action = persisted.reviewedCandidates.find((c) => c.type === "ACTION")!;
    expect(action.title).toBe("Call the district ASAP"); // edit survives, not "Call the district about the venue"
    expect(action.ownerText).toBeNull(); // removed by the human, not the AI

    const decision = persisted.reviewedCandidates.find((c) => c.type === "DECISION")!;
    expect(decision.selected).toBe(false); // human deselected it
  });

  it("reflects the human's edits again after the meeting is reopened from the saved list", async () => {
    // Simulate what IUWorkTracker does: the saveRecord result becomes the next `records` entry.
    const editedAndSaved = savedMeeting({
      reviewedCandidates: [
        candidate({ title: "Call the district ASAP", ownerText: null }),
        candidate({ type: "DECISION", title: "Move the fall network meeting to October", ownerText: null, dueText: null, selected: false, sourceExcerpt: "" }),
      ],
    });
    const user = userEvent.setup();
    renderMeeting({ records: [editedAndSaved] });
    await user.click(screen.getByRole("button", { name: /STEELS quarterly planning/ }));

    expect(screen.getByDisplayValue("Call the district ASAP")).toBeTruthy();
    expect(screen.queryByDisplayValue("Call the district about the venue")).toBeFalsy();
    expect(screen.queryByText("Annie")).toBeFalsy(); // owner was removed by the human — not resurrected
    const decisionCard = screen.getByDisplayValue("Move the fall network meeting to October").closest(".candidate-card");
    expect(decisionCard?.className).toContain("deselected");
  });
});

describe("Pre-analysis save — a meeting may be saved before Analyze has ever run", () => {
  it("saves an agenda-only meeting with empty candidates and no fabricated analysis metadata", async () => {
    const user = userEvent.setup();
    const saveRecord = vi.fn().mockResolvedValue(successResult(savedMeeting({ reviewedCandidates: [], analysisModel: null, analyzedAt: null })));
    renderMeeting({ saveRecord });

    await user.type(screen.getByPlaceholderText(/paste or type agenda items/i), "1. Grant status\n2. Fall meeting date");
    await user.click(screen.getByRole("button", { name: "Save Meeting" }));

    await waitFor(() => expect(saveRecord).toHaveBeenCalledTimes(1));
    const persisted = saveRecord.mock.calls[0][0] as MeetingRecord;
    expect(persisted.reviewedCandidates).toEqual([]);
    expect(persisted.analysisModel).toBeNull();
    expect(persisted.analyzedAt).toBeNull();
    expect(persisted.agendaText).toContain("Grant status");
    expect(persisted.minutesText).toContain("Agenda");
  });

  it("saves a notes-only meeting the same way, with no Meeting Intelligence section shown before analysis", async () => {
    const user = userEvent.setup();
    renderMeeting();
    expect(screen.queryByText(/Meeting Intelligence/)).toBeFalsy(); // gated on analyzedAt, not just content
    await user.type(screen.getByPlaceholderText(/take notes during the meeting/i), "Some early notes ahead of the meeting.");
    expect((screen.getByRole("button", { name: "Save Meeting" }) as HTMLButtonElement).disabled).toBe(false);
  });
});

describe("Dirty-state tracking", () => {
  it("shows no 'Unsaved changes' indicator on a fresh, unedited screen", () => {
    renderMeeting();
    expect(screen.queryByText("Unsaved changes")).toBeFalsy();
  });

  it.each([
    ["title", () => screen.getByPlaceholderText(/steels quarterly planning/i)],
    ["attendees", () => screen.getByPlaceholderText(/greg, annie, kim/i)],
    ["agenda", () => screen.getByPlaceholderText(/paste or type agenda items/i)],
    ["notes", () => screen.getByPlaceholderText(/take notes during the meeting/i)],
  ] as const)("becomes dirty after editing %s", async (_label, getField) => {
    const user = userEvent.setup();
    renderMeeting();
    await user.type(getField(), "x");
    expect(screen.getByText("Unsaved changes")).toBeTruthy();
  });

  it("becomes dirty after a fresh analysis completes", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(SUCCESS_RESULT));
    renderMeeting();
    await user.type(screen.getByPlaceholderText(/take notes during the meeting/i), "Reviewed the budget, decided to move the meeting.");
    await user.click(screen.getByRole("button", { name: "Analyze Meeting" }));
    await waitFor(() => expect(screen.getByText("3 candidates")).toBeTruthy());
    expect(screen.getByText("Unsaved changes")).toBeTruthy();
  });

  it("becomes dirty after editing, deselecting, or removing a candidate", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(SUCCESS_RESULT));
    renderMeeting();
    await user.type(screen.getByPlaceholderText(/take notes during the meeting/i), "Reviewed the budget, decided to move the meeting.");
    await user.click(screen.getByRole("button", { name: "Analyze Meeting" }));
    await waitFor(() => expect(screen.getByText("3 candidates")).toBeTruthy());

    const record = savedMeeting();
    // Save to clear dirty state first, then verify a further candidate edit re-dirties it.
    void record;
    const removeButtons = screen.getAllByRole("button", { name: "Remove" });
    await user.click(removeButtons[0]);
    expect(screen.getByText("Unsaved changes")).toBeTruthy();
  });

  it("clears the indicator and disables Save Meeting immediately after a successful save", async () => {
    const user = userEvent.setup();
    const saveRecord = vi.fn().mockResolvedValue(successResult(savedMeeting()));
    renderMeeting({ saveRecord });
    await user.type(screen.getByPlaceholderText(/steels quarterly planning/i), "Quick check-in");
    expect(screen.getByText("Unsaved changes")).toBeTruthy();
    await user.click(screen.getByRole("button", { name: "Save Meeting" }));
    await waitFor(() => expect(saveRecord).toHaveBeenCalledTimes(1));
    await waitFor(() => expect(screen.queryByText("Unsaved changes")).toBeFalsy());
    expect((screen.getByRole("button", { name: "Save Meeting" }) as HTMLButtonElement).disabled).toBe(true);
  });

  it("New Meeting asks for confirmation when there are unsaved changes, and does nothing if declined", async () => {
    const user = userEvent.setup();
    const confirmSpy = vi.spyOn(window, "confirm").mockReturnValue(false);
    renderMeeting();
    const titleInput = screen.getByPlaceholderText(/steels quarterly planning/i) as HTMLInputElement;
    await user.type(titleInput, "Quick check-in");
    await user.click(screen.getByRole("button", { name: "New Meeting" }));
    expect(confirmSpy).toHaveBeenCalledTimes(1);
    expect(titleInput.value).toBe("Quick check-in"); // declined — nothing discarded
  });

  it("New Meeting discards the current draft when confirmed", async () => {
    const user = userEvent.setup();
    vi.spyOn(window, "confirm").mockReturnValue(true);
    renderMeeting();
    const titleInput = screen.getByPlaceholderText(/steels quarterly planning/i) as HTMLInputElement;
    await user.type(titleInput, "Quick check-in");
    await user.click(screen.getByRole("button", { name: "New Meeting" }));
    expect(titleInput.value).toBe("");
    expect(screen.queryByText("Unsaved changes")).toBeFalsy();
  });

  it("New Meeting never prompts when there is nothing unsaved", async () => {
    const user = userEvent.setup();
    const confirmSpy = vi.spyOn(window, "confirm");
    renderMeeting();
    await user.click(screen.getByRole("button", { name: "New Meeting" }));
    expect(confirmSpy).not.toHaveBeenCalled();
  });

  it("Re-analyze asks for confirmation once candidates already exist, and leaves them in place if declined", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(SUCCESS_RESULT));
    renderMeeting();
    await user.type(screen.getByPlaceholderText(/take notes during the meeting/i), "Reviewed the budget, decided to move the meeting.");
    await user.click(screen.getByRole("button", { name: "Analyze Meeting" }));
    await waitFor(() => expect(screen.getByText("3 candidates")).toBeTruthy());

    const confirmSpy = vi.spyOn(window, "confirm").mockReturnValue(false);
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    fetchSpy.mockClear(); // discard the call count from the initial Analyze above
    await user.click(screen.getByRole("button", { name: "Re-analyze meeting" }));
    expect(confirmSpy).toHaveBeenCalledWith(
      "Re-analyzing will replace the current Meeting Intelligence candidates with a new AI analysis. Your agenda and notes will remain.",
    );
    expect(fetchSpy).not.toHaveBeenCalled(); // declined — no second AI call
    expect(screen.getByText("3 candidates")).toBeTruthy(); // unchanged
  });
});
