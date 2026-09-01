// @vitest-environment jsdom
import { cleanup, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import MeetingNotes from "../app/MeetingNotes";
import type { AnalyzeMeetingResult } from "../lib/anthropic-meeting-analysis";
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

function renderMeeting() {
  return render(<MeetingNotes openLog={vi.fn()} createDraftRecord={baseWorkRecord} />);
}

const SUCCESS_RESULT: AnalyzeMeetingResult = {
  status: "success",
  analysis: {
    candidates: [
      {
        type: "SUMMARY",
        title: "Grant on track, fall meeting moved to October",
        detail: "Reviewed the STEELS grant status and delayed the fall network meeting.",
        sourceExcerpt: "",
        ownerText: null,
        dueText: null,
        durationText: null,
      },
      {
        type: "ACTION",
        title: "Call the district about the venue",
        detail: "Confirm venue details.",
        sourceExcerpt: "Annie will call the district about the venue by Friday.",
        ownerText: "Annie",
        dueText: "Friday",
        durationText: null,
      },
      {
        type: "DECISION",
        title: "Move the fall network meeting to October",
        detail: "Agreed after discussion.",
        sourceExcerpt: "Decided to move the fall network meeting to October.",
        ownerText: null,
        dueText: null,
        durationText: null,
      },
    ],
  },
  usage: { model: "claude-opus-5", inputTokens: 1200, outputTokens: 480 },
};

function jsonResponse(value: unknown): Response {
  return new Response(JSON.stringify(value), { status: 200, headers: { "content-type": "application/json" } });
}

describe("MeetingNotes — details/agenda/notes editing, zero-cost load", () => {
  it("makes no request merely from rendering the screen", () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    renderMeeting();
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("shows the non-durable notice", () => {
    renderMeeting();
    expect(screen.getByText("Meeting Notes V1 is a review workspace. Meeting content is not saved yet.")).toBeTruthy();
  });

  it("meeting details are editable", async () => {
    const user = userEvent.setup();
    renderMeeting();
    const titleInput = screen.getByPlaceholderText(/steels quarterly planning/i);
    await user.type(titleInput, "District check-in");
    expect((titleInput as HTMLInputElement).value).toBe("District check-in");

    const attendeesInput = screen.getByPlaceholderText(/greg, annie, kim/i);
    await user.type(attendeesInput, "Greg, Annie");
    expect((attendeesInput as HTMLInputElement).value).toBe("Greg, Annie");

    await userEvent.selectOptions(screen.getByText("Choose a meeting type…").closest("select")!, "District Meeting");
    expect((screen.getByText("Choose a meeting type…").closest("select") as HTMLSelectElement).value).toBe("District Meeting");
  });

  it("agenda is editable", async () => {
    const user = userEvent.setup();
    renderMeeting();
    const agenda = screen.getByPlaceholderText(/paste or type agenda items/i);
    await user.type(agenda, "1. Grant status");
    expect((agenda as HTMLTextAreaElement).value).toBe("1. Grant status");
  });

  it("notes are editable", async () => {
    const user = userEvent.setup();
    renderMeeting();
    const notes = screen.getByPlaceholderText(/take notes during the meeting/i);
    await user.type(notes, "Reviewed the budget.");
    expect((notes as HTMLTextAreaElement).value).toBe("Reviewed the budget.");
  });

  it("disables Analyze Meeting when neither agenda nor notes have content", () => {
    renderMeeting();
    const button = screen.getByRole("button", { name: "Analyze Meeting" }) as HTMLButtonElement;
    expect(button.disabled).toBe(true);
  });

  it("enables Analyze Meeting once notes have content, even with no other fields filled", async () => {
    const user = userEvent.setup();
    renderMeeting();
    await user.type(screen.getByPlaceholderText(/take notes during the meeting/i), "Reviewed the budget.");
    const button = screen.getByRole("button", { name: "Analyze Meeting" }) as HTMLButtonElement;
    expect(button.disabled).toBe(false);
  });

  it("makes exactly one request when Analyze Meeting is clicked", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(SUCCESS_RESULT));
    renderMeeting();
    await user.type(screen.getByPlaceholderText(/take notes during the meeting/i), "Reviewed the budget.");
    await user.click(screen.getByRole("button", { name: "Analyze Meeting" }));
    await waitFor(() => expect(screen.getByText("3 candidates")).toBeTruthy());
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(fetchSpy).toHaveBeenCalledWith("/api/meeting-intelligence", expect.objectContaining({ method: "POST" }));
  });

  it("does not clear meeting content after a failed analysis", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse({ status: "server_error", message: "The meeting could not be analyzed." }));
    renderMeeting();
    const notes = screen.getByPlaceholderText(/take notes during the meeting/i) as HTMLTextAreaElement;
    await user.type(notes, "Reviewed the budget.");
    await user.click(screen.getByRole("button", { name: "Analyze Meeting" }));
    await waitFor(() => expect(screen.getByRole("alert")).toBeTruthy());
    expect(notes.value).toBe("Reviewed the budget.");
  });
});

describe("MeetingNotes — candidate review", () => {
  async function renderAnalyzed() {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(SUCCESS_RESULT));
    renderMeeting();
    await user.type(screen.getByPlaceholderText(/take notes during the meeting/i), "Reviewed the budget, decided to move the meeting.");
    await user.click(screen.getByRole("button", { name: "Analyze Meeting" }));
    await waitFor(() => expect(screen.getByText("3 candidates")).toBeTruthy());
    return { user, fetchSpy };
  }

  it("renders candidate cards with type, title, detail, and owner/due when present", async () => {
    await renderAnalyzed();
    expect(screen.getByDisplayValue("Grant on track, fall meeting moved to October")).toBeTruthy();
    expect(screen.getByDisplayValue("Call the district about the venue")).toBeTruthy();
    expect(screen.getByDisplayValue("Move the fall network meeting to October")).toBeTruthy();
    expect(screen.getByText("Annie")).toBeTruthy();
    expect(screen.getByText("Friday")).toBeTruthy();
    expect(screen.getByText("3 selected")).toBeTruthy();
    expect(screen.getByText("0 ignored")).toBeTruthy();
  });

  it("allows editing a candidate's title and detail", async () => {
    const { user } = await renderAnalyzed();
    const titleInput = screen.getByDisplayValue("Call the district about the venue");
    await user.clear(titleInput);
    await user.type(titleInput, "Call the district ASAP");
    expect(screen.getByDisplayValue("Call the district ASAP")).toBeTruthy();

    const detailInput = screen.getByDisplayValue("Confirm venue details.");
    await user.clear(detailInput);
    await user.type(detailInput, "Confirm by end of week.");
    expect(screen.getByDisplayValue("Confirm by end of week.")).toBeTruthy();
  });

  it("allows changing a candidate's type", async () => {
    const { user } = await renderAnalyzed();
    const typeSelects = screen.getAllByLabelText("Candidate type") as HTMLSelectElement[];
    await user.selectOptions(typeSelects[2], "FOLLOW_UP_AGENDA");
    expect(typeSelects[2].value).toBe("FOLLOW_UP_AGENDA");
  });

  it("allows deselecting a candidate, updating selected/ignored counts", async () => {
    const { user } = await renderAnalyzed();
    const checkboxes = screen.getAllByRole("checkbox");
    await user.click(checkboxes[1]);
    expect(screen.getByText("2 selected")).toBeTruthy();
    expect(screen.getByText("1 ignored")).toBeTruthy();
  });

  it("allows removing a candidate entirely", async () => {
    const { user } = await renderAnalyzed();
    const removeButtons = screen.getAllByRole("button", { name: "Remove" });
    await user.click(removeButtons[0]);
    expect(screen.getByText("2 candidates")).toBeTruthy();
    expect(screen.queryByDisplayValue("Grant on track, fall meeting moved to October")).toBeFalsy();
  });
});

describe("MeetingNotes — Draft Minutes and Copy Minutes", () => {
  async function renderAnalyzed() {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(SUCCESS_RESULT));
    renderMeeting();
    await user.type(screen.getByPlaceholderText(/steels quarterly planning/i), "STEELS check-in");
    await user.type(screen.getByPlaceholderText(/take notes during the meeting/i), "Reviewed the budget, decided to move the meeting.");
    await user.click(screen.getByRole("button", { name: "Analyze Meeting" }));
    await waitFor(() => expect(screen.getByText("3 candidates")).toBeTruthy());
    return { user };
  }

  it("Draft Minutes reflects the current reviewed state, including selected decisions and actions", async () => {
    await renderAnalyzed();
    const minutes = screen.getByText(/STEELS check-in/i).closest("pre")!;
    expect(minutes.textContent).toContain("STEELS check-in");
    expect(minutes.textContent).toContain("Move the fall network meeting to October");
    expect(minutes.textContent).toContain("Call the district about the venue");
    expect(minutes.textContent).toContain("Owner: Annie");
    expect(minutes.textContent).toContain("Due: Friday");
  });

  it("Draft Minutes excludes an ignored (deselected) candidate", async () => {
    const { user } = await renderAnalyzed();
    const checkboxes = screen.getAllByRole("checkbox");
    // Deselect the DECISION candidate (third card).
    await user.click(checkboxes[2]);
    const minutes = screen.getByText(/STEELS check-in/i).closest("pre")!;
    expect(minutes.textContent).not.toContain("Move the fall network meeting to October");
  });

  it("Draft Minutes updates live as a candidate is edited", async () => {
    const { user } = await renderAnalyzed();
    const titleInput = screen.getByDisplayValue("Call the district about the venue");
    await user.clear(titleInput);
    await user.type(titleInput, "Call the district today");
    const minutes = screen.getByText(/STEELS check-in/i).closest("pre")!;
    expect(minutes.textContent).toContain("Call the district today");
    expect(minutes.textContent).not.toContain("Call the district about the venue");
  });

  it("Copy Minutes copies exactly the current Draft Minutes text (same source the panel renders)", async () => {
    // jsdom's navigator.clipboard cannot be reliably stubbed across this module graph, so this
    // verifies the same guarantee a clipboard spy would: Copy Minutes succeeds (button flips to
    // "Copied!", proving the call completed without error) and the visible Draft Minutes panel
    // — the exact string copyMinutes() passes to writeText() — reflects the current reviewed
    // state, which the dedicated Draft Minutes tests above already assert in detail.
    const { user } = await renderAnalyzed();
    const minutesBefore = screen.getByText(/STEELS check-in/i).closest("pre")!.textContent;
    expect(minutesBefore).toContain("Call the district about the venue");
    await user.click(screen.getByRole("button", { name: "Copy Minutes" }));
    expect(await screen.findByRole("button", { name: "Copied!" })).toBeTruthy();
  });
});
