// @vitest-environment jsdom
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import VoiceIntelligence from "../app/VoiceIntelligence";
import type { AnalyzeTranscriptResult } from "../lib/anthropic-voice-analysis";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import { REFERENCE_DATA } from "../lib/reference-data";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

// This file exercises the paste/analyze/review/edit surface only. The "Log as work" handoff
// itself (openLog/createDraftRecord wiring, eligibility, prefill mapping) is covered
// separately in tests/voice-work-record-handoff.test.tsx — these render calls just need
// harmless stand-in props to satisfy the component's now-required props.
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

const SUCCESS_RESULT: AnalyzeTranscriptResult = {
  status: "success",
  analysis: {
    candidates: [
      {
        type: "COMPLETED_WORK",
        title: "Met with North Schuylkill about science resources",
        detail: "Went well, about an hour long.",
        sourceExcerpt: "I met with North Schuylkill this morning about science resources",
        durationText: "about an hour",
      },
      {
        type: "ACTION",
        title: "Send Kim the Discovery materials",
        detail: "Follow-up from the conversation.",
        sourceExcerpt: "I need to send Kim the Discovery materials",
        durationText: null,
      },
    ],
  },
  usage: { model: "claude-opus-5", inputTokens: 900, outputTokens: 420 },
};

function jsonResponse(value: unknown): Response {
  return new Response(JSON.stringify(value), { status: 200, headers: { "content-type": "application/json" } });
}

describe("VoiceIntelligence — zero-cost load and explicit analysis boundary", () => {
  it("makes no request merely from rendering the screen", () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    renderVoice();
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("rejects an empty transcript client-side without any request", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    renderVoice();
    const analyzeButton = screen.getByRole("button", { name: "Analyze transcript" }) as HTMLButtonElement;
    expect(analyzeButton.disabled).toBe(true);
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("makes exactly one request when Analyze transcript is clicked", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(SUCCESS_RESULT));
    renderVoice();
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "Met with the team this morning.");
    await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
    await waitFor(() => expect(screen.getByText(/2 candidates/)).toBeTruthy());
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(fetchSpy).toHaveBeenCalledWith("/api/voice-intelligence", expect.objectContaining({ method: "POST" }));
  });

  it("does not clear the transcript after a failed analysis", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      jsonResponse({ status: "server_error", message: "The transcript could not be analyzed." }),
    );
    renderVoice();
    const textarea = screen.getByPlaceholderText(/paste the transcript/i) as HTMLTextAreaElement;
    await user.type(textarea, "Met with the team this morning.");
    await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
    await waitFor(() => expect(screen.getByRole("alert")).toBeTruthy());
    expect(textarea.value).toBe("Met with the team this morning.");
  });
});

describe("VoiceIntelligence — review and edit, zero persistence", () => {
  async function renderInReview() {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(SUCCESS_RESULT));
    renderVoice();
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "Met with the team this morning.");
    await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
    await waitFor(() => expect(screen.getByText(/2 candidates/)).toBeTruthy());
    return { user, fetchSpy };
  }

  it("shows each candidate with type, title, detail, source excerpt, and duration", async () => {
    await renderInReview();
    expect(screen.getByDisplayValue("Met with North Schuylkill about science resources")).toBeTruthy();
    expect(screen.getByDisplayValue("Send Kim the Discovery materials")).toBeTruthy();
    expect(screen.getByText(/I met with North Schuylkill this morning/)).toBeTruthy();
    expect(screen.getByText("about an hour")).toBeTruthy();
    expect(screen.getByText("2 candidates")).toBeTruthy();
    expect(screen.getByText("2 selected")).toBeTruthy();
    expect(screen.getByText("0 ignored")).toBeTruthy();
  });

  it("allows editing a candidate's title", async () => {
    const { user } = await renderInReview();
    const titleInput = screen.getByDisplayValue("Send Kim the Discovery materials");
    await user.clear(titleInput);
    await user.type(titleInput, "Send Kim the Discovery packet");
    expect(screen.getByDisplayValue("Send Kim the Discovery packet")).toBeTruthy();
  });

  it("allows editing a candidate's detail", async () => {
    const { user } = await renderInReview();
    const detailInput = screen.getByDisplayValue("Follow-up from the conversation.");
    await user.clear(detailInput);
    await user.type(detailInput, "Send by Friday.");
    expect(screen.getByDisplayValue("Send by Friday.")).toBeTruthy();
  });

  it("allows changing a candidate's type", async () => {
    const { user } = await renderInReview();
    const typeSelects = screen.getAllByLabelText("Candidate type") as HTMLSelectElement[];
    await user.selectOptions(typeSelects[1], "IDEA");
    expect(typeSelects[1].value).toBe("IDEA");
  });

  it("allows deselecting and reselecting a candidate, updating the summary counts", async () => {
    const { user } = await renderInReview();
    const checkboxes = screen.getAllByRole("checkbox");
    await user.click(checkboxes[0]);
    expect(screen.getByText("1 selected")).toBeTruthy();
    expect(screen.getByText("1 ignored")).toBeTruthy();
    await user.click(checkboxes[0]);
    expect(screen.getByText("2 selected")).toBeTruthy();
  });

  it("allows removing a candidate entirely", async () => {
    const { user } = await renderInReview();
    const removeButtons = screen.getAllByRole("button", { name: "Remove" });
    await user.click(removeButtons[0]);
    expect(screen.getByText("1 candidate")).toBeTruthy();
    expect(screen.queryByDisplayValue("Met with North Schuylkill about science resources")).toBeFalsy();
  });

  it("allows removing an unsupported duration from a candidate", async () => {
    const { user } = await renderInReview();
    await user.click(screen.getByRole("button", { name: "Remove duration" }));
    expect(screen.queryByText("about an hour")).toBeFalsy();
  });

  it("candidate edits cause zero additional network requests", async () => {
    const { user, fetchSpy } = await renderInReview();
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    const titleInput = screen.getByDisplayValue("Send Kim the Discovery materials");
    await user.type(titleInput, " today");
    const checkboxes = screen.getAllByRole("checkbox");
    await user.click(checkboxes[0]);
    const typeSelects = screen.getAllByLabelText("Candidate type") as HTMLSelectElement[];
    await user.selectOptions(typeSelects[0], "DECISION");
    await user.click(screen.getAllByRole("button", { name: "Remove" })[1]);
    expect(fetchSpy).toHaveBeenCalledTimes(1); // still exactly the one analyze() request
  });

  it("never writes to localStorage or sessionStorage through the full paste → analyze → edit flow", async () => {
    const setItem = vi.spyOn(Storage.prototype, "setItem");
    const removeItem = vi.spyOn(Storage.prototype, "removeItem");
    const clear = vi.spyOn(Storage.prototype, "clear");
    const { user } = await renderInReview();
    const titleInput = screen.getByDisplayValue("Send Kim the Discovery materials");
    await user.type(titleInput, " today");
    await user.click(screen.getAllByRole("checkbox")[0]);
    await user.click(screen.getAllByRole("button", { name: "Remove" })[0]);
    expect(setItem).not.toHaveBeenCalled();
    expect(removeItem).not.toHaveBeenCalled();
    expect(clear).not.toHaveBeenCalled();
  });

  it("shows a friendly empty state when no candidates are returned", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse({ ...SUCCESS_RESULT, analysis: { candidates: [] } }));
    renderVoice();
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "Just testing this thing out.");
    await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
    await waitFor(() => expect(screen.getByText("No useful candidates found")).toBeTruthy());
    expect(within(screen.getByText("No useful candidates found").closest("div")!).getByRole("button", { name: "Edit transcript" })).toBeTruthy();
  });

  it("Analyze another transcript fully resets to a blank paste screen", async () => {
    const { user } = await renderInReview();
    await user.click(screen.getByRole("button", { name: "Analyze another transcript" }));
    const textarea = await screen.findByPlaceholderText(/paste the transcript/i);
    expect((textarea as HTMLTextAreaElement).value).toBe("");
    expect(screen.queryByDisplayValue("Send Kim the Discovery materials")).toBeFalsy();
    expect(screen.queryByText("2 selected")).toBeFalsy();
  });
});
