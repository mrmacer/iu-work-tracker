// @vitest-environment jsdom
//
// Patch 7.2 — Voice Intelligence WORKING SESSION persistence, end to end through the real
// VoiceIntelligence component (component-level tests use the injected `storage` seam and a
// small in-memory fake) and through the real IUWorkTracker shell, Log Work wizard, and
// MemoryDataProvider (full-app tests — window.sessionStorage is replaced with the same fake).
// Every test mocks fetch: ZERO real Anthropic calls, zero SharePoint.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { readFileSync } from "node:fs";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import IUWorkTracker from "../app/IUWorkTracker";
import VoiceIntelligence from "../app/VoiceIntelligence";
import type { AnalyzeTranscriptResult } from "../lib/anthropic-voice-analysis";
import { MemoryDataProvider } from "../lib/data-provider";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import { REFERENCE_DATA } from "../lib/reference-data";
import { VOICE_SESSION_STORAGE_KEY, type VoiceSessionStorage } from "../lib/voice-intelligence-session";
import { createFakeStorage, installWindowStorage, type FakeStorage } from "./voice-session-test-utils";

const restores: (() => void)[] = [];

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
  while (restores.length) restores.pop()!();
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

const ANALYSIS: AnalyzeTranscriptResult = {
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
        type: "COMPLETED_WORK",
        title: "Drafted the STEELS newsletter",
        detail: "Finished the draft.",
        sourceExcerpt: "I drafted the STEELS newsletter",
        durationText: null,
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

const STATUS_SAVED = "Working session saved in this tab.";
const STATUS_RESTORED = "Working session restored.";
const TAB_COPY = "Closing this tab clears the temporary Voice Intelligence session.";

// ---------------------------------------------------------------------------------------------
// Component-level tests — the `storage` seam.
// ---------------------------------------------------------------------------------------------

function renderVoice(storage: VoiceSessionStorage | null, openLog = vi.fn()) {
  return {
    openLog,
    ...render(
      <VoiceIntelligence
        openLog={openLog}
        createDraftRecord={baseWorkRecord}
        references={REFERENCE_DATA}
        saveContact={vi.fn()}
        updateContact={vi.fn()}
        storage={storage}
        saveProject={vi.fn()}
        updateProject={vi.fn()}
        saveOrganization={vi.fn()}
        updateOrganization={vi.fn()}
      />,
    ),
  };
}

async function analyzeInVoice(user: ReturnType<typeof userEvent.setup>) {
  await user.type(screen.getByPlaceholderText(/paste the transcript/i), "I met with North Schuylkill this morning.");
  await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
  await waitFor(() => expect(screen.getByText("3 candidates")).toBeTruthy());
}

describe("restore — transcript and analysis survive unmount/remount", () => {
  it("restores an un-analyzed pasted transcript on remount, without analyzing it", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const storage = createFakeStorage();
    const first = renderVoice(storage);
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "A ramble I have not analyzed yet.");
    first.unmount();

    renderVoice(storage);
    expect((screen.getByPlaceholderText(/paste the transcript/i) as HTMLTextAreaElement).value).toBe("A ramble I have not analyzed yet.");
    expect(screen.queryByText(/^\d+ candidates?$/)).toBeFalsy(); // no fake candidates, still the paste screen
    expect(fetchSpy).not.toHaveBeenCalled(); // never re-analyzes on restore
  });

  it("restores the analyzed candidates (and the review screen) on remount with ZERO fetch calls", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const storage = createFakeStorage();
    const first = renderVoice(storage);
    await analyzeInVoice(user);
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    first.unmount();
    fetchSpy.mockClear();

    renderVoice(storage);
    expect(screen.getByText("3 candidates")).toBeTruthy();
    expect(screen.getByDisplayValue("Met with North Schuylkill about science resources")).toBeTruthy();
    expect(screen.getByDisplayValue("Send Kim the Discovery materials")).toBeTruthy();
    expect(screen.getByText(/claude-opus-5 · 900 in \/ 420 out tokens/)).toBeTruthy();
    expect(fetchSpy).not.toHaveBeenCalled(); // restoring is a pure sessionStorage read — no Anthropic call
  });

  it("restores the same candidate ids — they are not regenerated on mount", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const storage = createFakeStorage();
    const first = renderVoice(storage);
    await analyzeInVoice(user);
    const idsBefore = (JSON.parse(storage.data.get(VOICE_SESSION_STORAGE_KEY)!).candidates as { id: string }[]).map((c) => c.id);
    first.unmount();
    renderVoice(storage);
    await user.click(screen.getAllByRole("checkbox")[0]); // any change re-persists the restored ids
    const idsAfter = (JSON.parse(storage.data.get(VOICE_SESSION_STORAGE_KEY)!).candidates as { id: string }[]).map((c) => c.id);
    expect(idsAfter).toEqual(idsBefore);
  });

  it("restores an edited title and detail", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const storage = createFakeStorage();
    const first = renderVoice(storage);
    await analyzeInVoice(user);
    await user.type(screen.getByDisplayValue("Send Kim the Discovery materials"), " by Friday");
    const detail = screen.getByDisplayValue("Follow-up from the conversation.");
    await user.clear(detail);
    await user.type(detail, "Human-edited detail");
    first.unmount();

    renderVoice(storage);
    expect(screen.getByDisplayValue("Send Kim the Discovery materials by Friday")).toBeTruthy();
    expect(screen.getByDisplayValue("Human-edited detail")).toBeTruthy();
  });

  it("restores a changed candidate type", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const storage = createFakeStorage();
    const first = renderVoice(storage);
    await analyzeInVoice(user);
    await user.selectOptions(screen.getAllByLabelText("Candidate type")[2], "DECISION");
    first.unmount();

    renderVoice(storage);
    expect((screen.getAllByLabelText("Candidate type")[2] as HTMLSelectElement).value).toBe("DECISION");
  });

  it("restores selection state and a removed duration", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const storage = createFakeStorage();
    const first = renderVoice(storage);
    await analyzeInVoice(user);
    await user.click(screen.getAllByRole("checkbox")[1]); // deselect the second candidate
    await user.click(screen.getByRole("button", { name: "Remove duration" }));
    first.unmount();

    renderVoice(storage);
    const checkboxes = screen.getAllByRole("checkbox") as HTMLInputElement[];
    expect(checkboxes.map((box) => box.checked)).toEqual([true, false, true]);
    expect(screen.getByText("2 selected")).toBeTruthy();
    expect(screen.getByText("1 not selected")).toBeTruthy();
    expect(screen.queryByText("about an hour")).toBeFalsy();
  });

  it("a removed candidate stays removed after remount", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const storage = createFakeStorage();
    const first = renderVoice(storage);
    await analyzeInVoice(user);
    await user.click(screen.getAllByRole("button", { name: "Remove" })[1]);
    expect(screen.queryByDisplayValue("Drafted the STEELS newsletter")).toBeFalsy();
    first.unmount();

    renderVoice(storage);
    expect(screen.getByText("2 candidates")).toBeTruthy();
    expect(screen.queryByDisplayValue("Drafted the STEELS newsletter")).toBeFalsy();
  });

  it("falls back to the normal empty screen — no crash, no raw parse error — for a malformed stored session", () => {
    const storage = createFakeStorage();
    storage.data.set(VOICE_SESSION_STORAGE_KEY, "{ definitely not json");
    renderVoice(storage);
    expect((screen.getByPlaceholderText(/paste the transcript/i) as HTMLTextAreaElement).value).toBe("");
    expect(screen.queryByText(/Unexpected token|JSON|SyntaxError/i)).toBeFalsy();
    expect(screen.queryByText(STATUS_RESTORED)).toBeFalsy();
    expect(storage.data.has(VOICE_SESSION_STORAGE_KEY)).toBe(false);
  });

  it("falls back to the normal empty screen for a wrong-schema-version session", () => {
    const storage = createFakeStorage();
    storage.data.set(VOICE_SESSION_STORAGE_KEY, JSON.stringify({ schemaVersion: 99, transcript: "x", candidates: [] }));
    renderVoice(storage);
    expect((screen.getByPlaceholderText(/paste the transcript/i) as HTMLTextAreaElement).value).toBe("");
  });
});

describe("Analyze remains the only AI trigger", () => {
  it("a failed re-analysis leaves the stored transcript in place and adds no fake candidates", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse({ status: "server_error", message: "The service is unavailable." }));
    const storage = createFakeStorage();
    renderVoice(storage);
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "Keep this transcript.");
    await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
    await waitFor(() => expect(screen.getByText("The service is unavailable.")).toBeTruthy());
    const stored = JSON.parse(storage.data.get(VOICE_SESSION_STORAGE_KEY)!);
    expect(stored.transcript).toBe("Keep this transcript.");
    expect(stored.analyzed).toBe(false);
    expect(stored.candidates).toEqual([]);
  });

  it("a successful analysis persists the session as analyzed", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const storage = createFakeStorage();
    renderVoice(storage);
    await analyzeInVoice(user);
    const stored = JSON.parse(storage.data.get(VOICE_SESSION_STORAGE_KEY)!);
    expect(stored.analyzed).toBe(true);
    expect(stored.phase).toBe("review");
    expect(stored.candidates).toHaveLength(3);
  });
});

describe("Clear destroys the temporary session", () => {
  it("Clear on the paste screen empties the transcript AND removes the session", async () => {
    const user = userEvent.setup();
    const storage = createFakeStorage();
    const first = renderVoice(storage);
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "Something to clear.");
    expect(storage.data.has(VOICE_SESSION_STORAGE_KEY)).toBe(true);
    await user.click(screen.getByRole("button", { name: "Clear" }));
    expect((screen.getByPlaceholderText(/paste the transcript/i) as HTMLTextAreaElement).value).toBe("");
    expect(storage.data.has(VOICE_SESSION_STORAGE_KEY)).toBe(false);
    first.unmount();
    renderVoice(storage);
    expect((screen.getByPlaceholderText(/paste the transcript/i) as HTMLTextAreaElement).value).toBe("");
  });

  it("Analyze another transcript clears every candidate and review state and removes the session", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const storage = createFakeStorage();
    const first = renderVoice(storage);
    await analyzeInVoice(user);
    await user.click(screen.getByRole("button", { name: "Analyze another transcript" }));
    expect(screen.queryByText("3 candidates")).toBeFalsy();
    expect(storage.data.has(VOICE_SESSION_STORAGE_KEY)).toBe(false);
    first.unmount();
    renderVoice(storage);
    expect(screen.queryByText("3 candidates")).toBeFalsy();
    expect((screen.getByPlaceholderText(/paste the transcript/i) as HTMLTextAreaElement).value).toBe("");
  });
});

describe("session status copy", () => {
  it("does not appear on a fresh, empty Voice screen", () => {
    renderVoice(createFakeStorage());
    expect(screen.queryByText(STATUS_SAVED)).toBeFalsy();
    expect(screen.queryByText(STATUS_RESTORED)).toBeFalsy();
    expect(screen.queryByText(new RegExp(TAB_COPY))).toBeFalsy();
  });

  it("appears once there is a session, is honest about being tab-scoped, and claims no durability", async () => {
    const user = userEvent.setup();
    renderVoice(createFakeStorage());
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "x");
    expect(screen.getByText(STATUS_SAVED)).toBeTruthy();
    const status = screen.getByRole("status");
    expect(status.textContent).toContain(TAB_COPY);
    expect(status.textContent).not.toMatch(/cloud|backed up|securely|your account|sharepoint/i);
  });

  it("says 'restored' after a remount and 'saved' again after the next change", async () => {
    const user = userEvent.setup();
    const storage = createFakeStorage();
    const first = renderVoice(storage);
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "hello");
    first.unmount();
    renderVoice(storage);
    expect(screen.getByText(STATUS_RESTORED)).toBeTruthy();
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "!");
    expect(screen.getByText(STATUS_SAVED)).toBeTruthy();
    expect(screen.queryByText(STATUS_RESTORED)).toBeFalsy();
  });

  it("disappears after Clear", async () => {
    const user = userEvent.setup();
    renderVoice(createFakeStorage());
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "x");
    await user.click(screen.getByRole("button", { name: "Clear" }));
    expect(screen.queryByText(STATUS_SAVED)).toBeFalsy();
  });
});

describe("storage failure never blocks the screen", () => {
  const failingStorage = (): VoiceSessionStorage => ({
    getItem: () => {
      throw new Error("SecurityError");
    },
    setItem: () => {
      throw new Error("QuotaExceededError");
    },
    removeItem: () => {
      throw new Error("SecurityError");
    },
  });

  it("analysis, editing, and Log as work all still work — and the UI does not claim the session was saved", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const openLog = vi.fn();
    renderVoice(failingStorage(), openLog);
    await analyzeInVoice(user);
    await user.type(screen.getByDisplayValue("Send Kim the Discovery materials"), " today");
    expect(screen.getByDisplayValue("Send Kim the Discovery materials today")).toBeTruthy();
    await user.click(screen.getAllByRole("button", { name: "Log as work" })[0]);
    expect(openLog).toHaveBeenCalledTimes(1);
    expect(screen.queryByText(STATUS_SAVED)).toBeFalsy(); // never overclaims when nothing was actually saved
  });

  it("works with no storage available at all (null)", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    renderVoice(null);
    await analyzeInVoice(user);
    expect(screen.getByText("3 candidates")).toBeTruthy();
    expect(screen.queryByText(STATUS_SAVED)).toBeFalsy();
  });
});

describe("Log as work — component handoff", () => {
  it("persists the current (edited) session BEFORE handing off, and passes an onSaved callback", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const storage = createFakeStorage();
    let storedAtHandoff: string | null = null;
    const openLog = vi.fn(() => {
      storedAtHandoff = storage.data.get(VOICE_SESSION_STORAGE_KEY) ?? null;
    });
    renderVoice(storage, openLog);
    await analyzeInVoice(user);
    await user.type(screen.getByDisplayValue("Drafted the STEELS newsletter"), " (v2)");
    await user.click(screen.getAllByRole("button", { name: "Log as work" })[1]);

    expect(openLog).toHaveBeenCalledTimes(1);
    expect(storedAtHandoff).not.toBeNull();
    expect(JSON.parse(storedAtHandoff!).candidates[1].title).toBe("Drafted the STEELS newsletter (v2)");
    expect(typeof (openLog.mock.calls[0] as unknown[])[1]).toBe("function");
  });

  it("does not remove the candidate, clear Voice state, or call Anthropic when Log as work is clicked", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    renderVoice(createFakeStorage());
    await analyzeInVoice(user);
    fetchSpy.mockClear();
    await user.click(screen.getAllByRole("button", { name: "Log as work" })[0]);
    expect(fetchSpy).not.toHaveBeenCalled();
    expect(screen.getByText("3 candidates")).toBeTruthy();
    expect(screen.getByDisplayValue("Met with North Schuylkill about science resources")).toBeTruthy();
  });

  it("clicking Log as work alone never marks a candidate Logged", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    renderVoice(createFakeStorage());
    await analyzeInVoice(user);
    await user.click(screen.getAllByRole("button", { name: "Log as work" })[0]);
    expect(screen.queryByText(/Logged/)).toBeFalsy();
  });

  it("marks ONLY the originating candidate Logged when the onSaved callback fires, and that survives remount", async () => {
    const user = userEvent.setup();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const storage = createFakeStorage();
    let onSaved: ((saved: WorkRecord) => void) | undefined;
    const openLog = vi.fn((_record?: WorkRecord, callback?: (saved: WorkRecord) => void) => {
      onSaved = callback;
    });
    const first = renderVoice(storage, openLog);
    await analyzeInVoice(user);
    await user.click(screen.getAllByRole("button", { name: "Log as work" })[1]); // the SECOND completed-work candidate
    expect(onSaved).toBeDefined();
    onSaved!({ ...baseWorkRecord(), appId: "saved-work-record-1" });

    await waitFor(() => expect(screen.getAllByText("Logged ✓")).toHaveLength(1));
    const cards = screen.getAllByLabelText("Candidate title").map((input) => input.closest(".candidate-card") as HTMLElement);
    expect(within(cards[1]).getByText("Logged ✓")).toBeTruthy();
    expect(within(cards[1]).getByRole("button", { name: "Log again" })).toBeTruthy(); // de-emphasized, not disabled
    expect(within(cards[0]).queryByText("Logged ✓")).toBeFalsy();
    expect(within(cards[0]).getByRole("button", { name: "Log as work" })).toBeTruthy();
    expect(cards[1].classList.contains("deselected")).toBe(false); // logging never silently deselects

    first.unmount();
    renderVoice(storage);
    const restoredCards = screen.getAllByLabelText("Candidate title").map((input) => input.closest(".candidate-card") as HTMLElement);
    expect(within(restoredCards[1]).getByText("Logged ✓")).toBeTruthy();
    expect(screen.getAllByText("Logged ✓")).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------------------------
// Full-app tests — real IUWorkTracker shell, real Log Work wizard, real MemoryDataProvider.
// ---------------------------------------------------------------------------------------------

describe("full app — navigation, cancel, and the real Work Record save path", () => {
  let sessionFake: FakeStorage;
  let localFake: FakeStorage;

  beforeEach(() => {
    sessionFake = createFakeStorage();
    localFake = createFakeStorage();
    restores.push(installWindowStorage("sessionStorage", sessionFake));
    restores.push(installWindowStorage("localStorage", localFake));
  });

  async function setup() {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const dataProvider = new MemoryDataProvider([]);
    const createSpy = vi.spyOn(dataProvider, "createWorkRecord");
    render(<IUWorkTracker dataProvider={dataProvider} inboxDataProvider={new SessionInboxIntelligenceProvider()} />);
    const nav = await screen.findByRole("navigation");
    const goTo = async (name: RegExp) => user.click(within(nav).getByRole("button", { name }));
    await goTo(/voice intelligence/i);
    await user.type(screen.getByPlaceholderText(/paste the transcript/i), "I met with North Schuylkill this morning.");
    await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
    await waitFor(() => expect(screen.getByText("3 candidates")).toBeTruthy());
    return { user, fetchSpy, createSpy, goTo, nav };
  }

  async function completeWizardAndSave(user: ReturnType<typeof userEvent.setup>) {
    const dialog = await screen.findByRole("dialog");
    const activityType = within(dialog).getByLabelText(/Activity type/i) as HTMLSelectElement;
    await user.selectOptions(activityType, activityType.options[1].value);
    for (let i = 0; i < 4; i++) await user.click(within(dialog).getByRole("button", { name: /continue/i }));
    await user.click(within(dialog).getByRole("button", { name: "Save & done" }));
  }

  it("normal navigation away and back restores the analyzed session with zero extra Anthropic calls", async () => {
    const { user, fetchSpy, goTo } = await setup();
    await user.type(screen.getByDisplayValue("Send Kim the Discovery materials"), " today");
    await user.click(screen.getAllByRole("checkbox")[1]);
    expect(fetchSpy).toHaveBeenCalledTimes(1);

    await goTo(/projects/i);
    expect(screen.queryByText("3 candidates")).toBeFalsy(); // Voice really unmounted
    await goTo(/voice intelligence/i);

    expect(screen.getByText("3 candidates")).toBeTruthy();
    expect(screen.getByDisplayValue("Send Kim the Discovery materials today")).toBeTruthy();
    expect((screen.getAllByRole("checkbox") as HTMLInputElement[]).map((box) => box.checked)).toEqual([true, false, true]);
    expect(screen.getByText(STATUS_RESTORED)).toBeTruthy();
    expect(fetchSpy).toHaveBeenCalledTimes(1); // still exactly the one original Analyze call
  });

  it("Log as work opens the EXISTING wizard, performs zero provider writes until explicit Save, and keeps the candidate", async () => {
    const { user, createSpy } = await setup();
    await user.click(screen.getAllByRole("button", { name: "Log as work" })[0]);
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByDisplayValue("Met with North Schuylkill about science resources")).toBeTruthy();
    expect(createSpy).not.toHaveBeenCalled();
    expect(sessionFake.data.has(VOICE_SESSION_STORAGE_KEY)).toBe(true); // persisted before/while the wizard is open
    expect(screen.getByText("3 candidates")).toBeTruthy(); // Voice state untouched behind the wizard
  });

  it("cancelling the Work Record wizard leaves the Voice session recoverable and marks nothing Logged", async () => {
    const { user, createSpy, goTo } = await setup();
    await user.click(screen.getAllByRole("button", { name: "Log as work" })[0]);
    const dialog = await screen.findByRole("dialog");
    await user.click(within(dialog).getByRole("button", { name: "Close work entry" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(createSpy).not.toHaveBeenCalled();
    expect(screen.queryByText(/Logged/)).toBeFalsy();

    await goTo(/projects/i);
    await goTo(/voice intelligence/i);
    expect(screen.getByText("3 candidates")).toBeTruthy();
    expect(screen.queryByText(/Logged/)).toBeFalsy();
  });

  it("a SUCCESSFUL save marks only the originating candidate Logged, via exactly one existing createWorkRecord call, and Logged survives returning to Voice", async () => {
    const { user, fetchSpy, createSpy, goTo } = await setup();
    await user.click(screen.getAllByRole("button", { name: "Log as work" })[1]); // "Drafted the STEELS newsletter"
    await completeWizardAndSave(user);
    await waitFor(() => expect(createSpy).toHaveBeenCalledTimes(1));
    expect(createSpy.mock.calls[0][0].title).toBe("Drafted the STEELS newsletter");

    // The app moves to Today after a successful save (existing behavior); return to Voice.
    await waitFor(() => expect(screen.queryByText("3 candidates")).toBeFalsy());
    await goTo(/voice intelligence/i);

    expect(screen.getByText("3 candidates")).toBeTruthy();
    const cards = screen.getAllByLabelText("Candidate title").map((input) => input.closest(".candidate-card") as HTMLElement);
    expect(within(cards[1]).getByText("Logged ✓")).toBeTruthy();
    expect(within(cards[0]).queryByText("Logged ✓")).toBeFalsy();
    expect(within(cards[2]).queryByText("Logged ✓")).toBeFalsy();
    expect(screen.getAllByText("Logged ✓")).toHaveLength(1);
    expect(createSpy).toHaveBeenCalledTimes(1); // Logged tracking created no second Work Record
    expect(fetchSpy).toHaveBeenCalledTimes(1); // ...and made no Anthropic call
  });

  it("does not touch localStorage anywhere in the flow, and the only sessionStorage key written is the Voice session key", async () => {
    const { user } = await setup();
    await user.click(screen.getAllByRole("button", { name: "Log as work" })[0]);
    await completeWizardAndSave(user);
    await waitFor(() => expect(screen.queryByText("3 candidates")).toBeFalsy());
    expect(localFake.setItemCalls).toEqual([]);
    expect(localFake.removeItemCalls).toEqual([]);
    expect(sessionFake.setItemCalls.length).toBeGreaterThan(0);
    for (const [key] of sessionFake.setItemCalls) expect(key).toBe(VOICE_SESSION_STORAGE_KEY);
  });
});

describe("architecture boundaries", () => {
  it("Voice session persistence involves no SharePoint/Graph/DataProvider import and no localStorage", () => {
    for (const path of ["app/VoiceIntelligence.tsx", "lib/voice-intelligence-session.ts"]) {
      const code = readFileSync(path, "utf-8").replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
      expect(code, path).not.toMatch(/from ["'][^"']*(sharepoint|microsoft-graph|microsoft-auth|data-provider)/i);
      expect(code, path).not.toMatch(/localStorage|indexedDB|document\.cookie/);
    }
  });

  it("the only network call in VoiceIntelligence is the explicit analyze() request to /api/voice-intelligence", () => {
    const code = readFileSync("app/VoiceIntelligence.tsx", "utf-8");
    expect(code.match(/fetch\(/g)).toHaveLength(1);
    expect(code).toContain('fetch("/api/voice-intelligence"');
  });

  it("the Log Work save path is unchanged: Voice adds no Work Record fields and the Work Record model has no Voice source id", () => {
    expect(readFileSync("lib/models.ts", "utf-8")).not.toMatch(/voice|loggedWorkRecordAppId|sourceCandidate/i);
    expect(readFileSync("lib/voice-intelligence-work-record.ts", "utf-8")).not.toMatch(/loggedWorkRecordAppId|sessionStorage/);
  });
});
