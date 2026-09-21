// @vitest-environment jsdom
//
// Patch 7.3 — Voice Intelligence SELECT + ROUTE. AI EXTRACTS. HUMAN SELECTS. THE SYSTEM ROUTES.
// Component-level tests use the injected `storage` seam and the REAL Memory Contact / Project /
// Organization providers behind spies (proving the existing providers are the only write path);
// full-app tests run through the real IUWorkTracker shell and Log Work wizard. Every test mocks
// fetch — ZERO real Anthropic calls, zero SharePoint.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { readFileSync } from "node:fs";
import { useRef, useState } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import IUWorkTracker from "../app/IUWorkTracker";
import VoiceIntelligence from "../app/VoiceIntelligence";
import type { AnalyzeTranscriptResult } from "../lib/anthropic-voice-analysis";
import { MemoryContactProvider } from "../lib/contact-provider";
import { MemoryDataProvider } from "../lib/data-provider";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { WORK_RECORD_SCHEMA_VERSION, type Contact, type Organization, type Project, type WorkRecord } from "../lib/models";
import { MemoryOrganizationProvider } from "../lib/organization-provider";
import { MemoryProjectProvider } from "../lib/project-provider";
import { REFERENCE_DATA } from "../lib/reference-data";
import { VOICE_SESSION_STORAGE_KEY, type VoiceSessionStorage } from "../lib/voice-intelligence-session";
import { createFakeStorage, installWindowStorage } from "./voice-session-test-utils";

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

const cand = (type: string, title: string, detail = "") => ({ type, title, detail, sourceExcerpt: `excerpt for ${title}`, durationText: null });

// Order matters — the tests index checkboxes by position.
const ANALYSIS = {
  status: "success",
  analysis: {
    candidates: [
      cand("COMPLETED_WORK", "Met with North Schuylkill about science resources", "Went well."), // 0
      cand("COMPLETED_WORK", "Drafted the STEELS newsletter", "Finished the draft."), //            1
      cand("PERSON", "Development District Lead"), //                                               2 (seeded Contact)
      cand("PERSON", "Annie Milewski"), //                                                          3 (no match)
      cand("ORGANIZATION", "FutureWorks Partnership"), //                                           4 (exact match)
      cand("DISTRICT", "Cedar Ridge SD"), //                                                        5 (no match)
      cand("PROJECT", "STEELS Implementation"), //                                                  6 (exact match)
      cand("PROJECT", "Brand New Initiative", "A fresh effort."), //                                7 (no match)
      cand("ACTION", "Send Kim the Discovery materials", "Follow-up."), //                          8
      cand("KNOWLEDGE", "Lab kits need labeled bins", "Label every bin."), //                       9
      cand("IDEA", "Shared lab-kit calendar"), //                                                   10
      cand("DECISION", "Use the spring window for pilots"), //                                      11
      cand("QUESTION", "Who owns the kit budget?"), //                                              12
    ],
  },
  usage: { model: "claude-opus-5", inputTokens: 900, outputTokens: 420 },
} as unknown as AnalyzeTranscriptResult;

const jsonResponse = (value: unknown) => new Response(JSON.stringify(value), { status: 200, headers: { "content-type": "application/json" } });

type OpenLog = (record?: WorkRecord, onSaved?: (saved: WorkRecord) => void) => void;
type Spy = ReturnType<typeof vi.fn<(entity: unknown) => void>>;
type Spies = { saveContact: Spy; saveProject: Spy; saveOrganization: Spy };
const makeSpies = (): Spies => ({ saveContact: vi.fn(), saveProject: vi.fn(), saveOrganization: vi.fn() });

/** Voice wired to the REAL Memory providers (behind spies), with references kept in sync after a create — as IUWorkTracker does. */
function Harness({ storage, openLog, spies }: { storage: VoiceSessionStorage | null; openLog: OpenLog; spies: Spies }) {
  const [contacts, setContacts] = useState<Contact[]>(REFERENCE_DATA.contacts);
  const [projects, setProjects] = useState<Project[]>(REFERENCE_DATA.projects);
  const [organizations, setOrganizations] = useState<Organization[]>(REFERENCE_DATA.organizations);
  const contactProvider = useRef(new MemoryContactProvider()).current;
  const projectProvider = useRef(new MemoryProjectProvider()).current;
  const organizationProvider = useRef(new MemoryOrganizationProvider()).current;
  return (
    <VoiceIntelligence
      openLog={openLog}
      createDraftRecord={baseWorkRecord}
      references={{ ...REFERENCE_DATA, contacts, projects, organizations }}
      storage={storage}
      saveContact={async (contact) => {
        spies.saveContact(contact);
        const result = await contactProvider.create(contact);
        if (result.status === "success") setContacts((current) => [...current, result.value]);
        return result;
      }}
      updateContact={vi.fn()}
      saveProject={async (project) => {
        spies.saveProject(project);
        const result = await projectProvider.create(project);
        if (result.status === "success") setProjects((current) => [...current, result.value]);
        return result;
      }}
      updateProject={vi.fn()}
      saveOrganization={async (organization) => {
        spies.saveOrganization(organization);
        const result = await organizationProvider.create(organization);
        if (result.status === "success") setOrganizations((current) => [...current, result.value]);
        return result;
      }}
      updateOrganization={vi.fn()}
    />
  );
}

function setup(storage: VoiceSessionStorage | null = createFakeStorage(), openLog: OpenLog = vi.fn()) {
  const spies = makeSpies();
  const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
  const view = render(<Harness storage={storage} openLog={openLog} spies={spies} />);
  const remount = () => {
    view.unmount();
    return render(<Harness storage={storage} openLog={openLog} spies={spies} />);
  };
  return { storage, openLog, spies, fetchSpy, remount, user: userEvent.setup() };
}

async function analyze(user: ReturnType<typeof userEvent.setup>) {
  await user.type(screen.getByPlaceholderText(/paste the transcript/i), "A long synthetic voice note.");
  await user.click(screen.getByRole("button", { name: "Analyze transcript" }));
  await waitFor(() => expect(screen.getByText("13 candidates")).toBeTruthy());
}

const checkboxes = () => screen.getAllByRole("checkbox") as HTMLInputElement[];

/** Deselect everything, then select exactly the given candidate positions. */
async function selectOnly(user: ReturnType<typeof userEvent.setup>, indexes: number[]) {
  await user.click(screen.getByRole("button", { name: "Clear selection" }));
  for (const index of indexes) await user.click(checkboxes()[index]);
}

const panel = () => screen.getByRole("region", { name: "Process selected" });
async function openPanel(user: ReturnType<typeof userEvent.setup>) {
  await user.click(screen.getByRole("button", { name: "Process selected" }));
  return panel();
}

// ---------------------------------------------------------------------------------------------
describe("selection tray and checkbox semantics", () => {
  it("shows no tray when nothing is selected", async () => {
    const { user } = setup();
    await analyze(user);
    await user.click(screen.getByRole("button", { name: "Clear selection" }));
    expect(screen.getByText("0 selected")).toBeTruthy();
    expect(screen.queryByRole("group", { name: "Selected candidates" })).toBeFalsy();
    expect(screen.queryByRole("button", { name: "Process selected" })).toBeFalsy();
  });

  it("shows the tray with the selected count once anything is selected, and the count follows select/unselect", async () => {
    const { user } = setup();
    await analyze(user);
    expect(screen.getByText("13 selected")).toBeTruthy();
    expect(within(screen.getByRole("group", { name: "Selected candidates" })).getByRole("button", { name: "Process selected" })).toBeTruthy();
    await user.click(checkboxes()[0]);
    expect(screen.getByText("12 selected")).toBeTruthy();
    expect(screen.getByText("1 not selected")).toBeTruthy();
    await user.click(checkboxes()[0]);
    expect(screen.getByText("13 selected")).toBeTruthy();
  });

  it("Select all appears only when something is unselected, and Clear selection only deselects", async () => {
    const { user } = setup();
    await analyze(user);
    expect(screen.queryByRole("button", { name: "Select all" })).toBeFalsy();
    await user.click(checkboxes()[3]);
    await user.click(screen.getByRole("button", { name: "Select all" }));
    expect(screen.getByText("13 selected")).toBeTruthy();
    await user.click(screen.getByRole("button", { name: "Clear selection" }));
    expect(screen.getByText("13 candidates")).toBeTruthy(); // nothing was deleted
    expect(screen.getAllByLabelText("Candidate title")).toHaveLength(13);
  });

  it("an unchecked candidate stays visible and intact; only Remove deletes", async () => {
    const { user } = setup();
    await analyze(user);
    await user.click(checkboxes()[8]); // uncheck the ACTION
    expect(screen.getByDisplayValue("Send Kim the Discovery materials")).toBeTruthy();
    expect(screen.getByText("13 candidates")).toBeTruthy();
    await user.click(screen.getAllByRole("button", { name: "Remove" })[8]);
    expect(screen.queryByDisplayValue("Send Kim the Discovery materials")).toBeFalsy();
    expect(screen.getByText("12 candidates")).toBeTruthy();
  });

  it("selection survives a remount (Patch 7.2 session persistence is preserved)", async () => {
    const { user, remount } = setup();
    await analyze(user);
    await selectOnly(user, [0, 2, 6]);
    remount();
    expect(checkboxes().map((box, index) => (box.checked ? index : -1)).filter((index) => index >= 0)).toEqual([0, 2, 6]);
    expect(screen.getByText("3 selected")).toBeTruthy();
    expect(screen.getByText("10 not selected")).toBeTruthy();
  });
});

// ---------------------------------------------------------------------------------------------
describe("Process selected — grouping and honest destinations", () => {
  it("opening the panel writes nothing and makes no AI call", async () => {
    const { user, spies, openLog, fetchSpy } = setup();
    await analyze(user);
    fetchSpy.mockClear();
    await openPanel(user);
    expect(openLog).not.toHaveBeenCalled();
    expect(spies.saveContact).not.toHaveBeenCalled();
    expect(spies.saveProject).not.toHaveBeenCalled();
    expect(spies.saveOrganization).not.toHaveBeenCalled();
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("groups the selected candidates by type with counts", async () => {
    const { user } = setup();
    await analyze(user);
    const p = await openPanel(user);
    const chips = within(within(p).getByLabelText("Selected by type"));
    for (const text of ["Completed work — 2", "Person — 2", "Organization — 1", "District — 1", "Project — 2", "Action — 1", "Knowledge — 1", "Idea — 1", "Decision — 1", "Question — 1"]) {
      expect(chips.getByText(text)).toBeTruthy();
    }
  });

  it("mixed selection: ready destinations are separated from types with no destination yet, and nothing is hidden", async () => {
    const { user } = setup();
    await analyze(user);
    await selectOnly(user, [0, 2, 4, 8, 9]); // completed work, person, organization, action, knowledge
    const p = await openPanel(user);
    expect(within(p).getByText("5 selected. Nothing is written until you review and explicitly save each item — opening this panel changes nothing.")).toBeTruthy();
    expect(within(p).getByText("Ready destinations")).toBeTruthy();
    expect(within(p).getByRole("region", { name: "Completed work destination" })).toBeTruthy();
    expect(within(p).getByRole("region", { name: "Person destination" })).toBeTruthy();
    expect(within(p).getByRole("region", { name: "Organization destination" })).toBeTruthy();
    expect(within(p).getByText("No destination yet")).toBeTruthy();
    expect(within(p).getByRole("region", { name: "Action — no destination yet" })).toBeTruthy();
    expect(within(p).getByRole("region", { name: "Knowledge — no destination yet" })).toBeTruthy();
    // Unselected types never appear.
    expect(within(p).queryByRole("region", { name: /Project destination/ })).toBeFalsy();
    expect(within(p).queryByRole("region", { name: /Idea/ })).toBeFalsy();
  });

  it.each([
    ["Action", 8],
    ["Knowledge", 9],
    ["Idea", 10],
    ["Decision", 11],
    ["Question", 12],
  ])("%s says its destination is not yet implemented and stays intact with zero provider writes", async (label, index) => {
    const { user, spies, openLog } = setup();
    await analyze(user);
    const title = (screen.getAllByLabelText("Candidate title")[index] as HTMLInputElement).value;
    await selectOnly(user, [index]);
    const p = await openPanel(user);
    expect(within(p).getByText(`${label} — destination not yet implemented`)).toBeTruthy();
    expect(within(p).getByText(title)).toBeTruthy();
    await user.click(within(p).getByRole("button", { name: "Done" }));
    expect(screen.getByDisplayValue(title)).toBeTruthy();
    expect(checkboxes()[index].checked).toBe(true); // still selected
    expect(spies.saveContact).not.toHaveBeenCalled();
    expect(spies.saveProject).not.toHaveBeenCalled();
    expect(spies.saveOrganization).not.toHaveBeenCalled();
    expect(openLog).not.toHaveBeenCalled();
  });

  it("Copy puts a no-destination candidate on the clipboard as plain text (and writes nothing durable)", async () => {
    const { user, spies } = setup(); // userEvent.setup() installs navigator.clipboard, so spy AFTER setup
    const writeText = vi.spyOn(navigator.clipboard, "writeText").mockResolvedValue(undefined);
    await analyze(user);
    await selectOnly(user, [9]);
    const p = await openPanel(user);
    await user.click(within(p).getByRole("button", { name: "Copy" }));
    expect(writeText).toHaveBeenCalledWith("Knowledge: Lab kits need labeled bins\nLabel every bin.");
    expect(await within(p).findByRole("button", { name: "Copied" })).toBeTruthy();
    expect(spies.saveProject).not.toHaveBeenCalled();
  });

  it("Done closes the panel without changing selection or any candidate", async () => {
    const { user } = setup();
    await analyze(user);
    await selectOnly(user, [0, 8]);
    const p = await openPanel(user);
    await user.click(within(p).getByRole("button", { name: "Done" }));
    expect(screen.queryByRole("region", { name: "Process selected" })).toBeFalsy();
    expect(screen.getByText("2 selected")).toBeTruthy();
    expect(screen.getByText("13 candidates")).toBeTruthy();
  });

  it("deselecting the last selected candidate closes the panel (nothing left to process)", async () => {
    const { user } = setup();
    await analyze(user);
    await selectOnly(user, [8]);
    await openPanel(user);
    await user.click(checkboxes()[8]);
    expect(screen.queryByRole("region", { name: "Process selected" })).toBeFalsy();
  });
});

// ---------------------------------------------------------------------------------------------
describe("Completed work → the existing Work Record wizard, sequentially", () => {
  it("offers the Work Records destination and creates nothing on Process selected alone", async () => {
    const { user, openLog } = setup();
    await analyze(user);
    await selectOnly(user, [0, 1]);
    const p = await openPanel(user);
    const region = within(within(p).getByRole("region", { name: "Completed work destination" }));
    expect(region.getByText(/Work Records/)).toBeTruthy();
    expect(region.getByText("Completed work 1 of 2")).toBeTruthy();
    expect(region.getByRole("button", { name: "Log this work" })).toBeTruthy();
    expect(openLog).not.toHaveBeenCalled();
  });

  it("Log this work uses the EXISTING openLog handoff for the current item, then advances only after a real successful save", async () => {
    let onSaved: ((saved: WorkRecord) => void) | undefined;
    const openLog = vi.fn((_record?: WorkRecord, callback?: (saved: WorkRecord) => void) => {
      onSaved = callback;
    });
    const { user } = setup(createFakeStorage(), openLog);
    await analyze(user);
    await selectOnly(user, [0, 1]);
    const p = await openPanel(user);
    const region = () => within(within(panel()).getByRole("region", { name: "Completed work destination" }));

    await user.click(region().getByRole("button", { name: "Log this work" }));
    expect(openLog).toHaveBeenCalledTimes(1);
    expect(openLog.mock.calls[0][0]!.title).toBe("Met with North Schuylkill about science resources");
    expect(screen.queryByText(/Logged/)).toBeFalsy(); // clicking alone marks nothing
    expect(region().getByText("Completed work 1 of 2")).toBeTruthy(); // cancel/no save → still the current item

    onSaved!({ ...baseWorkRecord(), appId: "work-record-1" });
    await waitFor(() => expect(region().getByText("Completed work 2 of 2")).toBeTruthy());
    expect(region().getAllByText("Drafted the STEELS newsletter").length).toBeGreaterThan(0); // current card + queue list
    expect(within(p).getAllByText("Logged ✓").length).toBeGreaterThan(0);
    expect(screen.getAllByLabelText("Candidate title").map((i) => i.closest(".candidate-card")!.textContent).filter((t) => t!.includes("Logged ✓"))).toHaveLength(1);
  });

  it("Skip moves to the next selected item, loses nothing, and skipped items can be shown again", async () => {
    const { user, openLog } = setup();
    await analyze(user);
    await selectOnly(user, [0, 1]);
    await openPanel(user);
    const region = () => within(within(panel()).getByRole("region", { name: "Completed work destination" }));
    await user.click(region().getByRole("button", { name: "Skip" }));
    expect(region().getByText("Completed work 2 of 2")).toBeTruthy();
    await user.click(region().getByRole("button", { name: "Skip" }));
    expect(region().getByText(/2 skipped for now/)).toBeTruthy();
    await user.click(region().getByRole("button", { name: "Show skipped again" }));
    expect(region().getByText("Completed work 1 of 2")).toBeTruthy();
    expect(openLog).not.toHaveBeenCalled();
    expect(screen.getByDisplayValue("Met with North Schuylkill about science resources")).toBeTruthy(); // nothing removed
  });

  it("when every selected completed-work item is logged, the panel says so", async () => {
    let onSaved: ((saved: WorkRecord) => void) | undefined;
    const openLog = vi.fn((_r?: WorkRecord, cb?: (saved: WorkRecord) => void) => {
      onSaved = cb;
    });
    const { user } = setup(createFakeStorage(), openLog);
    await analyze(user);
    await selectOnly(user, [0]);
    await openPanel(user);
    await user.click(within(panel()).getByRole("button", { name: "Log this work" }));
    onSaved!({ ...baseWorkRecord(), appId: "wr" });
    expect(await within(panel()).findByText("All selected completed work is logged ✓")).toBeTruthy();
  });
});

// ---------------------------------------------------------------------------------------------
describe("Person → the existing Contact matching (no second workflow)", () => {
  it("shows the existing Match Existing / Add Person / Ignore review and never creates a Contact by itself", async () => {
    const { user, spies } = setup();
    await analyze(user);
    await selectOnly(user, [2, 3]);
    const region = within(within(await openPanel(user)).getByRole("region", { name: "Person destination" }));
    expect(region.getAllByRole("button", { name: "Match Existing" }).length).toBeGreaterThan(0);
    expect(region.getAllByRole("button", { name: "Add Person" }).length).toBeGreaterThan(0);
    expect(region.getAllByRole("button", { name: "Ignore" }).length).toBeGreaterThan(0);
    expect(spies.saveContact).not.toHaveBeenCalled();
  });

  it("Match Existing in the panel records the reviewed decision in the Voice session, with no Contact write", async () => {
    const { user, spies, storage, remount } = setup();
    await analyze(user);
    await selectOnly(user, [2]);
    const region = within(within(await openPanel(user)).getByRole("region", { name: "Person destination" }));
    await user.click(region.getByRole("button", { name: "Match Existing" }));
    expect(region.getByText(/Matched Contact: Development District Lead/)).toBeTruthy();
    expect(spies.saveContact).not.toHaveBeenCalled();
    const stored = JSON.parse((storage as ReturnType<typeof createFakeStorage>).data.get(VOICE_SESSION_STORAGE_KEY)!);
    expect(stored.candidates[2].contactDecision.type).toBe("matched");
    remount();
    expect(within(screen.getByRole("region", { name: "Person destination" })).getByText(/Matched Contact: Development District Lead/)).toBeTruthy();
  });

  it("an existing reviewed match (decided in the card before processing) is preserved and shown", async () => {
    const { user } = setup();
    await analyze(user);
    const card = screen.getAllByLabelText("Candidate title")[2].closest(".candidate-card") as HTMLElement;
    await user.click(within(card).getByRole("button", { name: "Match Existing" }));
    await selectOnly(user, [2]);
    // selectOnly toggles selection but must not disturb the decision:
    const region = within(within(await openPanel(user)).getByRole("region", { name: "Person destination" }));
    expect(region.getByText(/Matched Contact: Development District Lead/)).toBeTruthy();
  });

  it("Add Person from the panel uses the existing Contact form + provider once, and only an explicit Save marks it created", async () => {
    const { user, spies } = setup();
    await analyze(user);
    await selectOnly(user, [3]);
    const region = within(within(await openPanel(user)).getByRole("region", { name: "Person destination" }));
    await user.click(region.getByRole("button", { name: "Add Person" }));
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByDisplayValue("Annie Milewski")).toBeTruthy();
    expect(spies.saveContact).not.toHaveBeenCalled(); // opening the form persists nothing
    await user.click(within(dialog).getByRole("button", { name: "Add Contact" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(spies.saveContact).toHaveBeenCalledTimes(1);
    expect(region.getByText(/Matched Contact: Annie Milewski/)).toBeTruthy();
    const card = screen.getAllByLabelText("Candidate title")[3].closest(".candidate-card") as HTMLElement;
    expect(within(card).getByText("Created Contact ✓")).toBeTruthy();
  });

  it("a stale 'Created Contact ✓' never survives a later Match Existing made from the candidate card", async () => {
    const { user } = setup();
    await analyze(user);
    const card = () => screen.getAllByLabelText("Candidate title")[2].closest(".candidate-card") as HTMLElement; // "Development District Lead" (has a seeded match)
    await user.click(within(card()).getByRole("button", { name: "Add Person" }));
    await user.click(within(await screen.findByRole("dialog")).getByRole("button", { name: "Add Contact" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(within(card()).getByText("Created Contact ✓")).toBeTruthy();
    await user.click(within(card()).getByRole("button", { name: "Change" })); // reset the decision from the card
    await user.click(within(card()).getAllByRole("button", { name: "Match Existing" })[0]); // now matched to an EXISTING Contact
    expect(within(card()).getByText(/Matched Contact:/)).toBeTruthy();
    expect(within(card()).queryByText("Created Contact ✓")).toBeFalsy(); // must not still claim "created"
  });

  it("cancelling Add Person creates nothing and marks nothing", async () => {
    const { user, spies } = setup();
    await analyze(user);
    await selectOnly(user, [3]);
    const region = within(within(await openPanel(user)).getByRole("region", { name: "Person destination" }));
    await user.click(region.getByRole("button", { name: "Add Person" }));
    await user.click(within(await screen.findByRole("dialog")).getByRole("button", { name: "Cancel" }));
    expect(spies.saveContact).not.toHaveBeenCalled();
    expect(screen.queryByText("Created Contact ✓")).toBeFalsy();
  });
});

// ---------------------------------------------------------------------------------------------
describe("Organization / District → the existing Organization path", () => {
  it("shows an exact-name existing Organization for review and never creates one automatically", async () => {
    const { user, spies } = setup();
    await analyze(user);
    await selectOnly(user, [4, 5]);
    const p = await openPanel(user);
    const org = within(within(p).getByRole("region", { name: "Organization destination" }));
    expect(org.getByText("FutureWorks Partnership", { selector: "span" })).toBeTruthy();
    expect(org.getByRole("button", { name: "Match Existing" })).toBeTruthy();
    const district = within(within(p).getByRole("region", { name: "District destination" }));
    expect(district.getByText("No exact match found.")).toBeTruthy();
    expect(spies.saveOrganization).not.toHaveBeenCalled();
  });

  it("Match Existing links the existing Organization in the session only — no write", async () => {
    const { user, spies } = setup();
    await analyze(user);
    await selectOnly(user, [4]);
    const org = within(within(await openPanel(user)).getByRole("region", { name: "Organization destination" }));
    await user.click(org.getByRole("button", { name: "Match Existing" }));
    expect(org.getByText("Matched Organization ✓: FutureWorks Partnership")).toBeTruthy();
    expect(spies.saveOrganization).not.toHaveBeenCalled();
    const card = screen.getAllByLabelText("Candidate title")[4].closest(".candidate-card") as HTMLElement;
    expect(within(card).getByText("Matched Organization ✓")).toBeTruthy();
    await user.click(org.getByRole("button", { name: "Change" })); // reversible: nothing durable happened
    expect(org.getByRole("button", { name: "Match Existing" })).toBeTruthy();
  });

  it("Create Organization opens the existing form (District → type district), and only Save writes — once, through the existing provider", async () => {
    const { user, spies, remount } = setup();
    await analyze(user);
    await selectOnly(user, [5]);
    const district = within(within(await openPanel(user)).getByRole("region", { name: "District destination" }));
    await user.click(district.getByRole("button", { name: "Create Organization" }));
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByRole("heading", { name: "Add Organization" })).toBeTruthy();
    expect((within(dialog).getByLabelText(/^Name/) as HTMLInputElement).value).toBe("Cedar Ridge SD");
    expect((within(dialog).getByLabelText("Type") as HTMLSelectElement).value).toBe("district");
    expect(spies.saveOrganization).not.toHaveBeenCalled();
    await user.click(within(dialog).getByRole("button", { name: "Add Organization" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(spies.saveOrganization).toHaveBeenCalledTimes(1);
    expect(spies.saveOrganization.mock.calls[0][0]).toMatchObject({ name: "Cedar Ridge SD", type: "district" });
    expect(district.getByText("Created Organization ✓: Cedar Ridge SD")).toBeTruthy();
    remount(); // processed state survives navigation/remount
    expect(within(screen.getByRole("region", { name: "District destination" })).getByText("Created Organization ✓: Cedar Ridge SD")).toBeTruthy();
  });

  it("Ignore records a session-only decision", async () => {
    const { user, spies } = setup();
    await analyze(user);
    await selectOnly(user, [5]);
    const district = within(within(await openPanel(user)).getByRole("region", { name: "District destination" }));
    await user.click(district.getByRole("button", { name: "Ignore" }));
    expect(district.getByText("Ignored for this item")).toBeTruthy();
    expect(spies.saveOrganization).not.toHaveBeenCalled();
  });
});

// ---------------------------------------------------------------------------------------------
describe("Project → the existing Project path with exact-name matching only", () => {
  it("proposes the exact-name existing Project, never fuzzy, and never creates a Project automatically", async () => {
    const { user, spies } = setup();
    await analyze(user);
    await selectOnly(user, [6, 7]);
    const project = within(within(await openPanel(user)).getByRole("region", { name: "Project destination" }));
    expect(project.getByText("STEELS Implementation", { selector: "li > span" })).toBeTruthy(); // the exact match
    expect(project.getByText("No exact match found.")).toBeTruthy(); // "Brand New Initiative"
    expect(spies.saveProject).not.toHaveBeenCalled();
  });

  it("Match Existing links the Project in the session only", async () => {
    const { user, spies } = setup();
    await analyze(user);
    await selectOnly(user, [6]);
    const project = within(within(await openPanel(user)).getByRole("region", { name: "Project destination" }));
    await user.click(project.getByRole("button", { name: "Match Existing" }));
    expect(project.getByText("Matched Project ✓: STEELS Implementation")).toBeTruthy();
    expect(spies.saveProject).not.toHaveBeenCalled();
  });

  it("Create Project opens the existing form prefilled with the candidate, and only Save writes — once", async () => {
    const { user, spies } = setup();
    await analyze(user);
    await selectOnly(user, [7]);
    const project = within(within(await openPanel(user)).getByRole("region", { name: "Project destination" }));
    await user.click(project.getByRole("button", { name: "Create Project" }));
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByRole("heading", { name: "Create Project" })).toBeTruthy();
    expect((within(dialog).getByLabelText(/Project name/) as HTMLInputElement).value).toBe("Brand New Initiative");
    expect((within(dialog).getByLabelText("Description") as HTMLTextAreaElement).value).toBe("A fresh effort.");
    expect(spies.saveProject).not.toHaveBeenCalled();
    await user.click(within(dialog).getByRole("button", { name: "Create Project" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(spies.saveProject).toHaveBeenCalledTimes(1);
    expect(project.getByText("Created Project ✓: Brand New Initiative")).toBeTruthy();
  });

  it("cancelling Create Project writes nothing and marks nothing", async () => {
    const { user, spies } = setup();
    await analyze(user);
    await selectOnly(user, [7]);
    const project = within(within(await openPanel(user)).getByRole("region", { name: "Project destination" }));
    await user.click(project.getByRole("button", { name: "Create Project" }));
    await user.click(within(await screen.findByRole("dialog")).getByRole("button", { name: "Cancel" }));
    expect(spies.saveProject).not.toHaveBeenCalled();
    expect(screen.queryByText("Created Project ✓")).toBeFalsy();
  });
});

// ---------------------------------------------------------------------------------------------
describe("session persistence of processing state, Clear, and boundaries", () => {
  it("the open panel, selection, and every processed state survive a remount", async () => {
    const { user, remount } = setup();
    await analyze(user);
    await selectOnly(user, [4, 6]);
    await openPanel(user);
    await user.click(within(within(panel()).getByRole("region", { name: "Organization destination" })).getByRole("button", { name: "Match Existing" }));
    await user.click(within(within(panel()).getByRole("region", { name: "Project destination" })).getByRole("button", { name: "Ignore" }));
    remount();
    expect(screen.getByRole("region", { name: "Process selected" })).toBeTruthy(); // panel reopened
    expect(screen.getByText("Matched Organization ✓: FutureWorks Partnership")).toBeTruthy();
    expect(within(screen.getByRole("region", { name: "Project destination" })).getByText("Ignored for this item")).toBeTruthy();
    expect(screen.getByText("2 selected")).toBeTruthy();
  });

  it("Analyze another transcript destroys the whole working session, routing state included", async () => {
    const { user, storage } = setup();
    await analyze(user);
    await selectOnly(user, [4]);
    await openPanel(user);
    await user.click(screen.getByRole("button", { name: "Analyze another transcript" }));
    expect((storage as ReturnType<typeof createFakeStorage>).data.has(VOICE_SESSION_STORAGE_KEY)).toBe(false);
    expect(screen.queryByRole("region", { name: "Process selected" })).toBeFalsy();
  });

  it("processing makes ZERO Anthropic calls — the single analyze fetch is the only one", async () => {
    const { user, fetchSpy } = setup();
    await analyze(user);
    await openPanel(user);
    await user.click(within(within(panel()).getByRole("region", { name: "Organization destination" })).getByRole("button", { name: "Match Existing" }));
    await user.click(within(within(panel()).getByRole("region", { name: "Project destination" })).getAllByRole("button", { name: "Ignore" })[0]);
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });

  it("touches no localStorage, and writes only the Voice session key to sessionStorage", async () => {
    const localFake = createFakeStorage();
    const sessionFake = createFakeStorage();
    restores.push(installWindowStorage("localStorage", localFake));
    restores.push(installWindowStorage("sessionStorage", sessionFake));
    const spies = makeSpies();
    vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    render(<Harness storage={undefined as unknown as null} openLog={vi.fn()} spies={spies} />); // undefined → default window.sessionStorage
    const user = userEvent.setup();
    await analyze(user);
    await openPanel(user);
    await user.click(within(within(panel()).getByRole("region", { name: "Organization destination" })).getByRole("button", { name: "Match Existing" }));
    expect(localFake.setItemCalls).toEqual([]);
    expect(sessionFake.setItemCalls.length).toBeGreaterThan(0);
    for (const [key] of sessionFake.setItemCalls) expect(key).toBe(VOICE_SESSION_STORAGE_KEY);
  });
});

describe("no duplicate paths (static)", () => {
  const strip = (path: string) => readFileSync(path, "utf-8").replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");

  it("Voice + the routing panel construct no provider and call no create/save function directly", () => {
    for (const path of ["app/VoiceIntelligence.tsx", "app/VoiceRoutingPanel.tsx", "lib/voice-routing.ts"]) {
      const code = strip(path);
      expect(code, path).not.toMatch(/new (Memory|Delegated)\w*Provider|select\w+Provider\(/);
      expect(code, path).not.toMatch(/\b(saveProject|saveOrganization|saveContact|updateProject|updateOrganization|updateContact|createWorkRecord)\s*\(/);
      expect(code, path).not.toMatch(/localStorage|indexedDB|document\.cookie|sharepoint/i);
    }
  });

  it("the Project/Organization forms live in ONE extracted file each, shared with the Projects/Organizations screens", () => {
    const tracker = strip("app/IUWorkTracker.tsx");
    expect(tracker).not.toMatch(/function (ProjectFormModal|OrganizationFormModal)\b/);
    expect(tracker).toMatch(/import ProjectFormModal/);
    expect(tracker).toMatch(/import OrganizationFormModal/);
    expect(strip("app/VoiceIntelligence.tsx")).toMatch(/import ProjectFormModal/);
    expect(strip("app/VoiceIntelligence.tsx")).toMatch(/import OrganizationFormModal/);
    expect(strip("app/ProjectFormModal.tsx")).toMatch(/export default function ProjectFormModal/);
    expect(strip("app/OrganizationFormModal.tsx")).toMatch(/export default function OrganizationFormModal/);
  });

  it("Voice still adds no Voice field to the durable Work Record model and no Anthropic prompt/model change", () => {
    expect(readFileSync("lib/models.ts", "utf-8")).not.toMatch(/voice|loggedWorkRecordAppId|routing/i);
    expect(strip("app/VoiceIntelligence.tsx").match(/fetch\(/g)).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------------------------
describe("full app — the real wizard, sequentially, and the real Project provider", () => {
  let sessionFake: ReturnType<typeof createFakeStorage>;
  beforeEach(() => {
    sessionFake = createFakeStorage();
    restores.push(installWindowStorage("sessionStorage", sessionFake));
    restores.push(installWindowStorage("localStorage", createFakeStorage()));
  });

  async function setupApp() {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(jsonResponse(ANALYSIS));
    const dataProvider = new MemoryDataProvider([]);
    const createSpy = vi.spyOn(dataProvider, "createWorkRecord");
    render(
      <IUWorkTracker
        dataProvider={dataProvider}
        inboxDataProvider={new SessionInboxIntelligenceProvider()}
        projectDataProvider={new MemoryProjectProvider()}
        organizationDataProvider={new MemoryOrganizationProvider()}
        contactDataProvider={new MemoryContactProvider()}
      />,
    );
    const nav = await screen.findByRole("navigation");
    const goTo = async (name: RegExp) => user.click(within(nav).getByRole("button", { name }));
    await goTo(/voice intelligence/i);
    await analyze(user);
    return { user, fetchSpy, createSpy, goTo };
  }

  async function completeWizard(user: ReturnType<typeof userEvent.setup>) {
    const dialog = await screen.findByRole("dialog");
    const type = within(dialog).getByLabelText(/Activity type/i) as HTMLSelectElement;
    await user.selectOptions(type, type.options[1].value);
    for (let i = 0; i < 4; i++) await user.click(within(dialog).getByRole("button", { name: /continue/i }));
    await user.click(within(dialog).getByRole("button", { name: "Save & done" }));
  }

  it("processes two completed-work items one at a time through the existing wizard: cancel marks nothing, a real save marks only its source, and the next item is waiting on return", async () => {
    const { user, fetchSpy, createSpy, goTo } = await setupApp();
    await selectOnly(user, [0, 1]);
    await openPanel(user);
    const region = () => within(within(panel()).getByRole("region", { name: "Completed work destination" }));

    // 1) open the wizard for item 1, then CANCEL → nothing logged, still item 1
    await user.click(region().getByRole("button", { name: "Log this work" }));
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByDisplayValue("Met with North Schuylkill about science resources")).toBeTruthy();
    expect(createSpy).not.toHaveBeenCalled();
    await user.click(within(dialog).getByRole("button", { name: "Close work entry" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(screen.queryByText(/Logged/)).toBeFalsy();
    expect(region().getByText("Completed work 1 of 2")).toBeTruthy();

    // 2) log item 1 for real
    await user.click(region().getByRole("button", { name: "Log this work" }));
    await completeWizard(user);
    await waitFor(() => expect(createSpy).toHaveBeenCalledTimes(1));
    expect(createSpy.mock.calls[0][0].title).toBe("Met with North Schuylkill about science resources");

    // 3) the app moves to Today (existing behavior); back in Voice the panel is still open on item 2
    await waitFor(() => expect(screen.queryByText("13 candidates")).toBeFalsy());
    await goTo(/voice intelligence/i);
    expect(screen.getByRole("region", { name: "Process selected" })).toBeTruthy();
    expect(region().getByText("Completed work 2 of 2")).toBeTruthy();
    expect(region().getAllByText("Drafted the STEELS newsletter").length).toBeGreaterThan(0);
    const logged = screen.getAllByLabelText("Candidate title").map((i) => i.closest(".candidate-card") as HTMLElement).filter((c) => c.textContent!.includes("Logged ✓"));
    expect(logged).toHaveLength(1);
    expect((within(logged[0]).getByLabelText("Candidate title") as HTMLInputElement).value).toBe("Met with North Schuylkill about science resources");
    expect(createSpy).toHaveBeenCalledTimes(1); // no bulk write, no second Work Record
    expect(fetchSpy).toHaveBeenCalledTimes(1); // still only the one Analyze
  });

  it("Create Project from Voice goes through the SAME Project provider the Projects screen uses", async () => {
    const { user, goTo } = await setupApp();
    await selectOnly(user, [7]);
    await openPanel(user);
    const project = () => within(within(panel()).getByRole("region", { name: "Project destination" }));
    await user.click(project().getByRole("button", { name: "Create Project" }));
    const dialog = await screen.findByRole("dialog");
    await user.click(within(dialog).getByRole("button", { name: "Create Project" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(project().getByText("Created Project ✓: Brand New Initiative")).toBeTruthy();

    await goTo(/^▤?\s*projects/i);
    expect(await screen.findByText("Brand New Initiative")).toBeTruthy(); // durable Project now on the Projects screen
  });
});
