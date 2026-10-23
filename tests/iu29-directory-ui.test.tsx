// @vitest-environment jsdom
//
// Patch 9B — the real IU29 directory as the human sees it, across every consumer of the
// Organization set: Log Work (district + partner pickers), editing a HISTORIC record that links
// to a retired placeholder, Contacts, Inbox Intelligence matching/display, Contact Import, and a
// durable Organization that shares a seed's name. All in-memory providers — no SharePoint, no AI.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import IUWorkTracker from "../app/IUWorkTracker";
import { buildContactDraft, MemoryContactProvider } from "../lib/contact-provider";
import { MemoryDataProvider } from "../lib/data-provider";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { resolveEmailAnalysisEntities, type EmailAnalysis } from "../lib/inbox-intelligence-models";
import { resolveOrganizationByExactName } from "../lib/contact-import";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import { buildOrganizationDraft, MemoryOrganizationProvider } from "../lib/organization-provider";
import { REFERENCE_DATA } from "../lib/reference-data";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

const DISTRICTS = [
  "Blue Mountain School District",
  "Mahanoy Area School District",
  "Minersville Area School District",
  "North Schuylkill School District",
  "Pine Grove Area School District",
  "Pottsville Area School District",
  "Saint Clair Area School District",
  "Schuylkill Haven Area School District",
  "Shenandoah Valley School District",
  "Tamaqua Area School District",
  "Tri-Valley School District",
  "Williams Valley School District",
];
const RETIRED = ["North Valley SD", "Riverbend Area SD", "FutureWorks Partnership"];

function historic(overrides: Partial<WorkRecord> = {}): WorkRecord {
  return {
    appId: "historic-record",
    title: "Historic planning meeting",
    activityDate: "2026-08-20",
    activityType: "District meeting",
    description: "Logged before the directory cleanup",
    detailedNotes: "",
    durationMinutes: 60,
    status: "complete",
    engagementScope: "specific",
    projectIds: [],
    organizationIds: ["org-north-valley", "org-futureworks"],
    contactIds: [],
    categoryIds: [],
    reach: { educatorsLeaders: 0, studentsFamilies: 0, workforceCommunity: 0, other: 0 },
    evidenceSummary: "",
    evidenceReferenceIds: [],
    output: "",
    outcome: "",
    nextStep: "",
    followUpNeeded: false,
    followUpDate: null,
    orbit: { reportable: false, primaryDeliverable: null, supportingDeliverables: [], stemPocMinutes: 0, tacMinutes: 0, evidence: "" },
    schemaVersion: WORK_RECORD_SCHEMA_VERSION,
    metadata: { providerId: "memory:1", version: 1, createdAt: "2026-08-20T12:00:00Z", modifiedAt: "2026-08-20T12:00:00Z", syncState: "saved" },
    isSample: false,
    ...overrides,
  };
}

function renderApp(
  overrides: {
    dataProvider?: MemoryDataProvider;
    contactDataProvider?: MemoryContactProvider;
    organizationDataProvider?: MemoryOrganizationProvider;
    inboxDataProvider?: SessionInboxIntelligenceProvider;
  } = {},
) {
  const dataProvider = overrides.dataProvider ?? new MemoryDataProvider([]);
  const contactDataProvider = overrides.contactDataProvider ?? new MemoryContactProvider();
  const organizationDataProvider = overrides.organizationDataProvider ?? new MemoryOrganizationProvider();
  const inboxDataProvider = overrides.inboxDataProvider ?? new SessionInboxIntelligenceProvider();
  render(
    <IUWorkTracker
      dataProvider={dataProvider}
      contactDataProvider={contactDataProvider}
      organizationDataProvider={organizationDataProvider}
      inboxDataProvider={inboxDataProvider}
    />,
  );
  return { dataProvider, contactDataProvider, organizationDataProvider, inboxDataProvider, user: userEvent.setup() };
}

const goTo = async (user: ReturnType<typeof userEvent.setup>, name: RegExp) =>
  user.click(within(await screen.findByRole("navigation")).getByRole("button", { name }));

/** New Log Work entry → step 2 ("Who was it for or with?") with the District(s) picker open. */
async function openNewEntryStep2(user: ReturnType<typeof userEvent.setup>) {
  await user.click((await screen.findAllByRole("button", { name: /log work/i }))[0]);
  const dialog = await screen.findByRole("dialog");
  await user.type(within(dialog).getByPlaceholderText(/district steels planning/i), "A new entry");
  const type = within(dialog).getByLabelText(/Activity type/i) as HTMLSelectElement;
  await user.selectOptions(type, type.options[1].value);
  await user.click(within(dialog).getByRole("button", { name: /continue/i }));
  return dialog;
}

const pickerButtons = (fieldset: HTMLElement) => within(fieldset).getAllByRole("button").map((button) => button.textContent!.replace("✓", "").replace(/^[▦◇○]/, "").trim());

describe("Log Work — the new directory in the pickers", () => {
  it("the District(s) picker (Specific scope) lists exactly the twelve IU29 districts — no IU, no PAIMS, no fictional ones", async () => {
    const { user } = renderApp();
    const dialog = await openNewEntryStep2(user);
    await user.click(within(dialog).getByRole("button", { name: /specific/i }));
    const fieldset = within(dialog).getByText(/District\(s\)/).closest("fieldset")!;
    expect(pickerButtons(fieldset)).toEqual(DISTRICTS);
  });

  it("the Organization / partner picker shows PAIMS and Schuylkill Intermediate Unit 29 — and no districts, no fictional placeholders", async () => {
    const { user } = renderApp();
    const dialog = await openNewEntryStep2(user);
    const fieldset = within(dialog).getByText(/Organization \/ partner/).closest("fieldset")!;
    const names = pickerButtons(fieldset);
    expect(names).toEqual(["Schuylkill Intermediate Unit 29", "PAIMS"]);
    for (const district of DISTRICTS) expect(names).not.toContain(district);
  });

  it("PAIMS is selectable as a partner on a NEW entry, with Specific scope not required", async () => {
    const { user } = renderApp();
    const dialog = await openNewEntryStep2(user);
    await user.click(within(dialog).getByRole("button", { name: /PAIMS/ }));
    expect(within(dialog).getByRole("button", { name: /PAIMS/ }).className).toContain("selected");
  });

  it("no retired placeholder is offered anywhere on a new entry", async () => {
    const { user } = renderApp();
    const dialog = await openNewEntryStep2(user);
    await user.click(within(dialog).getByRole("button", { name: /specific/i }));
    for (const name of RETIRED) expect(within(dialog).queryByText(new RegExp(name))).toBeFalsy();
  });

  it("a durable Organization that shares a seed's name appears ONCE in the picker (it replaces the seed)", async () => {
    const organizationDataProvider = new MemoryOrganizationProvider();
    await organizationDataProvider.create(buildOrganizationDraft({ appId: "durable-paims", name: "PAIMS", type: "partner" })); // pre-seed BEFORE render
    const { user } = renderApp({ organizationDataProvider });
    const dialog = await openNewEntryStep2(user);
    await waitFor(() => expect(within(dialog).getAllByRole("button", { name: /PAIMS/ })).toHaveLength(1));
    expect(pickerButtons(within(dialog).getByText(/Organization \/ partner/).closest("fieldset")!)).toEqual(["Schuylkill Intermediate Unit 29", "PAIMS"]);
  });
});

describe("a historic record that links to retired placeholders", () => {
  async function openHistoric(user: ReturnType<typeof userEvent.setup>) {
    await goTo(user, /history/i);
    await user.click((await screen.findAllByText("Historic planning meeting"))[0]);
    const dialog = await screen.findByRole("dialog");
    await user.click(within(dialog).getByRole("button", { name: /continue/i })); // → step 2
    return dialog;
  }

  it("still READS correctly: the original names resolve in the history list (never 'Unassigned')", async () => {
    const { user } = renderApp({ dataProvider: new MemoryDataProvider([historic()]) });
    await goTo(user, /history/i);
    expect((await screen.findAllByText(/North Valley SD/)).length).toBeGreaterThan(0);
    expect(screen.queryByText("Unassigned")).toBeFalsy();
  });

  it("shows ONLY the retired placeholders it already has, clearly labelled, selected — and not the other retired ones", async () => {
    const { user } = renderApp({ dataProvider: new MemoryDataProvider([historic()]) });
    const dialog = await openHistoric(user);
    const districts = within(dialog).getByText(/District\(s\)/).closest("fieldset")!;
    expect(pickerButtons(districts)).toEqual([...DISTRICTS, "North Valley SD (retired placeholder)"]);
    expect(within(districts).getByRole("button", { name: /North Valley SD/ }).className).toContain("selected");
    const partners = within(dialog).getByText(/Organization \/ partner/).closest("fieldset")!;
    expect(pickerButtons(partners)).toContain("FutureWorks Partnership (retired placeholder)");
    expect(within(dialog).queryByText(/Riverbend Area SD/)).toBeFalsy(); // not on this record → never offered
  });

  it("can be edited and saved with the retired placeholders left exactly as they were (never silently remapped)", async () => {
    const dataProvider = new MemoryDataProvider([historic()]);
    const update = vi.spyOn(dataProvider, "updateWorkRecord");
    const { user } = renderApp({ dataProvider });
    const dialog = await openHistoric(user);
    for (let i = 0; i < 3; i++) await user.click(within(dialog).getByRole("button", { name: /continue/i }));
    await user.click(within(dialog).getByRole("button", { name: "Save & done" }));
    await waitFor(() => expect(update).toHaveBeenCalledTimes(1));
    expect(update.mock.calls[0][0].organizationIds).toEqual(["org-north-valley", "org-futureworks"]);
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
  });

  it("the human can deselect a retired placeholder and choose a real district instead; once removed it cannot be re-added", async () => {
    const dataProvider = new MemoryDataProvider([historic()]);
    const update = vi.spyOn(dataProvider, "updateWorkRecord");
    const { user } = renderApp({ dataProvider });
    const dialog = await openHistoric(user);
    await user.click(within(dialog).getByRole("button", { name: /North Valley SD/ })); // deselect
    await user.click(within(dialog).getByRole("button", { name: /Blue Mountain School District/ }));
    expect(within(dialog).queryByText(/North Valley SD/)).toBeFalsy(); // gone from the picker — not re-selectable
    for (let i = 0; i < 3; i++) await user.click(within(dialog).getByRole("button", { name: /continue/i }));
    await user.click(within(dialog).getByRole("button", { name: "Save & done" }));
    await waitFor(() => expect(update).toHaveBeenCalledTimes(1));
    expect(update.mock.calls[0][0].organizationIds).toEqual(["org-futureworks", "org-blue-mountain"]);
  });
});

describe("Contacts", () => {
  async function seededContacts() {
    const contactDataProvider = new MemoryContactProvider();
    await contactDataProvider.create(buildContactDraft({ appId: "c-legacy", displayName: "Legacy Linked Person", role: undefined, organizationId: "org-north-valley", email: undefined, status: "active", notes: undefined }));
    await contactDataProvider.create(buildContactDraft({ appId: "c-real", displayName: "Real Linked Person", role: undefined, organizationId: "org-pottsville-area", email: undefined, status: "active", notes: undefined }));
    return contactDataProvider;
  }

  it("a contact already linked to a retired placeholder still shows its organization name (not 'Unknown organization')", async () => {
    const { user } = renderApp({ contactDataProvider: await seededContacts() });
    await goTo(user, /contacts/i);
    const card = (await screen.findByText("Legacy Linked Person")).closest("article")!;
    expect(within(card).getByText(/North Valley SD/)).toBeTruthy();
    expect(screen.queryByText(/Unknown organization/)).toBeFalsy();
    expect(within((await screen.findByText("Real Linked Person")).closest("article")!).getByText(/Pottsville Area School District/)).toBeTruthy();
  });

  it("editing that contact keeps its retired link visible and labelled — it is not silently lost", async () => {
    const { user } = renderApp({ contactDataProvider: await seededContacts() });
    await goTo(user, /contacts/i);
    await user.click(within((await screen.findByText("Legacy Linked Person")).closest("article")!).getByRole("button", { name: "Edit" }));
    const select = within(await screen.findByRole("dialog")).getByLabelText("Organization") as HTMLSelectElement;
    expect(select.value).toBe("org-north-valley");
    expect([...select.options].map((option) => option.textContent)).toContain("North Valley SD (retired placeholder)");
  });

  it("a NEW contact's Organization list offers the real directory only — never a retired placeholder", async () => {
    const { user } = renderApp();
    await goTo(user, /contacts/i);
    await user.click(await screen.findByRole("button", { name: /Add Contact/ }));
    const select = within(await screen.findByRole("dialog")).getByLabelText("Organization") as HTMLSelectElement;
    const names = [...select.options].map((option) => option.textContent);
    expect(names).toEqual(["No organization", "Schuylkill Intermediate Unit 29", ...DISTRICTS, "PAIMS"]);
    for (const retired of RETIRED) expect(names.join("|")).not.toContain(retired);
  });
});

describe("Inbox Intelligence matching", () => {
  const analysis = (overrides: Partial<EmailAnalysis>): EmailAnalysis => ({
    summary: "s", priority: "medium", needsAttention: false, actionItems: [], followUp: "", people: [], organizations: [], districts: [], projects: [], tags: [],
    suggestedWorkType: null, suggestedWorkRecord: { title: "t", description: "d" }, ...overrides,
  });

  it("matches an official district name and PAIMS exactly, never a retired placeholder, never across kinds", () => {
    const matches = resolveEmailAnalysisEntities(
      analysis({ districts: ["Blue Mountain School District", "North Valley SD", "PAIMS"], organizations: ["PAIMS", "Schuylkill Intermediate Unit 29", "FutureWorks Partnership", "Tamaqua Area School District"] }),
      REFERENCE_DATA,
    );
    expect(matches.districtIds).toEqual(["org-blue-mountain"]); // "PAIMS" is not a district; "North Valley SD" is retired
    expect(matches.organizationIds).toEqual(["org-paims", "org-iu"]); // a district name is not an organization; FutureWorks is retired
  });

  it("exact names only — the short 'SD' style no longer matches the official full name", () => {
    const matches = resolveEmailAnalysisEntities(analysis({ districts: ["Blue Mountain SD", "Blue Mountain"] }), REFERENCE_DATA);
    expect(matches.districtIds).toEqual([]);
  });

  it("an existing Inbox record that matched a retired district still displays its name", async () => {
    const inboxDataProvider = new SessionInboxIntelligenceProvider();
    const { buildInboxIntelligenceRecord } = await import("../lib/inbox-intelligence-models");
    const draft = buildInboxIntelligenceRecord(analysis({ suggestedWorkRecord: { title: "Legacy matched email", description: "d" } }), "x", REFERENCE_DATA, "2026-08-20T12:00:00.000Z");
    await inboxDataProvider.create({ ...draft, matchedDistrictIds: ["org-north-valley"] });
    const { user } = renderApp({ inboxDataProvider });
    await goTo(user, /inbox intelligence/i);
    const row = (await screen.findByText("Legacy matched email")).closest(".inbox-row")!;
    expect(within(row as HTMLElement).getByText("North Valley SD")).toBeTruthy();
  });
});

describe("Contact Import", () => {
  it("resolves an official district / IU / PAIMS name exactly, and leaves a retired or short-form name unresolved (never guessed)", () => {
    const resolve = (text: string) => resolveOrganizationByExactName(text, REFERENCE_DATA.organizations);
    expect(resolve("Blue Mountain School District")).toBe("org-blue-mountain");
    expect(resolve("  schuylkill   intermediate unit 29 ")).toBe("org-iu");
    expect(resolve("PAIMS")).toBe("org-paims");
    expect(resolve("North Valley SD")).toBeNull();
    expect(resolve("FutureWorks Partnership")).toBeNull();
    expect(resolve("Blue Mountain SD")).toBeNull();
    expect(resolve("Intermediate Unit")).toBeNull();
  });
});
