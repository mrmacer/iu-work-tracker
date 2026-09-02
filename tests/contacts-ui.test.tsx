// @vitest-environment jsdom
//
// Patch 8B — Durable Contacts. Covers the Contacts directory, Create/Edit UI, deterministic
// search, duplicate detection, Universal Work Record integration, legacy seeded-contact
// regression, and the zero-AI / zero-cross-resource-write boundaries. All against
// MemoryDataProvider/MemoryContactProvider — zero real SharePoint writes, zero real Anthropic
// calls.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import IUWorkTracker from "../app/IUWorkTracker";
import { MemoryContactProvider } from "../lib/contact-provider";
import { MemoryDataProvider } from "../lib/data-provider";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { MemoryMeetingRecordProvider } from "../lib/meeting-record-provider";
import { MICROSOFT_GRAPH_SCOPES } from "../lib/microsoft-auth-config";
import { MemoryProjectProvider } from "../lib/project-provider";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

function renderApp(overrides: { dataProvider?: MemoryDataProvider; contactDataProvider?: MemoryContactProvider } = {}) {
  const dataProvider = overrides.dataProvider ?? new MemoryDataProvider([]);
  const contactDataProvider = overrides.contactDataProvider ?? new MemoryContactProvider();
  const projectDataProvider = new MemoryProjectProvider();
  const meetingDataProvider = new MemoryMeetingRecordProvider();
  const inboxDataProvider = new SessionInboxIntelligenceProvider();
  const utils = render(
    <IUWorkTracker
      dataProvider={dataProvider}
      contactDataProvider={contactDataProvider}
      projectDataProvider={projectDataProvider}
      meetingDataProvider={meetingDataProvider}
      inboxDataProvider={inboxDataProvider}
    />,
  );
  return { ...utils, dataProvider, contactDataProvider, projectDataProvider, meetingDataProvider, inboxDataProvider };
}

async function openContacts(user: ReturnType<typeof userEvent.setup>) {
  const nav = await screen.findByRole("navigation");
  await user.click(within(nav).getByRole("button", { name: /contacts/i }));
}

async function addContactFromScreen(
  user: ReturnType<typeof userEvent.setup>,
  fields: { name: string; email?: string; role?: string },
) {
  await openContacts(user);
  await user.click(screen.getByRole("button", { name: "Add Contact" }));
  const dialog = await screen.findByRole("dialog");
  await user.type(within(dialog).getByPlaceholderText(/annie milewski/i), fields.name);
  if (fields.role) await user.type(within(dialog).getByPlaceholderText(/superintendent/i), fields.role);
  if (fields.email) await user.type(within(dialog).getByPlaceholderText(/name@example\.org/i), fields.email);
  await user.click(within(dialog).getByRole("button", { name: "Add Contact" }));
  await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
}

/** Step 1 requires a title + activity type before "Continue" is enabled; the Contacts picker
 * lives on step 2 ("Who was it for or with?") — see app/IUWorkTracker.tsx LogWizard. */
async function openLogWizardAtContactStep(user: ReturnType<typeof userEvent.setup>) {
  await user.click(within(screen.getByRole("complementary")).getByRole("button", { name: /log work/i }));
  const wizard = await screen.findByRole("dialog");
  await user.type(within(wizard).getByRole("textbox", { name: /activity title/i }), "Robotics kickoff planning");
  await user.selectOptions(within(wizard).getByRole("combobox", { name: /activity type/i }), "Professional learning");
  await user.click(within(wizard).getByRole("button", { name: /continue/i }));
  return wizard;
}

describe("Legacy seeded contacts (regression)", () => {
  it("still render before any durable contact exists", async () => {
    const user = userEvent.setup();
    renderApp();
    await openContacts(user);
    for (const name of ["Development District Lead", "Development Partner Contact", "Development IU Colleague"]) {
      expect(await screen.findByText(name)).toBeTruthy();
    }
    expect(screen.getAllByRole("article")).toHaveLength(3);
  });

  it("keep their existing Work Record relationships intact", async () => {
    const user = userEvent.setup();
    renderApp({ dataProvider: new MemoryDataProvider() }); // default sample records
    await openContacts(user);
    expect(await screen.findByText("Development Partner Contact")).toBeTruthy();
  });

  it("shows no duplicate cards after a durable contact is created — three seeded plus exactly one new", async () => {
    const user = userEvent.setup();
    renderApp();
    await addContactFromScreen(user, { name: "Annie Milewski" });
    await waitFor(() => expect(screen.getAllByRole("article")).toHaveLength(4));
    expect(screen.getAllByText("Annie Milewski")).toHaveLength(1);
  });
});

describe("Add Contact", () => {
  it("has an Add Contact control that opens a compact form", async () => {
    const user = userEvent.setup();
    renderApp();
    await openContacts(user);
    expect(screen.queryByRole("dialog")).toBeFalsy();
    await user.click(screen.getByRole("button", { name: "Add Contact" }));
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByRole("heading", { name: "Add Contact" })).toBeTruthy();
  });

  it("defaults status to Active and disables submit until a name is entered", async () => {
    const user = userEvent.setup();
    renderApp();
    await openContacts(user);
    await user.click(screen.getByRole("button", { name: "Add Contact" }));
    const dialog = await screen.findByRole("dialog");
    const statusSelect = within(dialog).getByDisplayValue("Active") as HTMLSelectElement;
    expect(statusSelect.value).toBe("active");
    expect((within(dialog).getByRole("button", { name: "Add Contact" }) as HTMLButtonElement).disabled).toBe(true);
  });

  it("Cancel closes the form and writes nothing", async () => {
    const user = userEvent.setup();
    const { contactDataProvider } = renderApp();
    const createSpy = vi.spyOn(contactDataProvider, "create");
    await openContacts(user);
    await user.click(screen.getByRole("button", { name: "Add Contact" }));
    const dialog = await screen.findByRole("dialog");
    await user.type(within(dialog).getByPlaceholderText(/annie milewski/i), "Abandoned Draft");
    await user.click(within(dialog).getByRole("button", { name: "Cancel" }));
    expect(screen.queryByRole("dialog")).toBeFalsy();
    expect(createSpy).not.toHaveBeenCalled();
    expect(screen.queryByText("Abandoned Draft")).toBeFalsy();
  });

  it("typing into the form fields alone never calls create() — no autosave", async () => {
    const user = userEvent.setup();
    const { contactDataProvider } = renderApp();
    const createSpy = vi.spyOn(contactDataProvider, "create");
    await openContacts(user);
    await user.click(screen.getByRole("button", { name: "Add Contact" }));
    const dialog = await screen.findByRole("dialog");
    await user.type(within(dialog).getByPlaceholderText(/annie milewski/i), "Draft Only");
    await user.type(within(dialog).getByPlaceholderText(/a sentence or two/i), "Some notes");
    expect(createSpy).not.toHaveBeenCalled();
  });

  it("creates exactly one contact through the provider, and it appears as a normal directory card", async () => {
    const user = userEvent.setup();
    const { contactDataProvider } = renderApp();
    const createSpy = vi.spyOn(contactDataProvider, "create");
    await addContactFromScreen(user, { name: "Annie Milewski", role: "Superintendent" });

    await waitFor(() => expect(createSpy).toHaveBeenCalledTimes(1));
    const card = (await screen.findByText("Annie Milewski")).closest("article")!;
    expect(within(card).getByText("Active")).toBeTruthy();
    expect(within(card).getByText(/Superintendent/)).toBeTruthy();
  });

  it("makes zero Anthropic requests", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    renderApp();
    await addContactFromScreen(user, { name: "Annie Milewski" });
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});

describe("Edit Contact", () => {
  async function addAndOpenEdit(user: ReturnType<typeof userEvent.setup>) {
    await addContactFromScreen(user, { name: "Annie Milewski" });
    const card = (await screen.findByText("Annie Milewski")).closest("article")!;
    await user.click(within(card).getByRole("button", { name: "Edit" }));
    const dialog = await screen.findByRole("dialog");
    return { dialog, card };
  }

  it("only offers Edit on durable (newly created) contacts, not the seeded ones", async () => {
    const user = userEvent.setup();
    renderApp();
    await openContacts(user);
    const seededCard = (await screen.findByText("Development District Lead")).closest("article")!;
    expect(within(seededCard).queryByRole("button", { name: "Edit" })).toBeFalsy();
  });

  it("routes a change through update(), not a second create()", async () => {
    const user = userEvent.setup();
    const { contactDataProvider } = renderApp();
    const createSpy = vi.spyOn(contactDataProvider, "create");
    const updateSpy = vi.spyOn(contactDataProvider, "update");
    const { dialog } = await addAndOpenEdit(user);
    expect(createSpy).toHaveBeenCalledTimes(1);

    const nameInput = within(dialog).getByDisplayValue("Annie Milewski");
    await user.clear(nameInput);
    await user.type(nameInput, "Annie M. Milewski");
    await user.click(within(dialog).getByRole("button", { name: "Save Changes" }));

    await waitFor(() => expect(updateSpy).toHaveBeenCalledTimes(1));
    expect(createSpy).toHaveBeenCalledTimes(1); // still exactly once
    expect(updateSpy.mock.calls[0][1]).toBe(1); // expectedVersion from the create response
    expect(await screen.findByText("Annie M. Milewski")).toBeTruthy();
  });

  it("never exposes appId, RecordVersion, or provider metadata as editable fields", async () => {
    const user = userEvent.setup();
    renderApp();
    const { dialog } = await addAndOpenEdit(user);
    expect(within(dialog).queryByText(/appId/i)).toBeFalsy();
    expect(within(dialog).queryByText(/RecordVersion/i)).toBeFalsy();
    expect(within(dialog).queryByText(/ETag/i)).toBeFalsy();
  });

  it("updating a contact never touches Work Records", async () => {
    const user = userEvent.setup();
    const { dataProvider } = renderApp();
    const workCreateSpy = vi.spyOn(dataProvider, "createWorkRecord");
    const { dialog } = await addAndOpenEdit(user);
    await user.click(within(dialog).getByRole("button", { name: "Save Changes" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(workCreateSpy).not.toHaveBeenCalled();
  });
});

describe("Directory search and archived filtering", () => {
  it("filters by name, role, and email substring, case-insensitively", async () => {
    const user = userEvent.setup();
    renderApp();
    await addContactFromScreen(user, { name: "Annie Milewski", role: "Superintendent", email: "annie@northvalley.org" });
    await addContactFromScreen(user, { name: "Kim Rivera", role: "Curriculum Director" });

    const search = screen.getByRole("textbox", { name: /search contacts/i });
    await user.type(search, "SUPERINTENDENT");
    expect(screen.getByText("Annie Milewski")).toBeTruthy();
    expect(screen.queryByText("Kim Rivera")).toBeFalsy();

    await user.clear(search);
    await user.type(search, "northvalley.org");
    expect(screen.getByText("Annie Milewski")).toBeTruthy();
    expect(screen.queryByText("Kim Rivera")).toBeFalsy();
  });

  it("hides archived contacts by default, shows them when the archived toggle is checked", async () => {
    const user = userEvent.setup();
    renderApp();
    await addContactFromScreen(user, { name: "Retired Partner" });
    const card = (await screen.findByText("Retired Partner")).closest("article")!;
    await user.click(within(card).getByRole("button", { name: "Edit" }));
    const dialog = await screen.findByRole("dialog");
    await user.selectOptions(within(dialog).getByDisplayValue("Active"), "archived");
    await user.click(within(dialog).getByRole("button", { name: "Save Changes" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());

    expect(screen.queryByText("Retired Partner")).toBeFalsy();
    await user.click(screen.getByRole("checkbox", { name: /show archived/i }));
    expect(screen.getByText("Retired Partner")).toBeTruthy();
  });
});

describe("Duplicate detection", () => {
  it("warns but allows creation when the normalized display name already exists", async () => {
    const user = userEvent.setup();
    const { contactDataProvider } = renderApp();
    const createSpy = vi.spyOn(contactDataProvider, "create");
    await addContactFromScreen(user, { name: "Kim Rivera" });

    await openContacts(user);
    await user.click(screen.getByRole("button", { name: "Add Contact" }));
    const dialog = await screen.findByRole("dialog");
    await user.type(within(dialog).getByPlaceholderText(/annie milewski/i), "kim rivera"); // same name, different case
    expect(within(dialog).getByText(/already named/i)).toBeTruthy();
    await user.click(within(dialog).getByRole("button", { name: "Add Contact" }));
    await waitFor(() => expect(createSpy).toHaveBeenCalledTimes(2)); // both created — name collisions are informational only
  });

  it("requires an explicit second confirmation before creating a second contact with the same normalized email", async () => {
    const user = userEvent.setup();
    const { contactDataProvider } = renderApp();
    const createSpy = vi.spyOn(contactDataProvider, "create");
    await addContactFromScreen(user, { name: "Kim Rivera", email: "kim@example.org" });

    await openContacts(user);
    await user.click(screen.getByRole("button", { name: "Add Contact" }));
    const dialog = await screen.findByRole("dialog");
    await user.type(within(dialog).getByPlaceholderText(/annie milewski/i), "Kimberly Rivera");
    await user.type(within(dialog).getByPlaceholderText(/name@example\.org/i), "KIM@EXAMPLE.ORG"); // same email, different case

    await user.click(within(dialog).getByRole("button", { name: "Add Contact" }));
    expect(createSpy).toHaveBeenCalledTimes(1); // still just the first contact — the click was consumed by the warning
    expect(within(dialog).getByRole("alert")).toBeTruthy();

    await user.click(within(dialog).getByRole("button", { name: "Add Contact" }));
    await waitFor(() => expect(createSpy).toHaveBeenCalledTimes(2)); // explicit second click proceeds
  });

  it("different people with different names/emails are still freely creatable", async () => {
    const user = userEvent.setup();
    const { contactDataProvider } = renderApp();
    const createSpy = vi.spyOn(contactDataProvider, "create");
    await addContactFromScreen(user, { name: "Annie Milewski", email: "annie@example.org" });
    await addContactFromScreen(user, { name: "Kim Rivera", email: "kim@example.org" });
    expect(createSpy).toHaveBeenCalledTimes(2);
  });
});

describe("Universal Work Record integration", () => {
  it("a newly created contact becomes selectable through the existing Log Work contact picker", async () => {
    const user = userEvent.setup();
    renderApp();
    await addContactFromScreen(user, { name: "Annie Milewski" });
    const wizard = await openLogWizardAtContactStep(user);
    expect(within(wizard).getByRole("button", { name: "Annie Milewski" })).toBeTruthy();
  });

  it("logging work against the new contact creates exactly one Work Record through the existing provider, using the existing contactIds field", async () => {
    const user = userEvent.setup();
    const { dataProvider, contactDataProvider } = renderApp();
    const createSpy = vi.spyOn(dataProvider, "createWorkRecord");
    await addContactFromScreen(user, { name: "Annie Milewski" });
    expect(createSpy).not.toHaveBeenCalled(); // contact creation itself makes zero Work Record writes

    const wizard = await openLogWizardAtContactStep(user);
    await user.click(within(wizard).getByRole("button", { name: "Annie Milewski" }));
    for (let i = 0; i < 3; i++) {
      await user.click(within(wizard).getByRole("button", { name: /continue/i }));
    }
    await user.click(within(wizard).getByRole("button", { name: /save & done/i }));

    await waitFor(() => expect(createSpy).toHaveBeenCalledTimes(1));
    const [savedRecord] = createSpy.mock.calls[0];
    const listed = await contactDataProvider.list();
    const contact = listed.status === "success" ? listed.value.find((c) => c.displayName === "Annie Milewski") : undefined;
    expect(contact).toBeTruthy();
    expect(savedRecord.contactIds).toEqual([contact!.appId]);
    expect(savedRecord.projectIds).toEqual([]); // no second/parallel relationship mechanism was introduced
  });

  it("an unknown contact ID still fails Work Record validation", async () => {
    const { dataProvider } = renderApp();
    const result = await dataProvider.createWorkRecord({
      appId: "wr-1",
      title: "Test",
      activityDate: "2026-09-01",
      activityType: "Professional learning",
      description: "",
      detailedNotes: "",
      durationMinutes: 30,
      status: "complete",
      engagementScope: "none",
      projectIds: [],
      organizationIds: [],
      contactIds: ["contact-does-not-exist"],
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
      schemaVersion: 2,
      metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
      isSample: false,
    });
    expect(result.status).toBe("validation_error");
  });

  it("a static seeded contact still validates as a legal Work Record relationship", async () => {
    const user = userEvent.setup();
    const { dataProvider } = renderApp();
    await openContacts(user); // ensure the app has loaded references
    const seeded = (await dataProvider.getContacts()).find((c) => c.appId === "contact-futureworks");
    expect(seeded).toBeTruthy();
    const result = await dataProvider.createWorkRecord({
      appId: "wr-2",
      title: "Test",
      activityDate: "2026-09-01",
      activityType: "Professional learning",
      description: "",
      detailedNotes: "",
      durationMinutes: 30,
      status: "complete",
      engagementScope: "none",
      projectIds: [],
      organizationIds: [],
      contactIds: [seeded!.appId],
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
      schemaVersion: 2,
      metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
      isSample: false,
    });
    expect(result.status).toBe("success");
  });
});

describe("Boundaries", () => {
  it("creating and editing a contact causes zero Meeting Record, Inbox Intelligence, and Project writes", async () => {
    const user = userEvent.setup();
    const { meetingDataProvider, inboxDataProvider, projectDataProvider } = renderApp();
    const meetingCreateSpy = vi.spyOn(meetingDataProvider, "create");
    const meetingUpdateSpy = vi.spyOn(meetingDataProvider, "update");
    const inboxCreateSpy = vi.spyOn(inboxDataProvider, "create");
    const inboxUpdateSpy = vi.spyOn(inboxDataProvider, "update");
    const projectCreateSpy = vi.spyOn(projectDataProvider, "create");

    await addContactFromScreen(user, { name: "Annie Milewski" });

    expect(meetingCreateSpy).not.toHaveBeenCalled();
    expect(meetingUpdateSpy).not.toHaveBeenCalled();
    expect(inboxCreateSpy).not.toHaveBeenCalled();
    expect(inboxUpdateSpy).not.toHaveBeenCalled();
    expect(projectCreateSpy).not.toHaveBeenCalled();
  });

  it("requests no Microsoft Graph permission beyond the existing User.Read and Sites.ReadWrite.All", () => {
    expect(MICROSOFT_GRAPH_SCOPES).toEqual(["User.Read", "Sites.ReadWrite.All"]);
  });

  it("no Contact provider/codec file imports the Anthropic SDK or an AI analysis route", async () => {
    // Static architecture check: grep the actual source text for forbidden imports, rather
    // than relying on runtime behavior alone — Contact browsing/search/create/edit should
    // never be able to trigger an AI call because no such import path exists at all.
    const fs = await import("node:fs/promises");
    const path = await import("node:path");
    for (const file of ["lib/contact-provider.ts", "lib/sharepoint-contacts.ts"]) {
      const source = await fs.readFile(path.join(process.cwd(), file), "utf-8");
      expect(source).not.toMatch(/@anthropic-ai\/sdk/);
      expect(source).not.toMatch(/\/api\/(inbox|meeting|voice)-intelligence/);
    }
  });
});
