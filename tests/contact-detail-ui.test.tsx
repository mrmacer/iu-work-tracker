// @vitest-environment jsdom
//
// Patch 8C — Contact Detail + Connected Work. Covers opening a Contact into its detail view,
// the derived Relationship Snapshot / Connected Projects / Recent Work sections, empty states,
// seed-contact and archived-contact behavior, back navigation, and the zero-AI /
// zero-new-persistence boundaries. All against MemoryDataProvider/MemoryContactProvider — zero
// real SharePoint writes, zero real Anthropic calls.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import IUWorkTracker from "../app/IUWorkTracker";
import { MemoryContactProvider } from "../lib/contact-provider";
import { MemoryDataProvider } from "../lib/data-provider";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { MemoryMeetingRecordProvider } from "../lib/meeting-record-provider";
import { MemoryProjectProvider } from "../lib/project-provider";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

function renderApp(overrides: { dataProvider?: MemoryDataProvider } = {}) {
  const dataProvider = overrides.dataProvider ?? new MemoryDataProvider([]);
  const contactDataProvider = new MemoryContactProvider();
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

async function openContactDetail(user: ReturnType<typeof userEvent.setup>, displayName: string) {
  const card = (await screen.findByText(displayName)).closest("article")!;
  await user.click(within(card).getByRole("button", { name: new RegExp(displayName) }));
}

async function addContactFromScreen(
  user: ReturnType<typeof userEvent.setup>,
  fields: { name: string; email?: string; role?: string; notes?: string },
) {
  await openContacts(user);
  await user.click(screen.getByRole("button", { name: "Add Contact" }));
  const dialog = await screen.findByRole("dialog");
  await user.type(within(dialog).getByPlaceholderText(/annie milewski/i), fields.name);
  if (fields.role) await user.type(within(dialog).getByPlaceholderText(/superintendent/i), fields.role);
  if (fields.email) await user.type(within(dialog).getByPlaceholderText(/name@example\.org/i), fields.email);
  if (fields.notes) await user.type(within(dialog).getByPlaceholderText(/not a biography/i), fields.notes);
  await user.click(within(dialog).getByRole("button", { name: "Add Contact" }));
  await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
}

describe("Opening Contact Detail", () => {
  it("clicking a seeded contact card opens its detail view", async () => {
    const user = userEvent.setup();
    renderApp();
    await openContacts(user);
    await openContactDetail(user, "Development Partner Contact");
    expect(await screen.findByRole("heading", { name: "Development Partner Contact" })).toBeTruthy();
    expect(screen.getByRole("button", { name: /back to contacts/i })).toBeTruthy();
  });

  it("resolves the organization name from organizationId", async () => {
    const user = userEvent.setup();
    renderApp();
    await openContacts(user);
    await openContactDetail(user, "Development Partner Contact");
    expect(screen.getByText("FutureWorks Partnership")).toBeTruthy();
  });

  it("displays the contact's status", async () => {
    const user = userEvent.setup();
    renderApp();
    await openContacts(user);
    await openContactDetail(user, "Development Partner Contact");
    expect(screen.getByText(/Active/)).toBeTruthy();
  });

  it("back navigation returns to the directory", async () => {
    const user = userEvent.setup();
    renderApp();
    await openContacts(user);
    await openContactDetail(user, "Development Partner Contact");
    await user.click(screen.getByRole("button", { name: /back to contacts/i }));
    expect(await screen.findByRole("heading", { name: "Contacts" })).toBeTruthy();
    expect(screen.getByText("Development District Lead")).toBeTruthy();
  });
});

describe("Relationship Snapshot, Connected Projects, and Recent Work (derived)", () => {
  it("renders connected time, work record count, last interaction, connected project names, and recent work newest-first", async () => {
    const user = userEvent.setup();
    renderApp({ dataProvider: new MemoryDataProvider() }); // default sample records — sample-steels contactIds includes contact-north-valley-lead
    await openContacts(user);
    await openContactDetail(user, "Development District Lead");

    // sample-steels: contactIds ["contact-north-valley-lead"], projectIds ["project-steels"], durationMinutes 60, activityDate 2026-08-26
    const snapshot = screen.getByText("Relationship snapshot").closest("section")!;
    expect(within(snapshot).getByText("Aug 26, 2026")).toBeTruthy(); // last interaction
    expect(within(snapshot).getByText("1h")).toBeTruthy(); // connected work total
    expect(within(snapshot).getByText("1")).toBeTruthy(); // work record count

    expect(screen.getByText("STEELS Implementation")).toBeTruthy(); // connected project
    expect(screen.getByText("District STEELS planning meeting")).toBeTruthy(); // recent work title
  });

  it("clicking a recent Work Record row opens the existing Log Work edit experience for that record", async () => {
    const user = userEvent.setup();
    renderApp({ dataProvider: new MemoryDataProvider() });
    await openContacts(user);
    await openContactDetail(user, "Development District Lead");
    await user.click(screen.getByText("District STEELS planning meeting"));
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByDisplayValue("District STEELS planning meeting")).toBeTruthy();
  });

  it("displays the contact's notes when present", async () => {
    const user = userEvent.setup();
    renderApp();
    await addContactFromScreen(user, { name: "Notes Test Contact", notes: "Leads district STEELS implementation." });
    await openContactDetail(user, "Notes Test Contact");
    expect(screen.getByText("Leads district STEELS implementation.")).toBeTruthy();
  });
});

describe("Empty states", () => {
  it("a contact with no connected Work Records shows calm empty-state language, not a crash", async () => {
    const user = userEvent.setup();
    renderApp(); // no sample records at all
    await openContacts(user);
    await openContactDetail(user, "Development District Lead");
    expect(screen.getByText("No recorded interaction yet.")).toBeTruthy();
    expect(screen.getByText("No connected projects yet.")).toBeTruthy();
    expect(screen.getByText("No work has been logged with this contact yet.")).toBeTruthy();
    expect(screen.getByText("0m")).toBeTruthy(); // connected time
  });
});

describe("Seed and archived Contact behavior", () => {
  it("a seed/reference contact (no metadata) can open detail and shows no Edit button", async () => {
    const user = userEvent.setup();
    renderApp();
    await openContacts(user);
    await openContactDetail(user, "Development District Lead");
    expect(screen.queryByRole("button", { name: "Edit Contact" })).toBeFalsy();
  });

  it("an archived contact still displays its connected work when explicitly opened", async () => {
    const user = userEvent.setup();
    renderApp();
    await addContactFromScreen(user, { name: "Archived Person Test" });

    // Edit the just-created contact to archived status.
    const card = (await screen.findByText("Archived Person Test")).closest("article")!;
    await user.click(within(card).getByRole("button", { name: "Edit" }));
    const editDialog = await screen.findByRole("dialog");
    await user.selectOptions(within(editDialog).getByLabelText(/relationship status/i), "archived");
    await user.click(within(editDialog).getByRole("button", { name: "Save Changes" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());

    // Log a Work Record connected to this contact via the real Log Work flow (LogWizard step 1
    // title+type -> step 2 scope/contacts (contacts live under the "Add classification / reach"
    // details) -> steps 3/4 accept defaults -> step 5 "Save & done").
    await user.click(within(screen.getByRole("complementary")).getByRole("button", { name: /log work/i }));
    const wizard = await screen.findByRole("dialog");
    await user.type(within(wizard).getByRole("textbox", { name: /activity title/i }), "Work with archived contact");
    await user.selectOptions(within(wizard).getByRole("combobox", { name: /activity type/i }), "Professional learning");
    await user.click(within(wizard).getByRole("button", { name: /continue/i }));

    await user.click(within(wizard).getByText("Add classification / reach"));
    await user.click(within(wizard).getByRole("button", { name: "Archived Person Test" }));
    await user.click(within(wizard).getByRole("button", { name: /continue/i }));
    await user.click(within(wizard).getByRole("button", { name: /continue/i }));
    await user.click(within(wizard).getByRole("button", { name: /continue/i }));
    await user.click(within(wizard).getByRole("button", { name: /save & done/i }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());

    await openContacts(user);
    await user.click(screen.getByRole("checkbox", { name: /show archived/i }));
    await openContactDetail(user, "Archived Person Test");
    expect(screen.getByText("Work with archived contact")).toBeTruthy();
  });
});

describe("Boundaries", () => {
  it("opening Contact Detail and viewing connected work make zero Anthropic requests", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    renderApp({ dataProvider: new MemoryDataProvider() });
    await openContacts(user);
    await openContactDetail(user, "Development District Lead");
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("Contact still has no projectIds and Project still has no contactIds", async () => {
    const fs = await import("node:fs/promises");
    const path = await import("node:path");
    const modelsSource = await fs.readFile(path.join(process.cwd(), "lib/models.ts"), "utf-8");
    const contactBlock = modelsSource.slice(modelsSource.indexOf("export type Contact"), modelsSource.indexOf("export type Category"));
    const projectBlock = modelsSource.slice(modelsSource.indexOf("export type Project"), modelsSource.indexOf("export type Organization"));
    expect(contactBlock).not.toMatch(/projectIds/);
    expect(projectBlock).not.toMatch(/contactIds/);
  });

  it("the derivation module performs no I/O — no fetch/network calls and no provider imports", async () => {
    const fs = await import("node:fs/promises");
    const path = await import("node:path");
    const source = await fs.readFile(path.join(process.cwd(), "lib/contact-relationships.ts"), "utf-8");
    expect(source).not.toMatch(/fetch\(/);
    expect(source).not.toMatch(/@anthropic-ai\/sdk/);
    expect(source).not.toMatch(/from ["'].*(provider|sharepoint|graph)/i); // no data-access imports at all
    expect(source).toMatch(/^import type \{ Project, WorkRecord \} from "\.\/models";$/m); // the only import: types
  });
});
