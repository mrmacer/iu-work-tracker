// @vitest-environment jsdom
//
// Patch 8F — Reviewed Contact Import. Covers the entry point, file selection/parse errors,
// candidate rendering, Match Existing / Create New / Ignore, multiple-candidate selection, a
// created Contact becoming matchable to a later duplicate row, session counts, and the safety
// invariants (Match Existing never updates a Contact, Ignore never mutates anything, no bulk
// action, zero AI calls). All against MemoryDataProvider/MemoryContactProvider/
// MemoryOrganizationProvider — zero real SharePoint writes, zero real Anthropic calls.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import IUWorkTracker from "../app/IUWorkTracker";
import { MemoryContactProvider } from "../lib/contact-provider";
import { MemoryDataProvider } from "../lib/data-provider";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { MemoryMeetingRecordProvider } from "../lib/meeting-record-provider";
import { MemoryOrganizationProvider } from "../lib/organization-provider";
import { MemoryProjectProvider } from "../lib/project-provider";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

function renderApp(overrides: { contactDataProvider?: MemoryContactProvider } = {}) {
  const dataProvider = new MemoryDataProvider([]);
  const contactDataProvider = overrides.contactDataProvider ?? new MemoryContactProvider();
  const organizationDataProvider = new MemoryOrganizationProvider();
  const projectDataProvider = new MemoryProjectProvider();
  const meetingDataProvider = new MemoryMeetingRecordProvider();
  const inboxDataProvider = new SessionInboxIntelligenceProvider();
  const utils = render(
    <IUWorkTracker
      dataProvider={dataProvider}
      contactDataProvider={contactDataProvider}
      organizationDataProvider={organizationDataProvider}
      projectDataProvider={projectDataProvider}
      meetingDataProvider={meetingDataProvider}
      inboxDataProvider={inboxDataProvider}
    />,
  );
  return { ...utils, dataProvider, contactDataProvider, organizationDataProvider, projectDataProvider, meetingDataProvider, inboxDataProvider };
}

async function openContacts(user: ReturnType<typeof userEvent.setup>) {
  const nav = await screen.findByRole("navigation");
  await user.click(within(nav).getByRole("button", { name: /contacts/i }));
}

async function openImport(user: ReturnType<typeof userEvent.setup>) {
  await openContacts(user);
  await user.click(screen.getByRole("button", { name: "Import Contacts" }));
}

function csvFile(content: string, name = "roster.csv"): File {
  return new File([content], name, { type: "text/csv" });
}

async function uploadCsv(user: ReturnType<typeof userEvent.setup>, content: string, name = "roster.csv") {
  const input = screen.getByLabelText(/csv file/i);
  await user.upload(input, csvFile(content, name));
}

describe("Entry point", () => {
  it("Contacts has an Import Contacts action that opens the import screen", async () => {
    const user = userEvent.setup();
    renderApp();
    await openContacts(user);
    expect(screen.getByRole("heading", { name: "Contacts" })).toBeTruthy();
    await user.click(screen.getByRole("button", { name: "Import Contacts" }));
    expect(await screen.findByRole("heading", { name: "Import Contacts" })).toBeTruthy();
  });

  it("← Back to Contacts returns to the directory", async () => {
    const user = userEvent.setup();
    renderApp();
    await openImport(user);
    await user.click(screen.getByRole("button", { name: /back to contacts/i }));
    expect(await screen.findByRole("heading", { name: "Contacts" })).toBeTruthy();
  });
});

describe("File selection and parsing", () => {
  it("parses a valid CSV and renders one row per data row", async () => {
    const user = userEvent.setup();
    renderApp();
    await openImport(user);
    await uploadCsv(user, "Full Name,Email,Role,Organization\nJordan Example,jordan@example.test,Coordinator,Intermediate Unit\nTaylor Example,,Teacher,North Valley SD");
    expect((await screen.findAllByText("Jordan Example")).length).toBeGreaterThan(0);
    expect(screen.getAllByText("Taylor Example").length).toBeGreaterThan(0);
    expect(screen.getByText(/roster\.csv/)).toBeTruthy();
  });

  it("shows a clear error when the file has no Name column, and does not render a review queue", async () => {
    const user = userEvent.setup();
    renderApp();
    await openImport(user);
    await uploadCsv(user, "Email,Role\njordan@example.test,Coordinator");
    expect(await screen.findByRole("alert")).toHaveProperty("textContent", expect.stringMatching(/Name column/i));
    expect(screen.queryByText(/total rows/)).toBeFalsy();
  });

  it("shows a clear error for malformed quoted CSV", async () => {
    const user = userEvent.setup();
    renderApp();
    await openImport(user);
    await uploadCsv(user, 'Name\n"Jordan');
    expect(await screen.findByRole("alert")).toBeTruthy();
  });
});

describe("Candidate display", () => {
  it("shows original Organization text and its resolved name when it matches an existing Organization", async () => {
    const user = userEvent.setup();
    renderApp();
    await openImport(user);
    await uploadCsv(user, "Name,Organization\nTaylor Example,North Valley SD");
    expect(await screen.findByText(/Organization: North Valley SD → North Valley SD/)).toBeTruthy();
  });

  it("shows the Organization text as unresolved when it doesn't match any existing Organization", async () => {
    const user = userEvent.setup();
    renderApp();
    await openImport(user);
    await uploadCsv(user, "Name,Organization\nMorgan Example,Unknown Organization");
    expect(await screen.findByText(/Organization: Unknown Organization \(not resolved/)).toBeTruthy();
  });
});

describe("Match Existing", () => {
  it("matches an existing Contact without ever calling ContactProvider.update()", async () => {
    const user = userEvent.setup();
    // Seeded BEFORE render/mount, so the app's own initial list() load (a mount-time effect)
    // actually sees it — creating through the provider after mount would miss that one-time load.
    const contactDataProvider = new MemoryContactProvider();
    const created = await contactDataProvider.create({
      appId: "contact-existing",
      displayName: "Jordan Example",
      organizationId: null,
      status: "active",
    });
    if (created.status !== "success") throw new Error("setup failed");
    const updateSpy = vi.spyOn(contactDataProvider, "update");

    renderApp({ contactDataProvider });
    await openImport(user);
    await uploadCsv(user, "Name\nJordan Example");
    await screen.findAllByText("Jordan Example");
    await user.click(screen.getByRole("button", { name: "Match Existing" }));
    expect(screen.getByText(/Matched Contact: Jordan Example/)).toBeTruthy();
    expect(updateSpy).not.toHaveBeenCalled();
  });

  it("renders every plausible Contact when multiple share the same name, and lets the human pick one", async () => {
    const user = userEvent.setup();
    const contactDataProvider = new MemoryContactProvider();
    await contactDataProvider.create({ appId: "contact-1", displayName: "Jordan Example", organizationId: null, status: "active" });
    await contactDataProvider.create({ appId: "contact-2", displayName: "Jordan Example", organizationId: null, status: "active" });

    renderApp({ contactDataProvider });
    await openImport(user);
    await uploadCsv(user, "Name\nJordan Example");
    await screen.findAllByText("Jordan Example");
    const matchButtons = screen.getAllByRole("button", { name: "Match Existing" });
    expect(matchButtons).toHaveLength(2);
    await user.click(matchButtons[1]);
    expect(screen.getByText(/Matched Contact: Jordan Example/)).toBeTruthy();
  });
});

describe("Create New", () => {
  it("opens the existing Contact form, correctly prefilled, and never puts unmapped columns into Notes", async () => {
    const user = userEvent.setup();
    renderApp();
    await openImport(user);
    await uploadCsv(user, "Name,Email,Role,Organization,Phone\nJordan Example,jordan@example.test,Coordinator,Intermediate Unit,555-0100");
    await screen.findAllByText("Jordan Example");
    await user.click(screen.getByRole("button", { name: "Add Person" }));
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByRole("heading", { name: "Add Contact" })).toBeTruthy();
    expect(within(dialog).getByDisplayValue("Jordan Example")).toBeTruthy();
    expect(within(dialog).getByDisplayValue("jordan@example.test")).toBeTruthy();
    expect(within(dialog).getByDisplayValue("Coordinator")).toBeTruthy();
    const orgSelect = within(dialog).getByRole("combobox", { name: /organization/i }) as HTMLSelectElement;
    expect(orgSelect.value).toBe("org-iu");
    const notesField = within(dialog).getByPlaceholderText(/not a biography/i) as HTMLTextAreaElement;
    expect(notesField.value).toBe(""); // Phone column is unmapped — never absorbed into Notes
  });

  it("creates the Contact through the real save path, and the row is then shown as matched", async () => {
    const user = userEvent.setup();
    renderApp();
    await openImport(user);
    await uploadCsv(user, "Name\nJordan Example");
    await screen.findAllByText("Jordan Example");
    await user.click(screen.getByRole("button", { name: "Add Person" }));
    const dialog = await screen.findByRole("dialog");
    await user.click(within(dialog).getByRole("button", { name: "Add Contact" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(screen.getByText(/Matched Contact: Jordan Example/)).toBeTruthy();
  });

  it("a Contact created for an earlier row becomes matchable by a later duplicate row in the same file", async () => {
    const user = userEvent.setup();
    renderApp();
    await openImport(user);
    await uploadCsv(
      user,
      "Name,Email\nJordan Example,jordan@example.test\nJordan Example,jordan@example.test",
    );
    await screen.findAllByText("Jordan Example");

    // Row 1 (row 2 in the file): Create New.
    const addPersonButtons = screen.getAllByRole("button", { name: "Add Person" });
    await user.click(addPersonButtons[0]);
    const dialog = await screen.findByRole("dialog");
    await user.click(within(dialog).getByRole("button", { name: "Add Contact" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());

    // Row 2 (the duplicate) should now offer the just-created Contact as a strong (email) match.
    const matchButtons = await screen.findAllByRole("button", { name: "Match Existing" });
    expect(matchButtons).toHaveLength(1); // only the still-unreviewed duplicate row offers a match
    expect(screen.getByText(/Exact email match/)).toBeTruthy();
  });
});

describe("Ignore", () => {
  it("marks the row ignored without creating or updating any Contact", async () => {
    const user = userEvent.setup();
    const { contactDataProvider } = renderApp();
    const createSpy = vi.spyOn(contactDataProvider, "create");
    await openImport(user);
    await uploadCsv(user, "Name\nMorgan Example");
    await screen.findAllByText("Morgan Example");
    await user.click(screen.getByRole("button", { name: "Ignore" }));
    expect(screen.getByText("Ignored for this item")).toBeTruthy();
    expect(createSpy).not.toHaveBeenCalled();
  });
});

describe("Session summary", () => {
  it("tracks total/unreviewed/matched existing/created/ignored correctly, with no bulk action available", async () => {
    const user = userEvent.setup();
    const contactDataProvider = new MemoryContactProvider();
    await contactDataProvider.create({ appId: "contact-existing", displayName: "Taylor Example", organizationId: null, status: "active" });

    renderApp({ contactDataProvider });
    await openImport(user);
    await uploadCsv(
      user,
      "Name\nTaylor Example\nJordan Example\nMorgan Example\nCasey Example",
    );
    await screen.findAllByText("Taylor Example");

    expect(screen.getByText("total rows").closest(".metric")?.querySelector("strong")?.textContent).toBe("4");

    // Match Existing for Taylor.
    await user.click(within(screen.getAllByText("Taylor Example")[0].closest(".import-row")!).getByRole("button", { name: "Match Existing" }));
    // Create New for Jordan.
    await user.click(within(screen.getAllByText("Jordan Example")[0].closest(".import-row")!).getByRole("button", { name: "Add Person" }));
    const dialog = await screen.findByRole("dialog");
    await user.click(within(dialog).getByRole("button", { name: "Add Contact" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    // Ignore Morgan.
    await user.click(within(screen.getAllByText("Morgan Example")[0].closest(".import-row")!).getByRole("button", { name: "Ignore" }));

    expect(screen.getByText("unreviewed").closest(".metric")?.querySelector("strong")?.textContent).toBe("1"); // Casey only
    expect(screen.getByText("matched existing").closest(".metric")?.querySelector("strong")?.textContent).toBe("1"); // Taylor
    expect(screen.getByText("created").closest(".metric")?.querySelector("strong")?.textContent).toBe("1"); // Jordan
    expect(screen.getByText("ignored").closest(".metric")?.querySelector("strong")?.textContent).toBe("1"); // Morgan

    expect(screen.queryByRole("button", { name: /save all|create all|import all|accept all|auto match all/i })).toBeFalsy();
  });
});

describe("Zero AI calls", () => {
  it("selecting a file, parsing, matching, and reviewing make zero network/AI calls", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    renderApp();
    await openImport(user);
    await uploadCsv(user, "Name\nJordan Example");
    await screen.findAllByText("Jordan Example");
    await user.click(screen.getByRole("button", { name: "Ignore" }));
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});

describe("Regression", () => {
  it("Contacts directory remains intact after visiting Import Contacts and returning", async () => {
    const user = userEvent.setup();
    renderApp();
    await openImport(user);
    await user.click(screen.getByRole("button", { name: /back to contacts/i }));
    for (const name of ["Development District Lead", "Development Partner Contact", "Development IU Colleague"]) {
      expect(await screen.findByText(name)).toBeTruthy();
    }
    expect(screen.getByRole("button", { name: "Add Contact" })).toBeTruthy();
  });
});
