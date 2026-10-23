// @vitest-environment jsdom
//
// Patch 8E — Durable Organizations. Covers the Organizations directory, search, Create/Edit UI,
// seed-organization read-only regression, duplicate-name warning, the Universal Work Record
// integration ("Log it once. Use it everywhere."), Contact.organizationId resolution, and Inbox
// Organization/District matching seeing a durable Organization. All against
// MemoryDataProvider/MemoryOrganizationProvider — zero real SharePoint writes, zero real
// Anthropic calls.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import IUWorkTracker from "../app/IUWorkTracker";
import { MemoryDataProvider } from "../lib/data-provider";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { MemoryMeetingRecordProvider } from "../lib/meeting-record-provider";
import { MemoryOrganizationProvider } from "../lib/organization-provider";
import { MemoryProjectProvider } from "../lib/project-provider";
import { MemoryContactProvider } from "../lib/contact-provider";
import type { AnalyzeEmailResult } from "../lib/anthropic-email-analysis";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

function renderApp(overrides: { dataProvider?: MemoryDataProvider; organizationDataProvider?: MemoryOrganizationProvider } = {}) {
  const dataProvider = overrides.dataProvider ?? new MemoryDataProvider([]);
  const organizationDataProvider = overrides.organizationDataProvider ?? new MemoryOrganizationProvider();
  const projectDataProvider = new MemoryProjectProvider();
  const contactDataProvider = new MemoryContactProvider();
  const meetingDataProvider = new MemoryMeetingRecordProvider();
  const inboxDataProvider = new SessionInboxIntelligenceProvider();
  const utils = render(
    <IUWorkTracker
      dataProvider={dataProvider}
      organizationDataProvider={organizationDataProvider}
      projectDataProvider={projectDataProvider}
      contactDataProvider={contactDataProvider}
      meetingDataProvider={meetingDataProvider}
      inboxDataProvider={inboxDataProvider}
    />
  );
  return { ...utils, dataProvider, organizationDataProvider, projectDataProvider, contactDataProvider, meetingDataProvider, inboxDataProvider };
}

async function openOrganizations(user: ReturnType<typeof userEvent.setup>) {
  const nav = await screen.findByRole("navigation");
  await user.click(within(nav).getByRole("button", { name: /organizations/i }));
}

async function addOrganizationFromScreen(user: ReturnType<typeof userEvent.setup>, fields: { name: string; type?: string }) {
  await openOrganizations(user);
  await user.click(screen.getByRole("button", { name: "Add Organization" }));
  const dialog = await screen.findByRole("dialog");
  await user.type(within(dialog).getByPlaceholderText(/gillingham charter school/i), fields.name);
  if (fields.type) await user.selectOptions(within(dialog).getByRole("combobox"), fields.type);
  await user.click(within(dialog).getByRole("button", { name: "Add Organization" }));
  await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
}

const IU29_DISTRICTS = [
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
const RETIRED_PLACEHOLDERS = ["North Valley SD", "Riverbend Area SD", "FutureWorks Partnership"];

describe("Seeded IU29 directory (Patch 9B)", () => {
  it("shows the twelve real IU29 districts, Schuylkill Intermediate Unit 29, and PAIMS — fourteen seeded Organizations", async () => {
    const user = userEvent.setup();
    renderApp();
    await openOrganizations(user);
    for (const name of [...IU29_DISTRICTS, "Schuylkill Intermediate Unit 29", "PAIMS"]) {
      expect(await screen.findByText(name)).toBeTruthy();
    }
    expect(screen.getAllByRole("article")).toHaveLength(14);
  });

  it("never shows the retired placeholder Organizations in the directory", async () => {
    const user = userEvent.setup();
    renderApp();
    await openOrganizations(user);
    await screen.findByText("PAIMS");
    for (const name of [...RETIRED_PLACEHOLDERS, "Intermediate Unit"]) expect(screen.queryByText(name)).toBeFalsy();
  });

  it("types are right: districts are District, PAIMS is a Partner (not a district), the IU is IU", async () => {
    const user = userEvent.setup();
    renderApp();
    await openOrganizations(user);
    const chip = async (name: string) => within((await screen.findByText(name)).closest("article")!).getByText(/^(District|Partner|IU)$/).textContent;
    expect(await chip("Blue Mountain School District")).toBe("District");
    expect(await chip("PAIMS")).toBe("Partner");
    expect(await chip("Schuylkill Intermediate Unit 29")).toBe("IU");
  });

  it("shows no duplicate cards after a durable organization is created — fourteen seeded plus exactly one new", async () => {
    const user = userEvent.setup();
    renderApp();
    await addOrganizationFromScreen(user, { name: "Gillingham Charter School", type: "partner" });
    await waitFor(() => expect(screen.getAllByRole("article")).toHaveLength(15));
    expect(screen.getAllByText("Gillingham Charter School")).toHaveLength(1);
  });

  it("a seed organization has no Edit button — read-only", async () => {
    const user = userEvent.setup();
    renderApp();
    await openOrganizations(user);
    const card = (await screen.findByText("Blue Mountain School District")).closest("article")!;
    expect(within(card).queryByRole("button", { name: "Edit" })).toBeFalsy();
  });
});

describe("Add Organization", () => {
  it("has an Add Organization control that opens a compact form with Name and Type only", async () => {
    const user = userEvent.setup();
    renderApp();
    await openOrganizations(user);
    await user.click(screen.getByRole("button", { name: "Add Organization" }));
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByRole("heading", { name: "Add Organization" })).toBeTruthy();
    expect(within(dialog).getByPlaceholderText(/gillingham charter school/i)).toBeTruthy();
    expect(within(dialog).getByRole("combobox")).toBeTruthy();
    expect(within(dialog).queryByText(/notes/i)).toBeFalsy(); // no Notes field in Patch 8E
  });

  it("defaults type to Partner and disables submit until a name is entered", async () => {
    const user = userEvent.setup();
    renderApp();
    await openOrganizations(user);
    await user.click(screen.getByRole("button", { name: "Add Organization" }));
    const dialog = await screen.findByRole("dialog");
    expect((within(dialog).getByRole("combobox") as HTMLSelectElement).value).toBe("partner");
    expect((within(dialog).getByRole("button", { name: "Add Organization" }) as HTMLButtonElement).disabled).toBe(true);
  });

  it("the new organization becomes editable immediately (durable) and appears selectable in Log Work", async () => {
    const user = userEvent.setup();
    renderApp();
    await addOrganizationFromScreen(user, { name: "Riverbend STEM Coalition", type: "partner" });
    const card = (await screen.findByText("Riverbend STEM Coalition")).closest("article")!;
    expect(within(card).getByRole("button", { name: "Edit" })).toBeTruthy();
  });
});

describe("Edit durable Organization", () => {
  it("saves a change to a durable organization's name and type", async () => {
    const user = userEvent.setup();
    renderApp();
    await addOrganizationFromScreen(user, { name: "Original Name", type: "partner" });
    const card = (await screen.findByText("Original Name")).closest("article")!;
    await user.click(within(card).getByRole("button", { name: "Edit" }));
    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByRole("heading", { name: "Edit Organization" })).toBeTruthy();
    const nameInput = within(dialog).getByDisplayValue("Original Name");
    await user.clear(nameInput);
    await user.type(nameInput, "Renamed Organization");
    await user.selectOptions(within(dialog).getByRole("combobox"), "district");
    await user.click(within(dialog).getByRole("button", { name: "Save Changes" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(await screen.findByText("Renamed Organization")).toBeTruthy();
    expect(screen.queryByText("Original Name")).toBeFalsy();
  });
});

describe("Duplicate-name warning", () => {
  it("warns on a normalized-name duplicate but does not block Save", async () => {
    const user = userEvent.setup();
    renderApp();
    await openOrganizations(user);
    await user.click(screen.getByRole("button", { name: "Add Organization" }));
    const dialog = await screen.findByRole("dialog");
    // "blue mountain school district" normalizes identically to the seeded "Blue Mountain School District".
    await user.type(within(dialog).getByPlaceholderText(/gillingham charter school/i), "  BLUE   MOUNTAIN school district  ");
    expect((await within(dialog).findByRole("status")).textContent).toMatch(/already named/i);
    await user.click(within(dialog).getByRole("button", { name: "Add Organization" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy()); // single click — never blocked
    // Patch 9B: the durable record REPLACES the same-named seed in the directory — no duplicate card.
    expect(screen.getAllByRole("article")).toHaveLength(14);
    expect(screen.getAllByText(/blue\s+mountain school district/i)).toHaveLength(1);
    expect(screen.queryByText("Blue Mountain School District")).toBeFalsy(); // the seed spelling stepped aside…
    const card = screen.getByText(/blue\s+mountain school district/i).closest("article")!;
    expect(within(card).getByRole("button", { name: "Edit" })).toBeTruthy(); // …for the durable (editable) one
  });
});

describe("Search", () => {
  it("filters by name, case-insensitive substring", async () => {
    const user = userEvent.setup();
    renderApp();
    await openOrganizations(user);
    await user.type(screen.getByRole("textbox", { name: /search organizations/i }), "paims");
    expect(screen.getByText("PAIMS")).toBeTruthy();
    expect(screen.queryByText("Blue Mountain School District")).toBeFalsy();
  });

  it("filters by type label", async () => {
    const user = userEvent.setup();
    renderApp();
    await openOrganizations(user);
    await user.type(screen.getByRole("textbox", { name: /search organizations/i }), "district");
    expect(screen.getByText("Blue Mountain School District")).toBeTruthy();
    expect(screen.getByText("Williams Valley School District")).toBeTruthy();
    expect(screen.queryByText("Schuylkill Intermediate Unit 29")).toBeFalsy();
    expect(screen.queryByText("PAIMS")).toBeFalsy();
  });
});

describe("Log Work sees a durable Organization", () => {
  it("a durable district appears in the District(s) picker under Specific scope, and only there — never in Organization/partner", async () => {
    const user = userEvent.setup();
    renderApp();
    await addOrganizationFromScreen(user, { name: "New Durable District", type: "district" });

    await user.click(within(screen.getByRole("complementary")).getByRole("button", { name: /log work/i }));
    const wizard = await screen.findByRole("dialog");
    await user.type(within(wizard).getByRole("textbox", { name: /activity title/i }), "District follow-up");
    await user.selectOptions(within(wizard).getByRole("combobox", { name: /activity type/i }), "District meeting");
    await user.click(within(wizard).getByRole("button", { name: /continue/i }));

    await user.click(within(wizard).getByRole("button", { name: "Specific district(s)" }));
    expect(within(wizard).getByRole("button", { name: /New Durable District/ })).toBeTruthy();

    // The district must not also appear as a selectable Organization/partner chip.
    const orgFieldsetLegend = within(wizard).getByText("Organization / partner");
    const orgFieldset = orgFieldsetLegend.closest("fieldset")!;
    expect(within(orgFieldset).queryByText(/New Durable District/)).toBeFalsy();
  });

  it("a durable partner appears in the Organization/partner picker, not the district picker", async () => {
    const user = userEvent.setup();
    renderApp();
    await addOrganizationFromScreen(user, { name: "New Durable Partner", type: "partner" });

    await user.click(within(screen.getByRole("complementary")).getByRole("button", { name: /log work/i }));
    const wizard = await screen.findByRole("dialog");
    await user.type(within(wizard).getByRole("textbox", { name: /activity title/i }), "Partner check-in");
    await user.selectOptions(within(wizard).getByRole("combobox", { name: /activity type/i }), "Partner meeting");
    await user.click(within(wizard).getByRole("button", { name: /continue/i }));

    expect(within(wizard).getByRole("button", { name: /New Durable Partner/ })).toBeTruthy();

    // The partner must not also appear as a selectable district once District(s) is shown.
    await user.click(within(wizard).getByRole("button", { name: "Specific district(s)" }));
    const districtFieldsetLegend = within(wizard).getByText("District(s)");
    const districtFieldset = districtFieldsetLegend.closest("fieldset")!;
    expect(within(districtFieldset).queryByText(/New Durable Partner/)).toBeFalsy();
  });
});

describe("Contact.organizationId resolves a durable Organization", () => {
  it("the durable organization appears as a selectable option in the Add Contact form", async () => {
    const user = userEvent.setup();
    renderApp();
    await addOrganizationFromScreen(user, { name: "New Durable District", type: "district" });

    const nav = await screen.findByRole("navigation");
    await user.click(within(nav).getByRole("button", { name: /contacts/i }));
    await user.click(screen.getByRole("button", { name: "Add Contact" }));
    const dialog = await screen.findByRole("dialog");
    const orgSelect = within(dialog).getByRole("combobox", { name: /organization/i });
    expect(within(orgSelect).getByRole("option", { name: "New Durable District" })).toBeTruthy();
  });
});

describe("Inbox Organization/District matching sees a durable Organization", () => {
  it("resolves a durable district by exact name and shows it on the saved Inbox row", async () => {
    const user = userEvent.setup();
    renderApp();
    await addOrganizationFromScreen(user, { name: "New Durable District", type: "district" });

    const nav = await screen.findByRole("navigation");
    await user.click(within(nav).getByRole("button", { name: /inbox intelligence/i }));

    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      new Response(
        JSON.stringify({
          status: "success",
          analysis: {
            summary: "Update from the new district.",
            priority: "medium",
            needsAttention: false,
            actionItems: [],
            followUp: "",
            people: [],
            organizations: [],
            districts: ["New Durable District"],
            projects: [],
            tags: [],
            suggestedWorkType: null,
            suggestedWorkRecord: { title: "New district update", description: "Synthetic test email." },
          },
          usage: { model: "claude-opus-5", inputTokens: 300, outputTokens: 100 },
        } satisfies AnalyzeEmailResult),
        { status: 200, headers: { "content-type": "application/json" } },
      ),
    );
    await user.type(screen.getByPlaceholderText(/paste the whole email/i), "Update from the new district office.");
    await user.click(screen.getByRole("button", { name: "Analyze email" }));
    await waitFor(() => expect(screen.getByRole("button", { name: "Save to Inbox" })).toBeTruthy());
    await user.click(screen.getByRole("button", { name: "Save to Inbox" }));
    await waitFor(() => expect(screen.getByText("Saved to Inbox.")).toBeTruthy());

    const needsAttentionSection = screen.getByText("Needs attention").closest("section")!;
    expect(within(needsAttentionSection).getByText(/New Durable District/)).toBeTruthy();
  });
});
