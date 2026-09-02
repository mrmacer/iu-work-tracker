// @vitest-environment jsdom
//
// Patch 8D — the shared review presentational component (app/ContactMatchPanel.tsx). Proves:
// AI MAY PROPOSE, THE HUMAN DECIDES — no candidate, regardless of strength, is ever rendered as
// already-decided; every decision requires an explicit click the test itself performs.
import { cleanup, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import type { ComponentProps } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import ContactMatchPanel from "../app/ContactMatchPanel";
import type { ContactMatchCandidate } from "../lib/contact-matching";
import type { Contact, Organization } from "../lib/models";

afterEach(cleanup);

const ORGANIZATIONS: Organization[] = [{ appId: "org-north-valley", name: "North Valley SD", type: "district" }];

function contact(overrides: Partial<Contact> = {}): Contact {
  return { appId: "contact-annie", displayName: "Annie Milewski", organizationId: "org-north-valley", status: "active", ...overrides };
}

function candidate(overrides: Partial<ContactMatchCandidate> = {}): ContactMatchCandidate {
  return { contact: contact(), strength: "review", reasons: ["Name match only — organization not confirmed"], ...overrides };
}

function renderPanel(props: Partial<ComponentProps<typeof ContactMatchPanel>> = {}) {
  const onMatch = vi.fn();
  const onIgnore = vi.fn();
  const onReset = vi.fn();
  const onAddPerson = vi.fn();
  const utils = render(
    <ContactMatchPanel
      personName="Annie Milewski"
      candidates={[candidate()]}
      decision={undefined}
      contacts={[contact()]}
      organizations={ORGANIZATIONS}
      onMatch={onMatch}
      onIgnore={onIgnore}
      onReset={onReset}
      onAddPerson={onAddPerson}
      {...props}
    />,
  );
  return { ...utils, onMatch, onIgnore, onReset, onAddPerson };
}

describe("Unreviewed state (no decision yet)", () => {
  it("renders a candidate but requires an explicit click before anything is decided", () => {
    const { onMatch } = renderPanel();
    expect(screen.getByText("Annie Milewski · North Valley SD", { exact: false })).toBeTruthy();
    expect(onMatch).not.toHaveBeenCalled();
    expect(screen.queryByText(/Matched Contact/)).toBeFalsy();
  });

  it("a strong (exact email) candidate is still only a suggestion — clicking Match Existing is required to commit it", async () => {
    const user = userEvent.setup();
    const strong = candidate({ strength: "strong", reasons: ["Exact email match"] });
    const { onMatch } = renderPanel({ candidates: [strong] });
    expect(screen.getByText(/Exact email match/)).toBeTruthy();
    expect(onMatch).not.toHaveBeenCalled(); // rendering the suggestion never calls onMatch by itself
    await user.click(screen.getByRole("button", { name: "Match Existing" }));
    expect(onMatch).toHaveBeenCalledWith("contact-annie");
  });

  it("clicking Match Existing on a specific candidate passes that candidate's contactAppId", async () => {
    const user = userEvent.setup();
    const first = candidate({ contact: contact({ appId: "contact-annie-1" }) });
    const second = candidate({ contact: contact({ appId: "contact-annie-2", organizationId: "org-other" }) });
    const { onMatch } = renderPanel({ candidates: [first, second] });
    const buttons = screen.getAllByRole("button", { name: "Match Existing" });
    expect(buttons).toHaveLength(2);
    await user.click(buttons[1]);
    expect(onMatch).toHaveBeenCalledWith("contact-annie-2");
    expect(onMatch).not.toHaveBeenCalledWith("contact-annie-1");
  });

  it("clicking Ignore calls onIgnore, not onMatch", async () => {
    const user = userEvent.setup();
    const { onIgnore, onMatch } = renderPanel();
    await user.click(screen.getByRole("button", { name: "Ignore" }));
    expect(onIgnore).toHaveBeenCalledTimes(1);
    expect(onMatch).not.toHaveBeenCalled();
  });

  it("clicking Add Person calls onAddPerson, not onMatch", async () => {
    const user = userEvent.setup();
    const { onAddPerson, onMatch } = renderPanel();
    await user.click(screen.getByRole("button", { name: "Add Person" }));
    expect(onAddPerson).toHaveBeenCalledTimes(1);
    expect(onMatch).not.toHaveBeenCalled();
  });

  it("zero candidates renders calm 'no reliable match' language, with no Match Existing button anywhere — Add Person and Ignore remain available", () => {
    renderPanel({ candidates: [] });
    expect(screen.getByText("No reliable match found.")).toBeTruthy();
    expect(screen.queryByRole("button", { name: "Match Existing" })).toBeFalsy();
    expect(screen.getByRole("button", { name: "Add Person" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Ignore" })).toBeTruthy();
  });

  it("multiple same-name candidates are all rendered, not silently narrowed to one", () => {
    const first = candidate({ contact: contact({ appId: "contact-annie-1" }) });
    const second = candidate({ contact: contact({ appId: "contact-annie-2", organizationId: "org-other" }) });
    renderPanel({ candidates: [first, second] });
    expect(screen.getByText(/2 possible matches/)).toBeTruthy();
  });
});

describe("Matched decision", () => {
  it("renders the matched Contact's name and organization, and offers Change instead of the picker", () => {
    renderPanel({ decision: { type: "matched", contactAppId: "contact-annie" } });
    expect(screen.getByText(/Matched Contact: Annie Milewski · North Valley SD/)).toBeTruthy();
    expect(screen.queryByRole("button", { name: "Match Existing" })).toBeFalsy();
    expect(screen.getByRole("button", { name: "Change" })).toBeTruthy();
  });

  it("clicking Change calls onReset, allowing the human to reconsider before anything is saved", async () => {
    const user = userEvent.setup();
    const { onReset } = renderPanel({ decision: { type: "matched", contactAppId: "contact-annie" } });
    await user.click(screen.getByRole("button", { name: "Change" }));
    expect(onReset).toHaveBeenCalledTimes(1);
  });
});

describe("Ignored decision", () => {
  it("renders 'Ignored for this item' and offers Reconsider", () => {
    renderPanel({ decision: { type: "ignored" } });
    expect(screen.getByText("Ignored for this item")).toBeTruthy();
    expect(screen.getByRole("button", { name: "Reconsider" })).toBeTruthy();
  });

  it("clicking Reconsider calls onReset", async () => {
    const user = userEvent.setup();
    const { onReset } = renderPanel({ decision: { type: "ignored" } });
    await user.click(screen.getByRole("button", { name: "Reconsider" }));
    expect(onReset).toHaveBeenCalledTimes(1);
  });
});
