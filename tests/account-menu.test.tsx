// @vitest-environment jsdom
//
// Focused tests for the header avatar/account popover (docs/AI_HANDOFF.md "Account popover").
// These exercise DevMicrosoftConnection directly rather than the full IUWorkTracker shell,
// since the behavior under test is entirely local to this component. No NEXT_PUBLIC_MS_ENTRA_*
// env vars are set in this test environment, so publicConfig.status is "disabled" — exactly
// the previously-dead branch this patch restores. The DEV Microsoft-connected branch is
// unmodified aside from one added aria-haspopup attribute and is not re-tested here.
import { cleanup, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import DevMicrosoftConnection from "../app/DevMicrosoftConnection";
import type { ChatGPTUser } from "../app/chatgpt-auth";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

const USER: ChatGPTUser = {
  userId: "user-1",
  displayName: "Greg Macer",
  email: "greg@example.org",
  fullName: "Greg Macer",
};

function renderMenu(user: ChatGPTUser | null = USER) {
  return render(<DevMicrosoftConnection chatGPTUser={user} chatGPTSignOutHref="/signout-with-chatgpt?return_to=%2F" />);
}

describe("account popover", () => {
  it("is a real button, closed by default, with popup semantics", () => {
    renderMenu();
    const trigger = screen.getByRole("button", { name: "Account options" });
    expect(trigger.getAttribute("aria-haspopup")).toBe("true");
    expect(trigger.getAttribute("aria-expanded")).toBe("false");
    expect(screen.queryByText("Signed in as")).toBeFalsy();
  });

  it("opens the account menu on click", async () => {
    const user = userEvent.setup();
    renderMenu();
    await user.click(screen.getByRole("button", { name: "Account options" }));
    expect(screen.getByText("Signed in as")).toBeTruthy();
    expect(screen.getByRole("button", { name: "Account options" }).getAttribute("aria-expanded")).toBe("true");
  });

  it("closes on a second click of the avatar", async () => {
    const user = userEvent.setup();
    renderMenu();
    const trigger = screen.getByRole("button", { name: "Account options" });
    await user.click(trigger);
    expect(screen.getByText("Signed in as")).toBeTruthy();
    await user.click(trigger);
    expect(screen.queryByText("Signed in as")).toBeFalsy();
  });

  it("closes on outside interaction", async () => {
    const user = userEvent.setup();
    render(
      <div>
        <DevMicrosoftConnection chatGPTUser={USER} chatGPTSignOutHref="/signout-with-chatgpt" />
        <button>Elsewhere</button>
      </div>,
    );
    await user.click(screen.getByRole("button", { name: "Account options" }));
    expect(screen.getByText("Signed in as")).toBeTruthy();
    await user.click(screen.getByRole("button", { name: "Elsewhere" }));
    await waitFor(() => expect(screen.queryByText("Signed in as")).toBeFalsy());
  });

  it("closes on Escape", async () => {
    const user = userEvent.setup();
    renderMenu();
    await user.click(screen.getByRole("button", { name: "Account options" }));
    expect(screen.getByText("Signed in as")).toBeTruthy();
    await user.keyboard("{Escape}");
    await waitFor(() => expect(screen.queryByText("Signed in as")).toBeFalsy());
  });

  it("displays the authenticated identity from the existing session, with email only when distinct", async () => {
    const user = userEvent.setup();
    renderMenu();
    await user.click(screen.getByRole("button", { name: "Account options" }));
    expect(screen.getByText("Greg Macer")).toBeTruthy();
    expect(screen.getByText("greg@example.org")).toBeTruthy();
  });

  it("does not duplicate the email when displayName already is the email", async () => {
    const user = userEvent.setup();
    renderMenu({ userId: "u2", displayName: "greg@example.org", email: "greg@example.org", fullName: null });
    await user.click(screen.getByRole("button", { name: "Account options" }));
    expect(screen.getAllByText("greg@example.org")).toHaveLength(1);
  });

  it("shows 'Not signed in' and no Sign out action when no session user is available", async () => {
    const user = userEvent.setup();
    renderMenu(null);
    await user.click(screen.getByRole("button", { name: "Account options" }));
    expect(screen.getByText("Not signed in")).toBeTruthy();
    expect(screen.queryByRole("link", { name: "Sign out" })).toBeFalsy();
  });

  it("Sign out reuses the existing sign-out pathway's exact URL", async () => {
    const user = userEvent.setup();
    render(<DevMicrosoftConnection chatGPTUser={USER} chatGPTSignOutHref="/signout-with-chatgpt?return_to=%2F" />);
    await user.click(screen.getByRole("button", { name: "Account options" }));
    const signOut = screen.getByRole("link", { name: "Sign out" });
    expect(signOut.getAttribute("href")).toBe("/signout-with-chatgpt?return_to=%2F");
  });

  it("makes no network request merely from opening or closing the menu", async () => {
    const user = userEvent.setup();
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    renderMenu();
    const trigger = screen.getByRole("button", { name: "Account options" });
    await user.click(trigger);
    await user.click(trigger);
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});
