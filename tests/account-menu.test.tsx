// @vitest-environment jsdom
//
// Focused tests for the header's account popover (docs/AI_HANDOFF.md "Account menu"). The
// menu is driven entirely by the app's own Microsoft-authenticated identity — never a
// hosting platform's identity (Patch 3A retired the ChatGPT Sites identity wiring that
// Patch 2.6 had introduced). `publicConfig` inside app/DevMicrosoftConnection.tsx is
// evaluated once at module load from NEXT_PUBLIC_MS_ENTRA_*/SHAREPOINT_* env vars, so the
// "configured" scenarios stub those env vars and re-import the module fresh via
// vi.resetModules() rather than relying on any runtime prop. The mocks below are registered
// unconditionally at file scope (the standard vi.mock/vi.hoisted pattern); they are inert
// for the "no configuration" describe block below since that code path never calls them.
import { cleanup, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import * as fs from "node:fs";
import * as path from "node:path";

const mocks = vi.hoisted(() => ({
  ACCOUNT: {
    homeAccountId: "home",
    environment: "login.microsoftonline.com",
    tenantId: "3276761c-22db-462b-a930-172d155bd795",
    username: "greg.macer@example.edu",
    localAccountId: "local",
    name: "Greg Macer",
  } as AccountInfo,
  acquireGraphToken: vi.fn(async () => "graph-token"),
  signIn: vi.fn(async () => undefined),
  signOut: vi.fn(async () => undefined),
  runDevConnectionDiagnostic: vi.fn(async () => ({
    user: { id: "u1", displayName: "Greg Macer", userPrincipalName: "greg.macer@example.edu", mail: "greg.macer@example.edu" },
    site: { id: "site-1", displayName: "IU Work Tracker DEV", webUrl: "https://siu29.sharepoint.com/sites/IUWorkTrackerDEV" },
    lists: [],
  })),
}));

vi.mock("../lib/microsoft-auth", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../lib/microsoft-auth")>();
  return {
    ...actual,
    createBrowserMicrosoftAuthController: vi.fn(() => ({
      initialize: vi.fn(async () => mocks.ACCOUNT),
      acquireGraphToken: mocks.acquireGraphToken,
      signIn: mocks.signIn,
      signOut: mocks.signOut,
    })),
  };
});

vi.mock("../lib/microsoft-graph", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../lib/microsoft-graph")>();
  return { ...actual, runDevConnectionDiagnostic: mocks.runDevConnectionDiagnostic };
});

afterEach(() => {
  cleanup();
});

describe("account popover — no Microsoft configuration (default test/dev env)", () => {
  it("is a real button, closed by default, with popup semantics, showing no fabricated identity", async () => {
    const { default: DevMicrosoftConnection } = await import("../app/DevMicrosoftConnection");
    render(<DevMicrosoftConnection />);
    const trigger = screen.getByRole("button", { name: "Account options" });
    expect(trigger.getAttribute("aria-haspopup")).toBe("true");
    expect(trigger.getAttribute("aria-expanded")).toBe("false");
    expect(screen.queryByText("Signed in as")).toBeFalsy();
  });

  it("opens to show 'Not signed in' and no sign-in/out action when Microsoft is not configured", async () => {
    const user = userEvent.setup();
    const { default: DevMicrosoftConnection } = await import("../app/DevMicrosoftConnection");
    render(<DevMicrosoftConnection />);
    await user.click(screen.getByRole("button", { name: "Account options" }));
    expect(screen.getByText("Signed in as")).toBeTruthy();
    expect(screen.getByText("Not signed in")).toBeTruthy();
    expect(screen.queryByRole("button", { name: /sign in with microsoft/i })).toBeFalsy();
    expect(screen.queryByRole("button", { name: "Sign out" })).toBeFalsy();
    expect(screen.queryByText("Microsoft")).toBeFalsy();
  });

  it("closes on a second click, outside interaction, and Escape", async () => {
    const user = userEvent.setup();
    const { default: DevMicrosoftConnection } = await import("../app/DevMicrosoftConnection");
    render(
      <div>
        <DevMicrosoftConnection />
        <button>Elsewhere</button>
      </div>,
    );
    const trigger = screen.getByRole("button", { name: "Account options" });

    await user.click(trigger);
    expect(screen.getByText("Signed in as")).toBeTruthy();
    await user.click(trigger);
    expect(screen.queryByText("Signed in as")).toBeFalsy();

    await user.click(trigger);
    await user.click(screen.getByRole("button", { name: "Elsewhere" }));
    await waitFor(() => expect(screen.queryByText("Signed in as")).toBeFalsy());

    await user.click(trigger);
    await user.keyboard("{Escape}");
    await waitFor(() => expect(screen.queryByText("Signed in as")).toBeFalsy());
  });

  it("no longer imports or depends on a ChatGPT Sites identity module", () => {
    const chatgptAuthPath = path.resolve(__dirname, "../app/chatgpt-auth.ts");
    expect(fs.existsSync(chatgptAuthPath)).toBe(false);
    const source = fs.readFileSync(path.resolve(__dirname, "../app/DevMicrosoftConnection.tsx"), "utf8");
    expect(source).not.toMatch(/chatgpt/i);
  });
});

describe("account popover — Microsoft configured and signed in", () => {
  beforeEach(() => {
    vi.stubEnv("NEXT_PUBLIC_MS_ENTRA_CLIENT_ID", "11111111-1111-4111-8111-111111111111");
    vi.stubEnv("NEXT_PUBLIC_MS_ENTRA_TENANT_ID", "3276761c-22db-462b-a930-172d155bd795");
    vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_HOSTNAME", "siu29.sharepoint.com");
    vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_SITE_PATH", "/sites/IUWorkTrackerDEV");
    vi.resetModules();
    mocks.acquireGraphToken.mockClear();
    mocks.signIn.mockClear();
    mocks.signOut.mockClear();
    mocks.runDevConnectionDiagnostic.mockClear();
  });

  afterEach(() => {
    vi.unstubAllEnvs();
  });

  it("shows the Microsoft account's name and username/UPN as 'Signed in as' — no extra Graph request to populate it", async () => {
    const user = userEvent.setup();
    const { default: DevMicrosoftConnection } = await import("../app/DevMicrosoftConnection");
    render(<DevMicrosoftConnection />);
    await user.click(screen.getByRole("button", { name: "Account options" }));
    await waitFor(() => expect(screen.getByText("Greg Macer")).toBeTruthy());
    expect(screen.getByText("greg.macer@example.edu")).toBeTruthy();
  });

  it("Sign out invokes the existing Microsoft logoutRedirect pathway", async () => {
    const user = userEvent.setup();
    const { default: DevMicrosoftConnection } = await import("../app/DevMicrosoftConnection");
    render(<DevMicrosoftConnection />);
    await user.click(screen.getByRole("button", { name: "Account options" }));
    const signOutButton = await screen.findByRole("button", { name: "Sign out" });
    await user.click(signOutButton);
    expect(mocks.signOut).toHaveBeenCalledTimes(1);
  });

  it("makes no additional Graph diagnostic call merely from opening/closing the menu", async () => {
    const user = userEvent.setup();
    const { default: DevMicrosoftConnection } = await import("../app/DevMicrosoftConnection");
    render(<DevMicrosoftConnection />);
    await waitFor(() => expect(mocks.runDevConnectionDiagnostic).toHaveBeenCalledTimes(1));
    const trigger = screen.getByRole("button", { name: "Account options" });
    await user.click(trigger);
    await user.click(trigger);
    await user.click(trigger);
    expect(mocks.runDevConnectionDiagnostic).toHaveBeenCalledTimes(1);
  });
});
