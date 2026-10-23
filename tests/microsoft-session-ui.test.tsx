// @vitest-environment jsdom
//
// Patch 9A — persistent Microsoft sign-in, as the human sees it: the account menu's distinct
// states (signed out / sign-in required / choose an account / authentication failed / SharePoint
// unavailable / connected), a Retry that really re-initializes, explicit sign-in and sign-out,
// no automatic redirect, and the app's storage banner + footer never describing a failed
// authenticated load as an ordinary Preview session. The MSAL controller is a controllable fake
// (module-scoped env stubs + vi.resetModules, exactly like tests/account-menu.test.tsx); no real
// Microsoft, Graph, SharePoint, or Anthropic call is ever made.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import type { ProviderResult } from "../lib/data-provider";
import type { WorkRecord } from "../lib/models";

const mocks = vi.hoisted(() => ({
  ACCOUNT: {
    homeAccountId: "home",
    environment: "login.microsoftonline.com",
    tenantId: "3276761c-22db-462b-a930-172d155bd795",
    username: "greg.macer@example.edu",
    localAccountId: "local",
    name: "Greg Macer",
  } as AccountInfo,
  OTHER: {
    homeAccountId: "other",
    environment: "login.microsoftonline.com",
    tenantId: "3276761c-22db-462b-a930-172d155bd795",
    username: "someone.else@example.edu",
    localAccountId: "other-local",
    name: "Someone Else",
  } as AccountInfo,
  initialize: vi.fn(),
  listCachedAccounts: vi.fn(),
  acquireGraphToken: vi.fn(),
  signIn: vi.fn(),
  signOut: vi.fn(),
  runDevConnectionDiagnostic: vi.fn(),
}));

vi.mock("../lib/microsoft-auth", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../lib/microsoft-auth")>();
  return {
    ...actual,
    createBrowserMicrosoftAuthController: vi.fn(() => ({
      initialize: mocks.initialize,
      listCachedAccounts: mocks.listCachedAccounts,
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

const DIAGNOSTIC = {
  user: { id: "u1", displayName: "Greg Macer", userPrincipalName: "greg.macer@example.edu", mail: "greg.macer@example.edu" },
  site: { id: "site-1", displayName: "IU Work Tracker DEV", webUrl: "https://siu29.sharepoint.com/sites/IUWorkTrackerDEV" },
  lists: [],
};

beforeEach(() => {
  vi.stubEnv("NEXT_PUBLIC_MS_ENTRA_CLIENT_ID", "11111111-1111-4111-8111-111111111111");
  vi.stubEnv("NEXT_PUBLIC_MS_ENTRA_TENANT_ID", "3276761c-22db-462b-a930-172d155bd795");
  vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_HOSTNAME", "siu29.sharepoint.com");
  vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_SITE_PATH", "/sites/IUWorkTrackerDEV");
  vi.resetModules();
  for (const fn of [mocks.initialize, mocks.listCachedAccounts, mocks.acquireGraphToken, mocks.signIn, mocks.signOut, mocks.runDevConnectionDiagnostic]) fn.mockReset();
  // Defaults: a returning, healthy user.
  mocks.initialize.mockResolvedValue(mocks.ACCOUNT);
  mocks.listCachedAccounts.mockReturnValue([]);
  mocks.acquireGraphToken.mockResolvedValue("graph-token");
  mocks.signIn.mockResolvedValue(undefined);
  mocks.signOut.mockResolvedValue(undefined);
  mocks.runDevConnectionDiagnostic.mockResolvedValue(DIAGNOSTIC);
});

afterEach(() => {
  cleanup();
  vi.unstubAllEnvs();
});

// Error classes must come from the SAME module instance the app code gets (vi.resetModules runs
// per test), so tests import them here, before rendering, and configure the fake with them.
const importAuth = () => import("../lib/microsoft-auth");

async function loadMenu() {
  const { default: DevMicrosoftConnection } = await import("../app/DevMicrosoftConnection");
  const user = userEvent.setup();
  render(<DevMicrosoftConnection />);
  const open = async () => {
    await user.click(screen.getByRole("button", { name: "Account options" }));
    return screen.getByRole("region", { name: "Account options" });
  };
  return { user, open };
}

const row = (panel: HTMLElement, label: string) => within(panel).getByText(label).closest(".dev-ms-row") as HTMLElement;

describe("account menu — returning user", () => {
  it("restores a cached session with no sign-in click: Microsoft + SharePoint connected, and nothing ever starts a redirect", async () => {
    const { open } = await loadMenu();
    await waitFor(() => expect(mocks.runDevConnectionDiagnostic).toHaveBeenCalledTimes(1));
    const panel = await open();
    await waitFor(() => expect(within(row(panel, "SharePoint DEV")).getByText("Connected")).toBeTruthy());
    expect(within(row(panel, "Microsoft")).getByText("Connected")).toBeTruthy();
    expect(within(panel).queryByRole("button", { name: "Sign in with Microsoft" })).toBeFalsy();
    expect(within(panel).getByRole("button", { name: "Sign out" })).toBeTruthy();
    expect(mocks.signIn).not.toHaveBeenCalled();
  });
});

describe("account menu — distinct states", () => {
  it("signed out: offers an explicit sign-in, opens the account chooser (no cached account), and starts nothing by itself", async () => {
    mocks.initialize.mockResolvedValue(null);
    const { user, open } = await loadMenu();
    const panel = await open();
    await waitFor(() => expect(within(panel).getByRole("button", { name: "Sign in with Microsoft" })).toBeTruthy());
    expect(mocks.signIn).not.toHaveBeenCalled();
    expect(within(panel).queryByRole("alert")).toBeFalsy();
    await user.click(within(panel).getByRole("button", { name: "Sign in with Microsoft" }));
    expect(mocks.signIn).toHaveBeenCalledTimes(1);
    expect(mocks.signIn).toHaveBeenCalledWith(null);
  });

  it("sign-in required: says the sign-in expired, is NOT shown as connected, and Sign in re-authenticates that account", async () => {
    const auth = await importAuth();
    mocks.acquireGraphToken.mockRejectedValue(new auth.InteractionRequiredError());
    const { user, open } = await loadMenu();
    const panel = await open();
    await waitFor(() => expect(within(panel).getByRole("alert").textContent).toMatch(/sign-in has expired/i));
    expect(within(row(panel, "Microsoft")).getByText("Sign-in required")).toBeTruthy();
    expect(within(row(panel, "SharePoint DEV")).getByText("Not connected")).toBeTruthy();
    expect(mocks.signIn).not.toHaveBeenCalled(); // never automatic
    expect(mocks.runDevConnectionDiagnostic).not.toHaveBeenCalled();
    await user.click(within(panel).getByRole("button", { name: "Sign in with Microsoft" }));
    expect(mocks.signIn).toHaveBeenCalledTimes(1);
    expect(mocks.signIn).toHaveBeenCalledWith(mocks.ACCOUNT);
  });

  it("several cached accounts: asks the human to choose and offers no Sign out for a guessed account", async () => {
    mocks.initialize.mockResolvedValue(null);
    mocks.listCachedAccounts.mockReturnValue([mocks.ACCOUNT, mocks.OTHER]);
    const { user, open } = await loadMenu();
    const panel = await open();
    await waitFor(() => expect(within(panel).getByRole("alert").textContent).toMatch(/several microsoft accounts/i));
    expect(within(panel).queryByRole("button", { name: "Sign out" })).toBeFalsy();
    expect(within(panel).getByText("Not signed in")).toBeTruthy();
    expect(mocks.signIn).not.toHaveBeenCalled();
    await user.click(within(panel).getByRole("button", { name: "Sign in with Microsoft" }));
    expect(mocks.signIn).toHaveBeenCalledWith(null); // account chooser, never an assumed account
  });

  it("authentication failed: shows the real error and the account stays identified but NOT connected", async () => {
    const auth = await importAuth();
    mocks.acquireGraphToken.mockRejectedValue(new auth.MicrosoftAuthenticationError("A Microsoft access token could not be acquired."));
    const { open } = await loadMenu();
    const panel = await open();
    await waitFor(() => expect(within(panel).getByRole("alert").textContent).toBe("A Microsoft access token could not be acquired."));
    expect(within(row(panel, "Microsoft")).getByText("Not connected")).toBeTruthy();
    expect(within(panel).getByText("Greg Macer")).toBeTruthy();
    expect(within(panel).getByRole("button", { name: "Sign in with Microsoft" })).toBeTruthy();
    expect(within(panel).getByRole("button", { name: "Sign out" })).toBeTruthy();
  });

  it("SharePoint unavailable is NOT an authentication problem: Microsoft stays Connected, SharePoint is not, no sign-in prompt", async () => {
    mocks.runDevConnectionDiagnostic.mockRejectedValue(new Error("SharePoint returned HTTP 503."));
    const { open } = await loadMenu();
    const panel = await open();
    await waitFor(() => expect(within(panel).getByRole("alert").textContent).toBe("SharePoint returned HTTP 503."));
    expect(within(row(panel, "Microsoft")).getByText("Connected")).toBeTruthy();
    expect(within(row(panel, "SharePoint DEV")).getByText("Not connected")).toBeTruthy();
    expect(within(panel).queryByRole("button", { name: "Sign in with Microsoft" })).toBeFalsy();
    expect(within(panel).getByRole("button", { name: "Retry check" })).toBeTruthy();
  });
});

describe("account menu — initialization failure then retry", () => {
  it("a failed initialization is shown as an error with Try again, and Try again re-initializes and recovers", async () => {
    const auth = await importAuth();
    mocks.initialize.mockReset();
    mocks.initialize.mockRejectedValueOnce(new auth.MicrosoftAuthenticationError()).mockResolvedValue(mocks.ACCOUNT);
    const { user, open } = await loadMenu();
    const panel = await open();
    await waitFor(() => expect(within(panel).getByRole("alert").textContent).toBe("Microsoft authentication could not be completed."));
    expect(mocks.signIn).not.toHaveBeenCalled();
    await user.click(within(panel).getByRole("button", { name: "Try again" }));
    await waitFor(() => expect(within(row(panel, "SharePoint DEV")).getByText("Connected")).toBeTruthy());
    expect(mocks.initialize).toHaveBeenCalledTimes(2);
    expect(within(panel).queryByRole("alert")).toBeFalsy();
  });

  it("a failing sign-in start is reported, not swallowed", async () => {
    const auth = await importAuth();
    mocks.initialize.mockResolvedValue(null);
    mocks.signIn.mockRejectedValue(new auth.MicrosoftAuthenticationError("Microsoft sign-in could not be started."));
    const { user, open } = await loadMenu();
    const panel = await open();
    await user.click(await within(panel).findByRole("button", { name: "Sign in with Microsoft" }));
    await waitFor(() => expect(within(panel).getByRole("alert").textContent).toBe("Microsoft sign-in could not be started."));
  });
});

describe("explicit sign-out", () => {
  it("Sign out calls the existing sign-out with the signed-in account, exactly once, and never signs in", async () => {
    const { user, open } = await loadMenu();
    const panel = await open();
    await user.click(await within(panel).findByRole("button", { name: "Sign out" }));
    expect(mocks.signOut).toHaveBeenCalledTimes(1);
    expect(mocks.signOut).toHaveBeenCalledWith(mocks.ACCOUNT);
    expect(mocks.signIn).not.toHaveBeenCalled();
  });
});

// ---------------------------------------------------------------------------------------------
// The whole app: storage banner + footer.
// ---------------------------------------------------------------------------------------------

async function loadApp(options: { failWorkRecordsWith?: string } = {}) {
  const { default: IUWorkTracker } = await import("../app/IUWorkTracker");
  const { MemoryDataProvider } = await import("../lib/data-provider");
  const dataProvider = new MemoryDataProvider([]);
  if (options.failWorkRecordsWith) {
    vi.spyOn(dataProvider, "getWorkRecords").mockResolvedValue({
      status: "network_error",
      message: options.failWorkRecordsWith,
    } satisfies ProviderResult<WorkRecord[]>);
  }
  render(<IUWorkTracker dataProvider={dataProvider} />);
}

const banner = () => document.querySelector(".storage-banner") as HTMLElement | null;
const footer = () => document.querySelector(".os-status") as HTMLElement;

describe("storage banner and footer — Preview vs. failed authenticated load", () => {
  it("signed out: the ordinary Preview banner (unchanged wording), not an alert, and the plain footer", async () => {
    mocks.initialize.mockResolvedValue(null);
    await loadApp();
    await waitFor(() => expect(banner()?.textContent).toMatch(/Preview mode/));
    expect(banner()!.textContent).toBe("Preview mode Changes in this session are not saved. Sign in with Microsoft (top right) to connect SharePoint and save durably.");
    expect(banner()!.getAttribute("role")).toBeNull();
    expect(footer().textContent).toContain("Preview session active");
    expect(footer().textContent).not.toContain("SharePoint not connected");
  });

  it("signed in but the SharePoint load FAILED: an alert naming SharePoint and the cause — never the ordinary Preview text", async () => {
    await loadApp({ failWorkRecordsWith: "The DEV SharePoint data store could not be reached." });
    await waitFor(() => expect(banner()?.textContent).toMatch(/SharePoint unavailable/));
    expect(banner()!.getAttribute("role")).toBe("alert");
    expect(banner()!.className).toContain("attention");
    expect(banner()!.textContent).toContain("You are signed in, but SharePoint could not be loaded (The DEV SharePoint data store could not be reached.)");
    expect(banner()!.textContent).toMatch(/not saved/i);
    expect(banner()!.textContent).not.toMatch(/Preview mode/);
    expect(footer().textContent).toContain("Preview session active — SharePoint not connected");
  });

  it("expired sign-in during load: 'Sign-in required' — and the app does NOT redirect on its own", async () => {
    const auth = await importAuth();
    mocks.acquireGraphToken.mockRejectedValue(new auth.InteractionRequiredError());
    await loadApp({ failWorkRecordsWith: "Microsoft sign-in has expired. Sign in again to reconnect SharePoint." });
    await waitFor(() => expect(banner()?.textContent).toMatch(/Sign-in required/));
    expect(banner()!.getAttribute("role")).toBe("alert");
    expect(banner()!.textContent).toMatch(/not saved/i);
    expect(banner()!.textContent).not.toMatch(/SharePoint could not be loaded/); // the cause is the sign-in, not SharePoint
    expect(mocks.signIn).not.toHaveBeenCalled();
  });

  it("several cached accounts: 'Choose an account' rather than a silent Preview", async () => {
    mocks.initialize.mockResolvedValue(null);
    mocks.listCachedAccounts.mockReturnValue([mocks.ACCOUNT, mocks.OTHER]);
    await loadApp();
    await waitFor(() => expect(banner()?.textContent).toMatch(/Choose an account/));
    expect(banner()!.getAttribute("role")).toBe("alert");
    expect(mocks.signIn).not.toHaveBeenCalled();
  });

  it("authentication failure: shows the real authentication message", async () => {
    const auth = await importAuth();
    mocks.acquireGraphToken.mockRejectedValue(new auth.MicrosoftAuthenticationError("A Microsoft access token could not be acquired."));
    await loadApp({ failWorkRecordsWith: "A Microsoft access token could not be acquired." });
    await waitFor(() => expect(banner()?.textContent).toMatch(/Microsoft sign-in problem/));
    expect(banner()!.textContent).toContain("A Microsoft access token could not be acquired.");
  });

  it("never claims SharePoint is connected while running in memory", async () => {
    await loadApp({ failWorkRecordsWith: "The DEV SharePoint data store could not be reached." });
    await waitFor(() => expect(banner()?.textContent).toMatch(/SharePoint unavailable/));
    expect(screen.queryByText(/SharePoint DEV connected/)).toBeFalsy();
  });
});
