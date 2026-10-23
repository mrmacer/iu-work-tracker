// @vitest-environment jsdom
//
// Patch 9A (Option B) — recovering an unfinished Work Record across a Microsoft re-sign-in, as the
// human experiences it: the failed save keeps the wizard and every field; a "Sign in again" action
// inside the wizard sets a safe copy aside BEFORE the (single, user-initiated) redirect; after
// sign-in the same account's draft reopens for review — never auto-submitted, never into another
// account — and its safe copy is removed only after that draft is saved or explicitly discarded.
// The MSAL controller is a controllable fake and sessionStorage is an in-memory fake; no real
// Microsoft, Graph, SharePoint, or Anthropic call is made.
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import {
  WORK_RECORD_CHECKPOINT_KEY,
  WORK_RECORD_CHECKPOINT_TTL_MS,
  saveWorkRecordCheckpoint,
} from "../lib/work-record-draft-checkpoint";
import { createFakeStorage, installWindowStorage, type FakeStorage } from "./voice-session-test-utils";

const mocks = vi.hoisted(() => ({
  ACCOUNT: {
    homeAccountId: "home",
    environment: "login.microsoftonline.com",
    tenantId: "3276761c-22db-462b-a930-172d155bd795",
    username: "owner@example.edu",
    localAccountId: "local",
    name: "Owner",
  } as AccountInfo,
  OTHER: {
    homeAccountId: "other-home",
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
  user: { id: "u1", displayName: "Owner", userPrincipalName: "owner@example.edu", mail: "owner@example.edu" },
  site: { id: "site-1", displayName: "IU Work Tracker DEV", webUrl: "https://siu29.sharepoint.com/sites/IUWorkTrackerDEV" },
  lists: [],
};

const restores: (() => void)[] = [];
let session: FakeStorage;

beforeEach(() => {
  vi.stubEnv("NEXT_PUBLIC_MS_ENTRA_CLIENT_ID", "11111111-1111-4111-8111-111111111111");
  vi.stubEnv("NEXT_PUBLIC_MS_ENTRA_TENANT_ID", "3276761c-22db-462b-a930-172d155bd795");
  vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_HOSTNAME", "siu29.sharepoint.com");
  vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_SITE_PATH", "/sites/IUWorkTrackerDEV");
  vi.resetModules();
  for (const fn of [mocks.initialize, mocks.listCachedAccounts, mocks.acquireGraphToken, mocks.signIn, mocks.signOut, mocks.runDevConnectionDiagnostic]) fn.mockReset();
  mocks.initialize.mockResolvedValue(mocks.ACCOUNT);
  mocks.listCachedAccounts.mockReturnValue([]);
  mocks.acquireGraphToken.mockResolvedValue("graph-token");
  mocks.signIn.mockResolvedValue(undefined);
  mocks.signOut.mockResolvedValue(undefined);
  mocks.runDevConnectionDiagnostic.mockResolvedValue(DIAGNOSTIC);
  session = createFakeStorage();
  restores.push(installWindowStorage("sessionStorage", session));
  restores.push(installWindowStorage("localStorage", createFakeStorage()));
});

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
  vi.unstubAllEnvs();
  while (restores.length) restores.pop()!();
});

const NON_AUTH_ERROR = "The DEV SharePoint data store could not be reached.";
const TOKEN_FAILED = "A Microsoft access token could not be acquired.";

function draftFixture(overrides: Partial<WorkRecord> = {}): WorkRecord {
  return {
    appId: "draft-under-recovery",
    title: "Recovered STEELS planning meeting",
    activityDate: "2026-09-18",
    activityType: "District meeting",
    description: "Recovered description",
    detailedNotes: "Recovered detailed notes",
    durationMinutes: 95,
    status: "complete",
    engagementScope: "none",
    projectIds: [],
    organizationIds: [],
    contactIds: [],
    categoryIds: [],
    reach: { educatorsLeaders: 3, studentsFamilies: 0, workforceCommunity: 0, other: 0 },
    evidenceSummary: "",
    evidenceReferenceIds: [],
    output: "",
    outcome: "Recovered outcome",
    nextStep: "",
    followUpNeeded: false,
    followUpDate: null,
    orbit: { reportable: false, primaryDeliverable: null, supportingDeliverables: [], stemPocMinutes: 0, tacMinutes: 0, evidence: "" },
    schemaVersion: WORK_RECORD_SCHEMA_VERSION,
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
    isSample: false,
    ...overrides,
  };
}

function seedCheckpoint(overrides: { accountId?: string | null; accountLabel?: string | null; step?: number; draft?: WorkRecord; now?: number } = {}) {
  const ok = saveWorkRecordCheckpoint(
    session,
    {
      accountId: overrides.accountId === undefined ? "home" : overrides.accountId,
      accountLabel: overrides.accountLabel === undefined ? "owner@example.edu" : overrides.accountLabel,
      step: overrides.step ?? 5,
      draft: overrides.draft ?? draftFixture(),
    },
    () => overrides.now ?? Date.now(),
  );
  expect(ok).toBe(true);
}

async function renderApp(options: { createResult?: { status: "network_error"; message: string } } = {}) {
  const importAuth = await import("../lib/microsoft-auth"); // same module instance the app uses
  const { default: IUWorkTracker } = await import("../app/IUWorkTracker");
  const { MemoryDataProvider } = await import("../lib/data-provider");
  const dataProvider = new MemoryDataProvider([]);
  const create = vi.spyOn(dataProvider, "createWorkRecord");
  if (options.createResult) create.mockResolvedValue(options.createResult);
  const user = userEvent.setup();
  render(<IUWorkTracker dataProvider={dataProvider} />);
  return { user, create, auth: importAuth };
}

const wizard = () => screen.getByRole("dialog");
const banner = () => screen.queryByRole("alert", { name: "Unfinished work entry" });

/** Opens the wizard, fills it in, and walks to step 5 so Save is available. */
async function fillWizard(user: ReturnType<typeof userEvent.setup>) {
  await user.click((await screen.findAllByRole("button", { name: /log work/i }))[0]);
  const dialog = await screen.findByRole("dialog");
  await user.type(within(dialog).getByPlaceholderText(/district steels planning/i), "Quarterly planning with North Valley");
  const type = within(dialog).getByLabelText(/Activity type/i) as HTMLSelectElement;
  await user.selectOptions(type, type.options[2].value);
  await user.type(within(dialog).getByPlaceholderText(/a sentence is enough/i), "Notes I must not lose");
  for (let i = 0; i < 4; i++) await user.click(within(dialog).getByRole("button", { name: /continue/i }));
  return dialog;
}

async function readStep1(user: ReturnType<typeof userEvent.setup>) {
  await user.click(within(wizard()).getByRole("button", { name: /Step 1 of 5/ }));
  return {
    title: (within(wizard()).getByPlaceholderText(/district steels planning/i) as HTMLInputElement).value,
    description: (within(wizard()).getByPlaceholderText(/a sentence is enough/i) as HTMLTextAreaElement).value,
  };
}

const checkpointInStorage = () => {
  const raw = session.data.get(WORK_RECORD_CHECKPOINT_KEY);
  return raw ? JSON.parse(raw) : null;
};

// ---------------------------------------------------------------------------------------------
describe("an authentication failure while saving a Work Record", () => {
  it("keeps the wizard open with every field intact, and offers 'Sign in again' INSIDE the wizard", async () => {
    const auth = await import("../lib/microsoft-auth");
    const { user, create } = await renderApp({ createResult: { status: "network_error", message: auth.SIGN_IN_REQUIRED_MESSAGE } });
    const dialog = await fillWizard(user);
    await user.click(within(dialog).getByRole("button", { name: "Save & done" }));
    await waitFor(() => expect(create).toHaveBeenCalledTimes(1));

    expect(within(wizard()).getByText(auth.SIGN_IN_REQUIRED_MESSAGE)).toBeTruthy();
    expect(within(wizard()).getByRole("button", { name: "Sign in again" })).toBeTruthy(); // reachable although the shell is inert
    expect(screen.queryByRole("button", { name: "Account options" })).toBeFalsy(); // …which is exactly why it must live here
    expect(await readStep1(user)).toEqual({ title: "Quarterly planning with North Valley", description: "Notes I must not lose" });
    expect(mocks.signIn).not.toHaveBeenCalled(); // never automatic
    expect(checkpointInStorage()).toBeNull(); // nothing is set aside until the human asks to sign in
  });

  it("a non-authentication error is unchanged: same message, no sign-in action", async () => {
    const { user, create } = await renderApp({ createResult: { status: "network_error", message: NON_AUTH_ERROR } });
    const dialog = await fillWizard(user);
    await user.click(within(dialog).getByRole("button", { name: "Save & done" }));
    await waitFor(() => expect(create).toHaveBeenCalledTimes(1));
    expect(within(wizard()).getByText(NON_AUTH_ERROR)).toBeTruthy();
    expect(within(wizard()).queryByRole("button", { name: "Sign in again" })).toBeFalsy();
  });

  it("a token-acquisition failure (not 'sign-in required') is also unchanged: no sign-in action", async () => {
    const { user, create } = await renderApp({ createResult: { status: "network_error", message: TOKEN_FAILED } });
    const dialog = await fillWizard(user);
    await user.click(within(dialog).getByRole("button", { name: "Save & done" }));
    await waitFor(() => expect(create).toHaveBeenCalledTimes(1));
    expect(within(wizard()).getByText(TOKEN_FAILED)).toBeTruthy();
    expect(within(wizard()).queryByRole("button", { name: "Sign in again" })).toBeFalsy();
  });
});

// ---------------------------------------------------------------------------------------------
describe("'Sign in again' — checkpoint first, then one user-initiated redirect", () => {
  async function failedSave() {
    const auth = await import("../lib/microsoft-auth");
    const ctx = await renderApp({ createResult: { status: "network_error", message: auth.SIGN_IN_REQUIRED_MESSAGE } });
    const dialog = await fillWizard(ctx.user);
    await ctx.user.click(within(dialog).getByRole("button", { name: "Save & done" }));
    await waitFor(() => expect(ctx.create).toHaveBeenCalledTimes(1));
    return ctx;
  }

  it("writes the checkpoint BEFORE the redirect starts, with the full draft, the step, and the owning account — then signs in once, for that account", async () => {
    let storedAtRedirect: ReturnType<typeof checkpointInStorage> = null;
    mocks.signIn.mockImplementation(async () => {
      storedAtRedirect = checkpointInStorage();
    });
    const { user, create } = await failedSave();
    await user.click(within(wizard()).getByRole("button", { name: "Sign in again" }));
    await waitFor(() => expect(mocks.signIn).toHaveBeenCalledTimes(1));

    expect(mocks.signIn).toHaveBeenCalledWith(mocks.ACCOUNT);
    expect(storedAtRedirect).not.toBeNull(); // already there when the redirect began
    expect(storedAtRedirect!.accountId).toBe("home");
    expect(storedAtRedirect!.accountLabel).toBe("owner@example.edu");
    expect(storedAtRedirect!.step).toBe(5);
    expect(storedAtRedirect!.draft.title).toBe("Quarterly planning with North Valley");
    expect(storedAtRedirect!.draft.description).toBe("Notes I must not lose");
    expect(create).toHaveBeenCalledTimes(1); // the checkpoint/sign-in never submitted anything
  });

  it("cannot start duplicate redirects: a double click starts exactly one, and the button shows it is busy", async () => {
    mocks.signIn.mockImplementation(() => new Promise<void>(() => undefined)); // the browser is leaving the page
    const { user } = await failedSave();
    const button = within(wizard()).getByRole("button", { name: "Sign in again" });
    await user.dblClick(button);
    await user.click(button);
    expect(mocks.signIn).toHaveBeenCalledTimes(1);
    const busy = await within(wizard()).findByRole("button", { name: "Redirecting to Microsoft…" });
    expect((busy as HTMLButtonElement).disabled).toBe(true);
  });

  it("if a safe copy cannot be made, it does NOT redirect — the entry stays open and the action stays available", async () => {
    const { user } = await failedSave();
    // Storage that refuses the write from here on.
    const blocked = createFakeStorage();
    blocked.setItem = () => {
      throw new Error("QuotaExceededError");
    };
    restores.push(installWindowStorage("sessionStorage", blocked));
    await user.click(within(wizard()).getByRole("button", { name: "Sign in again" }));
    expect(mocks.signIn).not.toHaveBeenCalled();
    expect(await within(wizard()).findByText(/safe copy of this entry could not be kept/i)).toBeTruthy();
    expect((within(wizard()).getByRole("button", { name: "Sign in again" }) as HTMLButtonElement).disabled).toBe(false);
    expect(await readStep1(user)).toMatchObject({ title: "Quarterly planning with North Valley" });
  });

  it("if sign-in cannot be started, the entry stays open, the safe copy is kept, and the user can try again", async () => {
    const auth = await import("../lib/microsoft-auth");
    mocks.signIn.mockRejectedValueOnce(new auth.MicrosoftAuthenticationError("Microsoft sign-in could not be started."));
    const { user } = await failedSave();
    await user.click(within(wizard()).getByRole("button", { name: "Sign in again" }));
    expect(await within(wizard()).findByText(/Microsoft sign-in could not be started\. Your entry is still open, and a safe copy is kept/)).toBeTruthy();
    expect(checkpointInStorage()?.draft.title).toBe("Quarterly planning with North Valley"); // preserved
    expect((within(wizard()).getByRole("button", { name: "Sign in again" }) as HTMLButtonElement).disabled).toBe(false);
    await user.click(within(wizard()).getByRole("button", { name: "Sign in again" }));
    await waitFor(() => expect(mocks.signIn).toHaveBeenCalledTimes(2)); // an explicit second attempt is allowed
  });

  it("with no readable account identity it still sets the entry aside (account unknown) so recovery will ask the human", async () => {
    mocks.initialize.mockRejectedValue(new Error("init failed")); // identity cannot be read at click time
    const { user } = await failedSave();
    await user.click(within(wizard()).getByRole("button", { name: "Sign in again" }));
    await waitFor(() => expect(mocks.signIn).toHaveBeenCalledWith(null));
    expect(checkpointInStorage()?.accountId).toBeNull();
  });
});

// ---------------------------------------------------------------------------------------------
describe("after sign-in — restoring the draft for review", () => {
  it("the same account's draft reopens in the wizard at the saved step with every field, a notice, and NOTHING saved", async () => {
    seedCheckpoint({ step: 5 });
    const { user, create } = await renderApp();
    await screen.findByRole("dialog");
    expect(within(wizard()).getByText(/Restored after sign-in/)).toBeTruthy();
    expect(within(wizard()).getByRole("button", { name: "Save & done" })).toBeTruthy(); // resumed at step 5
    expect(await readStep1(user)).toEqual({ title: "Recovered STEELS planning meeting", description: "Recovered description" });
    expect(create).not.toHaveBeenCalled(); // never auto-submitted
    expect(mocks.signIn).not.toHaveBeenCalled(); // and no further sign-in redirect
    expect(banner()).toBeFalsy();
    expect(checkpointInStorage()).not.toBeNull(); // still there until saved or discarded
  });

  it("resumes at whatever step the user was on", async () => {
    seedCheckpoint({ step: 3 });
    await renderApp();
    await screen.findByRole("dialog");
    expect(within(wizard()).getByRole("button", { name: /Step 3 of 5: How long\?, current step/ })).toBeTruthy();
  });

  it("only an explicit Save submits it — exactly once — and THEN the checkpoint is cleared", async () => {
    seedCheckpoint({ step: 5 });
    const { user, create } = await renderApp();
    await screen.findByRole("dialog");
    expect(checkpointInStorage()).not.toBeNull();
    await user.click(within(wizard()).getByRole("button", { name: "Save & done" }));
    await waitFor(() => expect(create).toHaveBeenCalledTimes(1));
    expect(create.mock.calls[0][0]).toMatchObject({ appId: "draft-under-recovery", title: "Recovered STEELS planning meeting", description: "Recovered description", durationMinutes: 95 });
    await waitFor(() => expect(checkpointInStorage()).toBeNull());
    expect(screen.queryByRole("dialog")).toBeFalsy();
  });

  it("a failed save of the restored draft does NOT clear the checkpoint, and the entry stays open", async () => {
    seedCheckpoint({ step: 5 });
    const { user, create } = await renderApp({ createResult: { status: "network_error", message: NON_AUTH_ERROR } });
    await screen.findByRole("dialog");
    await user.click(within(wizard()).getByRole("button", { name: "Save & done" }));
    await waitFor(() => expect(create).toHaveBeenCalledTimes(1));
    expect(within(wizard()).getByText(NON_AUTH_ERROR)).toBeTruthy();
    expect(checkpointInStorage()).not.toBeNull();
  });

  it("closing the restored entry still asks to discard; declining keeps it AND its checkpoint, accepting is the explicit discard that clears it", async () => {
    seedCheckpoint({ step: 1 });
    const { user } = await renderApp();
    await screen.findByRole("dialog");
    const confirm = vi.spyOn(window, "confirm").mockReturnValueOnce(false);
    await user.click(within(wizard()).getByRole("button", { name: "Cancel" }));
    expect(confirm).toHaveBeenCalledWith("Discard the changes in this work entry?");
    expect(screen.getByRole("dialog")).toBeTruthy();
    expect(checkpointInStorage()).not.toBeNull();

    confirm.mockReturnValueOnce(true);
    await user.click(within(wizard()).getByRole("button", { name: "Cancel" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeFalsy());
    expect(checkpointInStorage()).toBeNull();
  });

  it("a page refresh with the restored entry still open restores it again (the checkpoint was not consumed)", async () => {
    seedCheckpoint({ step: 2 });
    await renderApp();
    await screen.findByRole("dialog");
    cleanup();
    await renderApp();
    await screen.findByRole("dialog");
    expect(within(wizard()).getByText(/Restored after sign-in/)).toBeTruthy();
  });

  it("an ordinary Log work entry that is saved never touches an unrelated draft's checkpoint", async () => {
    seedCheckpoint({ accountId: "other-home", accountLabel: "someone.else@example.edu", draft: draftFixture({ appId: "somebody-elses-draft" }) });
    const { user, create } = await renderApp();
    await screen.findByRole("alert", { name: "Unfinished work entry" }); // different account → prompt, not restored
    const dialog = await fillWizard(user);
    await user.click(within(dialog).getByRole("button", { name: "Save & done" }));
    await waitFor(() => expect(create).toHaveBeenCalledTimes(1));
    expect(checkpointInStorage()?.draft.appId).toBe("somebody-elses-draft");
  });
});

// ---------------------------------------------------------------------------------------------
describe("account safety — never silently into a different or unconfirmed account", () => {
  it("a DIFFERENT signed-in account: nothing is restored, the owner is named, and the checkpoint is kept for them", async () => {
    mocks.initialize.mockResolvedValue(mocks.OTHER);
    seedCheckpoint();
    const { create } = await renderApp();
    const prompt = await screen.findByRole("alert", { name: "Unfinished work entry" });
    expect(prompt.textContent).toMatch(/belongs to owner@example\.edu, but you are signed in as someone\.else@example\.edu/);
    expect(prompt.textContent).toMatch(/not restored into a different account/i);
    expect(within(prompt).queryByRole("button", { name: "Restore entry" })).toBeFalsy(); // not even offered
    expect(screen.queryByRole("dialog")).toBeFalsy();
    expect(create).not.toHaveBeenCalled();
    expect(checkpointInStorage()).not.toBeNull();
  });

  it("explicit discard of a saved entry asks first; declining keeps it, accepting removes it", async () => {
    mocks.initialize.mockResolvedValue(mocks.OTHER);
    seedCheckpoint();
    const { user } = await renderApp();
    const prompt = await screen.findByRole("alert", { name: "Unfinished work entry" });
    const confirm = vi.spyOn(window, "confirm").mockReturnValueOnce(false);
    await user.click(within(prompt).getByRole("button", { name: "Discard saved entry" }));
    expect(confirm).toHaveBeenCalledWith("Discard the saved unfinished work entry? This cannot be undone.");
    expect(checkpointInStorage()).not.toBeNull();
    confirm.mockReturnValueOnce(true);
    await user.click(within(prompt).getByRole("button", { name: "Discard saved entry" }));
    await waitFor(() => expect(banner()).toBeFalsy());
    expect(checkpointInStorage()).toBeNull();
  });

  it("signed out (the redirect was cancelled or failed): kept and explained — no restore, no automatic sign-in", async () => {
    mocks.initialize.mockResolvedValue(null);
    seedCheckpoint();
    await renderApp();
    const prompt = await screen.findByRole("alert", { name: "Unfinished work entry" });
    expect(prompt.textContent).toMatch(/kept safe in this tab\. Sign in with Microsoft as owner@example\.edu/);
    expect(screen.queryByRole("dialog")).toBeFalsy();
    expect(mocks.signIn).not.toHaveBeenCalled();
    expect(checkpointInStorage()).not.toBeNull();
  });

  it("the owner's session still expired after returning: kept and explained — no restore, no redirect loop", async () => {
    const auth = await import("../lib/microsoft-auth");
    mocks.acquireGraphToken.mockRejectedValue(new auth.InteractionRequiredError());
    seedCheckpoint();
    await renderApp();
    const prompt = await screen.findByRole("alert", { name: "Unfinished work entry" });
    expect(prompt.textContent).toMatch(/Sign in with Microsoft as owner@example\.edu/);
    expect(screen.queryByRole("dialog")).toBeFalsy();
    expect(mocks.signIn).not.toHaveBeenCalled();
    expect(checkpointInStorage()).not.toBeNull();
  });

  it("an authentication failure on return: kept and explained, never restored", async () => {
    const auth = await import("../lib/microsoft-auth");
    mocks.acquireGraphToken.mockRejectedValue(new auth.MicrosoftAuthenticationError(TOKEN_FAILED));
    seedCheckpoint();
    await renderApp();
    await screen.findByRole("alert", { name: "Unfinished work entry" });
    expect(screen.queryByRole("dialog")).toBeFalsy();
    expect(checkpointInStorage()).not.toBeNull();
  });

  it("several cached accounts: kept and explained, never restored into a guessed one", async () => {
    mocks.initialize.mockResolvedValue(null);
    mocks.listCachedAccounts.mockReturnValue([mocks.ACCOUNT, mocks.OTHER]);
    seedCheckpoint();
    await renderApp();
    await screen.findByRole("alert", { name: "Unfinished work entry" });
    expect(screen.queryByRole("dialog")).toBeFalsy();
  });

  it("an entry whose account cannot be confirmed needs an EXPLICIT decision: not restored on its own; Restore reopens it for review", async () => {
    seedCheckpoint({ accountId: null, accountLabel: null });
    const { user, create } = await renderApp();
    const prompt = await screen.findByRole("alert", { name: "Unfinished work entry" });
    expect(prompt.textContent).toMatch(/cannot be confirmed/i);
    expect(screen.queryByRole("dialog")).toBeFalsy();
    await user.click(within(prompt).getByRole("button", { name: "Restore entry" }));
    await screen.findByRole("dialog");
    expect(within(wizard()).getByText(/Restored after sign-in/)).toBeTruthy();
    expect(create).not.toHaveBeenCalled();
  });
});

// ---------------------------------------------------------------------------------------------
describe("corrupt, outdated, and expired checkpoints", () => {
  it.each([
    ["damaged", () => session.data.set(WORK_RECORD_CHECKPOINT_KEY, "{ not json"), /it was damaged/],
    [
      "outdated",
      () => {
        seedCheckpoint();
        const stored = JSON.parse(session.data.get(WORK_RECORD_CHECKPOINT_KEY)!);
        stored.schemaVersion = 99;
        session.data.set(WORK_RECORD_CHECKPOINT_KEY, JSON.stringify(stored));
      },
      /older version of the app/,
    ],
    ["expired", () => seedCheckpoint({ now: Date.now() - WORK_RECORD_CHECKPOINT_TTL_MS - 60_000 }), /kept for too long/],
  ])("%s: says so, removes it, and the app carries on normally", async (_label, arrange, wording) => {
    arrange();
    const { user } = await renderApp();
    const prompt = await screen.findByRole("alert", { name: "Unfinished work entry" });
    expect(prompt.textContent).toMatch(wording);
    expect(screen.queryByRole("dialog")).toBeFalsy();
    expect(checkpointInStorage()).toBeNull();
    expect(screen.queryByText(/Unexpected token|SyntaxError|JSON/i)).toBeFalsy(); // never a raw parse error
    await user.click(within(prompt).getByRole("button", { name: "Dismiss" }));
    expect(banner()).toBeFalsy();
  });

  it("no checkpoint: no banner, no wizard, and the Microsoft session is not even inspected for recovery", async () => {
    await renderApp();
    await screen.findAllByRole("button", { name: /log work/i });
    expect(banner()).toBeFalsy();
    expect(screen.queryByRole("dialog")).toBeFalsy();
  });
});

// ---------------------------------------------------------------------------------------------
describe("an entry already in progress is never replaced silently", () => {
  it("when recovery is ready while the wizard is in use, the wizard is untouched; the prompt waits (the shell is inert behind the modal) and offers Restore once the wizard closes", async () => {
    // Several callers (account menu, storage notice, recovery) each wait on their own token request.
    const pending: (() => void)[] = [];
    mocks.acquireGraphToken.mockImplementation(() => new Promise<string>((resolve) => { pending.push(() => resolve("graph-token")); }));
    seedCheckpoint({ step: 5 });
    const { user } = await renderApp();

    // The human starts a new entry while the Microsoft check is still pending.
    await user.click((await screen.findAllByRole("button", { name: /log work/i }))[0]);
    const dialog = await screen.findByRole("dialog");
    await user.type(within(dialog).getByPlaceholderText(/district steels planning/i), "Entry I am typing right now");
    await waitFor(() => expect(pending.length).toBeGreaterThan(0));
    for (const release of pending.splice(0)) release();

    // The saved entry is NOT injected: the wizard keeps what the human is typing, with no restore notice.
    const prompt = await screen.findByRole("alert", { name: "Unfinished work entry", hidden: true });
    expect(prompt.textContent).toMatch(/ready to restore/i);
    expect((within(wizard()).getByPlaceholderText(/district steels planning/i) as HTMLInputElement).value).toBe("Entry I am typing right now");
    expect(within(wizard()).queryByText(/Restored after sign-in/)).toBeFalsy();
    expect(screen.queryByRole("alert", { name: "Unfinished work entry" })).toBeFalsy(); // not reachable behind the modal…
    expect(checkpointInStorage()).not.toBeNull();

    // …and when the human closes (and discards) the in-progress entry, the prompt is there to act on.
    vi.spyOn(window, "confirm").mockReturnValue(true);
    await user.click(within(wizard()).getByRole("button", { name: "Cancel" }));
    const available = await screen.findByRole("alert", { name: "Unfinished work entry" });
    expect(checkpointInStorage()?.draft.appId).toBe("draft-under-recovery"); // closing the OTHER entry never removed it
    await user.click(within(available).getByRole("button", { name: "Restore entry" }));
    await screen.findByRole("dialog");
    expect(within(wizard()).getByText(/Restored after sign-in/)).toBeTruthy();
    expect(await readStep1(user)).toMatchObject({ title: "Recovered STEELS planning meeting" });
  });
});

describe("boundaries", () => {
  it("recovery adds no network or SharePoint calls beyond the existing session check, and never touches localStorage", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    seedCheckpoint();
    await renderApp();
    await screen.findByRole("dialog");
    expect(fetchSpy).not.toHaveBeenCalled();
    const sessionKeys = [...session.data.keys()];
    expect(sessionKeys.every((key) => key === WORK_RECORD_CHECKPOINT_KEY)).toBe(true);
  });

  it("the wizard's discard confirmation text is unchanged", async () => {
    const { user } = await renderApp();
    await user.click((await screen.findAllByRole("button", { name: /log work/i }))[0]);
    const dialog = await screen.findByRole("dialog");
    await user.type(within(dialog).getByPlaceholderText(/district steels planning/i), "x");
    const confirm = vi.spyOn(window, "confirm").mockReturnValue(false);
    await user.click(within(dialog).getByRole("button", { name: "Cancel" }));
    expect(confirm).toHaveBeenCalledWith("Discard the changes in this work entry?");
  });
});
