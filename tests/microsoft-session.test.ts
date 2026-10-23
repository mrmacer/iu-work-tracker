// Patch 9A — session-state classification, the honest storage notice, and consistent
// authentication-error surfacing across every SharePoint-backed provider.
import { afterEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import {
  InteractionRequiredError,
  MicrosoftAuthenticationError,
  type MicrosoftAuthController,
} from "../lib/microsoft-auth";
import {
  describeStorageNotice,
  initialStorageNotice,
  inspectMicrosoftSession,
  resolveStorageNotice,
  sessionStatusFromTokenError,
  type MicrosoftSessionStatus,
} from "../lib/microsoft-session";
import { DelegatedSharePointContactProvider } from "../lib/contact-provider";
import { DelegatedSharePointDataProvider } from "../lib/data-provider";
import { DelegatedSharePointInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { DelegatedSharePointMeetingRecordProvider } from "../lib/meeting-record-provider";
import { DelegatedSharePointOrganizationProvider } from "../lib/organization-provider";
import { DelegatedSharePointProjectProvider } from "../lib/project-provider";

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllEnvs();
});

const account = { homeAccountId: "home", username: "dev@example.edu", name: "Dev User" } as AccountInfo;
const other = { homeAccountId: "other", username: "other@example.edu", name: "Other" } as AccountInfo;

function controller(overrides: Partial<Record<"initialize" | "listCachedAccounts" | "acquireGraphToken", unknown>> = {}) {
  return {
    initialize: vi.fn(async () => account),
    listCachedAccounts: vi.fn(() => [] as AccountInfo[]),
    acquireGraphToken: vi.fn(async () => "token"),
    ...overrides,
  } as unknown as Pick<MicrosoftAuthController, "initialize" | "listCachedAccounts" | "acquireGraphToken">;
}

describe("inspectMicrosoftSession — every state is distinguished", () => {
  it("signed-out: nothing cached", async () => {
    expect(await inspectMicrosoftSession(controller({ initialize: vi.fn(async () => null) }))).toEqual({ status: "signed-out" });
  });

  it("account-choice-required: several cached, none selected — never a guess", async () => {
    const status = await inspectMicrosoftSession(
      controller({ initialize: vi.fn(async () => null), listCachedAccounts: vi.fn(() => [account, other]) }),
    );
    expect(status).toEqual({ status: "account-choice-required", accounts: [account, other] });
  });

  it("connected: cached account + a silently acquired token", async () => {
    expect(await inspectMicrosoftSession(controller())).toEqual({ status: "connected", account });
  });

  it("sign-in-required: the cached account's session expired (interaction required)", async () => {
    const status = await inspectMicrosoftSession(controller({ acquireGraphToken: vi.fn(async () => { throw new InteractionRequiredError(); }) }));
    expect(status).toEqual({ status: "sign-in-required", account });
  });

  it("auth-failed: a non-interactive token failure keeps the account and a safe message", async () => {
    const status = await inspectMicrosoftSession(
      controller({ acquireGraphToken: vi.fn(async () => { throw new MicrosoftAuthenticationError("A Microsoft access token could not be acquired."); }) }),
    );
    expect(status).toEqual({ status: "auth-failed", account, message: "A Microsoft access token could not be acquired." });
  });

  it("auth-failed: a failed initialization has no account", async () => {
    const status = await inspectMicrosoftSession(
      controller({ initialize: vi.fn(async () => { throw new MicrosoftAuthenticationError(); }) }),
    );
    expect(status).toEqual({ status: "auth-failed", account: null, message: "Microsoft authentication could not be completed." });
  });

  it("never leaks an unknown error's raw text", async () => {
    const status = await inspectMicrosoftSession(controller({ acquireGraphToken: vi.fn(async () => { throw new Error("raw upstream secret"); }) }));
    expect(JSON.stringify(status)).not.toMatch(/secret/);
  });

  it("an initialization failure followed by a retry succeeds on a second inspection", async () => {
    const initialize = vi.fn().mockRejectedValueOnce(new MicrosoftAuthenticationError()).mockResolvedValue(account);
    const c = controller({ initialize });
    expect((await inspectMicrosoftSession(c)).status).toBe("auth-failed");
    expect((await inspectMicrosoftSession(c)).status).toBe("connected");
  });

  it("sessionStatusFromTokenError maps the two token failure kinds", () => {
    expect(sessionStatusFromTokenError(new InteractionRequiredError(), account).status).toBe("sign-in-required");
    expect(sessionStatusFromTokenError(new Error("x"), account).status).toBe("auth-failed");
  });
});

describe("storage notice — Preview is never confused with a failed authenticated load", () => {
  const sessions: Record<string, MicrosoftSessionStatus> = {
    signedOut: { status: "signed-out" },
    choice: { status: "account-choice-required", accounts: [account, other] },
    expired: { status: "sign-in-required", account },
    failed: { status: "auth-failed", account, message: "A Microsoft access token could not be acquired." },
    connected: { status: "connected", account },
  };

  it("signed out is the ordinary Preview notice (unchanged wording) and is not flagged for attention", () => {
    const notice = describeStorageNotice({ session: sessions.signedOut, loadFailureMessage: null });
    expect(notice).toEqual({
      tone: "preview",
      label: "Preview mode",
      message: "Changes in this session are not saved. Sign in with Microsoft (top right) to connect SharePoint and save durably.",
    });
    expect(initialStorageNotice(false)).toEqual(notice); // the pre-patch first paint is byte-identical
  });

  it("each distinct condition gets its own label and demands attention", () => {
    const labels = {
      choice: describeStorageNotice({ session: sessions.choice, loadFailureMessage: null }),
      expired: describeStorageNotice({ session: sessions.expired, loadFailureMessage: null }),
      failed: describeStorageNotice({ session: sessions.failed, loadFailureMessage: null }),
      sharePoint: describeStorageNotice({ session: sessions.connected, loadFailureMessage: "The DEV SharePoint data store could not be reached." }),
    };
    expect(labels.choice.label).toBe("Choose an account");
    expect(labels.expired.label).toBe("Sign-in required");
    expect(labels.failed.label).toBe("Microsoft sign-in problem");
    expect(labels.sharePoint.label).toBe("SharePoint unavailable");
    for (const notice of Object.values(labels)) {
      expect(notice.tone).toBe("attention");
      expect(notice.label).not.toBe("Preview mode");
    }
  });

  it("a signed-in session whose SharePoint load FAILED says so, quotes the cause, and is not 'Preview mode'", () => {
    const notice = describeStorageNotice({ session: sessions.connected, loadFailureMessage: "The DEV SharePoint data store could not be reached." });
    expect(notice.message).toMatch(/You are signed in, but SharePoint could not be loaded \(The DEV SharePoint data store could not be reached\.\)/);
    expect(notice.message).not.toMatch(/Sign in with Microsoft \(top right\) to connect SharePoint and save durably/);
  });

  it("the auth-failed notice shows the real authentication message", () => {
    expect(describeStorageNotice({ session: sessions.failed, loadFailureMessage: null }).message).toContain("A Microsoft access token could not be acquired.");
  });

  it("an expired sign-in outranks a load-failure message (the cause is the sign-in, not SharePoint)", () => {
    const notice = describeStorageNotice({ session: sessions.expired, loadFailureMessage: "The DEV SharePoint data store could not be reached." });
    expect(notice.label).toBe("Sign-in required");
  });

  it("EVERY notice states that changes are not saved — nothing ever implies durability", () => {
    for (const session of Object.values(sessions)) {
      for (const loadFailureMessage of [null, "boom"]) {
        expect(describeStorageNotice({ session, loadFailureMessage }).message).toMatch(/not saved/i);
      }
    }
    expect(initialStorageNotice(true).message).toMatch(/not saved/i);
    expect(initialStorageNotice(false).message).toMatch(/not saved/i);
  });

  it("while a configured session is still being inspected, the notice is a neutral 'checking', not a sign-in demand", () => {
    expect(initialStorageNotice(true).label).toBe("Checking Microsoft session…");
    expect(initialStorageNotice(true).tone).toBe("preview");
  });
});

describe("resolveStorageNotice", () => {
  it("with no Microsoft configuration it returns the plain Preview notice and never touches MSAL", async () => {
    expect(await resolveStorageNotice("anything")).toEqual(initialStorageNotice(false));
  });
});

// ---------------------------------------------------------------------------------------------
// Consistent authentication errors across all six providers.
// ---------------------------------------------------------------------------------------------

describe("every SharePoint-backed provider surfaces the authentication error, not the generic 'could not be reached'", () => {
  const GENERIC = "The DEV SharePoint data store could not be reached.";
  const failing = (error: Error) => ({ acquireGraphToken: vi.fn(async () => { throw error; }) }) as unknown as MicrosoftAuthController;

  const makers: [string, (c: MicrosoftAuthController) => () => Promise<{ status: string; message?: string }>][] = [
    ["Work Records", (c) => { const p = new DelegatedSharePointDataProvider(c, account, { siteId: "s", workRecordsListId: "l" }); return () => p.getWorkRecords(); }],
    ["Contacts", (c) => { const p = new DelegatedSharePointContactProvider(c, account, { siteId: "s", contactsListId: "l" }); return () => p.list(); }],
    ["Projects", (c) => { const p = new DelegatedSharePointProjectProvider(c, account, { siteId: "s", projectsListId: "l" }); return () => p.list(); }],
    ["Organizations", (c) => { const p = new DelegatedSharePointOrganizationProvider(c, account, { siteId: "s", organizationsListId: "l" }); return () => p.list(); }],
    ["Meeting Records", (c) => { const p = new DelegatedSharePointMeetingRecordProvider(c, account, { siteId: "s", meetingRecordsListId: "l" }); return () => p.list(); }],
    ["Inbox Intelligence", (c) => { const p = new DelegatedSharePointInboxIntelligenceProvider(c, account, { siteId: "s", inboxIntelligenceListId: "l" }); return () => p.list(); }],
  ];

  it.each(makers)("%s — a failed token", async (_name, make) => {
    const result = await make(failing(new MicrosoftAuthenticationError("A Microsoft access token could not be acquired.")))();
    expect(result).toEqual({ status: "network_error", message: "A Microsoft access token could not be acquired." });
  });

  it.each(makers)("%s — an expired sign-in", async (_name, make) => {
    const result = await make(failing(new InteractionRequiredError()))();
    expect(result.status).toBe("network_error");
    expect(result.message).toMatch(/sign-in has expired/i);
    expect(result.message).not.toBe(GENERIC);
  });

  it.each(makers)("%s — an unrelated error still uses the generic fallback (unchanged)", async (_name, make) => {
    expect(await make(failing(new Error("something else")))()).toEqual({ status: "network_error", message: GENERIC });
  });
});
