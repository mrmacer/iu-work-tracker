import type { AccountInfo } from "@azure/msal-browser";
import {
  createBrowserMicrosoftAuthController,
  InteractionRequiredError,
  MicrosoftAuthenticationError,
  type MicrosoftAuthController,
} from "./microsoft-auth";
import { readDevMicrosoftConfig } from "./microsoft-auth-config";

// Patch 9A — one shared vocabulary for "where is the Microsoft session?", so the account menu,
// the storage banner, and the footer never disagree and never describe an authentication
// problem as a harmless Preview session. Pure classification + one async inspector; nothing
// here stores a token, starts a redirect, or touches SharePoint.

export type MicrosoftSessionStatus =
  /** No Microsoft account is cached: the normal "not signed in yet" state. */
  | { status: "signed-out" }
  /** Several accounts are cached and none is selected — the human must choose; we never guess. */
  | { status: "account-choice-required"; accounts: AccountInfo[] }
  /** An account is cached but cannot renew a token silently: its sign-in expired. */
  | { status: "sign-in-required"; account: AccountInfo }
  /** Initialization or token acquisition failed for a reason other than "sign in again". */
  | { status: "auth-failed"; account: AccountInfo | null; message: string }
  /** A Graph token was acquired silently for the account. (SharePoint itself is checked separately.) */
  | { status: "connected"; account: AccountInfo };

const AUTH_FAILED_FALLBACK = "Microsoft authentication could not be completed.";

/** Maps a token-acquisition failure for a known account onto a session status. */
export function sessionStatusFromTokenError(error: unknown, account: AccountInfo): MicrosoftSessionStatus {
  if (error instanceof InteractionRequiredError) return { status: "sign-in-required", account };
  return {
    status: "auth-failed",
    account,
    message: error instanceof MicrosoftAuthenticationError ? error.message : AUTH_FAILED_FALLBACK,
  };
}

type InspectableController = Pick<MicrosoftAuthController, "initialize" | "listCachedAccounts" | "acquireGraphToken">;

/**
 * Initializes MSAL (restoring the cached account) and, when there is one, attempts a silent
 * token acquisition. Never redirects, never throws, never calls SharePoint.
 */
export async function inspectMicrosoftSession(controller: InspectableController): Promise<MicrosoftSessionStatus> {
  let account: AccountInfo | null;
  try {
    account = await controller.initialize();
  } catch (error) {
    return {
      status: "auth-failed",
      account: null,
      message: error instanceof MicrosoftAuthenticationError ? error.message : AUTH_FAILED_FALLBACK,
    };
  }
  if (!account) {
    const accounts = controller.listCachedAccounts();
    return accounts.length > 1 ? { status: "account-choice-required", accounts } : { status: "signed-out" };
  }
  try {
    await controller.acquireGraphToken(account);
    return { status: "connected", account };
  } catch (error) {
    return sessionStatusFromTokenError(error, account);
  }
}

// ---------------------------------------------------------------------------------------------
// Storage notice — what the banner/footer say when the app is NOT running on SharePoint.
// ---------------------------------------------------------------------------------------------

export type StorageNotice = {
  /** "preview" = an ordinary signed-out/unconfigured preview; "attention" = something needs fixing. */
  tone: "preview" | "attention";
  label: string;
  message: string;
};

const NOT_SAVED = "Changes in this session are not saved.";

const PREVIEW_NOTICE: StorageNotice = {
  tone: "preview",
  label: "Preview mode",
  message: `${NOT_SAVED} Sign in with Microsoft (top right) to connect SharePoint and save durably.`,
};

/** Synchronous first paint: the long-standing preview text, or a neutral "checking" while a configured session is being inspected. */
export function initialStorageNotice(microsoftConfigured: boolean): StorageNotice {
  return microsoftConfigured
    ? { tone: "preview", label: "Checking Microsoft session…", message: `${NOT_SAVED} SharePoint is not connected yet.` }
    : PREVIEW_NOTICE;
}

/**
 * The honest explanation for running in memory. A failed authenticated SharePoint load is NEVER
 * presented as an ordinary Preview session: the cause (sign-in expired, authentication failed,
 * several accounts, SharePoint unavailable) is stated, and every variant says changes are not saved.
 */
export function describeStorageNotice(input: {
  session: MicrosoftSessionStatus;
  /** The message from a failed SharePoint load, if the app tried to load and could not. */
  loadFailureMessage: string | null;
}): StorageNotice {
  const { session, loadFailureMessage } = input;
  switch (session.status) {
    case "account-choice-required":
      return {
        tone: "attention",
        label: "Choose an account",
        message: `Several Microsoft accounts are saved in this browser. Use Sign in with Microsoft (top right) to choose one. ${NOT_SAVED}`,
      };
    case "sign-in-required":
      return {
        tone: "attention",
        label: "Sign-in required",
        message: `Your Microsoft sign-in has expired. Use Sign in with Microsoft (top right) to reconnect SharePoint. This is a temporary preview session — ${NOT_SAVED.toLowerCase()}`,
      };
    case "auth-failed":
      return {
        tone: "attention",
        label: "Microsoft sign-in problem",
        message: `${session.message} Open the account menu (top right) to try again or sign in. This is a temporary preview session — ${NOT_SAVED.toLowerCase()}`,
      };
    case "connected":
      return loadFailureMessage
        ? {
            tone: "attention",
            label: "SharePoint unavailable",
            message: `You are signed in, but SharePoint could not be loaded (${loadFailureMessage}). This is a temporary preview session — ${NOT_SAVED.toLowerCase()}`,
          }
        : {
            tone: "preview",
            label: "Preview mode",
            message: `You are signed in, but SharePoint storage is not configured for this build. ${NOT_SAVED}`,
          };
    case "signed-out":
    default:
      return PREVIEW_NOTICE;
  }
}

/** Never rejects. Unconfigured Microsoft ⇒ the plain preview notice with no MSAL involvement. */
export async function resolveStorageNotice(loadFailureMessage: string | null): Promise<StorageNotice> {
  const config = readDevMicrosoftConfig();
  if (config.status !== "enabled") return PREVIEW_NOTICE;
  try {
    const controller = createBrowserMicrosoftAuthController(config.value, window.location.origin);
    return describeStorageNotice({ session: await inspectMicrosoftSession(controller), loadFailureMessage });
  } catch {
    return describeStorageNotice({
      session: { status: "auth-failed", account: null, message: AUTH_FAILED_FALLBACK },
      loadFailureMessage,
    });
  }
}

/**
 * The current Microsoft session for flows that must know WHICH account is signed in (Work Record
 * draft recovery). Null when Microsoft is not configured. Never rejects, never redirects.
 */
export async function inspectCurrentMicrosoftSession(): Promise<MicrosoftSessionStatus | null> {
  const config = readDevMicrosoftConfig();
  if (config.status !== "enabled") return null;
  try {
    return await inspectMicrosoftSession(createBrowserMicrosoftAuthController(config.value, window.location.origin));
  } catch {
    return { status: "auth-failed", account: null, message: AUTH_FAILED_FALLBACK };
  }
}
