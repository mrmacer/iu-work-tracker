"use client";

import { useEffect, useRef, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import {
  createBrowserMicrosoftAuthController,
  MicrosoftAuthenticationError,
  type MicrosoftAuthController,
} from "../lib/microsoft-auth";
import { readDevMicrosoftConfig, type DevMicrosoftPublicConfig } from "../lib/microsoft-auth-config";
import { sessionStatusFromTokenError } from "../lib/microsoft-session";
import {
  runDevConnectionDiagnostic,
  type DevConnectionDiagnostic,
} from "../lib/microsoft-graph";

// Patch 9A — authentication and SharePoint problems are distinct states (previously one
// catch-all "error"): "sign-in-required" = the cached account's sign-in expired (the human
// chooses when to sign in again — nothing redirects on its own), "auth-error" = Microsoft
// authentication itself failed (including a failed initialization, which Retry re-attempts),
// "sharepoint-error" = signed in fine but the SharePoint/Graph check failed.
type ConnectionState =
  | { status: "initializing" }
  | { status: "disconnected" }
  | { status: "account-choice"; accounts: AccountInfo[] }
  | { status: "checking"; account: AccountInfo }
  | { status: "sign-in-required"; account: AccountInfo }
  | {
      status: "connected";
      account: AccountInfo;
      diagnostic: DevConnectionDiagnostic;
    }
  | { status: "auth-error"; account: AccountInfo | null; message: string }
  | { status: "sharepoint-error"; account: AccountInfo; message: string }
  | { status: "config-error"; message: string };

const publicConfig = readDevMicrosoftConfig();

function safeErrorMessage(error: unknown): string {
  return error instanceof Error
    ? error.message
    : "The Microsoft connection check failed.";
}

// Only MicrosoftAuthenticationError text is written for users; anything else gets a safe default.
function authErrorMessage(error: unknown): string {
  return error instanceof MicrosoftAuthenticationError
    ? error.message
    : "Microsoft authentication could not be completed.";
}

/**
 * Restores the cached session, then verifies SharePoint. Re-callable: a failed initialization is
 * retried by the controller, so Retry genuinely starts over. It never redirects — sign-in is
 * always an explicit click — and it reports every outcome through `setState` as a distinct state.
 */
async function runConnectionCheck(
  auth: MicrosoftAuthController,
  config: DevMicrosoftPublicConfig,
  isActive: () => boolean,
  setState: (state: ConnectionState) => void,
): Promise<void> {
  let account: AccountInfo | null;
  try {
    account = await auth.initialize();
  } catch (error) {
    if (isActive()) setState({ status: "auth-error", account: null, message: authErrorMessage(error) });
    return;
  }
  if (!isActive()) return;
  if (!account) {
    const accounts = auth.listCachedAccounts();
    setState(accounts.length > 1 ? { status: "account-choice", accounts } : { status: "disconnected" });
    return;
  }
  setState({ status: "checking", account });
  let accessToken: string;
  try {
    accessToken = await auth.acquireGraphToken(account);
  } catch (error) {
    if (!isActive()) return;
    const session = sessionStatusFromTokenError(error, account);
    setState(
      session.status === "sign-in-required"
        ? { status: "sign-in-required", account }
        : { status: "auth-error", account, message: authErrorMessage(error) },
    );
    return;
  }
  try {
    const diagnostic = await runDevConnectionDiagnostic(accessToken, config);
    if (isActive()) setState({ status: "connected", account, diagnostic });
  } catch (error) {
    if (isActive()) setState({ status: "sharepoint-error", account, message: safeErrorMessage(error) });
  }
}

function initials(name: string | undefined): string {
  const parts = name?.trim().split(/\s+/).filter(Boolean) ?? [];
  if (parts.length === 0) return "MS";
  return `${parts[0][0]}${parts.at(-1)?.[0] ?? ""}`.toUpperCase();
}

/**
 * The header's account control. This is a single account menu, always driven by the app's
 * own Microsoft-authenticated identity (never a hosting platform's identity — see
 * docs/AI_HANDOFF.md). When DEV Microsoft/SharePoint configuration is present it also shows
 * the existing connection diagnostic rows; when it is absent there is simply no Microsoft
 * account to show, and the menu says so rather than inventing an identity from elsewhere.
 */
export default function DevMicrosoftConnection() {
  const [open, setOpen] = useState(false);
  const controller = useRef<MicrosoftAuthController | null>(null);
  const [state, setState] = useState<ConnectionState>(() =>
    publicConfig.status === "invalid"
      ? { status: "config-error", message: publicConfig.message }
      : { status: "initializing" },
  );
  const container = useRef<HTMLDivElement>(null);
  const mounted = useRef(true);

  useEffect(() => {
    mounted.current = true;
    if (publicConfig.status === "enabled") {
      const auth = createBrowserMicrosoftAuthController(publicConfig.value, window.location.origin);
      controller.current = auth;
      void runConnectionCheck(auth, publicConfig.value, () => mounted.current, setState);
    }
    return () => {
      mounted.current = false;
    };
  }, []);

  useEffect(() => {
    if (!open) return;
    const closeWhenOutside = (event: PointerEvent) => {
      if (!container.current?.contains(event.target as Node)) setOpen(false);
    };
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") setOpen(false);
    };
    document.addEventListener("pointerdown", closeWhenOutside);
    document.addEventListener("keydown", closeOnEscape);
    return () => {
      document.removeEventListener("pointerdown", closeWhenOutside);
      document.removeEventListener("keydown", closeOnEscape);
    };
  }, [open]);

  const configured = publicConfig.status !== "disabled";
  const account = "account" in state ? state.account : null;
  // "Microsoft connected" means a token can actually be acquired — a cached account whose
  // sign-in expired, or whose token failed, is not shown as connected.
  const microsoftConnected =
    state.status === "checking" || state.status === "connected" || state.status === "sharepoint-error";
  const sharePointConnected = state.status === "connected";
  const busy = state.status === "initializing" || state.status === "checking";
  const needsSignIn =
    state.status === "disconnected" ||
    state.status === "account-choice" ||
    state.status === "sign-in-required" ||
    state.status === "auth-error";

  const primaryName =
    state.status === "connected"
      ? state.diagnostic.user.displayName
      : (account?.name ?? account?.username ?? "Not signed in");
  // MSAL's cached AccountInfo already carries the username (typically the UPN/email) —
  // reused here as-is, never a new Graph request merely to populate this menu.
  const secondaryIdentity =
    account?.username && account.username !== primaryName ? account.username : null;

  const retry = () => {
    if (!controller.current) return;
    setState({ status: "initializing" });
    if (publicConfig.status === "enabled") {
      void runConnectionCheck(controller.current, publicConfig.value, () => mounted.current, setState);
    }
  };

  const signIn = async () => {
    if (!controller.current) return;
    try {
      // With an expired cached account, re-authenticate that account; otherwise open the
      // account chooser. The controller allows only one redirect at a time.
      await controller.current.signIn(state.status === "sign-in-required" ? state.account : null);
    } catch (error) {
      setState({ status: "auth-error", account, message: authErrorMessage(error) });
    }
  };

  const signOut = async () => {
    if (!controller.current) return;
    try {
      await controller.current.signOut(account);
    } catch (error) {
      setState({ status: "auth-error", account, message: authErrorMessage(error) });
    }
  };

  const attentionMessage =
    state.status === "sign-in-required"
      ? "Your Microsoft sign-in has expired. Sign in again to reconnect SharePoint."
      : state.status === "account-choice"
        ? "Several Microsoft accounts are saved in this browser. Sign in to choose which one to use."
        : state.status === "auth-error" || state.status === "sharepoint-error" || state.status === "config-error"
          ? state.message
          : null;

  return (
    <div className="dev-ms-connection" ref={container}>
      <button
        className={`avatar dev-ms-trigger ${sharePointConnected ? "connected" : ""}`}
        aria-label="Account options"
        aria-haspopup="true"
        aria-expanded={open}
        aria-controls="account-options-panel"
        onClick={() => setOpen((value) => !value)}
      >
        {initials(account?.name)}
      </button>
      {open && (
        <section
          className="dev-ms-panel"
          id="account-options-panel"
          aria-label="Account options"
        >
          <p className="dev-ms-kicker">Account</p>
          {configured && (
            <>
              <ConnectionRow
                label="Microsoft"
                connected={microsoftConnected}
                busy={busy}
                offLabel={state.status === "sign-in-required" ? "Sign-in required" : undefined}
              />
              <ConnectionRow
                label="SharePoint DEV"
                connected={sharePointConnected}
                busy={state.status === "checking"}
              />
            </>
          )}
          <div className="dev-ms-user">
            <span>Signed in as</span>
            <strong>{primaryName}</strong>
            {secondaryIdentity && <span>{secondaryIdentity}</span>}
          </div>
          {state.status === "connected" && (
            <div className="dev-ms-detail">
              <span>Graph site ID</span>
              <code>{state.diagnostic.site.id}</code>
              <span>{state.diagnostic.lists.length} existing lists readable</span>
            </div>
          )}
          {attentionMessage && (
            <p className="dev-ms-error" role="alert">
              {attentionMessage}
            </p>
          )}
          {configured && state.status !== "config-error" && (
            <div className="dev-ms-actions">
              {needsSignIn && (
                <button onClick={() => void signIn()} disabled={busy}>
                  Sign in with Microsoft
                </button>
              )}
              {(state.status === "auth-error" || account !== null) && !needsSignIn && (
                <button onClick={retry} disabled={busy}>
                  {busy ? "Checking…" : "Retry check"}
                </button>
              )}
              {state.status === "auth-error" && (
                <button className="secondary" onClick={retry} disabled={busy}>
                  Try again
                </button>
              )}
              {account !== null && state.status !== "account-choice" && (
                <button className="secondary" onClick={() => void signOut()}>
                  Sign out
                </button>
              )}
            </div>
          )}
        </section>
      )}
    </div>
  );
}

function ConnectionRow({
  label,
  connected,
  busy,
  offLabel = "Not connected",
}: {
  label: string;
  connected: boolean;
  busy: boolean;
  offLabel?: string;
}) {
  return (
    <div className="dev-ms-row">
      <span>{label}</span>
      <strong className={connected ? "is-connected" : ""}>
        {busy ? "Checking…" : connected ? "Connected" : offLabel}
      </strong>
    </div>
  );
}
