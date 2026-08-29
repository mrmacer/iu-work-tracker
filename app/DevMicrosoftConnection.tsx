"use client";

import { useEffect, useRef, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import {
  createBrowserMicrosoftAuthController,
  InteractiveRedirectStartedError,
  type MicrosoftAuthController,
} from "../lib/microsoft-auth";
import {
  readDevMicrosoftConfig,
  type DevMicrosoftPublicConfig,
} from "../lib/microsoft-auth-config";
import {
  runDevConnectionDiagnostic,
  type DevConnectionDiagnostic,
} from "../lib/microsoft-graph";

type ConnectionState =
  | { status: "initializing" }
  | { status: "disconnected" }
  | { status: "checking"; account: AccountInfo }
  | {
      status: "connected";
      account: AccountInfo;
      diagnostic: DevConnectionDiagnostic;
    }
  | { status: "error"; account: AccountInfo | null; message: string };

const publicConfig = readDevMicrosoftConfig();

function safeErrorMessage(error: unknown): string {
  return error instanceof Error
    ? error.message
    : "The DEV Microsoft connection check failed.";
}

function initials(name: string | undefined): string {
  const parts = name?.trim().split(/\s+/).filter(Boolean) ?? [];
  if (parts.length === 0) return "MS";
  return `${parts[0][0]}${parts.at(-1)?.[0] ?? ""}`.toUpperCase();
}

export default function DevMicrosoftConnection() {
  const [open, setOpen] = useState(false);
  const controller = useRef<MicrosoftAuthController | null>(null);
  const [state, setState] = useState<ConnectionState>(() =>
    publicConfig.status === "invalid"
      ? { status: "error", account: null, message: publicConfig.message }
      : { status: "initializing" },
  );
  const container = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (publicConfig.status !== "enabled") return;

    let current = true;
    const auth = createBrowserMicrosoftAuthController(
      publicConfig.value,
      window.location.origin,
    );
    controller.current = auth;

    const restore = async () => {
      let restoredAccount: AccountInfo | null = null;
      try {
        const account = await auth.initialize();
        restoredAccount = account;
        if (!current) return;
        if (!account) {
          setState({ status: "disconnected" });
          return;
        }
        setState({ status: "checking", account });
        const accessToken = await auth.acquireGraphToken(account);
        const diagnostic = await runDevConnectionDiagnostic(
          accessToken,
          publicConfig.value,
        );
        if (current) setState({ status: "connected", account, diagnostic });
      } catch (error) {
        if (!current || error instanceof InteractiveRedirectStartedError) return;
        setState({
          status: "error",
          account: restoredAccount,
          message: safeErrorMessage(error),
        });
      }
    };
    void restore();

    return () => {
      current = false;
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

  if (publicConfig.status === "disabled") {
    return (
      <button className="avatar" aria-label="Account options">
        GM
      </button>
    );
  }

  const account = "account" in state ? state.account : null;
  const microsoftConnected = account !== null;
  const sharePointConnected = state.status === "connected";
  const busy = state.status === "initializing" || state.status === "checking";

  const verifyConnection = async (
    auth: MicrosoftAuthController,
    selectedAccount: AccountInfo,
    config: DevMicrosoftPublicConfig,
  ) => {
    setState({ status: "checking", account: selectedAccount });
    try {
      const accessToken = await auth.acquireGraphToken(selectedAccount);
      const diagnostic = await runDevConnectionDiagnostic(accessToken, config);
      setState({ status: "connected", account: selectedAccount, diagnostic });
    } catch (error) {
      if (error instanceof InteractiveRedirectStartedError) return;
      setState({
        status: "error",
        account: selectedAccount,
        message: safeErrorMessage(error),
      });
    }
  };

  const signIn = async () => {
    if (!controller.current) return;
    try {
      await controller.current.signIn();
    } catch (error) {
      setState({ status: "error", account: null, message: safeErrorMessage(error) });
    }
  };

  const signOut = async () => {
    if (!controller.current) return;
    try {
      await controller.current.signOut(account);
    } catch (error) {
      setState({ status: "error", account, message: safeErrorMessage(error) });
    }
  };

  return (
    <div className="dev-ms-connection" ref={container}>
      <button
        className={`avatar dev-ms-trigger ${sharePointConnected ? "connected" : ""}`}
        aria-label="DEV Microsoft connection"
        aria-expanded={open}
        aria-controls="dev-microsoft-connection-panel"
        onClick={() => setOpen((value) => !value)}
      >
        {initials(account?.name)}
      </button>
      {open && (
        <section
          className="dev-ms-panel"
          id="dev-microsoft-connection-panel"
          aria-label="DEV Microsoft connection status"
        >
          <p className="dev-ms-kicker">DEV connection</p>
          <ConnectionRow label="Microsoft" connected={microsoftConnected} busy={busy} />
          <ConnectionRow
            label="SharePoint DEV"
            connected={sharePointConnected}
            busy={state.status === "checking"}
          />
          <div className="dev-ms-user">
            <span>Signed in as</span>
            <strong>
              {state.status === "connected"
                ? state.diagnostic.user.displayName
                : account?.name ?? account?.username ?? "Not signed in"}
            </strong>
          </div>
          {state.status === "connected" && (
            <div className="dev-ms-detail">
              <span>Graph site ID</span>
              <code>{state.diagnostic.site.id}</code>
              <span>{state.diagnostic.lists.length} existing lists readable</span>
            </div>
          )}
          {state.status === "error" && (
            <p className="dev-ms-error" role="alert">
              {state.message}
            </p>
          )}
          <div className="dev-ms-actions">
            {!account ? (
              <button
                onClick={() => void signIn()}
                disabled={publicConfig.status !== "enabled" || busy}
              >
                Sign in with Microsoft
              </button>
            ) : (
              <>
                <button
                  onClick={() =>
                    publicConfig.status === "enabled" &&
                    controller.current &&
                    void verifyConnection(
                      controller.current,
                      account,
                      publicConfig.value,
                    )
                  }
                  disabled={busy}
                >
                  {busy ? "Checking…" : "Retry check"}
                </button>
                <button className="secondary" onClick={() => void signOut()}>
                  Sign out
                </button>
              </>
            )}
          </div>
        </section>
      )}
    </div>
  );
}

function ConnectionRow({
  label,
  connected,
  busy,
}: {
  label: string;
  connected: boolean;
  busy: boolean;
}) {
  return (
    <div className="dev-ms-row">
      <span>{label}</span>
      <strong className={connected ? "is-connected" : ""}>
        {busy ? "Checking…" : connected ? "Connected" : "Not connected"}
      </strong>
    </div>
  );
}
