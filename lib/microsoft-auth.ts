import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration,
} from "@azure/msal-browser";
import {
  createMsalConfiguration,
  MICROSOFT_GRAPH_SCOPES,
  type DevMicrosoftPublicConfig,
} from "./microsoft-auth-config";

export interface MsalBrowserClient {
  initialize(): Promise<void>;
  handleRedirectPromise(): Promise<AuthenticationResult | null>;
  getActiveAccount(): AccountInfo | null;
  getAllAccounts(): AccountInfo[];
  setActiveAccount(account: AccountInfo | null): void;
  loginRedirect(request: { scopes: string[]; prompt?: string }): Promise<void>;
  acquireTokenSilent(request: {
    account: AccountInfo;
    scopes: string[];
  }): Promise<AuthenticationResult>;
  acquireTokenRedirect(request: {
    account: AccountInfo;
    scopes: string[];
  }): Promise<void>;
  logoutRedirect(request: {
    account?: AccountInfo;
    postLogoutRedirectUri: string;
  }): Promise<void>;
}

export class InteractiveRedirectStartedError extends Error {
  constructor() {
    super("Interactive Microsoft authentication has started.");
    this.name = "InteractiveRedirectStartedError";
  }
}

export class MicrosoftAuthenticationError extends Error {
  constructor(message = "Microsoft authentication could not be completed.") {
    super(message);
    this.name = "MicrosoftAuthenticationError";
  }
}

/**
 * Patch 9A — the cached Microsoft account can no longer renew a Graph token silently (the
 * Entra session or refresh token expired, consent changed, a Conditional Access policy needs
 * the user). This is a normal "sign in again" condition, distinct from an authentication
 * FAILURE: it is raised instead of starting a redirect on its own, so the human chooses when
 * to leave the page (an automatic redirect mid-save would discard unsaved work, and one that
 * fires on every load can loop). The UI offers an explicit "Sign in with Microsoft" action.
 */
export const SIGN_IN_REQUIRED_MESSAGE = "Microsoft sign-in has expired. Sign in again to reconnect SharePoint.";
const NO_ACCOUNT_MESSAGE = "Sign in with Microsoft to continue.";

/**
 * True when a provider's user-facing error message means "the human must sign in again" (an
 * expired session, or no account at all) as opposed to any other failure. Providers carry their
 * errors as message strings, so the Work Record wizard uses this to decide whether to offer the
 * "Sign in again" recovery action; both strings come from the constants above, never retyped.
 */
export function isSignInRequiredMessage(message: string): boolean {
  return message === SIGN_IN_REQUIRED_MESSAGE || message === NO_ACCOUNT_MESSAGE;
}

export class InteractionRequiredError extends MicrosoftAuthenticationError {
  constructor(message = SIGN_IN_REQUIRED_MESSAGE) {
    super(message);
    this.name = "InteractionRequiredError";
  }
}

/**
 * The user-facing text for a Microsoft authentication error, or null when `error` is not one.
 * Shared by every SharePoint-backed provider so an authentication problem is reported the
 * same, useful way everywhere instead of collapsing into a generic "could not be reached".
 */
export function microsoftAuthErrorMessage(error: unknown): string | null {
  if (error instanceof InteractiveRedirectStartedError) {
    return "Microsoft sign-in confirmation is required. Finish signing in, then try again.";
  }
  if (error instanceof MicrosoftAuthenticationError) return error.message;
  return null;
}

function isInteractionInProgress(error: unknown): boolean {
  return error instanceof Error && (error as { errorCode?: string }).errorCode === "interaction_in_progress";
}

export class MicrosoftAuthController {
  private initialization: Promise<AccountInfo | null> | null = null;
  // One interactive redirect at a time. Every sign-in request that arrives while one is being
  // started shares it, so several callers can never start competing redirects.
  private interactive: Promise<void> | null = null;

  constructor(
    private readonly client: MsalBrowserClient,
    private readonly origin: string,
  ) {}

  /**
   * Initializes MSAL once and restores the account (redirect result, else the active account,
   * else the sole cached account). Concurrent callers share one attempt. Patch 9A: a FAILED
   * attempt is forgotten so the next call retries — previously one transient failure left
   * every caller on the same rejected promise until a hard reload.
   */
  initialize(): Promise<AccountInfo | null> {
    if (!this.initialization) {
      const attempt = this.initializeOnce();
      this.initialization = attempt;
      attempt.catch(() => {
        if (this.initialization === attempt) this.initialization = null;
      });
    }
    return this.initialization;
  }

  /**
   * Every account MSAL has cached in this browser. Lets the UI explain the "several accounts,
   * none selected" case that initialize() deliberately resolves to null rather than guessing.
   */
  listCachedAccounts(): AccountInfo[] {
    try {
      return [...this.client.getAllAccounts()];
    } catch {
      return [];
    }
  }

  private async initializeOnce(): Promise<AccountInfo | null> {
    try {
      await this.client.initialize();
      const redirectResult = await this.client.handleRedirectPromise();
      const redirectAccount = redirectResult?.account ?? null;
      if (redirectAccount) {
        this.client.setActiveAccount(redirectAccount);
        return redirectAccount;
      }

      const activeAccount = this.client.getActiveAccount();
      if (activeAccount) return activeAccount;

      const cachedAccounts = this.client.getAllAccounts();
      if (cachedAccounts.length === 1) {
        this.client.setActiveAccount(cachedAccounts[0]);
        return cachedAccounts[0];
      }

      return null;
    } catch {
      throw new MicrosoftAuthenticationError();
    }
  }

  /**
   * The ONE way an interactive Microsoft redirect starts, always from an explicit user action.
   * With a cached `account` (its session expired) it re-authenticates that account; without
   * one it opens the normal account chooser. Concurrent calls share a single in-flight
   * redirect, and MSAL's own "interaction already in progress" lock is treated as "a redirect
   * is already underway" rather than as an error.
   */
  signIn(account?: AccountInfo | null): Promise<void> {
    this.interactive ??= this.startInteractiveSignIn(account ?? null).finally(() => {
      this.interactive = null;
    });
    return this.interactive;
  }

  private async startInteractiveSignIn(account: AccountInfo | null): Promise<void> {
    try {
      if (account) {
        await this.client.acquireTokenRedirect({ account, scopes: [...MICROSOFT_GRAPH_SCOPES] });
      } else {
        await this.client.loginRedirect({ scopes: [...MICROSOFT_GRAPH_SCOPES], prompt: "select_account" });
      }
    } catch (error) {
      if (isInteractionInProgress(error)) return;
      throw new MicrosoftAuthenticationError("Microsoft sign-in could not be started.");
    }
  }

  /**
   * Silent-only. Never starts a redirect: when MSAL says interaction is required this throws
   * InteractionRequiredError and leaves it to the user to sign in (see that class).
   */
  async acquireGraphToken(account?: AccountInfo | null): Promise<string> {
    const selectedAccount = account ?? (await this.initialize());
    if (!selectedAccount) {
      throw new MicrosoftAuthenticationError(NO_ACCOUNT_MESSAGE);
    }

    try {
      const result = await this.client.acquireTokenSilent({
        account: selectedAccount,
        scopes: [...MICROSOFT_GRAPH_SCOPES],
      });
      return result.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) throw new InteractionRequiredError();
      throw new MicrosoftAuthenticationError(
        "A Microsoft access token could not be acquired.",
      );
    }
  }

  async signOut(account?: AccountInfo | null): Promise<void> {
    try {
      await this.client.logoutRedirect({
        account: account ?? this.client.getActiveAccount() ?? undefined,
        postLogoutRedirectUri: `${this.origin}/`,
      });
    } catch {
      throw new MicrosoftAuthenticationError("Microsoft sign-out could not be started.");
    }
  }
}

let browserController:
  | { cacheKey: string; controller: MicrosoftAuthController }
  | undefined;

export function createBrowserMicrosoftAuthController(
  config: DevMicrosoftPublicConfig,
  origin: string,
): MicrosoftAuthController {
  const msalConfig: Configuration = createMsalConfiguration(config, origin);
  const cacheKey = `${config.clientId}|${config.tenantId}|${origin}`;
  if (browserController?.cacheKey !== cacheKey) {
    browserController = {
      cacheKey,
      controller: new MicrosoftAuthController(
        new PublicClientApplication(msalConfig),
        origin,
      ),
    };
  }
  return browserController.controller;
}
