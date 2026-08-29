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

export class MicrosoftAuthController {
  private initialization: Promise<AccountInfo | null> | null = null;

  constructor(
    private readonly client: MsalBrowserClient,
    private readonly origin: string,
  ) {}

  initialize(): Promise<AccountInfo | null> {
    this.initialization ??= this.initializeOnce();
    return this.initialization;
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

  async signIn(): Promise<void> {
    try {
      await this.client.loginRedirect({
        scopes: [...MICROSOFT_GRAPH_SCOPES],
        prompt: "select_account",
      });
    } catch {
      throw new MicrosoftAuthenticationError("Microsoft sign-in could not be started.");
    }
  }

  async acquireGraphToken(account?: AccountInfo | null): Promise<string> {
    const selectedAccount = account ?? (await this.initialize());
    if (!selectedAccount) {
      throw new MicrosoftAuthenticationError("Sign in with Microsoft to continue.");
    }

    try {
      const result = await this.client.acquireTokenSilent({
        account: selectedAccount,
        scopes: [...MICROSOFT_GRAPH_SCOPES],
      });
      return result.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        try {
          await this.client.acquireTokenRedirect({
            account: selectedAccount,
            scopes: [...MICROSOFT_GRAPH_SCOPES],
          });
        } catch {
          throw new MicrosoftAuthenticationError(
            "Microsoft permission confirmation could not be started.",
          );
        }
        throw new InteractiveRedirectStartedError();
      }
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
