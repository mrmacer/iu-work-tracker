import { describe, expect, it, vi } from "vitest";
import {
  BrowserCacheLocation,
  InteractionRequiredAuthError,
  type AccountInfo,
  type AuthenticationResult,
} from "@azure/msal-browser";
import {
  buildDevMicrosoftConfig,
  createMsalConfiguration,
  MICROSOFT_GRAPH_SCOPES,
  type DevMicrosoftPublicConfig,
} from "../lib/microsoft-auth-config";
import {
  InteractiveRedirectStartedError,
  MicrosoftAuthenticationError,
  MicrosoftAuthController,
  type MsalBrowserClient,
} from "../lib/microsoft-auth";

const publicConfig: DevMicrosoftPublicConfig = {
  clientId: "11111111-1111-4111-8111-111111111111",
  tenantId: "3276761c-22db-462b-a930-172d155bd795",
  sharePointHostname: "siu29.sharepoint.com",
  sharePointSitePath: "/sites/IUWorkTrackerDEV",
};

const account = {
  homeAccountId: "home",
  environment: "login.microsoftonline.com",
  tenantId: publicConfig.tenantId,
  username: "dev@example.edu",
  localAccountId: "local",
  name: "Dev User",
} as AccountInfo;

function authResult(overrides: Partial<AuthenticationResult> = {}) {
  return {
    authority: "https://login.microsoftonline.com/tenant",
    uniqueId: "unique",
    tenantId: publicConfig.tenantId,
    scopes: [...MICROSOFT_GRAPH_SCOPES],
    account,
    idToken: "id-token",
    idTokenClaims: {},
    accessToken: "graph-token",
    fromCache: false,
    expiresOn: new Date(Date.now() + 60_000),
    tokenType: "Bearer",
    correlationId: "correlation",
    ...overrides,
  } as AuthenticationResult;
}

function createClient(overrides: Partial<MsalBrowserClient> = {}) {
  const client: MsalBrowserClient = {
    initialize: vi.fn(async () => undefined),
    handleRedirectPromise: vi.fn(async () => null),
    getActiveAccount: vi.fn(() => null),
    getAllAccounts: vi.fn(() => []),
    setActiveAccount: vi.fn(() => undefined),
    loginRedirect: vi.fn(async () => undefined),
    acquireTokenSilent: vi.fn(async () => authResult()),
    acquireTokenRedirect: vi.fn(async () => undefined),
    logoutRedirect: vi.fn(async () => undefined),
    ...overrides,
  };
  return client;
}

describe("DEV Microsoft configuration", () => {
  it("is disabled when no public Microsoft values are present", () => {
    expect(buildDevMicrosoftConfig({})).toEqual({ status: "disabled" });
  });

  it("accepts the complete DEV configuration", () => {
    expect(
      buildDevMicrosoftConfig({
        NEXT_PUBLIC_MS_ENTRA_CLIENT_ID: publicConfig.clientId,
        NEXT_PUBLIC_MS_ENTRA_TENANT_ID: publicConfig.tenantId,
        NEXT_PUBLIC_SHAREPOINT_HOSTNAME: publicConfig.sharePointHostname,
        NEXT_PUBLIC_SHAREPOINT_SITE_PATH: publicConfig.sharePointSitePath,
      }),
    ).toEqual({ status: "enabled", value: publicConfig });
  });

  it("rejects a non-DEV SharePoint path", () => {
    const result = buildDevMicrosoftConfig({
      NEXT_PUBLIC_MS_ENTRA_CLIENT_ID: publicConfig.clientId,
      NEXT_PUBLIC_MS_ENTRA_TENANT_ID: publicConfig.tenantId,
      NEXT_PUBLIC_SHAREPOINT_HOSTNAME: publicConfig.sharePointHostname,
      NEXT_PUBLIC_SHAREPOINT_SITE_PATH: "/sites/IUWorkTracker",
    });
    expect(result.status).toBe("invalid");
  });

  it("uses the tenant-specific authority, SPA redirects, and session storage", () => {
    const configuration = createMsalConfiguration(
      publicConfig,
      "http://localhost:3000",
    );
    expect(configuration.auth.authority).toBe(
      `https://login.microsoftonline.com/${publicConfig.tenantId}`,
    );
    expect(configuration.auth.redirectUri).toBe("http://localhost:3000/");
    expect(configuration.auth.postLogoutRedirectUri).toBe(
      "http://localhost:3000/",
    );
    expect(configuration.cache?.cacheLocation).toBe(
      BrowserCacheLocation.SessionStorage,
    );
  });

  it("requests only the required delegated Graph scopes", () => {
    expect(MICROSOFT_GRAPH_SCOPES).toEqual([
      "User.Read",
      "Sites.ReadWrite.All",
    ]);
  });
});

describe("MicrosoftAuthController", () => {
  it("restores the account returned by the redirect", async () => {
    const client = createClient({
      handleRedirectPromise: vi.fn(async () => authResult()),
    });
    const controller = new MicrosoftAuthController(client, "http://localhost:3000");

    await expect(controller.initialize()).resolves.toBe(account);
    expect(client.setActiveAccount).toHaveBeenCalledWith(account);
  });

  it("restores the sole cached account after reload", async () => {
    const client = createClient({ getAllAccounts: vi.fn(() => [account]) });
    const controller = new MicrosoftAuthController(client, "http://localhost:3000");

    await expect(controller.initialize()).resolves.toBe(account);
    expect(client.setActiveAccount).toHaveBeenCalledWith(account);
  });

  it("does not choose an account implicitly when multiple are cached", async () => {
    const second = { ...account, homeAccountId: "other" };
    const client = createClient({
      getAllAccounts: vi.fn(() => [account, second]),
    });
    const controller = new MicrosoftAuthController(client, "http://localhost:3000");

    await expect(controller.initialize()).resolves.toBeNull();
    expect(client.setActiveAccount).not.toHaveBeenCalled();
  });

  it("acquires a Graph token silently", async () => {
    const client = createClient();
    const controller = new MicrosoftAuthController(client, "http://localhost:3000");

    await expect(controller.acquireGraphToken(account)).resolves.toBe("graph-token");
    expect(client.acquireTokenSilent).toHaveBeenCalledWith({
      account,
      scopes: [...MICROSOFT_GRAPH_SCOPES],
    });
  });

  it("uses redirect fallback only when interaction is required", async () => {
    const client = createClient({
      acquireTokenSilent: vi.fn(async () => {
        throw new InteractionRequiredAuthError(
          "interaction_required",
          "Interaction is required.",
        );
      }),
    });
    const controller = new MicrosoftAuthController(client, "http://localhost:3000");

    await expect(controller.acquireGraphToken(account)).rejects.toBeInstanceOf(
      InteractiveRedirectStartedError,
    );
    expect(client.acquireTokenRedirect).toHaveBeenCalledWith({
      account,
      scopes: [...MICROSOFT_GRAPH_SCOPES],
    });
  });

  it("surfaces a safe authentication error for non-interactive failures", async () => {
    const client = createClient({
      acquireTokenSilent: vi.fn(async () => {
        throw new Error("sensitive upstream detail");
      }),
    });
    const controller = new MicrosoftAuthController(client, "http://localhost:3000");

    await expect(controller.acquireGraphToken(account)).rejects.toEqual(
      new MicrosoftAuthenticationError(
        "A Microsoft access token could not be acquired.",
      ),
    );
  });
});
