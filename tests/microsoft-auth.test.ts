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
  InteractionRequiredError,
  InteractiveRedirectStartedError,
  MicrosoftAuthenticationError,
  MicrosoftAuthController,
  microsoftAuthErrorMessage,
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

  it("uses the tenant-specific authority, SPA redirects, and MSAL's supported localStorage cache", () => {
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
    // Patch 9A: persistent across refresh / new tab / tab close (MSAL encrypts the cache with a
    // session-cookie key, so it still ends with the browser session). Never sessionStorage or
    // memory, and no custom cache plugin.
    expect(configuration.cache?.cacheLocation).toBe(BrowserCacheLocation.LocalStorage);
    expect(configuration.cache?.cacheLocation).toBe("localStorage");
    expect(Object.keys(configuration.cache ?? {})).toEqual(["cacheLocation"]);
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

  it("never starts a redirect on its own: interaction-required surfaces as InteractionRequiredError for the user to act on", async () => {
    const client = createClient({
      acquireTokenSilent: vi.fn(async () => {
        throw new InteractionRequiredAuthError(
          "interaction_required",
          "Interaction is required.",
        );
      }),
    });
    const controller = new MicrosoftAuthController(client, "http://localhost:3000");

    const failure = await controller.acquireGraphToken(account).catch((error) => error);
    expect(failure).toBeInstanceOf(InteractionRequiredError);
    expect(failure).toBeInstanceOf(MicrosoftAuthenticationError);
    expect(failure.message).toMatch(/sign-in has expired/i);
    expect(client.acquireTokenRedirect).not.toHaveBeenCalled();
    expect(client.loginRedirect).not.toHaveBeenCalled();
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


// ---------------------------------------------------------------------------------------------
// Patch 9A — persistent sign-in and session restoration.
// ---------------------------------------------------------------------------------------------

const ORIGIN = "http://localhost:3000";

/** A stand-in for MSAL's persistent cache shared by every tab: accounts + the active account. */
function sharedCache(initial: AccountInfo[] = [], active: AccountInfo | null = null) {
  const cache = { accounts: [...initial], active };
  const clientFor = (overrides: Partial<MsalBrowserClient> = {}) =>
    createClient({
      getAllAccounts: vi.fn(() => cache.accounts),
      getActiveAccount: vi.fn(() => cache.active),
      setActiveAccount: vi.fn((next: AccountInfo | null) => {
        cache.active = next;
      }),
      ...overrides,
    });
  return { cache, clientFor };
}

const interactionRequired = () =>
  new InteractionRequiredAuthError("interaction_required", "Interaction is required.");

describe("returning user, refresh, and new tab (persistent cache)", () => {
  it("a first-time user with nothing cached is signed out — no redirect, no error", async () => {
    const { clientFor } = sharedCache();
    const client = clientFor();
    const controller = new MicrosoftAuthController(client, ORIGIN);
    await expect(controller.initialize()).resolves.toBeNull();
    expect(controller.listCachedAccounts()).toEqual([]);
    expect(client.loginRedirect).not.toHaveBeenCalled();
    expect(client.acquireTokenRedirect).not.toHaveBeenCalled();
  });

  it("a returning user with valid cached authentication is restored and gets a token silently", async () => {
    const { clientFor } = sharedCache([account], account);
    const client = clientFor();
    const controller = new MicrosoftAuthController(client, ORIGIN);
    await expect(controller.initialize()).resolves.toBe(account);
    await expect(controller.acquireGraphToken(account)).resolves.toBe("graph-token");
    expect(client.loginRedirect).not.toHaveBeenCalled();
    expect(client.acquireTokenRedirect).not.toHaveBeenCalled();
  });

  it("a browser refresh (a brand-new controller over the same cache) restores the same account without interaction", async () => {
    const { clientFor } = sharedCache([account], account);
    await new MicrosoftAuthController(clientFor(), ORIGIN).initialize();
    const afterRefresh = clientFor();
    const reloaded = new MicrosoftAuthController(afterRefresh, ORIGIN);
    await expect(reloaded.initialize()).resolves.toBe(account);
    await expect(reloaded.acquireGraphToken(account)).resolves.toBe("graph-token");
    expect(afterRefresh.loginRedirect).not.toHaveBeenCalled();
  });

  it("a NEW TAB sees the account the first tab signed in (shared persistent cache) — the sessionStorage-era failure", async () => {
    const { cache, clientFor } = sharedCache();
    const tabA = new MicrosoftAuthController(
      clientFor({ handleRedirectPromise: vi.fn(async () => authResult()) }),
      ORIGIN,
    );
    await expect(tabA.initialize()).resolves.toBe(account); // signs in → active account stored in the shared cache
    cache.accounts.push(account);
    const tabBClient = clientFor();
    const tabB = new MicrosoftAuthController(tabBClient, ORIGIN);
    await expect(tabB.initialize()).resolves.toBe(account);
    await expect(tabB.acquireGraphToken(account)).resolves.toBe("graph-token");
    expect(tabBClient.loginRedirect).not.toHaveBeenCalled();
  });

  it("an expired access token is renewed silently and the new token is used", async () => {
    const { clientFor } = sharedCache([account], account);
    const acquireTokenSilent = vi
      .fn()
      .mockResolvedValueOnce(authResult({ accessToken: "first-token" }))
      .mockResolvedValueOnce(authResult({ accessToken: "renewed-token", fromCache: false }));
    const controller = new MicrosoftAuthController(clientFor({ acquireTokenSilent }), ORIGIN);
    await expect(controller.acquireGraphToken(account)).resolves.toBe("first-token");
    await expect(controller.acquireGraphToken(account)).resolves.toBe("renewed-token"); // after the first token expired
    expect(acquireTokenSilent).toHaveBeenCalledTimes(2);
  });

  it("a failed silent renewal that is NOT 'interaction required' is an authentication failure with a safe message", async () => {
    const { clientFor } = sharedCache([account], account);
    const client = clientFor({ acquireTokenSilent: vi.fn(async () => { throw new Error("network down: secret-detail"); }) });
    const controller = new MicrosoftAuthController(client, ORIGIN);
    const failure = await controller.acquireGraphToken(account).catch((error) => error);
    expect(failure).toBeInstanceOf(MicrosoftAuthenticationError);
    expect(failure).not.toBeInstanceOf(InteractionRequiredError);
    expect(failure.message).toBe("A Microsoft access token could not be acquired.");
    expect(failure.message).not.toMatch(/secret-detail/);
    expect(client.acquireTokenRedirect).not.toHaveBeenCalled();
  });
});

describe("initialization failure and retry", () => {
  it("a failed initialization is retried by the next call instead of staying failed forever", async () => {
    const { clientFor } = sharedCache([account], account);
    const initialize = vi.fn().mockRejectedValueOnce(new Error("transient")).mockResolvedValue(undefined);
    const controller = new MicrosoftAuthController(clientFor({ initialize }), ORIGIN);
    await expect(controller.initialize()).rejects.toBeInstanceOf(MicrosoftAuthenticationError);
    await expect(controller.initialize()).resolves.toBe(account);
    expect(initialize).toHaveBeenCalledTimes(2);
  });

  it("concurrent callers share ONE attempt, and a success is cached (no re-initialization)", async () => {
    const { clientFor } = sharedCache([account], account);
    const client = clientFor();
    const controller = new MicrosoftAuthController(client, ORIGIN);
    const results = await Promise.all([controller.initialize(), controller.initialize(), controller.initialize()]);
    expect(results).toEqual([account, account, account]);
    await controller.initialize();
    expect(client.initialize).toHaveBeenCalledTimes(1);
    expect(client.handleRedirectPromise).toHaveBeenCalledTimes(1);
  });

  it("concurrent callers of a FAILING attempt all see the failure, then a later call can succeed", async () => {
    const { clientFor } = sharedCache([account], account);
    const initialize = vi.fn().mockRejectedValueOnce(new Error("boom")).mockResolvedValue(undefined);
    const controller = new MicrosoftAuthController(clientFor({ initialize }), ORIGIN);
    const settled = await Promise.allSettled([controller.initialize(), controller.initialize()]);
    expect(settled.map((entry) => entry.status)).toEqual(["rejected", "rejected"]);
    expect(initialize).toHaveBeenCalledTimes(1);
    await expect(controller.initialize()).resolves.toBe(account);
  });
});

describe("multiple cached accounts", () => {
  const second = { ...account, homeAccountId: "other", username: "other@example.edu", name: "Other User" } as AccountInfo;

  it("never picks one silently, and exposes the cached accounts so the UI can ask", async () => {
    const { clientFor } = sharedCache([account, second], null);
    const client = clientFor();
    const controller = new MicrosoftAuthController(client, ORIGIN);
    await expect(controller.initialize()).resolves.toBeNull();
    expect(client.setActiveAccount).not.toHaveBeenCalled();
    expect(controller.listCachedAccounts().map((entry) => entry.username)).toEqual(["dev@example.edu", "other@example.edu"]);
  });

  it("honors an account the human already chose (the active account), even with several cached", async () => {
    const { clientFor } = sharedCache([account, second], second);
    const controller = new MicrosoftAuthController(clientFor(), ORIGIN);
    await expect(controller.initialize()).resolves.toBe(second);
  });

  it("listCachedAccounts is safe if MSAL throws", () => {
    const controller = new MicrosoftAuthController(
      createClient({ getAllAccounts: vi.fn(() => { throw new Error("storage blocked"); }) }),
      ORIGIN,
    );
    expect(controller.listCachedAccounts()).toEqual([]);
  });
});

describe("interactive sign-in is explicit, single-flight, and never loops", () => {
  it("concurrent token requests that need interaction ALL fail with InteractionRequiredError and start ZERO redirects", async () => {
    const { clientFor } = sharedCache([account], account);
    const client = clientFor({ acquireTokenSilent: vi.fn(async () => { throw interactionRequired(); }) });
    const controller = new MicrosoftAuthController(client, ORIGIN);
    const settled = await Promise.allSettled(Array.from({ length: 7 }, () => controller.acquireGraphToken(account)));
    for (const entry of settled) {
      expect(entry.status).toBe("rejected");
      expect((entry as PromiseRejectedResult).reason).toBeInstanceOf(InteractionRequiredError);
    }
    expect(client.acquireTokenRedirect).not.toHaveBeenCalled();
    expect(client.loginRedirect).not.toHaveBeenCalled();
  });

  it("repeated load cycles with an expired session never start a redirect (no redirect loop)", async () => {
    const { clientFor } = sharedCache([account], account);
    const client = clientFor({ acquireTokenSilent: vi.fn(async () => { throw interactionRequired(); }) });
    for (let load = 0; load < 5; load++) {
      const controller = new MicrosoftAuthController(client, ORIGIN); // a fresh page load each time
      await controller.initialize();
      await controller.acquireGraphToken(account).catch(() => undefined);
    }
    expect(client.acquireTokenRedirect).not.toHaveBeenCalled();
    expect(client.loginRedirect).not.toHaveBeenCalled();
  });

  it("concurrent signIn() calls share ONE redirect (signed-out → account chooser)", async () => {
    const { clientFor } = sharedCache();
    const loginRedirect = vi.fn(async () => { await new Promise((resolve) => setTimeout(resolve, 5)); });
    const client = clientFor({ loginRedirect });
    const controller = new MicrosoftAuthController(client, ORIGIN);
    await Promise.all([controller.signIn(), controller.signIn(), controller.signIn()]);
    expect(loginRedirect).toHaveBeenCalledTimes(1);
    expect(loginRedirect).toHaveBeenCalledWith({ scopes: [...MICROSOFT_GRAPH_SCOPES], prompt: "select_account" });
    expect(client.acquireTokenRedirect).not.toHaveBeenCalled();
  });

  it("concurrent signIn(account) calls share ONE token redirect for an expired session", async () => {
    const { clientFor } = sharedCache([account], account);
    const acquireTokenRedirect = vi.fn(async () => { await new Promise((resolve) => setTimeout(resolve, 5)); });
    const client = clientFor({ acquireTokenRedirect });
    const controller = new MicrosoftAuthController(client, ORIGIN);
    await Promise.all([controller.signIn(account), controller.signIn(account)]);
    expect(acquireTokenRedirect).toHaveBeenCalledTimes(1);
    expect(acquireTokenRedirect).toHaveBeenCalledWith({ account, scopes: [...MICROSOFT_GRAPH_SCOPES] });
    expect(client.loginRedirect).not.toHaveBeenCalled();
  });

  it("a later, separate sign-in click is allowed once the first has settled", async () => {
    const { clientFor } = sharedCache();
    const client = clientFor();
    const controller = new MicrosoftAuthController(client, ORIGIN);
    await controller.signIn();
    await controller.signIn();
    expect(client.loginRedirect).toHaveBeenCalledTimes(2);
  });

  it("MSAL's own 'interaction already in progress' lock means a redirect is underway — not an error", async () => {
    const { clientFor } = sharedCache();
    const lock = Object.assign(new Error("interaction_in_progress"), { errorCode: "interaction_in_progress" });
    const controller = new MicrosoftAuthController(clientFor({ loginRedirect: vi.fn(async () => { throw lock; }) }), ORIGIN);
    await expect(controller.signIn()).resolves.toBeUndefined();
  });

  it("any other failure to start sign-in is a safe MicrosoftAuthenticationError", async () => {
    const { clientFor } = sharedCache();
    const controller = new MicrosoftAuthController(
      clientFor({ loginRedirect: vi.fn(async () => { throw new Error("popup blocked: detail"); }) }),
      ORIGIN,
    );
    await expect(controller.signIn()).rejects.toEqual(new MicrosoftAuthenticationError("Microsoft sign-in could not be started."));
  });

  it("redirect calls exist in exactly one place: the explicit signIn path", async () => {
    const { readFileSync } = await import("node:fs");
    const code = readFileSync("lib/microsoft-auth.ts", "utf-8").replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
    expect(code.match(/this\.client\.acquireTokenRedirect\(/g)).toHaveLength(1);
    expect(code.match(/this\.client\.loginRedirect\(/g)).toHaveLength(1);
    const acquire = code.slice(code.indexOf("async acquireGraphToken"), code.indexOf("async signOut"));
    expect(acquire).not.toMatch(/Redirect\(/);
  });
});

describe("explicit sign-out", () => {
  it("signs out the given account through logoutRedirect with the current-origin return URL", async () => {
    const { clientFor } = sharedCache([account], account);
    const client = clientFor();
    await new MicrosoftAuthController(client, ORIGIN).signOut(account);
    expect(client.logoutRedirect).toHaveBeenCalledWith({ account, postLogoutRedirectUri: "http://localhost:3000/" });
  });

  it("a sign-out that cannot start is a safe error, not a silent success", async () => {
    const { clientFor } = sharedCache([account], account);
    const controller = new MicrosoftAuthController(clientFor({ logoutRedirect: vi.fn(async () => { throw new Error("x"); }) }), ORIGIN);
    await expect(controller.signOut(account)).rejects.toEqual(new MicrosoftAuthenticationError("Microsoft sign-out could not be started."));
  });
});

describe("microsoftAuthErrorMessage — one shared mapping for every provider", () => {
  it("maps authentication errors to their own message and everything else to null", () => {
    expect(microsoftAuthErrorMessage(new InteractionRequiredError())).toMatch(/sign-in has expired/i);
    expect(microsoftAuthErrorMessage(new MicrosoftAuthenticationError("A Microsoft access token could not be acquired."))).toBe(
      "A Microsoft access token could not be acquired.",
    );
    expect(microsoftAuthErrorMessage(new InteractiveRedirectStartedError())).toBe(
      "Microsoft sign-in confirmation is required. Finish signing in, then try again.",
    );
    expect(microsoftAuthErrorMessage(new Error("unrelated"))).toBeNull();
    expect(microsoftAuthErrorMessage("string")).toBeNull();
  });
});

describe("no custom token storage", () => {
  it("auth code never touches web storage, cookies, or MSAL's cache keys directly — MSAL alone owns the cache", async () => {
    const { readFileSync } = await import("node:fs");
    for (const path of ["lib/microsoft-auth.ts", "lib/microsoft-session.ts", "app/DevMicrosoftConnection.tsx"]) {
      const code = readFileSync(path, "utf-8").replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
      expect(code, path).not.toMatch(/localStorage|sessionStorage|document\.cookie|indexedDB|msal\.\d|setItem\(|getItem\(/);
    }
  });
});
