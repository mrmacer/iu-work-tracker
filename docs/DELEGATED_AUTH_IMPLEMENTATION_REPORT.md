# IU Work Tracker — Delegated Auth Implementation Report

**Phase:** 2B.2 — DEV delegated Microsoft authentication  
**Date:** August 28, 2026  
**Scope:** Authentication and read-only connection diagnostics only

## Outcome

The minimum DEV-only Microsoft authentication layer is implemented. It uses MSAL Browser's SPA authorization-code flow with PKCE, restores a single cached account after reload, acquires Graph tokens silently when possible, uses redirect interaction when Microsoft requires it, and supports redirect-based sign-in and sign-out.

The complete DEV configuration is now present in the ignored `.env.local`, the configured application builds successfully, and the local server is ready for the live Microsoft sign-in test at `http://localhost:3000/`.

The diagnostic path performs only these reads:

1. Resolve the signed-in user with Microsoft Graph `/me`.
2. Resolve `https://siu29.sharepoint.com/sites/IUWorkTrackerDEV` by hostname and server-relative path.
3. Read the site's existing lists.

No SharePoint lists or list items are created, updated, or deleted. No DataProvider was added or changed. Work Record persistence remains on the existing V1.1 path.

## Files changed

| File | Change |
|---|---|
| `.env.example` | Documents the four public DEV configuration values without supplying or requiring a secret. |
| `.gitignore` | Allows the non-secret `.env.example` template to be tracked while continuing to ignore real environment files. |
| `package.json` | Adds an exact MSAL Browser dependency. |
| `package-lock.json` | Locks MSAL and its dependencies. |
| `lib/microsoft-auth-config.ts` | Validates DEV public configuration, defines the tenant authority and Graph scopes, and creates the MSAL SPA configuration. |
| `lib/microsoft-auth.ts` | Implements redirect handling, account restoration, sign-in, silent token acquisition, interactive redirect fallback, and sign-out. |
| `lib/microsoft-graph.ts` | Implements the three read-only Graph operations and safe Graph error handling. |
| `app/DevMicrosoftConnection.tsx` | Adds the compact, configuration-gated DEV connection diagnostic. |
| `app/IUWorkTracker.tsx` | Places the diagnostic behind the existing account control without changing application data flow. |
| `app/globals.css` | Styles the compact connection popover and status indicators. |
| `tests/microsoft-auth.test.ts` | Adds configuration, account restoration, silent-token, fallback, and authentication-failure coverage. |
| `tests/microsoft-graph.test.ts` | Adds `/me`, site-resolution, list-read, sequencing, authentication-failure, and permission-failure coverage. |

## MSAL package

- Package: `@azure/msal-browser`
- Version: `5.19.0`, exact-pinned
- Application type: SPA/public client
- OAuth flow: authorization code with PKCE, managed by MSAL Browser
- Client secret: none
- Implicit flow: not used

MSAL 5 no longer accepts the legacy `storeAuthStateInCookie` configuration used by older implementations. This implementation uses the current MSAL 5 configuration surface rather than copying obsolete MAC-Walkthrough settings.

## Configuration

The implementation reads these public build/runtime values:

```dotenv
NEXT_PUBLIC_MS_ENTRA_CLIENT_ID=a47025f6-b2c7-4628-9c1f-752970acec6b
NEXT_PUBLIC_MS_ENTRA_TENANT_ID=3276761c-22db-462b-a930-172d155bd795
NEXT_PUBLIC_SHAREPOINT_HOSTNAME=siu29.sharepoint.com
NEXT_PUBLIC_SHAREPOINT_SITE_PATH=/sites/IUWorkTrackerDEV
```

These values were verified in `.env.local` on August 28, 2026. The file remains ignored and is not committed.

The delegated feature is dormant when all four values are absent. Partial or invalid configuration produces a visible DEV diagnostic error. The configured SharePoint path must be exactly `/sites/IUWorkTrackerDEV`; another path is rejected before authentication starts.

For each active origin, the MSAL values are:

| Value | Configuration |
|---|---|
| Client ID | `NEXT_PUBLIC_MS_ENTRA_CLIENT_ID` |
| Authority | `https://login.microsoftonline.com/3276761c-22db-462b-a930-172d155bd795` |
| Redirect URI | `<current origin>/` |
| Post-logout redirect URI | `<current origin>/` |
| Cache | Browser `sessionStorage` through MSAL |

This resolves to the registered URLs `http://localhost:3000/` locally and `https://iu-work-tracker.gmacer.chatgpt.site/` on the owner-only deployed DEV site.

## Delegated scopes

The code requests exactly:

```text
User.Read
Sites.ReadWrite.All
```

`Sites.ReadWrite.All` is Microsoft's canonical delegated Graph permission name. `Sites.ReadWrite` is not a valid Graph scope name. The implementation and its scope regression test both use exactly `Sites.ReadWrite.All`; no scope correction was required in application code.

## Authentication flow

1. MSAL initializes and processes any authorization redirect response.
2. A redirect-returned account becomes active.
3. If there is no redirect result, MSAL restores the active account or the sole cached account.
4. If multiple accounts are cached and none is active, the implementation does not choose one implicitly.
5. The user starts sign-in with `loginRedirect` and Microsoft's account selector.
6. Graph token acquisition first uses `acquireTokenSilent` for the selected account and exact scopes.
7. Only an MSAL `InteractionRequiredAuthError` triggers `acquireTokenRedirect`.
8. Sign-out uses `logoutRedirect` with the active account and current-origin post-logout URI.

MSAL persists authentication state in `sessionStorage`, allowing restoration after reload in the same browser tab/session. MSAL manages cached token renewal; the application does not store, inspect, log, or manually refresh tokens.

## Microsoft Graph calls

All requests use `GET`, `Accept: application/json`, and a bearer access token.

```text
GET https://graph.microsoft.com/v1.0/me?$select=id,displayName,userPrincipalName,mail

GET https://graph.microsoft.com/v1.0/sites/siu29.sharepoint.com:/sites/IUWorkTrackerDEV?$select=id,displayName,webUrl

GET https://graph.microsoft.com/v1.0/sites/{opaque-site-id}/lists?$select=id,name,displayName
```

The returned opaque Graph site ID and existing-list count appear in the diagnostic output. Tokens and raw Graph error bodies do not.

## DEV connection UX

When configuration is enabled, the existing account control opens a small connection popover showing:

- Microsoft: Connected / Not connected / Checking
- SharePoint DEV: Connected / Not connected / Checking
- Signed-in user
- Opaque Graph site ID after success
- Count of readable existing lists
- Sign in, retry, and sign out actions

When Microsoft configuration is absent, the existing account control remains unchanged. This keeps the auth layer from turning the Work Tracker into an authentication dashboard or affecting its normal persistence flow.

## Error handling and security boundaries

- MSAL initialization, sign-in, token, and sign-out failures receive safe application messages.
- Graph `401` responses are identified as authentication failures.
- Graph `403` responses are identified as account/resource authorization failures.
- Raw Graph bodies and access tokens are never displayed or logged.
- A single cached account may be restored; multiple cached accounts require explicit user selection.
- Only the exact DEV SharePoint site path is accepted by configuration validation.
- No client secret exists or is required.
- No implicit-flow configuration or code is present.
- No SharePoint write endpoint is implemented in this phase.
- The existing DataProvider and optimistic-concurrency architecture are untouched.

## Tests

The suite now has **42 passing tests and 0 failures**, including the existing **26 V1.1 tests**.

New coverage verifies:

- complete, missing, and non-DEV configuration behavior;
- tenant-specific authority and SPA redirect values;
- exact delegated scopes;
- redirect account restoration;
- sole cached-account restoration after reload;
- safe behavior with multiple cached accounts;
- silent token acquisition;
- interaction-required redirect fallback;
- non-interactive authentication failure;
- Graph `/me` mapping;
- SharePoint site resolution and opaque site ID handling;
- existing-list reads using `GET`;
- complete diagnostic call order;
- Graph `401` authentication failure;
- Graph `403` permission failure.

## Validation status

| Check | Result |
|---|---|
| Production build | PASS |
| TypeScript type-check | PASS |
| ESLint | PASS |
| Vitest | PASS — 42/42 |
| Rendered HTML smoke test | PASS — 1/1 |
| Production dependency audit | PASS — 0 production vulnerabilities |
| Local DEV configuration | PASS — all four required values loaded from `.env.local` |
| Local server | PASS — `http://localhost:3000/` returned HTTP 200 |
| Live Entra sign-in | READY FOR USER TEST |
| Live Graph `/me` | NOT RUN — requires live Entra sign-in |
| Live SharePoint site resolution | NOT RUN — requires live Entra sign-in |
| Live existing-list read | NOT RUN — requires live Entra sign-in |

## Local DEV authentication test

The configured development server is running at:

```text
http://localhost:3000/
```

This exactly matches the registered SPA redirect URI. Open the account control, choose **Sign in with Microsoft**, and use the intended IU29 account. A complete successful diagnostic will show Microsoft and SharePoint DEV as connected, the signed-in user, the opaque Graph site ID, and the number of readable existing lists.

## Known limitations and remaining readiness conditions

1. Tenant sign-in, `/me`, site resolution, and list-read success cannot be claimed until the live user test completes with a token issued to client `a47025f6-b2c7-4628-9c1f-752970acec6b`.
2. The current hosted DEV environment was not changed in this local-test step. Hosted configuration and deployment remain separate follow-up actions.
3. The diagnostic deliberately does not prove write access because this phase prohibits SharePoint modification. It proves authentication and read access only.
4. Full `npm audit` reports advisories in development tooling dependencies; `npm audit --omit=dev` reports zero production dependency vulnerabilities. No automated dependency remediation was run because it is outside this phase.

The local application is ready for the user to perform the live Microsoft sign-in. To reach GREEN, complete the three-call diagnostic with the intended IU29 account and confirm `/me`, site resolution, and existing-list retrieval all succeed.

## DEV AUTH READINESS:

**YELLOW — READY FOR LIVE SIGN-IN VALIDATION**
