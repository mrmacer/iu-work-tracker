# IU Work Tracker DEV — Delegated Authentication Setup

**Phase:** 2B.1 — configuration plan only

**Status:** No application code, Microsoft Entra configuration, SharePoint resources, or persistence were changed by this document.

**Source of truth:** `docs/MAC_WALKTHROUGH_AUTH_REFERENCE.md` and the inspected working MAC/IEP Skook MSAL implementation.

**Target registration name:** `IU Work Tracker DEV`

**Environment limit:** This plan is for delegated Microsoft Graph authentication in DEV only. The production target remains server-side application authentication with Microsoft Graph `Sites.Selected`.

## Confirmed identifiers and endpoints

| Value | Exact setting |
| --- | --- |
| Microsoft tenant | IU29 single tenant |
| Directory (tenant) ID | `3276761c-22db-462b-a930-172d155bd795` |
| Authority | `https://login.microsoftonline.com/3276761c-22db-462b-a930-172d155bd795` |
| Graph API base | `https://graph.microsoft.com/v1.0` |
| SharePoint hostname | `siu29.sharepoint.com` |
| DEV site path | `/sites/IUWorkTrackerDEV` |
| DEV site URL | `https://siu29.sharepoint.com/sites/IUWorkTrackerDEV` |
| Current local origin | `http://localhost:3000` |
| Current deployed DEV origin | `https://iu-work-tracker.gmacer.chatgpt.site` |
| Client/application ID | Assigned by Entra when `IU Work Tracker DEV` is registered; not known yet |

Do not reuse the MAC/IEP Skook application client ID for a separately registered `IU Work Tracker DEV` application. The working registration supplies the architecture and tenant ID, not the new application's identity.

The Sites project ID in `.openai/hosting.json` and the hosting platform's own authentication client ID are also not Microsoft Entra application IDs and must never be placed in MSAL configuration.

## 1. Required Microsoft Graph delegated permissions

Configure only **delegated** Microsoft Graph permissions:

| Permission/scope | Required | Purpose |
| --- | --- | --- |
| `User.Read` | Yes for working-pattern parity and the `/me` verification call | Sign in and read the signed-in user's basic profile |
| `Sites.ReadWrite.All` | Yes | Resolve the DEV SharePoint site and read/create/update/delete list items on behalf of the signed-in user |

The exact working code requests `Sites.ReadWrite.All`. There is no Microsoft Graph delegated permission named simply `Sites.ReadWrite`; do not enter that shorthand in configuration or code.

Do **not** add these permissions for Phase 2B.1:

- Any Microsoft Graph application permission
- `Sites.Selected`
- `Sites.Manage.All`
- `Sites.FullControl.All`
- `Directory.Read.All`
- `User.Read.All`
- Any Files, Groups, Mail, or offline application permission

`Sites.ReadWrite.All` can write items in existing lists. Creating the seven lists is a separate provisioning action and is not part of authentication setup. Microsoft documents delegated `Sites.Manage.All`, not `Sites.ReadWrite.All`, as the permission for creating SharePoint lists.

Microsoft's current permission definitions are documented in the [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference).

## 2. Consent requirements

Adding a delegated permission to the app registration declares what the app may request; it does not itself grant consent.

| Permission | Microsoft `AdminConsentRequired` metadata | Consent behavior |
| --- | --- | --- |
| `User.Read` delegated | No | A user may consent if IU29's user-consent policy permits it; an administrator may instead grant tenant-wide consent |
| `Sites.ReadWrite.All` delegated | No | A user may consent if IU29's user-consent policy permits it; the tenant may still require an administrator through its consent policy |

The OpenID Connect protocol scopes used by MSAL are:

- `openid`
- `profile`
- `offline_access`

MSAL adds those standard scopes to its requests. They do not need separate application code entries in the Graph scope array. `offline_access` is what permits MSAL to receive the artifacts it needs for silent renewal.

The inspected working app registration also exposes `email`, but the working code does not explicitly request or use that scope. It uses `account.username` and can use `/me` under `User.Read`. Therefore `email` is **not required** for this baseline. Add it later only if approved implementation code actually consumes the ID-token email claim.

## 3. Is administrator consent required?

**Microsoft does not mark either required delegated permission as inherently admin-consent-only.** However, the repository cannot reveal IU29's current user-consent policy, permission classification, or approval workflow. Those tenant settings can require an administrator even when a permission's Microsoft metadata says `AdminConsentRequired: No`.

For this DEV registration:

1. Plan for tenant review because `Sites.ReadWrite.All` is broad.
2. If IU29 allows user consent for this permission, the assigned DEV user can consent at first sign-in.
3. If the sign-in displays **Need admin approval**, stop and use IU29's admin-consent request process.
4. A tenant administrator may use **Grant admin consent for IU29** to avoid per-user prompts, but this is an institutional decision, not a technical requirement established by the inspected implementation.

Do not click tenant-wide admin consent merely because this document exists. The person granting it must have an appropriate Microsoft Entra administrator role and IU authorization to approve the broad delegated permission. See [Microsoft's tenant-wide admin consent guidance](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/grant-admin-consent).

## 4. Exact redirect URI platform type

Use **Single-page application (SPA)** for every redirect URI in this DEV design.

Do not configure the same URI under:

- Web
- Mobile and desktop applications

Do not enable implicit grant access tokens or ID tokens. MSAL Browser uses the authorization code flow with PKCE for the SPA platform. Do not enable **Allow public client flows** under Advanced settings; that portal switch is for native/mobile/desktop public-client flows, not this browser SPA registration.

Microsoft's portal sequence is **Authentication → Add a platform → Single-page application**. See [Microsoft's app-registration guidance](https://learn.microsoft.com/en-us/graph/toolkit/get-started/add-aad-app-registration) and [authorization-code-with-PKCE guidance](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow).

## 5. Exact redirect URIs

Register these root URIs under the SPA platform:

| Environment | Exact SPA redirect URI | Status |
| --- | --- | --- |
| Local development | `http://localhost:3000/` | Required |
| Current owner-only deployed DEV | `https://iu-work-tracker.gmacer.chatgpt.site/` | Register for deployed DEV verification |

The local URI is based on the current project's `http://localhost:3000` development origin. Development must keep port `3000`; if the dev server runs on another port, `window.location.origin` changes and authentication will fail until the actual origin is registered or the server is returned to port `3000`.

Sites project metadata confirms that the current live deployment is `https://iu-work-tracker.gmacer.chatgpt.site` and its access policy currently admits only the owner. The URI is a valid HTTPS SPA redirect format.

The inspected MAC application did not run behind the Sites owner-only access gate. It therefore cannot prove that an Entra redirect response survives that additional hosting access flow. Register the deployed URI, but treat successful `handleRedirectPromise()` completion on the deployed Site as a required verification gate. If the hosting gate interrupts or loses the Entra response, use local DEV only until a supported hosted callback path is confirmed; do not invent a different callback URL.

Use root origins because the working pattern sets `redirectUri: window.location.origin`. Do not register `/auth/callback`, `/api/auth/callback`, or another path unless the later implementation explicitly changes that pattern.

Microsoft requires redirect URIs to be registered and to match the requested URI. Hosted URIs must use HTTPS; localhost is the development HTTP exception. See [redirect URI requirements](https://learn.microsoft.com/en-us/entra/identity-platform/reply-url).

## 6. Logout and post-logout requirements

The inspected working implementation calls `logoutRedirect({ account })` and does not configure a front-channel logout URL or explicit post-logout URI.

For IU Work Tracker DEV, configure MSAL with:

```text
postLogoutRedirectUri: window.location.origin
```

This returns the browser to the same environment root after Microsoft logout. Both possible roots are already registered as SPA redirect URIs:

- `http://localhost:3000/`
- `https://iu-work-tracker.gmacer.chatgpt.site/`

No separate Entra **Front-channel logout URL** is required for this phase. Leave it blank. Front-channel logout is for cross-application single sign-out and needs an application route that can process the notification; the inspected app provides no such route.

Sign-out navigation must be allowed to complete. Interrupting it can clear the local cache without completing the Microsoft server session. See [MSAL Browser sign-out guidance](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/logout).

## 7. MSAL configuration values

The later implementation should instantiate one browser `PublicClientApplication` with this configuration shape:

```ts
const msalConfig = {
  auth: {
    clientId: "<IU Work Tracker DEV Application (client) ID>",
    authority: "https://login.microsoftonline.com/3276761c-22db-462b-a930-172d155bd795",
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};
```

The scope request is separate:

```ts
const graphScopes = ["User.Read", "Sites.ReadWrite.All"];
```

This is a configuration specification, not implementation code. At implementation time, use a pinned current supported `@azure/msal-browser` package from the project package manager. Do not copy MAC's floating `@azure/msal-browser@2` CDN URL.

## 8. Tenant ID use

Use the exact IU29 directory tenant ID:

```text
3276761c-22db-462b-a930-172d155bd795
```

It fixes sign-in to the IU29 tenant. Do not use `common`, `organizations`, `consumers`, or a personal Microsoft-account authority. Configure the app registration for **Accounts in this organizational directory only**.

The tenant ID is an identifier, not a secret, and is safe in client-side configuration.

## 9. Client ID use

After creating `IU Work Tracker DEV`, copy its **Application (client) ID** exactly into DEV client configuration.

Do not use:

- The **Object ID** of the app registration
- The enterprise application's service-principal object ID
- The Sites `.openai/hosting.json` project ID
- The Sites hosting authentication client ID
- The existing `IEP Skook` client ID

The client ID is public application metadata and is safe in browser code. It does not authenticate the application by itself.

## 10. Authority URL

Use exactly:

```text
https://login.microsoftonline.com/3276761c-22db-462b-a930-172d155bd795
```

No trailing OAuth endpoint path is needed in MSAL configuration. MSAL performs OpenID Connect discovery and selects the appropriate authorization/token endpoints.

## 11. Required scopes

The application-request scope array is exactly:

```text
User.Read
Sites.ReadWrite.All
```

Use the same array for initial `loginRedirect` and Graph `acquireTokenSilent` calls so consent and cache behavior remain predictable.

Do not put `https://graph.microsoft.com/.default` in this delegated SPA scope array. `.default` is not the working incremental delegated pattern and is commonly associated with preconfigured/static consent and application flows.

## 12. Token acquisition pattern

Use the working redirect model, with the safety corrections identified in the reference report:

1. Initialize `PublicClientApplication` once in browser code.
2. Call `handleRedirectPromise()` once before protected data loading.
3. If the redirect returns an account, make it the active account.
4. Otherwise, restore a deliberately selected cached account; do not blindly choose `getAllAccounts()[0]` when multiple accounts exist.
5. If no account exists, show **Sign in with Microsoft**.
6. Sign-in calls `loginRedirect({ scopes: graphScopes })`.
7. Every Graph operation requests a token through `acquireTokenSilent({ scopes: graphScopes, account })`.
8. Graph receives the returned `accessToken` only in its `Authorization: Bearer` header.

The Graph access token must not be sent to `/api/records` and treated as an IU Work Tracker API session. Its audience is Microsoft Graph.

## 13. Token cache and storage behavior

Use `sessionStorage` for IU Work Tracker DEV.

This intentionally hardens MAC's `localStorage` choice:

- Authentication state survives page reloads in the same browser tab/session.
- It does not intentionally persist across browser restarts.
- It avoids MAC's cross-session local-storage persistence.
- It still requires strict XSS prevention because same-origin script can invoke MSAL even when storage persistence is shorter.

Set `storeAuthStateInCookie: false`. Do not create a separate custom token store. Do not copy tokens into application state, logs, D1, SharePoint, localStorage, cookies, or environment variables.

See [MSAL Browser caching options](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/caching).

## 14. Silent token refresh

The application does not handle refresh tokens directly. For every Graph request, call `acquireTokenSilent()`:

1. MSAL checks its cache for the requested account, authority, resource, and scopes.
2. If the access token is missing or nearing expiry, MSAL attempts silent renewal using its cached artifacts and the user's Entra session.
3. The application receives only the resulting access token.

Do not implement a refresh timer, store a refresh token, or call the token endpoint manually. See [MSAL token acquisition guidance](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/acquire-token).

## 15. Interactive fallback

If `acquireTokenSilent()` throws `InteractionRequiredAuthError`, call:

```text
acquireTokenRedirect({ scopes, account })
```

Then stop the pending provider operation. Do not continue a `fetch` with a null or missing token. This corrects the inspected MAC source's `Bearer null` fall-through risk.

Do not redirect for every silent failure indiscriminately. Non-interaction failures should become safe authentication/network errors, not redirect loops.

## 16. Sign-in and sign-out flow

### Sign-in

1. User opens local or owner-only deployed DEV.
2. The app initializes MSAL and completes any pending redirect.
3. If no account is active, the user selects **Sign in with Microsoft**.
4. `loginRedirect` sends the browser to the IU29 tenant with the two Graph scopes.
5. The user authenticates and completes any permitted consent prompt.
6. Entra returns to the exact current origin.
7. `handleRedirectPromise` restores the account.
8. The app verifies `/me`, then resolves the DEV SharePoint site.

### Sign-out

1. The user selects **Sign out**.
2. The app calls `logoutRedirect` for the active account with `postLogoutRedirectUri: window.location.origin`.
3. MSAL clears local account/token state and completes Microsoft logout navigation.
4. The browser returns to the same DEV origin in a signed-out state.

The application must distinguish a Microsoft-authenticated account from authorization to use IU Work Tracker. DEV admission should use enterprise-application assignment to the approved IU Work Tracker DEV security group where practical.

## 17. Client secret requirement

No client secret is required or allowed for this SPA delegated approach.

Leave **Certificates & secrets** empty. A browser cannot keep a secret, and adding one would not make the SPA a confidential client. No certificate, federated credential, managed identity, or application credential is used in Phase 2B.1.

## 18. Environment variables and configuration values

The later implementation should use these public DEV values:

| Proposed key | Value | Timing |
| --- | --- | --- |
| `NEXT_PUBLIC_MS_ENTRA_TENANT_ID` | `3276761c-22db-462b-a930-172d155bd795` | Known now |
| `NEXT_PUBLIC_MS_ENTRA_CLIENT_ID` | New `IU Work Tracker DEV` Application (client) ID | Record after registration |
| `NEXT_PUBLIC_SHAREPOINT_HOSTNAME` | `siu29.sharepoint.com` | Known now |
| `NEXT_PUBLIC_SHAREPOINT_SITE_PATH` | `/sites/IUWorkTrackerDEV` | Known now |
| `NEXT_PUBLIC_SHAREPOINT_SITE_ID` | Opaque Graph site ID returned by site resolution | Record after authorized verification |

Keep these as code constants rather than mutable environment variables:

| Constant | Exact value |
| --- | --- |
| Graph base URL | `https://graph.microsoft.com/v1.0` |
| Authority construction | `https://login.microsoftonline.com/{tenantId}` |
| Graph scopes | `User.Read`, `Sites.ReadWrite.All` |
| Redirect URI | `window.location.origin` |
| Post-logout URI | `window.location.origin` |
| Cache | `sessionStorage` |

Vinext supports browser-exposed `NEXT_PUBLIC_*` values. Anything with that prefix is included in client output and must be treated as public. Redirect origins are derived rather than configured so one build can use the exact current local or deployed origin.

The seven list GUID variables are not part of authentication setup. They are recorded only after list provisioning.

## 19. Values safe to expose client-side

These values are public identifiers/configuration and may be included in browser output:

- Entra tenant ID
- Entra application/client ID
- Tenant-specific authority URL
- Redirect and post-logout origins
- Graph base URL
- Requested scope names
- SharePoint hostname and site path
- Graph site ID
- SharePoint list GUIDs after provisioning
- Non-secret provider/environment selector names

Making them hosted environment variables does not make them secret if they are prefixed `NEXT_PUBLIC_` or otherwise delivered to browser code.

## 20. Values that must never be committed

Never commit or expose:

- Client secrets; none should be created for this SPA
- Certificates, private keys, federated credential material, or managed-identity tokens
- Access tokens, ID tokens, refresh tokens, authorization codes, or serialized MSAL caches
- Browser storage dumps or copied bearer-token values
- `VERCEL_OIDC_TOKEN`, Sites bypass tokens, source-repository credentials, or other hosting credentials
- Tenant administrator passwords, recovery codes, MFA material, or session cookies
- `.env.local` or any file containing credentials
- Raw Graph error bodies or debug captures containing IU user/SharePoint data

An `.env.example` may contain key names and non-sensitive placeholders. It must not contain a real token or credential.

## 21. Graph calls that verify authentication

Run these later through the `IU Work Tracker DEV` MSAL token, not through Graph Explorer's own app registration:

### A. Verify identity and `User.Read`

```http
GET https://graph.microsoft.com/v1.0/me?$select=id,displayName,userPrincipalName,mail
Authorization: Bearer {delegated-access-token}
```

Expected result: `200 OK`, an IU29 user object, and no raw token output in logs.

### B. Verify SharePoint site resolution

```http
GET https://graph.microsoft.com/v1.0/sites/siu29.sharepoint.com:/sites/IUWorkTrackerDEV?$select=id,displayName,webUrl
Authorization: Bearer {delegated-access-token}
```

Expected result: `200 OK`, the exact DEV `webUrl`, and an opaque `id`.

### C. Verify collection read access

After resolving the site ID:

```http
GET https://graph.microsoft.com/v1.0/sites/{site-id}/lists?$select=id,name,displayName&$top=10
Authorization: Bearer {delegated-access-token}
```

Expected result: `200 OK`. This is read-only and does not create the seven application lists.

Graph Explorer is useful for understanding endpoints, but a successful Graph Explorer call does **not** prove that `IU Work Tracker DEV` has the correct registration, redirect, consent, cache, or token. Final verification must use a token whose client ID is the new registration.

## 22. Verify access to the DEV SharePoint site

Verify all of the following:

1. The signed-in user can open `https://siu29.sharepoint.com/sites/IUWorkTrackerDEV` in SharePoint.
2. The `/me` call succeeds under the new application token.
3. The path-based site call returns `webUrl` exactly matching `https://siu29.sharepoint.com/sites/IUWorkTrackerDEV` after normal URL normalization.
4. The site response includes a non-empty `id`.
5. Reading `/sites/{site-id}/lists` succeeds.
6. A `403` is treated as missing consent and/or insufficient signed-in-user SharePoint rights; it is not bypassed.
7. A `404` requires checking the hostname, path, and the signed-in user's site visibility; do not assume the site is absent based on an unauthorized context alone.

Delegated Graph permission does not grant the user SharePoint rights they do not already have. Both the delegated scope and the user's site authorization must permit the operation.

## 23. Resolve and store the Graph site ID

Use this exact call once the user has a valid delegated token:

```http
GET /v1.0/sites/siu29.sharepoint.com:/sites/IUWorkTrackerDEV?$select=id,displayName,webUrl
Host: graph.microsoft.com
Authorization: Bearer {delegated-access-token}
```

Microsoft documents this hostname-plus-server-relative-path form in [Get a SharePoint site by path](https://learn.microsoft.com/en-us/graph/api/site-getbypath?view=graph-rest-1.0).

Record the entire returned `id` unchanged as `NEXT_PUBLIC_SHAREPOINT_SITE_ID` for DEV runtime configuration. A site ID commonly contains comma-separated components, but it is opaque: do not split, parse, regenerate, or substitute the SharePoint URL for it.

On later startup, configuration verification may read `/sites/{site-id}` and compare the returned `webUrl` to the expected DEV URL. Runtime requests should use the configured site ID rather than resolving by mutable path every time.

## 24. Verify read/write permission without provisioning the seven lists

Read access can be proven without a write by resolving the site and enumerating its lists.

Write access cannot be conclusively proven with a non-mutating request. The token's `scp` claim can show `Sites.ReadWrite.All`, but that does not prove the signed-in user has write rights to this particular SharePoint site.

Use one of these later verification paths:

### Preferred: approved existing scratch list

If the DEV site already contains a tenant-approved disposable scratch list with a writable `Title` field:

1. Resolve its list GUID.
2. Create one uniquely named item:

```http
POST /v1.0/sites/{site-id}/lists/{scratch-list-id}/items
Content-Type: application/json
Authorization: Bearer {delegated-access-token}

{"fields":{"Title":"IUWT auth probe {uuid}"}}
```

3. Confirm `201 Created` and record the returned item ID.
4. Delete only that exact probe item:

```http
DELETE /v1.0/sites/{site-id}/lists/{scratch-list-id}/items/{probe-item-id}
Authorization: Bearer {delegated-access-token}
```

5. Confirm `204 No Content` and verify that the item is gone.

Creating and deleting list items under delegated access requires `Sites.ReadWrite.All`; see Microsoft's [create list item](https://learn.microsoft.com/en-us/graph/api/listitem-create?view=graph-rest-1.0) and [delete list item](https://learn.microsoft.com/en-us/graph/api/listitem-delete?view=graph-rest-1.0) documentation.

### If no approved scratch list exists

Do not write to a system list, document library, or unrelated business list. Do not add `Sites.Manage.All` merely to create an authentication test list. Defer write proof until the first planned DEV application list is provisioned and its smoke-test item is authorized by the provisioning checklist.

The inspected working implementation offers no zero-impact Graph operation that proves SharePoint write access. Any claim of verified write access without an actual approved mutation would be a guess.

## 25. DEV-only safety boundaries

1. Register `IU Work Tracker DEV` as single-tenant and separate from the later production registration.
2. Use only delegated `User.Read` and `Sites.ReadWrite.All`; add no application permissions.
3. Limit enterprise-application assignment to the dedicated IU Work Tracker DEV security group where practical.
4. Require the signed-in users themselves to have only the SharePoint rights needed on the DEV site.
5. Use only `https://siu29.sharepoint.com/sites/IUWorkTrackerDEV`; never configure a production site ID or list GUID in the DEV provider.
6. Keep the Sites deployment owner-only during initial authentication verification.
7. Do not reuse MAC's `IEP_Users2` browser-side role scan as the security boundary.
8. Do not create a client secret.
9. Keep tokens in MSAL `sessionStorage`; never manually persist or log them.
10. Use a pinned package, strict CSP, and normal XSS protections before enabling real Graph tokens.
11. Keep Microsoft Graph calls inside the future DEV-only `DelegatedSharePointDataProvider`; UI components do not call Graph directly.
12. Preserve strict schema validation, `ProviderResult` mapping, pagination, and the approved `RecordVersion` plus ETag algorithm.
13. Do not treat Graph scope consent as proof of access to a particular SharePoint site.
14. Do not use Graph Explorer results as proof that this app registration works.
15. Keep D1 and current production persistence unchanged until a separately approved cutover.
16. Reassess the hosted callback before allowing any viewer beyond the owner.

## 26. Exact manual Entra clicks

The menu labels below reflect the current Microsoft Entra admin-center flow. A tenant administrator may see additional policy panels; do not change them unless specified.

### Create the registration

1. Open `https://entra.microsoft.com`.
2. Confirm the active directory is the IU29 tenant whose Directory ID is `3276761c-22db-462b-a930-172d155bd795`.
3. Select **Identity**.
4. Select **Applications**.
5. Select **App registrations**.
6. Select **New registration**.
7. Enter the name `IU Work Tracker DEV`.
8. Under **Supported account types**, select **Accounts in this organizational directory only**.
9. Leave the initial Redirect URI blank so platform configuration is performed explicitly on the next screen.
10. Select **Register**.
11. On **Overview**, copy and securely record:
    - **Application (client) ID**
    - **Directory (tenant) ID**
12. Verify the tenant ID is exactly `3276761c-22db-462b-a930-172d155bd795`.

### Configure SPA redirects

13. Under **Manage**, select **Authentication**.
14. Under **Platform configurations**, select **Add a platform**.
15. Select **Single-page application**.
16. Enter `http://localhost:3000/`.
17. Select **Configure**.
18. In the SPA platform, select **Add URI**.
19. Enter `https://iu-work-tracker.gmacer.chatgpt.site/`.
20. Select **Save**.
21. Verify both URIs appear under **Single-page application** and neither appears under **Web** or **Mobile and desktop applications**.
22. Under **Implicit grant and hybrid flows**, leave both **Access tokens** and **ID tokens** unchecked.
23. Leave **Front-channel logout URL** blank.
24. Under **Advanced settings**, leave **Allow public client flows** set to **No**.
25. Select **Save** if the page shows unsaved changes.

### Configure delegated Graph permissions

26. Under **Manage**, select **API permissions**.
27. Verify `Microsoft Graph — User.Read — Delegated` already exists. If it does not:
    1. Select **Add a permission**.
    2. Select **Microsoft Graph**.
    3. Select **Delegated permissions**.
    4. Search for `User.Read`.
    5. Select the exact `User.Read` permission.
    6. Select **Add permissions**.
28. Select **Add a permission**.
29. Select **Microsoft Graph**.
30. Select **Delegated permissions**.
31. Search for `Sites.ReadWrite.All`.
32. Select the exact `Sites.ReadWrite.All` delegated permission.
33. Select **Add permissions**.
34. Verify the table contains only the required Microsoft Graph delegated permissions for this phase:
    - `User.Read`
    - `Sites.ReadWrite.All`
35. Verify no Microsoft Graph **Application** permission was added.
36. Do not add `Sites.Manage.All`, `Sites.FullControl.All`, `Directory.Read.All`, or `User.Read.All`.

### Handle consent without guessing tenant policy

37. Inspect the **Status** column for both permissions.
38. If IU29 requires tenant-wide admin consent, have an authorized tenant administrator select **Grant admin consent for IU29**, review the two permissions, and confirm.
39. If IU29 permits user consent and does not require tenant-wide approval, leave tenant-wide admin consent unchanged; the assigned DEV user will consent during first sign-in.
40. If neither path is authorized, stop. Do not attempt to bypass the tenant consent policy.

### Leave confidential-client features empty

41. Select **Certificates & secrets**.
42. Verify there is no client secret, certificate, or federated credential for this SPA. Do not create one.
43. Leave **Expose an API** unchanged; no custom API scope is needed.
44. Leave **Token configuration** unchanged; no optional claim is required by the inspected flow.

### Establish ownership and DEV assignment

45. Select **Owners**.
46. Add the approved primary steward and backup steward if they are not already owners.
47. Open **Identity → Applications → Enterprise applications**.
48. Open `IU Work Tracker DEV`.
49. Under **Manage**, select **Properties**.
50. If IU governance approves group-restricted assignment, set **Assignment required?** to **Yes** and select **Save**.
51. Select **Users and groups**.
52. Select **Add user/group**.
53. Select the dedicated IU Work Tracker DEV security group.
54. Select **Assign**.
55. Verify the intended DEV tester is a member of that group before testing.

### Record configuration for the later coding phase

56. Record the new Application (client) ID as the future value of `NEXT_PUBLIC_MS_ENTRA_CLIENT_ID`.
57. Record `3276761c-22db-462b-a930-172d155bd795` as `NEXT_PUBLIC_MS_ENTRA_TENANT_ID`.
58. Record both exact SPA redirect URIs in the DEV configuration manifest.
59. Record whether consent is user-level, tenant-wide, pending, or denied.
60. Do not record a secret because this architecture has none.

# MANUAL ENTRA SETUP

- [ ] Sign in to `https://entra.microsoft.com` and switch to tenant `3276761c-22db-462b-a930-172d155bd795`.
- [ ] Open **Identity → Applications → App registrations → New registration**.
- [ ] Name the registration `IU Work Tracker DEV`.
- [ ] Select **Accounts in this organizational directory only**.
- [ ] Leave the initial redirect blank and select **Register**.
- [ ] Record the new **Application (client) ID** and confirm the **Directory (tenant) ID**.
- [ ] Open **Authentication → Add a platform → Single-page application**.
- [ ] Add `http://localhost:3000/`.
- [ ] Add `https://iu-work-tracker.gmacer.chatgpt.site/` to the same SPA platform.
- [ ] Confirm there are no Web or Mobile/Desktop redirect platforms for these URIs.
- [ ] Leave implicit **Access tokens** and **ID tokens** unchecked.
- [ ] Leave **Front-channel logout URL** blank.
- [ ] Leave **Allow public client flows** set to **No**.
- [ ] Open **API permissions** and confirm `User.Read` is **Delegated**.
- [ ] Add Microsoft Graph `Sites.ReadWrite.All` as **Delegated**.
- [ ] Confirm no Microsoft Graph **Application** permissions are present.
- [ ] Do not add `Sites.Manage.All`, `Sites.FullControl.All`, `Directory.Read.All`, or `User.Read.All`.
- [ ] Follow IU29 policy for consent; if **Need admin approval** appears, stop and request authorized approval.
- [ ] Leave **Certificates & secrets** empty.
- [ ] Leave **Expose an API** and **Token configuration** unchanged.
- [ ] Add the approved primary and backup application owners.
- [ ] In **Enterprise applications → IU Work Tracker DEV**, set **Assignment required?** to **Yes** only if IU governance approves group-restricted assignment.
- [ ] Assign the dedicated IU Work Tracker DEV security group.
- [ ] Record the client ID, tenant ID, redirect URIs, owners, group, and consent status in the DEV configuration manifest.
- [ ] After implementation, verify `/me` using a token issued to this client ID.
- [ ] Resolve `GET /sites/siu29.sharepoint.com:/sites/IUWorkTrackerDEV` and record the opaque site ID.
- [ ] Verify `/sites/{site-id}/lists` returns `200` without creating application lists.
- [ ] Prove write access only with an approved scratch-list probe, or defer it until the first authorized provisioning smoke test.
- [ ] Confirm the deployed owner-only Entra callback completes successfully before expanding Sites access.
