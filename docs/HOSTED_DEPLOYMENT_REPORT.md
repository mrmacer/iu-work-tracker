# IU Work Tracker — Hosted Deployment Report

## Deployment result

**HOSTED READINESS: GREEN**

The verified GREEN IU Work Tracker build is deployed to the existing owner-only Sites environment:

- URL: `https://iu-work-tracker.gmacer.chatgpt.site/`
- Source commit: `7e22e5909d1218c88a3be4deec45d3c61418da0f`
- Sites version: `3`
- Sites deployment: `appgdep_6a9490913b6881918fbeae33f15a4d24`
- Hosted environment revision: `2`
- Deployment date: 2026-08-30

The deployment archive was built from a temporary clean checkout of the exact source commit. It did not include `.env.local`, credentials, `app/dev-sharepoint-smoke/`, `docs/MAC_WALKTHROUGH_AUTH_REFERENCE.md`, or other untracked/experimental files. The existing commit was pushed only to the established Sites source repository as required by the existing deployment path. No new Git commit and no GitHub/origin push were made.

## Pre-deployment verification

- Local `main` and `origin/main`: both `7e22e5909d1218c88a3be4deec45d3c61418da0f`
- Type-check: PASS
- Relevant tracked TypeScript/TSX lint: PASS (0 errors; only ESLint's expected ignored-file warning for `next-env.d.ts`)
- Full automated tests: PASS — 10 Vitest files / 113 tests plus 1 server-render test, 114 total, 0 failed
- Production build: PASS
- Clean-checkout deployment build: PASS
- Client bundle secret scan: PASS
- Client bundle contains no `ANTHROPIC_API_KEY` identifier, Anthropic key value, or `api.anthropic.com` endpoint

The deployed route set is `/`, `/api/records`, and `/api/inbox-intelligence`. The excluded local-only smoke route was not deployed.

## Hosted environment configuration

The Sites environment supplies these public identifiers:

- `NEXT_PUBLIC_MS_ENTRA_CLIENT_ID`
- `NEXT_PUBLIC_MS_ENTRA_TENANT_ID`
- `NEXT_PUBLIC_SHAREPOINT_HOSTNAME`
- `NEXT_PUBLIC_SHAREPOINT_SITE_PATH`
- `NEXT_PUBLIC_SHAREPOINT_SITE_ID`
- `NEXT_PUBLIC_SHAREPOINT_IU_WORK_RECORDS_LIST_ID`
- `NEXT_PUBLIC_SHAREPOINT_IU_INBOX_INTELLIGENCE_LIST_ID`
- `SITE_ORIGIN`

The Sites secret mechanism supplies:

- `ANTHROPIC_API_KEY`

`ANTHROPIC_API_KEY` remains server-only. Its value was never printed, committed, rendered, logged, or placed in a `NEXT_PUBLIC_*` variable.

Verified SharePoint target:

- Hostname: `siu29.sharepoint.com`
- Site path: `/sites/IUWorkTrackerDEV`
- Graph site ID: `siu29.sharepoint.com,dbea0e1c-562b-4296-8e69-a820448a4acd,dc4fa421-a8b6-4a58-84d0-8d84cbd7a8ec`
- Work Records list ID: `db7540bb-4896-4eb8-aa5f-b911038d0460`
- Inbox Intelligence list ID: `892dbe47-6fa2-42f0-b9c4-1ed7a3664737`

No other SharePoint site was used.

## Microsoft authentication

Hosted delegated sign-in succeeded for `Macer, Gregory` using authorization code with PKCE.

The live authorization request verified:

- Redirect URI: `https://iu-work-tracker.gmacer.chatgpt.site/`
- Tenant authority: IU29 tenant `3276761c-22db-462b-a930-172d155bd795`
- Graph delegated scopes: `User.Read` and `Sites.ReadWrite.All`
- No implicit flow, client secret, application permission, `Sites.Manage.All`, or `Mail.Read`

The hosted diagnostic showed:

- Microsoft: Connected
- SharePoint DEV: Connected
- Signed in as: Macer, Gregory
- Correct opaque Graph site ID
- 10 existing lists readable

Authentication and account state survived a full hosted reload.

## Hosted provider status

The application footer reported `SharePoint DEV connected`. Connected operation loaded zero records from the DEV SharePoint lists before the smoke tests and did not fall back to the prototype provider.

Reference/configuration data remains static seed data by the existing DEV design; this deployment did not change that architecture.

## Work Record hosted smoke test

One synthetic Work Record was created through the normal Log Work UI:

- Title: `HOSTED DEPLOY TEST — DELETE ME`
- SharePoint item ID: `3`
- AppId: `8e889d56-026b-41b9-93c8-d91602b5ecd9`
- RecordVersion: `1`
- ETag: `"1"`
- Created: `2026-08-30T20:31:08Z`
- Modified: `2026-08-30T20:31:08Z`

The record was read directly from `IU_Work_Records`, survived a full hosted reload, and appeared in the normal Today/Home UI. The provider indicator remained `SharePoint DEV connected`.

Cleanup:

- Removed the exact synthetic item from the active list
- Permanently removed only that exact item from the SharePoint recycle bin
- Exact item lookup returned SharePoint's item-not-found/HTTP 404 response
- `IU_Work_Records` returned to its pre-test empty state

## Inbox Intelligence hosted smoke test

One synthetic, non-sensitive email was processed through the normal hosted flow:

`PASTE → ANALYZE → REVIEW → SAVE TO INBOX`

Exactly one Anthropic analysis request was made. Sites worker logs contain one successful request:

- Method/route: `POST /api/inbox-intelligence`
- Status: `200`
- Outcome: `ok`
- Request ID: `72f78be216ee4281bc4707fd9b4d1672`

No second analysis request was made. The browser sent the request only to the hosted application route; the client bundle contains no Anthropic endpoint or credential. The structured result rendered successfully before explicit user-directed persistence.

Durable SharePoint verification:

- SharePoint item ID: `2`
- AppId: `5254a7e1-2509-49c3-9408-21bf003ced86`
- Initial status: `open`
- Full hosted reload: PASS
- `open → waiting`: ETag/RecordVersion `1 → 2`
- `waiting → resolved`: ETag/RecordVersion `2 → 3`
- `resolved → open` (Reopen): ETag/RecordVersion `3 → 4`
- Created remained `2026-08-30T20:33:07Z`
- Modified advanced on every update, ending at `2026-08-30T20:33:55Z`
- `ResolvedAt` was set when resolved and cleared when reopened
- Action-item JSON and relationship/tag JSON round-tripped exactly

Privacy verification:

- Only the approved compact `SourceExcerpt` was persisted
- No raw email body/thread/signature field was persisted
- No Anthropic request or response payload was persisted
- No API credential was persisted

Cleanup:

- Removed the exact synthetic item from the active list
- Permanently removed only that exact item from the SharePoint recycle bin
- Exact item lookup returned SharePoint's item-not-found/HTTP 404 response
- `IU_Inbox_Intelligence` returned to its pre-test empty state

## Known limitations

- Hosted access is intentionally owner-only in the current Sites environment.
- SharePoint authentication remains DEV-only delegated authentication and is single-user verified.
- Production authentication still requires separate evaluation of the planned narrower `Sites.Selected` application architecture.
- Reference/configuration lists are not yet loaded from SharePoint.
- Interactive Microsoft password/MFA steps require the user and are not automated.

No SharePoint schema, data model, Anthropic behavior, Entra permission, persistence architecture, or application feature was changed during deployment.

