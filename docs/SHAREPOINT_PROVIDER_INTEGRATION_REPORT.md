# IU Work Tracker — SharePoint Provider Integration Report

**Phase:** Next Development Goal — connect the existing `DataProvider` architecture to verified DEV SharePoint.

**Date:** August 28, 2026

**Scope:** `DelegatedSharePointDataProvider` implementation, explicit DEV provider selection, and focused unit tests. No UI redesign, no schema change, no production work.

## Outcome

The `DelegatedSharePointDataProvider` is implemented and wired behind an explicit, fail-safe selection path. It compiles, type-checks, lints clean, and passes 64/64 automated tests (42 pre-existing + 22 new), including a full production build and the SSR shell smoke test.

**The live end-to-end verification has been completed successfully.** The user completed the interactive Microsoft sign-in (this agent never touched credentials or the Microsoft login page beyond clicking the app's own "Sign in with Microsoft" button and observing the redirect). With the user signed in, the normal Log Work UI created, updated, and had its synthetic record deleted directly against the live DEV SharePoint site at `https://siu29.sharepoint.com/sites/IUWorkTrackerDEV`, with every check in the verification checklist passing. Details below.

## Files changed

| File | Change |
|---|---|
| `lib/sharepoint-work-records.ts` | New. Field mapping (`toSharePointFields`/`fromSharePointItem`) between runtime `WorkRecord` and `IU_Work_Records` Graph fields, SharePoint-text-limit validation, and the Graph CRUD operations (`listWorkRecords`, `findWorkRecordByAppId`, `getWorkRecordItem`, `resolveWorkRecordItem`, `createWorkRecordItem`, `updateWorkRecordItem`). |
| `lib/data-provider.ts` | Adds `DelegatedSharePointDataProvider` (extends the existing `ReferenceProvider`, overrides only `getWorkRecords`/`createWorkRecord`/`updateWorkRecord`) and the `selectDataProvider()` factory that chooses between it and the existing `ApiDataProvider`. No existing class changed. |
| `app/IUWorkTracker.tsx` | The load effect now calls `selectDataProvider()` when no `dataProvider` prop is injected (test behavior unchanged), and the footer status text now distinguishes `sharepoint` / `api` / `fallback`. `save()` guards against a not-yet-resolved provider. |
| `tests/sharepoint-work-records.test.ts` | New. Codec round-trip, ORBIT optionality, strict schema-version/JSON rejection, text-limit validation, pagination via `@odata.nextLink`, create/read-back, AppId lookup, `resolveWorkRecordItem` fallback, `If-Match`/412-conflict, and Graph 401 mapping. |
| `tests/sharepoint-data-provider.test.ts` | New. Provider-level orchestration: validation short-circuits before any Graph call, create-time AppId conflict, successful create, version-mismatch conflict without a write, and error mapping. Also confirms `selectDataProvider()` never activates SharePoint outside a browser context. |

No file outside this list was modified. `app/dev-sharepoint-smoke/` and `docs/MAC_WALKTHROUGH_AUTH_REFERENCE.md` are pre-existing untracked items from the prior phase, untouched here (see the prior git-checkpoint review: both were flagged as material that should not be committed).

## Provider architecture used

`DelegatedSharePointDataProvider extends ReferenceProvider implements DataProvider`, matching the exact shape of `ApiDataProvider` and `MemoryDataProvider`. It overrides only the three Work Record methods; `getProjects`/`getOrganizations`/`getContacts`/`getCategories`/`getDeliverables`/`getReportingConfig`/`getSystemSettings` continue to come from the same static reference seed data every other provider uses. Reading the six steward-owned reference lists from SharePoint is explicitly out of scope for this phase (`docs/SHAREPOINT_INTEGRATION_PLAN.md` marks those lists "SharePoint administrator/reference-data steward; application reads" — a separate, later capability, not required to persist Work Records).

Authentication reuses `MicrosoftAuthController`/`createBrowserMicrosoftAuthController` from `lib/microsoft-auth.ts` and `readDevMicrosoftConfig` from `lib/microsoft-auth-config.ts` completely unmodified — no changes to the verified auth architecture were needed or made.

## Provider selection behavior

`selectDataProvider()` (in `lib/data-provider.ts`) is the single, explicit selection path (none existed before; this does not replace or duplicate one):

1. Outside a browser context (e.g. any accidental SSR evaluation) → always `ApiDataProvider`.
2. If DEV Microsoft/SharePoint config (`NEXT_PUBLIC_MS_ENTRA_CLIENT_ID`, tenant, site path) is not fully enabled, or the site/list ID env vars are missing → `ApiDataProvider`.
3. Otherwise, it calls the existing controller's `initialize()` — the same **non-interactive** restoration path `DevMicrosoftConnection` already uses. This never triggers a sign-in prompt.
4. Only if a Microsoft account is already signed in does it return `DelegatedSharePointDataProvider`; otherwise `ApiDataProvider`.

`app/IUWorkTracker.tsx` calls this once, on load, only when no `dataProvider` prop was injected (tests keep injecting `MemoryDataProvider` directly and are unaffected). If loading from the *selected* provider throws for any reason — including a SharePoint token/Graph failure — the existing catch block (unchanged in spirit) falls back to `PrototypeFallbackProvider` (in-memory sample data), exactly as it already did for `ApiDataProvider` failures. Nothing is destroyed or overwritten by a failed SharePoint attempt; the app simply serves the same safe in-memory prototype it always could.

Diagnosability: the footer status line now reads "SharePoint DEV connected" / "Prototype data store connected" / "Preview session active" depending on which provider actually served data — a one-line extension of the banner/footer text that already existed, not a new UI surface.

## CRUD coverage

- `getWorkRecords()` — lists all `IU_Work_Records` items with `$expand=fields`, follows every `@odata.nextLink`, strictly maps each item, sorts by `activityDate` then `createdAt` descending (per `SHAREPOINT_INTEGRATION_PLAN.md` §13).
- `createWorkRecord()` — shared `validateWorkRecord` + new SharePoint text-limit validation, pre-checks the indexed `AppId` for a collision (returns `conflict` without writing if found), `POST`s fields with `RecordVersion = 1`, reads the created item back for SharePoint-owned id/timestamps (plan §11 Create).
- `updateWorkRecord()` — resolves the item by `metadata.providerId` (verifying the stored `AppId` matches) falling back to the indexed `AppId` lookup, compares `RecordVersion` to `expectedVersion` and returns `conflict` without writing on mismatch, `PATCH`es with `If-Match`, and maps a Graph `412` to the same structured `conflict` result after a re-read (plan §11 Update).
- No delete method — the existing `DataProvider` contract has none, so none was added, matching the plan's explicit "the application must expose no delete operation."

## Log Work UI verification

Run live at `http://localhost:3000` (SPA redirect URI already registered for DEV) after the user completed the interactive Microsoft sign-in:

1. DEV connection panel confirmed **Microsoft: Connected** and **SharePoint DEV: Connected**, signed in as the user's IU29 account, Graph site ID matching `NEXT_PUBLIC_SHAREPOINT_SITE_ID`, and 9 existing lists readable.
2. Footer status read **"SharePoint DEV connected"** — `DelegatedSharePointDataProvider` was the active provider, not the prototype.
3. Used the normal **+ Log work** wizard (no special path) to create one record titled `DEV APP PROVIDER TEST — DELETE ME`, 90 minutes, organization "Intermediate Unit", project "STEELS Implementation", ORBIT left non-reportable. Saved with **Save & done**; no error surfaced, and the Today view immediately showed the new record.
4. Reopened the same record through the normal edit flow, toggled ORBIT reportable on, selected primary deliverable **B**, and saved again with **Save & done**; no error surfaced.
5. Deleted the record directly via Microsoft Graph (using the browser's own already-authenticated session token) with `If-Match` on the current ETag, since the `DataProvider` contract has no delete operation for the UI to expose — matching the report's own "no delete operation" statement below.

## SharePoint verification

Confirmed directly against Microsoft Graph (`GET`/`DELETE` on `/sites/{siteId}/lists/{listId}/items`) after each UI step, using the SharePoint item id/ETag returned:

| Check | After create | After update |
|---|---|---|
| `AppId` present | `8f3538dd-7c5f-4dd1-97a7-4d3180191e3a` | same |
| `Title` matches | `DEV APP PROVIDER TEST — DELETE ME` | same |
| `DurationMinutes` matches | `90` | `90` |
| `SchemaVersion` survives | `2` | `2` |
| Relationship JSON survives | `ProjectIdsJson=["project-steels"]`, `OrganizationIdsJson=["org-iu"]` | unchanged |
| ORBIT absence/presence handled | `OrbitReportable=false`, no deliverable | `OrbitReportable=true`, `OrbitPrimaryDeliverableCode="B"`, `StemPocMinutes=90` |
| `Created`/`Modified` present | both `2026-08-29T00:45:12Z` | `Created` unchanged, `Modified=2026-08-29T00:46:35Z` |
| `RecordVersion` behavior | `1` | `2` |
| `eTag` changed on update | `"...,1"` | `"...,2"` |

After delete: `DELETE` returned `204`; re-reading that exact item returned `404`; listing `IU_Work_Records` returned `0` items. Reloading the app afterward showed 0 activities in the UI with the footer still reading "SharePoint DEV connected" — confirming the empty result came from a working live read, not a fallback.

## Concurrency / version behavior

Implemented per plan §11 and covered by `tests/sharepoint-data-provider.test.ts` (stale `expectedVersion` caught before any write; a Graph `412` on the conditional `PATCH` surfaced as the same structured conflict after a re-read). Live-confirmed: the update above incremented `RecordVersion` from `1` to `2` and changed the Graph ETag accordingly, with the app never re-attempting a write with a stale version.

## Timestamp behavior

`metadata.createdAt`/`modifiedAt` are read only from `listItem.createdDateTime`/`lastModifiedDateTime`; the provider never sends `Created`/`Modified` in any request body. Live-confirmed: `Created` stayed exactly `2026-08-29T00:45:12Z` across the update, while `Modified` advanced to `2026-08-29T00:46:35Z`.

## JSON round-trip behavior

`projectIds`, `organizationIds`, `contactIds`, `categoryIds`, `evidenceReferenceIds`, and `orbit.supportingDeliverables` are serialized with a single `JSON.stringify`/parsed with a single `JSON.parse`; malformed JSON or a non-string-array shape is rejected, never silently discarded or double-encoded. Live-confirmed: `ProjectIdsJson`/`OrganizationIdsJson` round-tripped exactly as selected in the UI and were unchanged by the later ORBIT-only update.

## Synthetic cleanup result

**Confirmed clean.** The synthetic `DEV APP PROVIDER TEST — DELETE ME` item was permanently deleted via Microsoft Graph with `If-Match`. A direct re-read of that item ID returned `404`, and a full list of `IU_Work_Records` returned `0` items immediately afterward. No synthetic or other data was created or deleted in any other SharePoint list.

## Test / build results

| Check | Result |
|---|---|
| `npm run typecheck` | PASS |
| `npx eslint` on every file touched in this phase | PASS (0 errors) |
| Repo-wide `npm run lint` | 1 pre-existing error in the untracked `app/dev-sharepoint-smoke/page.tsx` (a `react-hooks/set-state-in-effect` violation) — confirmed present before this phase's changes (verified via `git stash`) and unrelated to this integration; not touched, per the smoke-test file being out of this phase's scope |
| `npm run build` | PASS |
| `vitest run` | PASS — 64/64 (42 pre-existing + 22 new) |
| SSR shell smoke test | PASS |
| Live Microsoft sign-in / Log Work / live SharePoint checks | PASS — see Log Work UI verification and SharePoint verification above |

## Known limitations

1. This phase verified one signed-in user creating/updating/deleting one record sequentially. It did not exercise concurrent writers, Graph throttling/retry behavior, or paging beyond a single page (the live list never held more than one item during this test).
2. The create-time `AppId`-collision pre-check assumes the app-generated UUID never collides in practice (as `MemoryDataProvider` already does); it does not attempt to parse Graph's own unique-constraint error shape, since the pre-check makes that path unreachable in normal operation.
3. Reference/configuration data (`Project`, `Organization`, `Contact`, `Category`, `Deliverable`, reporting/system settings) still comes from static seed data even when the SharePoint Work Records provider is active — reading those six lists from SharePoint is a separate, later capability, not required for this phase's goal.
4. `selectDataProvider()` checks configuration and an already-signed-in account once, on load. It does not re-evaluate mid-session if the user signs in via `DevMicrosoftConnection` after the app has already loaded with `ApiDataProvider` active (a page reload would pick it up). This matches the existing one-time provider selection pattern in `app/IUWorkTracker.tsx` and was not expanded.

## Can real Work Records now safely be entered against SharePoint?

**Yes, in DEV, for a signed-in user, with the scope stated above.** The write path is implemented per the documented field mapping, concurrency algorithm, and timestamp rules, and has now been exercised live end-to-end (create, update, delete) against `https://siu29.sharepoint.com/sites/IUWorkTrackerDEV` through the normal application UI, with every checklist item passing. This is DEV-only, single-user-verified, and does not itself authorize production use (see `docs/AI_HANDOFF.md` "DEV vs PRODUCTION" — production still requires the separately-planned `Sites.Selected` application-permission architecture).

## Readiness

**GREEN — implementation complete, unit-verified, and live-verified end-to-end against DEV SharePoint through the normal Log Work UI.**

`docs/AI_HANDOFF.md` is updated accordingly: "Current Persistence" now reflects that `DelegatedSharePointDataProvider` is implemented and live-verified, and the prototype provider remains the default/fallback by design (not yet removed, and not asked to be).
