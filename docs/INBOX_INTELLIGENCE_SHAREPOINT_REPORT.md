# IU Work Tracker — Inbox Intelligence V1.1: Durable SharePoint Persistence

**Phase:** Inbox Intelligence V1.1 — durable DEV SharePoint persistence for reviewed Inbox Intelligence.

**Status: GREEN — the approved DEV SharePoint schema is provisioned and verified, the durable provider is active under DEV configuration, and the complete create/reload/status/reopen/cleanup lifecycle has passed against live SharePoint.**

## What changed from V1

V1's workflow was PASTE EMAIL → ANALYZE → REVIEW, with session-only save (lost on reload). V1.1 implements the full target workflow end-to-end in code:

PASTE EMAIL → ANALYZE → REVIEW/EDIT → **SAVE TO INBOX** → **APPEARS IN INBOX** (Needs attention / Waiting / Recent-resolved) → optionally CREATE WORK RECORD → **TRACK UNTIL RESOLVED** (open ⇄ waiting → resolved ⇄ reopen)

The AI analysis pipeline itself (already GREEN in V1) was **not modified** — no prompt change, no schema change, no new Anthropic call sites. This phase is entirely about what happens *after* analysis: reviewed intelligence becomes a durable, trackable record instead of disappearing on reload.

## Runtime model

`InboxIntelligenceRecord` (`lib/inbox-intelligence-models.ts`) — the reviewed intelligence, never the email:

```ts
{
  appId: string;
  schemaVersion: 1;
  sourceType: "pasted-email";
  analyzedAt: string;              // ISO timestamp — when AI produced this analysis
  sourceExcerpt: string;           // short excerpt only, ≤500 chars — see "Privacy behavior"
  analysis: EmailAnalysis;         // V1's exact, unmodified schema (summary, priority, needsAttention,
                                    // actionItems, followUp, people, organizations, districts, projects,
                                    // tags, suggestedWorkType, suggestedWorkRecord)
  matchedOrganizationIds: string[]; // resolved canonical org AppIds (exact-name match only)
  matchedDistrictIds: string[];     // resolved canonical district AppIds
  matchedProjectIds: string[];      // resolved canonical project AppIds
  status: "open" | "waiting" | "resolved";
  resolvedAt: string | null;
  linkedWorkRecordAppId: string | null;
  metadata: ProviderMetadata;       // the exact same shape WorkRecord uses — providerId, version,
                                     // createdAt, modifiedAt, syncState — reused, not duplicated
}
```

Two deliberate simplifications versus the fields listed in the authorizing instructions, both because reusing existing structure fully covers them without a parallel concept:

- **`savedAt`** is not a separate column. SharePoint's own `Created` (via `metadata.createdAt`) already is "when this became durable" — adding a second timestamp for the same moment would duplicate what `docs/SHAREPOINT_PROVISIONING_SPEC.md` §15 already establishes as canonical.
- **`RecordVersion`**/**ETag** are not new concepts — they are `metadata.version` and a provider-private opaque token, exactly like `WorkRecord`. No parallel concurrency model was invented.

`resolveEmailAnalysisEntities()` (shared by the Work Record prefill mapping and the durable record's `matched*Ids`) performs the same exact-name, case-insensitive, non-fuzzy matching V1 already established — one implementation, reused by both callers, not duplicated.

## SharePoint schema (approved, see "Provisioning" for status)

Proposed and reviewed against the runtime model above — no material mismatch found, so no schema/model conflict needed to be escalated. Full column table added to `docs/SHAREPOINT_PROVISIONING_SPEC.md` as a new, additive section (existing sections untouched). Summary:

| Internal name | Type | Required | Notes |
|---|---|---|---|
| `Title` (built-in) | Single line text | Yes | `analysis.suggestedWorkRecord.title` — reused, not duplicated into a second column |
| `AppId` | Single line text, unique+indexed | Yes | Stable ID, same pattern as every other list |
| `SchemaVersion` | Number, 0 decimals | Yes | `1` |
| `SourceType` | Single line text | Yes | Always `pasted-email` for V1.1 |
| `AnalyzedAt` | Date and Time (full date+time, **not** date-only) | Yes | Business timestamp, distinct from SharePoint `Created` |
| `SourceExcerpt` | Multiple lines plain text, max 500 | No | Short excerpt only — see "Privacy behavior" |
| `SummaryText` | Multiple lines plain text, max 2000 | Yes | `analysis.summary` |
| `Priority` | Choice (`high`,`medium`,`low`) | Yes | |
| `NeedsAttention` | Yes/No | Yes | |
| `ActionItemsJson` | Multiple lines plain text, max 10000 | Yes | JSON array of `{action, dueDate, owner}` |
| `FollowUpText` | Multiple lines plain text, max 1000 | No | |
| `PeopleJson` / `OrganizationsJson` / `DistrictsJson` / `ProjectsJson` / `TagsJson` | Multiple lines plain text, max 10000 | Yes | Raw AI-extracted names/tags, JSON string arrays |
| `MatchedOrganizationIdsJson` / `MatchedDistrictIdsJson` / `MatchedProjectIdsJson` | Multiple lines plain text, max 10000 | Yes | Resolved canonical AppIds, JSON string arrays |
| `SuggestedWorkType` | Single line text, max 255 | No | Unvalidated free text |
| `SuggestedWorkRecordDescription` | Multiple lines plain text, max 2000 | Yes | |
| `LinkedWorkRecordAppId` | Single line text, max 255 | No | Stable `IU_Work_Records.AppId` — **not** a Lookup column |
| `Status` | Choice (`open`,`waiting`,`resolved`) | Yes | Default `open` |
| `ResolvedAt` | Date and Time | No | Null unless `Status = resolved` |
| `RecordVersion` | Number, 0 decimals | Yes | Default `1` |
| `Created` / `Modified` | Built-in system Date/Time | — | Canonical, never client-supplied |

No SharePoint Lookup columns. No raw email, no thread, no signatures, no Anthropic request/response payload, no API secret — none of these are columns, and none can be, because the codec (`lib/sharepoint-inbox-intelligence.ts`) only ever reads from an `InboxIntelligenceRecord`, whose type has no field capable of holding them (enforced structurally, and asserted by an exhaustive-key unit test).

## Provisioning

**Complete and verified in DEV.** The two earlier delegated provisioning attempts remain useful history: list creation and column creation each returned `403 accessDenied` because the runtime app has item-write permission (`Sites.ReadWrite.All`), not schema-management permission. The user then created the list and all approved columns manually without changing the Entra permission surface.

Final verified resource:

- Display/internal list name: `IU_Inbox_Intelligence`
- List ID: `892dbe47-6fa2-42f0-b9c4-1ed7a3664737`
- URL: `https://siu29.sharepoint.com/sites/IUWorkTrackerDEV/Lists/IU_Inbox_Intelligence`
- Built-in canonical title field: internal name `Title`, required single-line text
- Approved custom fields: all 24 present with exact internal names, types, required flags, defaults, choice values, date/time formats, and JSON/plain-text settings from §23
- `AppId`: unique and indexed
- Lookup fields: none among the approved application columns
- Accidental duplicate `Title1`: removed manually before live persistence testing
- Pre-test item count: `0`

`NEXT_PUBLIC_SHAREPOINT_IU_INBOX_INTELLIGENCE_LIST_ID=892dbe47-6fa2-42f0-b9c4-1ed7a3664737` is configured in `.env.local` only. No Entra permission, production configuration, or production persistence change was made.

## Provider architecture

`DelegatedSharePointInboxIntelligenceProvider` (`lib/inbox-intelligence-provider.ts`) — minimum operations only (`list`, `create`, `update`; no delete, since the open/waiting/resolved workflow never needs one):

- Reuses the exact same delegated `MicrosoftAuthController`/`createBrowserMicrosoftAuthController` as `DelegatedSharePointDataProvider` — zero new auth code.
- Reuses the exact same numeric-`RecordVersion` + `ETag`/`If-Match` algorithm as Work Records: read-current → compare version → conditional `PATCH` → `412` mapped to a structured conflict after a re-read.
- `selectInboxIntelligenceProvider()` mirrors `selectDataProvider()` exactly: activates only when DEV config is present **and** a Microsoft account is already signed in (checked non-interactively via `controller.initialize()` — never forces a sign-in prompt), otherwise falls back to `SessionInboxIntelligenceProvider` (in-memory, same fallback role `PrototypeFallbackProvider` plays for Work Records).
- `lib/sharepoint-inbox-intelligence.ts` (the Graph codec/CRUD layer) is a new, independent file rather than a shared abstraction with `lib/sharepoint-work-records.ts` — matching this codebase's existing one-file-per-resource precedent (`lib/microsoft-graph.ts` and `lib/sharepoint-work-records.ts` are likewise independent of each other), not a novel refactor of the already-verified Work Records path.
- No direct SharePoint/Graph calls from `app/InboxIntelligence.tsx` — it only calls the provider's `list`/`create`/`update` methods, passed down as props from `app/IUWorkTracker.tsx` (which owns the provider instance, mirroring exactly how it already owns the Work Record `DataProvider`).

## Privacy behavior / no-raw-email guarantee

Unchanged principle from V1, now enforced at the durable-persistence layer too:

- `buildInboxIntelligenceRecord()` (`lib/inbox-intelligence-models.ts`) takes only an already-derived `EmailAnalysis` and an already-truncated `sourceExcerpt` — it has no parameter through which a raw email could flow in.
- The `InboxIntelligenceRecord` type itself has no field capable of holding a full email, thread, signature, or Anthropic payload — asserted by an exhaustive-key unit test (`Object.keys(record)` matches the documented shape exactly, nothing more).
- `sourceExcerpt` is capped at 500 characters in the SharePoint schema; the app only ever writes the ≤200-character excerpt V1 already computed (whitespace-collapsed, ellipsized) — the extra headroom is defensive, not an invitation to store more.
- The codec test suite asserts the mapped Graph `fields` payload never contains a `RawEmail`/`EmailBody`-shaped key.

## CRUD / status behavior

- **Create** (`Save to Inbox`): validates the record (re-validates `analysis` against the exact same `EmailAnalysisSchema` the AI pipeline enforces, plus SharePoint text-limit checks), pre-checks the indexed `AppId` for a collision, `POST`s, reads the created item back for SharePoint-owned id/timestamps — identical algorithm shape to Work Record create.
- **List**: follows every `@odata.nextLink`; strictly maps every item; throws (never silently drops) on malformed JSON or an unsupported `SchemaVersion`.
- **Status change** (`Mark waiting` / `Resolve` / `Reopen`): a single `update()` call per click — same numeric-version/ETag algorithm as any other update. `resolvedAt` is set exactly when transitioning into `resolved` and cleared exactly when transitioning out (`Reopen`), enforced by `validateInboxIntelligenceRecord`'s cross-field check (resolved requires `resolvedAt`; non-resolved forbids it).
- No delete operation exists anywhere in the stack — the UX never needs one.

## Concurrency behavior

Identical semantics to Work Records, unit-tested directly against a mocked Graph layer:

- Create writes `RecordVersion = 1`.
- Update re-reads the current item, compares `RecordVersion` to the caller's `expectedVersion` **before** writing anything; a mismatch returns the current record as a structured conflict with zero writes.
- A successful update `PATCH`es with `If-Match: {current ETag}` and `RecordVersion = expectedVersion + 1`; a Graph `412` (a race after the pre-check) is caught, the item re-read, and returned as the same structured conflict.
- `Created` is never sent in any write; `Modified` is only ever read back from SharePoint after a write succeeds.

## Work Record linking behavior

**Implemented via a small, targeted extension to the existing callback path — not a parallel save path.**

`app/IUWorkTracker.tsx`'s `openLog()` gained one new optional parameter: `onSaved?: (saved: WorkRecord) => void`, invoked exactly once, immediately after the existing `save()` function's single `createWorkRecord`/`updateWorkRecord` call succeeds (whether the user then clicks "Save & done" or "Save & log another"). This is the same single write path every other Work Record creation already uses — no second persistence route was introduced.

`app/InboxIntelligence.tsx`'s "Create Work Record" button calls `openLog(draft, onSaved)`, where `onSaved` issues exactly one follow-up `updateRecord()` call setting `linkedWorkRecordAppId` on the originating Inbox Intelligence record. If that follow-up update fails (e.g. a conflict or network error), the Work Record itself is **not** rolled back — it was already created successfully — and a non-blocking error is surfaced (`"The Work Record was created, but linking it back to this Inbox item failed: ..."`) rather than silently losing the failure or duplicating anything. Once `linkedWorkRecordAppId` is set, the UI hides "Create Work Record" for that item and shows "Linked to a Work Record" instead, preventing a duplicate.

This was judged safely implementable within the existing architecture — no STOP condition was hit here.

## AI cost behavior

No persistence operation calls Anthropic. Save, list, status changes, reopen, and Work Record creation/linking are pure provider/DataProvider operations. For the live persistence test, one exact synthetic email received a temporary deterministic local analysis response so the real UI review and Save path could be exercised with zero Anthropic requests; that local response was removed immediately after Save and left no source change. The persisted record then traversed the normal SharePoint provider for every operation.

## Tests

**113/113 passing** (was 89 before this phase). All Graph calls are mocked; **zero automated tests make a paid AI call** — none of the new test files touch the Anthropic client at all.

- `tests/inbox-intelligence-provider.test.ts` (extended): `buildWorkRecordDraftFromAnalysis` (unchanged, re-verified against the refactored shared matcher), `buildInboxIntelligenceRecord` (exhaustive no-raw-email key check, matched-ID resolution), `computeInboxIntelligenceSummary`, and `SessionInboxIntelligenceProvider` (provider-metadata assignment, newest-first ordering, duplicate-AppId conflict, `RecordVersion` increment on update, stale-version conflict, per-instance isolation).
- `tests/sharepoint-inbox-intelligence.test.ts` (new): field-mapping round-trip (asserts no `RawEmail`/`EmailBody` key ever appears), unsupported-`SchemaVersion` rejection, malformed-JSON rejection, SharePoint text-limit validation, `status`/`resolvedAt` cross-field validation, re-validation against the shared `EmailAnalysisSchema`, `@odata.nextLink` pagination, create-then-read-back, indexed AppId lookup, providerId-then-AppId-fallback resolution, `If-Match`/`412`-as-conflict, successful version-incrementing update, and a Graph `401`-to-auth-error mapping.
- `tests/inbox-intelligence-sharepoint-provider.test.ts` (new): `DelegatedSharePointInboxIntelligenceProvider` create-conflict, successful create, version-mismatch conflict without a write, network-error mapping, and validation-before-token-acquisition; `selectInboxIntelligenceProvider()` falling back to the session provider outside a browser context.
- Build/typecheck/lint: `npm run build` succeeds (client, SSR, RSC); `tsc --noEmit` clean; `eslint` clean on every file touched this phase (the one pre-existing, unrelated error in the untracked `app/dev-sharepoint-smoke/page.tsx` is unchanged from prior phases).

## Live verification

**Passed against live DEV SharePoint through the normal application UI/provider path. Zero Anthropic API calls occurred.**

1. Restored the IU DEV Microsoft session and confirmed `Microsoft: Connected`, `SharePoint DEV: Connected`, signed in as `Macer, Gregory`.
2. Opened Inbox Intelligence with the live list empty.
3. Reviewed one non-sensitive synthetic analysis and clicked the real **Save to Inbox** button. The existing `DelegatedSharePointInboxIntelligenceProvider.create()` path created item `1`, AppId `2de8fcce-431b-40e7-a012-408de7a52d5e`.
4. Initial Graph read-back: `Status=open`, `RecordVersion=1`, ETag `"6fc9ca5f-ffd0-465d-8d22-f878e5f835f1,1"`, `Created=Modified=2026-08-29T16:53:07Z`.
5. Reloaded the entire application. The item survived and reappeared in **Needs attention**, proving the durable provider was active rather than the session fallback.
6. Clicked **Mark waiting**. Graph read-back: `Status=waiting`, `RecordVersion=2`, ETag suffix `,2`, `Modified=2026-08-29T16:53:34Z`.
7. Clicked **Resolve**. Graph read-back: `Status=resolved`, `RecordVersion=3`, ETag suffix `,3`, `Modified=2026-08-29T16:53:43Z`, `ResolvedAt=2026-08-29T16:53:42Z`.
8. Clicked **Reopen**. Graph read-back: `Status=open`, `RecordVersion=4`, ETag suffix `,4`, `Modified=2026-08-29T16:53:53Z`, `ResolvedAt=null`.
9. `Created` remained exactly `2026-08-29T16:53:07Z` across every update; `Modified` advanced on every transition.
10. Action-item JSON round-tripped exactly, including action, due date `2026-09-01`, and owner `me`.
11. Extracted and matched relationship JSON round-tripped exactly: organization `FutureWorks Partnership`/`org-futureworks`, district `North Valley SD`/`org-north-valley`, and project `AI in Education`/`project-ai`; people and tag arrays also round-tripped.
12. The live field set contained no raw-email/body/thread/signature column, no Anthropic request/response field, and no API-key/credential/secret/token field. Unique synthetic markers placed beyond the permitted source excerpt were absent from every stored field.
13. Server diagnostics and the deterministic local analysis path confirmed **zero Anthropic API calls** for this verification.

## Cleanup result

The exact synthetic item was removed. A subsequent exact item lookup for list item `1` returned SharePoint `itemNotFound` (HTTP 404), and a fresh list read returned zero items. `IU_Inbox_Intelligence` is back to its pre-test empty state.

## Known limitations

1. This verification is DEV-only and single-user. It does not authorize production SharePoint cutover or replace the planned production authentication review.
2. The runtime provider intentionally exposes no delete operation because deletion is not part of the Inbox workflow; authorized synthetic cleanup therefore remains an exceptional administrative action.
3. The "Create Work Record" link-back assumes the Inbox record's captured version remains current while the Work Record wizard is open; a concurrent edit surfaces as a conflict rather than being overwritten.
4. Action-item editing supports edit/remove, not adding a new item from scratch; the Inbox list intentionally remains compact rather than adding a separate detail screen.

## Readiness

**GREEN.** The approved schema, DEV delegated provider activation, normal UI create/reload/status/reopen workflow, live Graph metadata and JSON round-trips, privacy boundaries, zero-Anthropic-call constraint, and cleanup verification all passed. No synthetic Inbox Intelligence item remains.
