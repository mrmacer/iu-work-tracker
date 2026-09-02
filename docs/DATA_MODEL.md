# IU Work Tracker Data Model — V1.1

## Relationship overview

```text
Project 0..* <——> 0..* WorkRecord 0..* <——> Organization
                              | 0..* <——> Contact
                              | 0..* <——> Category
                              └ 0..1 ——> ORBIT classification

Reporting configuration ——> derived school year, quarter, and reporting days
Provider metadata ——> storage identity, version, timestamps, and sync state
```

`WorkRecord` is the single business event. Today, History, Projects, organization/LEA context, and ORBIT projections all read that same record; the application does not create reporting copies.

## Runtime WorkRecord

| Concern | Runtime fields and rules |
| --- | --- |
| Identity | `appId`, `title`, `activityDate`, explicit `activityType`, `status` |
| Description | `description`, `detailedNotes`, `output`, `outcome`, `nextStep` |
| Time | `durationMinutes` is the authoritative total activity duration and must be a positive whole number |
| Engagement | `engagementScope`: `none`, `specific`, `regional`, or `allDistricts` |
| Relationships | canonical `projectIds`, `organizationIds`, `contactIds`, and `categoryIds` arrays |
| Reach | four non-negative whole-number counts |
| General evidence | `evidenceSummary` and arbitrary stable `evidenceReferenceIds` |
| Follow-up | `followUpNeeded` and optional `followUpDate` |
| ORBIT | optional `orbit` object with `reportable`, one primary deliverable, supporting deliverables, PoC minutes, TaC minutes, and ORBIT evidence |
| Evolution | `schemaVersion`, currently `2` |
| Prototype marker | `isSample`; seeded scenarios remain clearly marked development data |
| Persistence | nested `metadata` with provider ID, optimistic version, server timestamps, and sync state |

Provider metadata is deliberately nested. SharePoint item IDs, ETags, D1 row IDs, and synchronization state do not become business classifications.

## Engagement scope

- `none`: no district/LEA audience is implied. IU or partner organizations may still be related.
- `specific`: at least one real organization of type `district` is required.
- `regional`: regional audience; district IDs are not attached merely to express the scope.
- `allDistricts`: all districts in the IU region; no fake or enumerated “all districts” organization is created.

District IDs are rejected outside `specific`. Partner and IU organizations may coexist with any scope because they describe real relationships rather than LEA audience scope.

The V1 regional development sample previously used `org-regional`. Migration `drizzle/0001_free_corsair.sql` deterministically changes that exact stored relationship to `engagement_scope = 'regional'` and `organization_ids_json = '[]'`. The pseudo-organization is absent from reference data.

## Reference and configuration entities

- `Project`: stable ID, name, description, status (`planning`/`active`/`paused`/`complete`), and display color — optionally durable as of Patch 7 (see below).
- `Organization`: stable ID, canonical name, and real type (`district`, `partner`, or `iu`).
- `Contact`: stable ID, display name, optional role, and optional organization relationship — optionally durable as of Patch 8B (see below). Not a CRM.
- `Category`: stable ID, name, and category group.
- `Deliverable`: canonical ORBIT code and label.
- `ReportingConfig`: minutes per day, July school-year boundary, and quarter month ranges.
- `SystemSettings`: current controlled activity-type vocabulary.

All are requested by the frontend through `DataProvider` methods. The seed arrays live behind the provider and are not imported by screen components.

## Validation and save ownership

`lib/validation.ts` is shared by provider and API paths. It validates required strings, canonical IDs, arrays, reach, duration, engagement scope, schema version, and nested ORBIT invariants. A create receives server timestamps and version `1`. An update supplies `expectedVersion`; successful updates increment the version, preserve `createdAt`, and receive a server-owned `modifiedAt`.

## MeetingRecord (Patch 6B — durable, SharePoint list provisioned and live)

A second durable resource, independent of `WorkRecord`, following the identical `DataProvider`-style boundary and optimistic-concurrency discipline described above. See `docs/AI_HANDOFF.md` "Meeting Notes durability (Patch 6B)" for the full design rationale.

| Concern | Runtime fields and rules |
| --- | --- |
| Identity | `appId`, `schemaVersion` (currently `1`) |
| Meeting details | `title`, `meetingDate`, `meetingType`, `attendeesText` (plain text — no entity resolution) |
| Source content | `agendaText`, `notesText` — the same free text the AI analysis pipeline reads, now durable |
| Reviewed intelligence | `reviewedCandidates: ReviewedMeetingCandidate[]` — **one JSON blob**, not per-type child lists. Each item is a `MeetingCandidate` (`type`, `title`, `detail`, `sourceExcerpt`, `ownerText`, `dueText`, `durationText`) plus the human's `selected` decision. This is always the CURRENT human-reviewed state, never the original AI output. |
| Minutes | `minutesText` — deterministic output of `buildDraftMinutes()` at save time; no AI call |
| Analysis provenance | `analysisModel`, `analyzedAt` — both `null` (no analysis has run yet) or both set together; never independently |
| Persistence | the same nested `metadata` (`ProviderMetadata`) as `WorkRecord` — provider ID, numeric version, SharePoint's own Created/Modified, sync state |

Relationships intentionally NOT present: no `projectIds`/`organizationIds`/`contactIds`/`categoryIds`, no district/people entity resolution on `attendeesText`. No delete operation exists for this resource in Patch 6B — `MeetingRecordProvider` exposes `list`/`create`/`update` only.

`MeetingRecord` and `WorkRecord` persistence are deliberately uncoupled: saving a meeting never creates or updates a `WorkRecord`, and the "Log as work" handoff (`buildWorkRecordDraftFromMeetingCandidate()`) never creates or updates a `MeetingRecord`.

## Durable Project (Patch 7 / 7B — GREEN, live-verified against real DEV SharePoint, committed)

Extends the existing `Project` reference-data type in place (`lib/models.ts`) rather than introducing a parallel model — see `docs/AI_HANDOFF.md` "Durable Projects (Patch 7 / 7B)" for the full design rationale. Durably persisted to the existing `IU_Projects` SharePoint list (extended in place, not replaced) via `NEXT_PUBLIC_SHAREPOINT_IU_PROJECTS_LIST_ID` — the single authoritative Project-list configuration; no second list or env var exists.

| Concern | Runtime fields and rules |
| --- | --- |
| Identity | `appId`, `name` |
| Description | `description` |
| Status | `status`: `planning`, `active`, `paused` (new in Patch 7), or `complete` — stored in `IU_Projects`'s existing `ProjectStatus` Choice column (kept as Choice, never converted to text; `"paused"` added as a fourth allowed value) |
| Visual | `color` — one of the five existing `project-mark` tokens; assigned deterministically on create, no color picker in the UI |
| Timeline | `startDate`, `targetDate` — both optional, `null` when unset |
| STEM/ORBIT | `stemOrbit` — optional boolean; a manual flag only, never AI-derived, never auto-classifying any Work Record |
| Persistence | optional nested `metadata` (`ProviderMetadata`) — **present only for a durable project** (created/loaded through `ProjectProvider`); **absent for the five seeded reference-data projects**, which is exactly how the UI decides whether to offer "Edit" |

No record count or duration field exists on `Project` — the Projects screen always derives both by filtering `WorkRecord[]` on `projectIds.includes(project.appId)` and summing `durationMinutes`, exactly as it did before this patch. No delete operation exists for this resource — `ProjectProvider` exposes `list`/`create`/`update` only. The five seeded projects (`project-steels`, `project-ai`, `project-keystone`, `project-ecosystem`, `project-makerspace`) are not migrated into the durable store in this patch and remain exactly as they were.

`DataProvider.setDurableProjects()` (`lib/data-provider.ts`) is the one deliberate coupling point: it keeps `createWorkRecord`/`updateWorkRecord`'s `projectIds` validation (`lib/validation.ts`) aware of durable projects the UI offers, without `WorkRecord` and `Project` persistence otherwise depending on each other in any way — creating or editing a Project never creates or updates a `WorkRecord`, and logging work never creates or updates a `Project`.

## Durable Contact (Patch 8B — GREEN, live-verified against real DEV SharePoint, uncommitted)

Extends the existing `Contact` reference-data type in place (`lib/models.ts`) rather than introducing a parallel model — see `docs/AI_HANDOFF.md` "Durable Contacts (Patch 8B)" for the full design rationale. Durably persisted to the existing `IU_Contacts` SharePoint list (extended in place, not replaced) via `NEXT_PUBLIC_SHAREPOINT_IU_CONTACTS_LIST_ID` — the single authoritative Contact-list configuration; no second list or env var exists.

| Concern | Runtime fields and rules |
| --- | --- |
| Identity | `appId` |
| Identity (display) | `displayName` |
| Role | `role` — optional (relaxed from required in Patch 8B) |
| Organization | `organizationId: string \| null` — singular, nullable, unchanged shape; points at the existing `Organization` reference set (District is `Organization.type === "district"`, not a separate model) |
| Contact evidence | `email` — optional; matching evidence only, never the Contact's identity (`appId` remains identity) |
| Status | `status`: `active`, `developing`, `occasional`, `dormant`, or `archived` — required, stored in `IU_Contacts`'s `Status` column as plain text (not a SharePoint Choice column); manually-created Contacts default to `active` |
| Notes | `notes` — optional, small manual relationship note; not a biography, not a Knowledge repository |
| Persistence | optional nested `metadata` (`ProviderMetadata`) — **present only for a durable contact** (created/loaded through `ContactProvider`); **absent for the three seeded reference-data contacts**, which is exactly how the UI decides whether to offer "Edit" |

No delete operation exists for this resource — `ContactProvider` exposes `list`/`create`/`update` only; `archived` status is the V1 substitute for deletion. No `projectIds` field exists on `Contact`, and no `contactIds` field exists on `Project` — a Contact's connected Projects derive transitively through `WorkRecord.contactIds` → `WorkRecord.projectIds` (see "Contact connected work" below), never a direct relationship. No connected-work totals, timelines, or "Last Interaction" are stored on `Contact`. The three seeded contacts (`contact-north-valley-lead`, `contact-futureworks`, `contact-iu-colleague`) are not migrated into the durable store and remain exactly as they were, now with `status: "active"` added for type compatibility.

`DataProvider.setDurableContacts()` (`lib/data-provider.ts`) is the one deliberate coupling point, mirroring `setDurableProjects()` exactly: it keeps `createWorkRecord`/`updateWorkRecord`'s `contactIds` validation (`lib/validation.ts`) aware of durable contacts the UI offers, without `WorkRecord` and `Contact` persistence otherwise depending on each other in any way — creating or editing a Contact never creates or updates a `WorkRecord`, and logging work never creates or updates a `Contact`. `WorkRecord.contactIds: string[]` itself is unchanged from before this patch — it already existed, was already validated/serialized, and remains the sole Contact relationship anywhere in the codebase.

## Contact connected work (Patch 8C — derived, never persisted)

Contact Detail (`app/IUWorkTracker.tsx`'s `ContactDetail`, opened from the Contacts directory) answers "what work connects to this person" entirely by deriving from records already loaded elsewhere — it adds **zero** fields to `Contact` and makes **zero** additional Graph/SharePoint requests. The whole derivation lives in `lib/contact-relationships.ts` (`buildContactRelationshipSummary()`), a pure function mirroring `lib/inbox-action-center.ts`'s role for the Home Action Center: no I/O, no AI, in-memory only, over the same `WorkRecord[]`/`Project[]` arrays the rest of the screen already has.

The relationship chain is exactly `Contact.appId → WorkRecord.contactIds → WorkRecord.projectIds → Project` — nothing else:

- **Connected Work Records**: every `WorkRecord` whose `contactIds` array includes the Contact's `appId`, sorted newest-first by `WorkRecord.activityDate` (the same business-date field and sort direction History already uses — `records.sort((a, b) => b.activityDate.localeCompare(a.activityDate))`). Never inferred from a name, email, organization, or free-text mention — only the explicit `contactIds` array counts.
- **Connected Projects**: the union of `projectIds` across those connected Work Records, deduplicated by `appId` and resolved against the current merged (seeded + durable) Project reference set. A historical `projectId` that no longer resolves is silently skipped — never fabricated, never a crash.
- **Total Connected Time / Work Record count**: `durationMinutes` summed, and a plain count, across the connected Work Records only.
- **Last Interaction**: the `activityDate` of the newest connected Work Record — explicitly **not** `ProviderMetadata.createdAt`/`modifiedAt` (SharePoint Created/Modified describe persistence timing, not interaction timing) and not "today". `null` (rendered as "No recorded interaction yet.") when there is no connected Work Record.

None of these values are ever written back to `Contact` or anywhere else — they are recomputed on every render from `records`/`projects` already held in `IUWorkTracker`'s top-level state. Recent Work is capped to the 5 most recent connected Work Records for display; clicking a row opens the existing Log Work edit dialog for that record (no new Work Record CRUD surface). Contact Detail works identically for durable and seed Contacts, and for archived Contacts when explicitly opened.

## Intelligence Contact matching (Patch 8D — deterministic, human-reviewed)

**AI may propose. The human decides.** Intelligence workflows detect people as plain evidence strings (`EmailAnalysis.people: string[]`, a Voice `PERSON` candidate's `title`); `lib/contact-matching.ts`'s `matchContactCandidates()` deterministically classifies how strong the evidence is against the current merged (seed + durable) Contact set — it never calls AI, never touches the network, and never itself decides anything. Matching strength, strongest to weakest:

- **`strong`** — exact normalized email match (trim + lowercase, reusing `normalizeContactEmail` from `lib/contact-provider.ts`).
- **`useful`** — exact normalized full name match AND the candidate Contact's `organizationId` is among organization/district IDs already resolved as relevant to the same Intelligence item (for Inbox: `resolveEmailAnalysisEntities()`'s own `organizationIds ∪ districtIds`).
- **`review`** — exact normalized full name match with no organization evidence.
- **`none`** — anything weaker, including a bare first name ("Annie" never matches "Annie Milewski", no matter how few Contacts share it) — rendered as "No reliable match found." Archived Contacts are never excluded from matching; `status` describes the durable relationship, not identity-eligibility.

No candidate at any strength — including `strong` — is ever auto-committed. The shared review component (`app/ContactMatchPanel.tsx`) always requires an explicit "Match Existing" / "Add Person" / "Ignore" click, and "Add Person" reuses the exact Patch 8B Contact creation path (`app/ContactFormModal.tsx`, `ContactProvider.create()`) — no second Contact form, no second provider, no direct SharePoint write from Intelligence UI. Matching an existing Contact never updates it — the Intelligence relationship and the Contact profile stay separate concerns.

**Persistence — Inbox Intelligence only.** `InboxIntelligenceRecord` gains `matchedContactIds: string[]`, alongside its existing `matchedOrganizationIds`/`matchedDistrictIds`/`matchedProjectIds` (`SharePoint MatchedContactIdsJson` column, same JSON-array-of-appId convention). Unlike those three siblings — silently auto-resolved by exact name match inside `buildInboxIntelligenceRecord()` at save time — `matchedContactIds` is populated only from the human's explicit review decisions in the People section, passed into `buildInboxIntelligenceRecord()` as an argument. Voice Intelligence has no durable persistence of any kind (see `app/VoiceIntelligence.tsx`), so its PERSON-candidate matching is real but entirely transient — a review decision, and any Contact created via "Add Person", lives only for that browser session; the created Contact itself IS durable (same creation path), but the fact that this transcript's candidate matched it is not. Meeting Intelligence has neither a `PERSON` candidate type nor any entity-matching precedent (`attendeesText`/`ownerText` remain plain free text) — Contact linkage there is explicitly deferred, not attempted in this patch.
