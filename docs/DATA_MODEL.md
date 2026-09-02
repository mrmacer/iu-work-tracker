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
- `Contact`: stable ID, display name, role, and optional organization relationship. V1.1 only supplies sample references; it is not a CRM.
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

## Durable Project (Patch 7 / 7B — GREEN, live-verified against real DEV SharePoint, uncommitted)

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
