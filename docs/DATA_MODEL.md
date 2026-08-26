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

- `Project`: stable ID, name, description, status, and display color.
- `Organization`: stable ID, canonical name, and real type (`district`, `partner`, or `iu`).
- `Contact`: stable ID, display name, role, and optional organization relationship. V1.1 only supplies sample references; it is not a CRM.
- `Category`: stable ID, name, and category group.
- `Deliverable`: canonical ORBIT code and label.
- `ReportingConfig`: minutes per day, July school-year boundary, and quarter month ranges.
- `SystemSettings`: current controlled activity-type vocabulary.

All are requested by the frontend through `DataProvider` methods. The seed arrays live behind the provider and are not imported by screen components.

## Validation and save ownership

`lib/validation.ts` is shared by provider and API paths. It validates required strings, canonical IDs, arrays, reach, duration, engagement scope, schema version, and nested ORBIT invariants. A create receives server timestamps and version `1`. An update supplies `expectedVersion`; successful updates increment the version, preserve `createdAt`, and receive a server-owned `modifiedAt`.
