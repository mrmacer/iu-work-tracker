# IU Work Tracker — SharePoint Integration Plan

**Status:** Final Phase 2 integration design; no SharePoint resources, Microsoft authentication, Graph calls, or persistence changes are implemented by this document.

**Baseline:** V1.1 runtime schema version `2`, the existing `DataProvider` interface, shared `validateWorkRecord` rules, and the current structured provider-result semantics.

## Design decisions

- Keep one canonical Work Record. SharePoint does not receive reporting copies or separate ORBIT records.
- Keep `DataProvider` as the frontend boundary. Screens continue to consume `WorkRecord` and the existing reference/configuration types.
- Keep the browser calling application API routes. Microsoft Graph credentials and tokens never enter browser storage.
- Store stable application IDs in ordinary text fields. SharePoint item IDs and ETags remain provider metadata.
- Use no SharePoint Lookup columns in the initial schema. Multi-value relationships remain JSON arrays of canonical IDs.
- Use seven lists. Do not create an evidence list, document library, user list, reporting-period list, or ORBIT-report list because none is required by the current runtime model.
- Keep D1 as the active production persistence provider until a separately approved cutover.

## 1. Final proposed SharePoint Lists

| List title | Runtime entity | Write ownership |
| --- | --- | --- |
| `IU_Work_Records` | `WorkRecord` | Application only during normal operation |
| `IU_Projects` | `Project` | SharePoint administrator/reference-data steward; application reads |
| `IU_Organizations` | `Organization` | SharePoint administrator/reference-data steward; application reads |
| `IU_Contacts` | `Contact` | SharePoint administrator/reference-data steward; application reads |
| `IU_Work_Categories` | `Category` | SharePoint administrator/reference-data steward; application reads |
| `IU_ORBIT_Deliverables` | `Deliverable` | SharePoint administrator/reference-data steward; application reads |
| `IU_Settings` | `ReportingConfig` and `SystemSettings` | SharePoint administrator/reference-data steward; application reads |

Use these same logical titles in separate development/test and production sites. Runtime configuration must identify lists by immutable SharePoint list GUID, not by title.

## 2. Exact purpose of each list

### `IU_Work_Records`

Stores the canonical business event represented by `WorkRecord`. History, project, organization/LEA, and ORBIT views all project from these same items. The list also stores the numeric provider version and record schema version needed by the existing contract.

### `IU_Projects`

Stores the canonical project references returned by `getProjects()`. A project is a selectable classification, not a container or duplicate record store.

### `IU_Organizations`

Stores real district, partner, and IU organizations returned by `getOrganizations()`. It must not contain pseudo-organizations for `regional` or `allDistricts` scope.

### `IU_Contacts`

Stores the lightweight contact references returned by `getContacts()`. It is not a CRM and has no contact history, ownership, or communication model.

### `IU_Work_Categories`

Stores the controlled work-area and topic categories returned by `getCategories()`.

### `IU_ORBIT_Deliverables`

Stores the canonical ORBIT deliverable code/label pairs returned by `getDeliverables()`. Work Records reference the stable code.

### `IU_Settings`

Stores exactly two keyed JSON configuration items: `ReportingConfig` and `SystemSettings`. It supplies `getReportingConfig()` and `getSystemSettings()` without introducing configuration entities that do not exist at runtime.

## 3. Column names

The names below are the required **internal names**. Create them without spaces so SharePoint does not encode or alter them. Friendly display names may be applied afterward, but code must use the internal names.

### `IU_Work_Records`

| Internal name | Required/default | Meaning |
| --- | --- | --- |
| `Title` | Required | Runtime title; built-in SharePoint Title column |
| `AppId` | Required, unique | Stable application identity |
| `ActivityDate` | Required | Business activity date |
| `ActivityType` | Required | Value validated against `SystemSettings.activityTypes` |
| `ShortDescription` | Default empty | Runtime `description` |
| `DetailedNotes` | Default empty | Runtime `detailedNotes` |
| `DurationMinutes` | Required | Authoritative total duration |
| `RecordStatus` | Required | `complete` or `draft` |
| `EngagementScope` | Required | `none`, `specific`, `regional`, or `allDistricts` |
| `ProjectIdsJson` | Required, default `[]` | JSON array of Project AppIds |
| `OrganizationIdsJson` | Required, default `[]` | JSON array of Organization AppIds |
| `ContactIdsJson` | Required, default `[]` | JSON array of Contact AppIds |
| `CategoryIdsJson` | Required, default `[]` | JSON array of Category AppIds |
| `EducatorsLeadersReach` | Required, default `0` | Reach count |
| `StudentsFamiliesReach` | Required, default `0` | Reach count |
| `WorkforceCommunityReach` | Required, default `0` | Reach count |
| `OtherReach` | Required, default `0` | Reach count |
| `EvidenceSummary` | Default empty | General evidence summary |
| `EvidenceReferenceIdsJson` | Required, default `[]` | JSON array of arbitrary stable evidence reference IDs |
| `WorkOutput` | Default empty | Runtime `output` |
| `WorkOutcome` | Default empty | Runtime `outcome` |
| `NextStep` | Default empty | Runtime `nextStep` |
| `FollowUpNeeded` | Required, default `No` | Follow-up flag |
| `FollowUpDate` | Optional | Follow-up date or null |
| `OrbitReportable` | Required, default `No` | ORBIT projection flag |
| `OrbitPrimaryDeliverableCode` | Optional | One canonical deliverable code or null |
| `OrbitSupportingCodesJson` | Required, default `[]` | JSON array of supporting deliverable codes |
| `StemPocMinutes` | Required, default `0` | Non-overlapping PoC allocation |
| `TacMinutes` | Required, default `0` | Non-overlapping TaC allocation |
| `OrbitEvidence` | Default empty | ORBIT-specific evidence |
| `SchemaVersion` | Required, default `2` | Runtime record schema version |
| `RecordVersion` | Required, default `1` | Numeric optimistic-concurrency version |
| `IsSample` | Required, default `No` | Development/sample marker |
| `Created` | SharePoint-managed | Provider creation timestamp |
| `Modified` | SharePoint-managed | Provider modification timestamp |

### Reference and settings lists

| List | Internal name | Required/default | Meaning |
| --- | --- | --- | --- |
| `IU_Projects` | `Title` | Required | Project name |
| `IU_Projects` | `AppId` | Required, unique | Stable project ID |
| `IU_Projects` | `ProjectDescription` | Default empty | Project description |
| `IU_Projects` | `ProjectStatus` | Required | `active`, `planning`, or `complete` |
| `IU_Projects` | `Color` | Required | Runtime display color token |
| `IU_Organizations` | `Title` | Required | Organization name |
| `IU_Organizations` | `AppId` | Required, unique | Stable organization ID |
| `IU_Organizations` | `OrganizationType` | Required | `district`, `partner`, or `iu` |
| `IU_Contacts` | `Title` | Required | Contact display name |
| `IU_Contacts` | `AppId` | Required, unique | Stable contact ID |
| `IU_Contacts` | `Role` | Required | Contact role text |
| `IU_Contacts` | `OrganizationAppId` | Optional | Stable organization ID or null |
| `IU_Work_Categories` | `Title` | Required | Category name |
| `IU_Work_Categories` | `AppId` | Required, unique | Stable category ID |
| `IU_Work_Categories` | `CategoryGroup` | Required | `work-area` or `topic` |
| `IU_ORBIT_Deliverables` | `Title` | Required | Deliverable label |
| `IU_ORBIT_Deliverables` | `Code` | Required, unique | Stable deliverable code |
| `IU_Settings` | `Title` | Required, unique | Exact key: `ReportingConfig` or `SystemSettings` |
| `IU_Settings` | `ValueJson` | Required | JSON object for the keyed runtime value |

`IU_Settings` contains exactly these shapes:

```json
{
  "Title": "ReportingConfig",
  "ValueJson": {
    "minutesPerReportingDay": 420,
    "schoolYearStartMonth": 7,
    "quarters": [
      { "code": "Q1", "startMonth": 7, "endMonth": 9 },
      { "code": "Q2", "startMonth": 10, "endMonth": 12 },
      { "code": "Q3", "startMonth": 1, "endMonth": 3 },
      { "code": "Q4", "startMonth": 4, "endMonth": 6 }
    ]
  }
}
```

```json
{
  "Title": "SystemSettings",
  "ValueJson": {
    "activityTypes": ["configured strings"]
  }
}
```

`ValueJson` is stored as compact JSON text, not as a SharePoint JSON object field.

## 4. Recommended SharePoint field types

| Field family | SharePoint field type | Configuration |
| --- | --- | --- |
| `Title`, `AppId`, `ActivityType`, `Color`, `Role`, `OrganizationAppId`, `Code`, `OrbitPrimaryDeliverableCode` | Single line of text | Plain text; no fill-in behavior |
| Description, notes, evidence, output/outcome/next-step fields | Multiple lines of text | Plain text; append changes off; rich text off |
| All `*Json` and `ValueJson` fields | Multiple lines of text | Plain text; append changes off; rich text off |
| `ActivityDate`, `FollowUpDate` | Date and Time | Date Only; store/return canonical `YYYY-MM-DD` |
| `DurationMinutes`, reach fields, ORBIT minute fields, `SchemaVersion`, `RecordVersion` | Number | Zero decimal places; provider enforces whole/non-negative rules and positive total duration |
| `FollowUpNeeded`, `OrbitReportable`, `IsSample` | Yes/No | Explicit defaults shown above |
| `RecordStatus` | Choice | Closed choices: `complete`, `draft`; fill-in off |
| `EngagementScope` | Choice | Closed choices: `none`, `specific`, `regional`, `allDistricts`; fill-in off |
| `ProjectStatus` | Choice | Closed choices: `active`, `planning`, `complete`; fill-in off |
| `OrganizationType` | Choice | Closed choices: `district`, `partner`, `iu`; fill-in off |
| `CategoryGroup` | Choice | Closed choices: `work-area`, `topic`; fill-in off |
| `Created`, `Modified` | Built-in SharePoint system Date/Time | Read-only through the provider |

`ActivityType` and ORBIT deliverable codes are deliberately text rather than SharePoint Choice fields. Their controlled vocabularies already live in `IU_Settings` and `IU_ORBIT_Deliverables`; duplicating those vocabularies in list-column schema would create two sources of truth.

## 5. Values that should use lookups

**None in the Phase 2 schema.** The current domain contract identifies related entities by stable string IDs. Native SharePoint Lookup fields are not required to express any current runtime value.

The provider performs logical resolution and referential validation:

- `projectIds` against `IU_Projects.AppId`
- `organizationIds` and `Contact.organizationId` against `IU_Organizations.AppId`
- `contactIds` against `IU_Contacts.AppId`
- `categoryIds` against `IU_Work_Categories.AppId`
- ORBIT codes against `IU_ORBIT_Deliverables.Code`

This is application-level resolution, not a SharePoint Lookup column.

## 6. Values that should NOT use lookups

- The four Work Record relationship arrays must not be multi-lookup fields.
- `Contact.organizationId` must remain `OrganizationAppId` text, not a lookup.
- ORBIT primary/supporting deliverables must remain stable code text/JSON, not lookup or multi-lookup fields.
- `ActivityType` must remain text validated against `SystemSettings`, not a lookup or duplicated Choice vocabulary.
- Evidence reference IDs must remain arbitrary stable text IDs. No evidence entity exists in V1.1.
- `regional` and `allDistricts` must never be represented by organization lookups or pseudo-organizations.

Reasons are contract preservation, portability, simpler Graph payloads, avoidance of lookup expansion/join limits, and prevention of SharePoint item IDs leaking into the business model. Microsoft also warns that lookup-heavy views and filters can encounter threshold constraints.

## 7. Arrays and multi-value relationships

Store each runtime array as a JSON string in a plain multiline-text column:

```json
["project-steels","project-ai"]
```

Rules:

- An empty array is exactly `[]`; null is not valid for an array field.
- Values are stable `AppId` strings, except ORBIT arrays use stable deliverable `Code` strings and evidence references use their existing arbitrary stable IDs.
- Do not store labels, names, SharePoint item IDs, or lookup objects in these arrays.
- Preserve runtime order on serialization. Consumers must not assign business meaning to that order.
- The adapter parses JSON strictly. Malformed JSON, non-string elements, duplicates, or unknown canonical IDs produce a persistence/validation failure; they are not silently discarded.
- Multi-list referential integrity is enforced by `validateWorkRecord` plus provider reference resolution, not by SharePoint.

## 8. `WorkRecord.appId` mapping

`WorkRecord.appId` maps one-to-one to `IU_Work_Records.AppId`.

- It is application-generated before create.
- It is required, unique, indexed, and immutable after create.
- It remains stable across D1, SharePoint environments, migration, export, and re-import.
- It is the record identity used inside application data and conflict recovery.
- SharePoint's list-item `id` is not copied into `appId`.
- The provider uses the SharePoint item ID for direct Graph addressing after load and verifies that the stored `AppId` still matches.

Hide `AppId` from normal SharePoint forms/views and prevent direct edits through governance. Unique-value enforcement prevents duplicates but does not itself make an existing value immutable.

## 9. `schemaVersion` storage

`WorkRecord.schemaVersion` maps to the required zero-decimal Number column `IU_Work_Records.SchemaVersion`.

- The current value is explicitly written as `2` on every create/update.
- The adapter rejects unsupported versions on read and write; it must not silently coerce them to `2`.
- `SchemaVersion` versions the Work Record payload, not the SharePoint list definition and not the provider implementation.
- Reference/configuration runtime types currently have no `schemaVersion`, so no speculative schema-version columns are added to their lists.

## 10. Provider metadata, item IDs, and ETags

| Runtime/provider value | SharePoint/Graph source | Rule |
| --- | --- | --- |
| `metadata.providerId` | `listItem.id` | Store as the Graph-returned string for the `IU_Work_Records` item |
| `metadata.version` | `fields.RecordVersion` | Positive integer; create at `1`; increment exactly once per successful provider update |
| `metadata.createdAt` | `listItem.createdDateTime` / SharePoint `Created` | Provider-owned ISO timestamp |
| `metadata.modifiedAt` | `listItem.lastModifiedDateTime` / SharePoint `Modified` | Provider-owned ISO timestamp |
| `metadata.syncState` | Not stored in SharePoint | Client/provider transient state; return `saved` after successful reads/writes |
| Graph ETag | `listItem.eTag` | Opaque, provider-private concurrency token; never parsed or exposed as a business field |

The numeric `RecordVersion` is necessary because the existing contract accepts `expectedVersion: number`; an ETag is an opaque string and cannot be placed faithfully in that field.

## 11. Optimistic concurrency with Microsoft Graph

### Create

1. Load current reference/configuration data and run shared validation.
2. Confirm no item exists for the indexed `AppId`.
3. `POST /sites/{site-id}/lists/{list-id}/items` with mapped fields, `SchemaVersion = 2`, and `RecordVersion = 1`.
4. Read the created item back with fields so the returned record uses SharePoint's item ID, timestamps, and ETag-derived state.
5. If SharePoint's unique constraint reports an `AppId` collision, load that item and return the existing structured `conflict` result.

### Update

1. Resolve the item by `metadata.providerId`; use indexed `AppId` lookup only when provider metadata is absent. Verify the stored `AppId` matches.
2. Read the current item and fields to obtain both `RecordVersion` and the current opaque ETag.
3. If `RecordVersion !== expectedVersion`, map the current item to `WorkRecord` and return `ProviderResult.conflict` without writing.
4. Validate the proposed Work Record against current reference data.
5. PATCH `/sites/{site-id}/lists/{list-id}/items/{item-id}/fields`, including `RecordVersion = expectedVersion + 1`, with `If-Match: {current eTag}`.
6. A Graph `412 Precondition Failed` means another write landed after step 2. Re-read the item and return the existing structured `conflict` result.
7. On success, read the item back and return the normalized record with the incremented numeric version and SharePoint timestamps.

This read/compare/conditional-write/read sequence preserves the numeric `DataProvider` contract while using Graph's ETag protection. It intentionally costs additional Graph reads; caching an ETag in a stateless web process is not a safe replacement.

The application must be the only normal writer to `IU_Work_Records`. A native SharePoint edit changes the ETag but does not increment the custom `RecordVersion`; without an opaque token in the current contract, such edits cannot be reconciled reliably with a previously loaded numeric version.

## 12. Created and modified timestamps

- Do not create custom `CreatedAt` or `ModifiedAt` columns.
- Ignore timestamp values supplied by a client on create/update.
- Map Graph `listItem.createdDateTime` and `lastModifiedDateTime` to ISO strings in nested provider metadata.
- On update, preserve `createdAt` by reading SharePoint's system value; do not PATCH `Created`.
- Use SharePoint/Graph time in UTC. `activityDate` and `followUpDate` remain date-only business fields and must not undergo timezone conversion.
- Prototype provider timestamps are provider metadata rather than activity history. On migration, `activityDate` is preserved exactly; SharePoint `Created`/`Modified` begin at migration/create time unless the organization explicitly requires legacy audit timestamps before provisioning.

## 13. Loading through `DataProvider`

The UI-facing interface and runtime types remain unchanged.

| `DataProvider` method | SharePoint operation | Mapping behavior |
| --- | --- | --- |
| `getWorkRecords()` | Page through all `IU_Work_Records` items with selected fields expanded | Strictly decode all columns, attach item metadata, sort by `activityDate` then created time as today, return one canonical record per `AppId` |
| `createWorkRecord(record)` | Create one `IU_Work_Records` item | Validate, map, set version `1`, normalize response |
| `updateWorkRecord(record, expectedVersion)` | Conditional update of one item | Use numeric comparison plus Graph ETag sequence in section 11 |
| `getProjects()` | Read all `IU_Projects` items | `AppId`, `Title`, description, status, color to `Project[]` |
| `getOrganizations()` | Read all `IU_Organizations` items | `AppId`, `Title`, type to `Organization[]` |
| `getContacts()` | Read all `IU_Contacts` items | `AppId`, `Title`, role, nullable `OrganizationAppId` to `Contact[]` |
| `getCategories()` | Read all `IU_Work_Categories` items | `AppId`, `Title`, group to `Category[]` |
| `getDeliverables()` | Read all `IU_ORBIT_Deliverables` items | `Code`, `Title` to `Deliverable[]` |
| `getReportingConfig()` | Read `IU_Settings` item whose `Title` is `ReportingConfig` | Strict JSON parse and exact shape validation |
| `getSystemSettings()` | Read `IU_Settings` item whose `Title` is `SystemSettings` | Strict JSON parse and exact shape validation |

Implementation boundary:

- The browser may continue using `ApiDataProvider` and the current `/api/records` behavior.
- Internal `ApiDataProvider` reference getters may call an application API instead of static seed data; callers and screens do not change.
- Server-side route code selects a D1 or SharePoint adapter from environment configuration.
- The SharePoint adapter implements the existing `DataProvider` behavior and maps Graph/network/validation/conflict failures into the current result semantics.
- Follow every Graph `@odata.nextLink`; do not assume a list response contains all items.

## 14. ORBIT nested-data mapping

The runtime nested object is flattened into the canonical Work Record item:

| Runtime ORBIT field | SharePoint column | Rule |
| --- | --- | --- |
| `orbit.reportable` | `OrbitReportable` | Boolean projection gate |
| `orbit.primaryDeliverable` | `OrbitPrimaryDeliverableCode` | Nullable stable code; required when reportable |
| `orbit.supportingDeliverables` | `OrbitSupportingCodesJson` | JSON array of distinct stable codes |
| `orbit.stemPocMinutes` | `StemPocMinutes` | Non-negative whole minutes |
| `orbit.tacMinutes` | `TacMinutes` | Non-negative whole minutes |
| `orbit.evidence` | `OrbitEvidence` | Plain multiline text |

Provider validation preserves the existing rules: a reportable record has exactly one valid primary code; primary cannot repeat as supporting; PoC and TaC are non-overlapping; their sum cannot exceed `DurationMinutes`; and non-reportable items cannot retain deliverables or allocated reporting time.

School year, quarter, reporting days, and PoC/TaC projections remain pure runtime derivations. Do not store them in SharePoint.

## 15. `engagementScope` storage

Store `engagementScope` in the closed Choice column `EngagementScope` with exact case-sensitive serialized values:

- `none`
- `specific`
- `regional`
- `allDistricts`

`OrganizationIdsJson` always contains real canonical organizations. `specific` requires at least one real organization whose type is `district`. District IDs are rejected for all other scopes. Partner and IU organization IDs may coexist with any scope. Never enumerate every district for `allDistricts` and never create regional/all-district pseudo-organizations.

## 16. Indexing recommendations

Create indexes at provisioning time, before data volume grows.

### `IU_Work_Records`

- `AppId`: unique value enforcement and index. This supports identity resolution and duplicate prevention.
- `ActivityDate`: simple index for current date/history/reporting access patterns.
- `FollowUpDate`: simple index for current follow-up access patterns.
- `ActivityDate` primary plus `OrbitReportable` secondary: compound index matching the current ORBIT-by-date projection. Do not rely on `OrbitReportable` alone as the first filter because a Boolean is low-selectivity.

### Reference/configuration lists

- `AppId`: unique/indexed on Projects, Organizations, Contacts, and Categories.
- `Code`: unique/indexed on Deliverables.
- `Title`: unique/indexed on Settings.

Do not index multiline JSON/text columns and do not add speculative indexes for fields the current provider never filters. Microsoft Graph filtering works best on indexed fields and can use only one indexed field at a time in a list-item filter. SharePoint's list-view threshold also makes selective indexed-first queries important as data grows.

## 17. Expected Microsoft Graph permissions

### Runtime application

- Microsoft Graph **application** permission: `Sites.Selected`.
- Explicit SharePoint site grant: role `write` on a dedicated IU Work Tracker site containing only these seven lists.
- Do not grant runtime `Sites.ReadWrite.All`, `Sites.Manage.All`, or `Sites.FullControl.All`.

`Sites.Selected` starts with no resource access; it requires Entra consent, an explicit permission grant to the selected site, and an access token carrying the selected scope. A dedicated site keeps the `write` grant bounded to IU Work Tracker content while avoiding seven list-specific permission breaks.

### Provisioning only

A tenant/SharePoint administrator uses a separate temporary administrative context to create the lists, columns, indexes, and the selected-site permission grant. Microsoft documents `Sites.FullControl.All` as required to grant a site-level selected permission. That elevated permission must not remain on the runtime application.

The selected `write` role is broader than the current `DataProvider` because Graph's selected roles do not provide a create/update-without-delete role. The application must expose no delete operation; the current contract has none.

## 18. Authentication approach

Use one single-tenant Microsoft Entra confidential web application registration per environment.

1. Authenticate IU users with OpenID Connect and OAuth 2.0 authorization-code flow with PKCE.
2. Restrict access through Entra enterprise-application assignment to the authorized IU user/group population; do not create a new application user entity because the runtime has none.
3. Keep the authenticated session server-side or in secure, HTTP-only session cookies.
4. Acquire the Microsoft Graph token server-to-server with the OAuth 2.0 client-credentials flow and the `Sites.Selected` application permission.
5. Prefer a certificate or federated credential in production; keep all credentials in deployment secrets, never source code or browser storage.
6. Call Graph only from authenticated application API routes. The frontend continues to call same-origin APIs through `ApiDataProvider`.

Application-only Graph access is intentional: the runtime has no author/owner field, and storage access should not depend on each user's direct SharePoint permissions. SharePoint `Created By`/`Modified By` will identify the application service principal; those fields are not part of the current runtime contract.

## 19. Development/test and production environment strategy

- Use separate SharePoint sites, list GUIDs, Entra app registrations, credentials, and configuration for development/test and production.
- Use the same seven list titles/internal column names in each site so the mapper is identical.
- Configure provider selection explicitly, for example `DATA_PROVIDER=d1|sharepoint`; defaulting silently to SharePoint is not allowed.
- Store tenant ID, client ID, credential reference, site ID, and all seven list GUIDs as environment configuration. Never discover production lists by display title on every request.
- Unit/provider tests continue to use `MemoryDataProvider` and mocked Graph responses.
- Graph integration and concurrency tests use only the dedicated development/test site and disposable records with unique AppIds. Mark development scenarios `IsSample = Yes`.
- Keep current D1 production persistence unchanged until migration verification and explicit cutover approval.
- Do not dual-write D1 and SharePoint. Dual writes would introduce ordering and reconciliation behavior absent from the current contract.
- Maintain a read-only D1 export/snapshot for rollback during the cutover window.

## 20. Migration path from current prototype records

1. Export a frozen D1 snapshot containing every Work Record column and provider metadata; retain an immutable backup.
2. Validate every source record with the current schema-version-2 validator and the exact reference/configuration seed set. Stop on any invalid or duplicate `AppId`.
3. Provision and validate the six reference/configuration lists first. Load projects, organizations, contacts, categories, deliverables, and the two settings items with their existing stable IDs/codes.
4. Verify referential closure: every canonical ID/code used by a Work Record exists; evidence reference IDs remain opaque and require no evidence list.
5. Transform each D1 row through the same explicit runtime-to-SharePoint mapper used by the provider. Do not copy D1 row IDs into `AppId` or SharePoint item IDs.
6. Create each Work Record with its original `AppId`, all business fields, `SchemaVersion = 2`, and `RecordVersion = 1`. Prototype `activityDate` and `followUpDate` are preserved exactly.
7. Exclude `isSample = true` records from production unless an explicit migration manifest authorizes individual samples. They may be migrated to development/test.
8. Record a migration manifest mapping `AppId` to the new SharePoint item ID plus success/failure state. The manifest is migration output, not a new runtime list.
9. Re-read all migrated items through the SharePoint adapter. Compare counts, AppId uniqueness, scalar values, JSON arrays, relationship resolution, total minutes, reach, scope rules, ORBIT projections, and schema versions with the frozen source.
10. Run the existing provider/validation/projection test suite plus Graph conflict tests against development/test.
11. During an approved production maintenance window, freeze D1 writes, migrate the final delta/snapshot, verify again, and switch only the provider configuration.
12. Keep the D1 snapshot read-only for rollback; do not delete it as part of the integration phase.

Provider metadata is rebound to the new provider: `providerId`, `createdAt`, and `modifiedAt` become SharePoint item metadata and the numeric version restarts at `1`. If institutional audit policy requires preservation of prototype storage timestamps, that requirement must be decided before list creation because it would require explicit legacy-audit columns not present in this plan.

## 21. Risks and SharePoint-specific constraints

| Risk/constraint | Design response |
| --- | --- |
| Graph ETags are opaque while `expectedVersion` is numeric | Keep `RecordVersion`, pre-read current ETag, and use `If-Match`; never parse an ETag |
| Native SharePoint edits do not increment `RecordVersion` | Make the app the sole normal Work Records writer; restrict normal users from direct list editing |
| Single-line SharePoint text is limited and the runtime currently validates no maximum title/AppId length | Approve and add provider validation for SharePoint-compatible limits before enabling writes; do not truncate silently |
| JSON relationship columns are not queryable joins and have no SharePoint referential integrity | Resolve/validate IDs in the provider; use indexed scalar fields for current queries |
| SharePoint list-view threshold and Graph pagination | Index current filters, follow `@odata.nextLink`, and avoid non-indexed scans where possible |
| `getWorkRecords()` currently loads the complete collection | Accept for the current contract and volume; define pagination/delta as a future contract change only when real volume requires it |
| Graph can throttle or transiently fail | Map failures to existing network/persistence results; use bounded retry with Graph guidance in implementation |
| SharePoint has no cross-list transaction for reference changes plus Work Record writes | Treat reference lists as stewarded configuration; validate before the single Work Record item write |
| Direct list edits can bypass ORBIT, scope, and duration validation | Closed Choice fields provide limited enforcement; sole-writer policy and shared application validation remain authoritative |
| Selected-site `write` permission can permit operations beyond current methods | Dedicated site, no delete API, least-privilege app code, audit, and admin-only provisioning |
| SharePoint system timestamps cannot preserve prototype storage timestamps on migration | Treat them as provider metadata, or approve legacy-audit columns before provisioning if policy requires preservation |
| Renaming columns can leave confusing display/internal-name pairs | Create exact internal names once; bind application configuration to list GUIDs and code to internal column names |
| Reference/config getter methods do not return structured `ProviderResult` values | Preserve the interface; reject failed getter promises through the existing application load path and test that behavior before cutover |

## 22. Exact implementation order for the next coding phase

1. Resolve the four pre-provisioning gates listed below: sole-writer policy, text-length limits, prototype timestamp policy, and tenant/site ownership.
2. Freeze this mapping as executable mapper/codec tests before connecting to Graph.
3. Build an idempotent provisioning definition/script for the **development/test** site only; review its dry-run output before creating resources.
4. Create the development/test Entra registration, authentication configuration, application credential, and `Sites.Selected` consent/grant.
5. Provision the seven development/test lists with exact internal names, types, choices, defaults, uniqueness, and indexes; record their GUIDs in environment configuration.
6. Seed and verify reference/configuration data using existing stable IDs/codes.
7. Implement the server-only Graph client, token acquisition/cache, paging, retry/error classification, and field selection.
8. Implement strict codecs between Graph field sets and the existing runtime types; reject malformed stored data.
9. Implement and test all reference/configuration getter methods through the existing `DataProvider` contract.
10. Implement `getWorkRecords()` with provider metadata mapping and deterministic ordering.
11. Implement create with shared validation, unique-AppId conflict mapping, server metadata, and read-back normalization.
12. Implement update with the numeric-version plus ETag algorithm and explicit `412` conflict mapping.
13. Wire server API provider selection behind explicit environment configuration; keep D1 as the default/production provider.
14. Update `ApiDataProvider` internals for reference loading without changing its public interface or any screen component.
15. Add unit tests for every field/array/null mapping, malformed JSON, Graph errors, duplicates, schema mismatch, timestamps, and metadata.
16. Add development/test Graph integration tests for create/read/update, stale numeric versions, ETag races, reference loads, paging, scope validation, and ORBIT invariants.
17. Build a dry-run migration tool and reconciliation report; execute it only against the frozen D1 export and development/test SharePoint site.
18. Complete user acceptance and end-to-end verification with SharePoint selected as the development/test provider.
19. Provision production with a separately reviewed administrative run; configure production credentials/list GUIDs without changing the active provider.
20. Execute the approved freeze, final migration/reconciliation, provider switch, smoke test, and rollback check. Do not remove D1 during this phase.

## Runtime-to-SharePoint field mapping

| Runtime field | SharePoint list | SharePoint column | Type | Notes |
| --- | --- | --- | --- | --- |
| `WorkRecord.appId` | `IU_Work_Records` | `AppId` | Single line text | Unique, indexed, immutable stable ID |
| `WorkRecord.title` | `IU_Work_Records` | `Title` | Single line text | Built-in Title column |
| `WorkRecord.activityDate` | `IU_Work_Records` | `ActivityDate` | Date Only | Preserve `YYYY-MM-DD`; no timezone shift |
| `WorkRecord.activityType` | `IU_Work_Records` | `ActivityType` | Single line text | Validated against SystemSettings |
| `WorkRecord.description` | `IU_Work_Records` | `ShortDescription` | Multiline plain text | Empty string allowed |
| `WorkRecord.detailedNotes` | `IU_Work_Records` | `DetailedNotes` | Multiline plain text | Empty string allowed |
| `WorkRecord.durationMinutes` | `IU_Work_Records` | `DurationMinutes` | Number, 0 decimals | Positive whole number |
| `WorkRecord.status` | `IU_Work_Records` | `RecordStatus` | Choice | `complete`, `draft` |
| `WorkRecord.engagementScope` | `IU_Work_Records` | `EngagementScope` | Choice | Exact four-value vocabulary |
| `WorkRecord.projectIds` | `IU_Work_Records` | `ProjectIdsJson` | Multiline plain text | JSON array of `IU_Projects.AppId` |
| `WorkRecord.organizationIds` | `IU_Work_Records` | `OrganizationIdsJson` | Multiline plain text | JSON array of real `IU_Organizations.AppId` |
| `WorkRecord.contactIds` | `IU_Work_Records` | `ContactIdsJson` | Multiline plain text | JSON array of `IU_Contacts.AppId` |
| `WorkRecord.categoryIds` | `IU_Work_Records` | `CategoryIdsJson` | Multiline plain text | JSON array of `IU_Work_Categories.AppId` |
| `reach.educatorsLeaders` | `IU_Work_Records` | `EducatorsLeadersReach` | Number, 0 decimals | Non-negative whole number |
| `reach.studentsFamilies` | `IU_Work_Records` | `StudentsFamiliesReach` | Number, 0 decimals | Non-negative whole number |
| `reach.workforceCommunity` | `IU_Work_Records` | `WorkforceCommunityReach` | Number, 0 decimals | Non-negative whole number |
| `reach.other` | `IU_Work_Records` | `OtherReach` | Number, 0 decimals | Non-negative whole number |
| `WorkRecord.evidenceSummary` | `IU_Work_Records` | `EvidenceSummary` | Multiline plain text | No evidence entity is introduced |
| `WorkRecord.evidenceReferenceIds` | `IU_Work_Records` | `EvidenceReferenceIdsJson` | Multiline plain text | JSON array of opaque stable IDs |
| `WorkRecord.output` | `IU_Work_Records` | `WorkOutput` | Multiline plain text | Empty string allowed |
| `WorkRecord.outcome` | `IU_Work_Records` | `WorkOutcome` | Multiline plain text | Empty string allowed |
| `WorkRecord.nextStep` | `IU_Work_Records` | `NextStep` | Multiline plain text | Empty string allowed |
| `WorkRecord.followUpNeeded` | `IU_Work_Records` | `FollowUpNeeded` | Yes/No | Provider writes explicit Boolean |
| `WorkRecord.followUpDate` | `IU_Work_Records` | `FollowUpDate` | Date Only | Nullable |
| `orbit.reportable` | `IU_Work_Records` | `OrbitReportable` | Yes/No | Gates ORBIT projection |
| `orbit.primaryDeliverable` | `IU_Work_Records` | `OrbitPrimaryDeliverableCode` | Single line text | Nullable code from Deliverables list |
| `orbit.supportingDeliverables` | `IU_Work_Records` | `OrbitSupportingCodesJson` | Multiline plain text | JSON code array |
| `orbit.stemPocMinutes` | `IU_Work_Records` | `StemPocMinutes` | Number, 0 decimals | Non-negative allocation |
| `orbit.tacMinutes` | `IU_Work_Records` | `TacMinutes` | Number, 0 decimals | Non-negative allocation |
| `orbit.evidence` | `IU_Work_Records` | `OrbitEvidence` | Multiline plain text | Empty string allowed |
| `WorkRecord.schemaVersion` | `IU_Work_Records` | `SchemaVersion` | Number, 0 decimals | Current supported value `2` |
| `metadata.providerId` | `IU_Work_Records` | SharePoint item `id` | Graph system string | Not a custom column |
| `metadata.version` | `IU_Work_Records` | `RecordVersion` | Number, 0 decimals | Numeric DataProvider version |
| `metadata.createdAt` | `IU_Work_Records` | `Created` | System Date/Time | From `createdDateTime` |
| `metadata.modifiedAt` | `IU_Work_Records` | `Modified` | System Date/Time | From `lastModifiedDateTime` |
| `metadata.syncState` | — | — | Provider-local | Not persisted |
| Graph concurrency token | `IU_Work_Records` | List-item ETag | Graph system string | Provider-private; sent in `If-Match` |
| `WorkRecord.isSample` | `IU_Work_Records` | `IsSample` | Yes/No | Excluded from production migration by default |
| `Project.appId` | `IU_Projects` | `AppId` | Single line text | Unique/indexed |
| `Project.name` | `IU_Projects` | `Title` | Single line text | Display name |
| `Project.description` | `IU_Projects` | `ProjectDescription` | Multiline plain text | Empty string allowed |
| `Project.status` | `IU_Projects` | `ProjectStatus` | Choice | Exact runtime vocabulary |
| `Project.color` | `IU_Projects` | `Color` | Single line text | Runtime display token |
| `Organization.appId` | `IU_Organizations` | `AppId` | Single line text | Unique/indexed |
| `Organization.name` | `IU_Organizations` | `Title` | Single line text | Canonical name |
| `Organization.type` | `IU_Organizations` | `OrganizationType` | Choice | `district`, `partner`, `iu` |
| `Contact.appId` | `IU_Contacts` | `AppId` | Single line text | Unique/indexed |
| `Contact.displayName` | `IU_Contacts` | `Title` | Single line text | Display name |
| `Contact.role` | `IU_Contacts` | `Role` | Single line text | Lightweight reference text |
| `Contact.organizationId` | `IU_Contacts` | `OrganizationAppId` | Single line text | Nullable stable ID; not lookup |
| `Category.appId` | `IU_Work_Categories` | `AppId` | Single line text | Unique/indexed |
| `Category.name` | `IU_Work_Categories` | `Title` | Single line text | Display name |
| `Category.group` | `IU_Work_Categories` | `CategoryGroup` | Choice | `work-area`, `topic` |
| `Deliverable.code` | `IU_ORBIT_Deliverables` | `Code` | Single line text | Unique/indexed stable code |
| `Deliverable.label` | `IU_ORBIT_Deliverables` | `Title` | Single line text | Display label |
| `ReportingConfig` | `IU_Settings` | `ValueJson` where `Title=ReportingConfig` | Multiline plain text | Exact JSON object; strict parse |
| `SystemSettings` | `IU_Settings` | `ValueJson` where `Title=SystemSettings` | Multiline plain text | Exact JSON object; strict parse |
| Derived school year | — | — | Derived | Compute from `ActivityDate` and ReportingConfig |
| Derived quarter | — | — | Derived | Compute from `ActivityDate` and ReportingConfig |
| Derived reporting days | — | — | Derived | Compute from PoC minutes and ReportingConfig |

## Remaining architectural concerns before list creation

The schema is ready to provision only after owners explicitly resolve these gates:

1. **Sole-writer policy:** confirm that ordinary users will not edit `IU_Work_Records` directly in SharePoint. The application-only Graph design and numeric-version/ETag algorithm depend on this. Emergency administrator corrections must use a controlled process that also increments `RecordVersion`.
2. **SharePoint text limits:** approve SharePoint-compatible maximum lengths—especially the 255-character single-line limit—and add matching shared/provider validation in the next coding phase. Silent truncation is prohibited.
3. **Prototype timestamp policy:** confirm that prototype `createdAt`/`modifiedAt` values are provider metadata and may rebind to SharePoint migration timestamps. If they are legally/audit-significant, legacy-audit columns must be designed before provisioning.
4. **Tenant ownership:** identify the production tenant, dedicated development/test and production sites, site owners, reference-data stewards, enterprise-app assignment group, retention policy, and administrator who will approve the selected-site grant. These are tenant governance decisions, not runtime-model fields.

No other runtime-model expansion is required before SharePoint list creation.

## Microsoft references used for this design

- [Update a SharePoint list item with `If-Match`](https://learn.microsoft.com/en-us/graph/api/listitem-update?view=graph-rest-1.0)
- [Microsoft Graph `listItem` resource: ID, ETag, and timestamps](https://learn.microsoft.com/en-us/graph/api/resources/listitem?view=graph-rest-1.0)
- [List SharePoint items, field expansion, indexed filtering, and paging](https://learn.microsoft.com/en-us/graph/api/listitem-list?view=graph-rest-1.0)
- [Selected permissions for SharePoint and OneDrive](https://learn.microsoft.com/en-us/graph/permissions-selected-overview)
- [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Microsoft identity authorization-code flow with PKCE](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)
- [Microsoft identity client-credentials flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow)
- [SharePoint list-view threshold and indexing guidance](https://support.microsoft.com/en-us/sharepoint/data-and-lists/working-with-the-list-view-threshold-limit-for-all-versions-of-sharepoint)
