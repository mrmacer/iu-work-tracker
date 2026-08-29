# IU Work Tracker — SharePoint Provisioning Specification

**Status:** Phase 2A specification only. This document does not authorize or perform list creation, authentication, permission grants, data seeding, migration, or persistence cutover.

**Approved baseline:** The application is the sole normal writer to Work Records; SharePoint `Created` and `Modified` are canonical; IU owns and governs the site and data; normal application access is assigned through a dedicated Entra security group where practical; retention follows existing IU Microsoft 365 policy.

**Provisioning target:** A dedicated SharePoint site per environment, using generic lists and the Microsoft Graph `v1.0` API. The application continues to use the existing runtime models and `DataProvider` contract.

## Conventions

- **Internal name** is the API-facing column name. Create custom columns with the exact internal name shown, without spaces. It must never be renamed or inferred from a later display name.
- **Display name** is the human-facing SharePoint label.
- **No default** means SharePoint must not synthesize a value. The application supplies the value or maps an absent optional field to the runtime empty string/null specified below.
- **App maximum** is enforced by shared/application validation even if SharePoint accepts more characters. Values must never be truncated.
- Multiline text is plain text with rich text and append-changes disabled.
- All Number fields use zero decimal places. Whole-number, positive, non-negative, and cross-field rules remain application validation responsibilities.
- All Choice fields are closed; fill-in choices are disabled.
- No custom Lookup, Person, Calculated, Managed Metadata, attachment, or multi-choice columns are part of this schema.
- Custom technical columns marked **Hidden** are still selected explicitly by the Graph adapter.

## 1. Exact seven SharePoint lists

| Provisioning order | List display title | Expected list name | Template | Purpose |
| ---: | --- | --- | --- | --- |
| 1 | `IU_Organizations` | `IU_Organizations` | `genericList` | Canonical district, partner, and IU organizations |
| 2 | `IU_Projects` | `IU_Projects` | `genericList` | Canonical project references |
| 3 | `IU_Contacts` | `IU_Contacts` | `genericList` | Lightweight contacts with optional organization AppId |
| 4 | `IU_Work_Categories` | `IU_Work_Categories` | `genericList` | Work-area and topic categories |
| 5 | `IU_ORBIT_Deliverables` | `IU_ORBIT_Deliverables` | `genericList` | Canonical ORBIT code/label vocabulary |
| 6 | `IU_Settings` | `IU_Settings` | `genericList` | Reporting configuration and activity-type vocabulary |
| 7 | `IU_Work_Records` | `IU_Work_Records` | `genericList` | Canonical Work Records and provider version data |

Microsoft Graph creates a generic list from `displayName`; its returned list GUID is the authoritative configuration key. Capture each GUID immediately. Runtime code must address lists by GUID, not title or expected name.

For all seven lists:

- Content approval/moderation: off.
- Attachments: off; evidence bytes are not part of the runtime model.
- Folders: not used.
- List item version history: retain the site/tenant default; it is audit history, not `RecordVersion`.
- Permissions: inherit from the dedicated site. Do not create per-item permissions.
- Retention: inherit/apply existing IU Microsoft 365 policy; do not create a tracker-specific retention policy in this phase.

## 2. Exact internal and display names

The seven matrices below are authoritative for custom-column internal names and human-facing display names. They also consolidate the remaining column settings so an operator can verify one row at a time.

### 2.1 `IU_Work_Records`

| Internal name | Display name | SharePoint / Graph type | Required | Default | App maximum | Index | Unique | UI |
| --- | --- | --- | --- | --- | ---: | --- | --- | --- |
| `Title` | Work title | Single line text / `text` | Yes | No default | 255 | No | No | Shown |
| `AppId` | Application ID | Single line text / `text` | Yes | No default | 255 | Yes | Yes | Hidden |
| `ActivityDate` | Activity date | Date and Time / `dateTime`, Date Only | Yes | No default | — | Yes | No | Shown |
| `ActivityType` | Activity type | Single line text / `text` | Yes | No default | 255 | No | No | Shown |
| `ShortDescription` | Short description | Multiple lines plain text / `text` | No | No default; runtime `""` | 1,000 | No | No | Shown |
| `DetailedNotes` | Detailed notes | Multiple lines plain text / `text` | No | No default; runtime `""` | 10,000 | No | No | Shown |
| `DurationMinutes` | Duration minutes | Number / `number`, 0 decimals | Yes | No default | — | No | No | Shown |
| `RecordStatus` | Record status | Choice / `choice` | Yes | `complete` | — | No | No | Shown |
| `EngagementScope` | Engagement scope | Choice / `choice` | Yes | `none` | — | No | No | Shown |
| `ProjectIdsJson` | Project IDs (JSON) | Multiple lines plain text / `text` | Yes | `[]` | 10,000 | No | No | Hidden |
| `OrganizationIdsJson` | Organization IDs (JSON) | Multiple lines plain text / `text` | Yes | `[]` | 10,000 | No | No | Hidden |
| `ContactIdsJson` | Contact IDs (JSON) | Multiple lines plain text / `text` | Yes | `[]` | 10,000 | No | No | Hidden |
| `CategoryIdsJson` | Category IDs (JSON) | Multiple lines plain text / `text` | Yes | `[]` | 10,000 | No | No | Hidden |
| `EducatorsLeadersReach` | Educators / leaders reach | Number / `number`, 0 decimals | Yes | `0` | — | No | No | Shown |
| `StudentsFamiliesReach` | Students / families reach | Number / `number`, 0 decimals | Yes | `0` | — | No | No | Shown |
| `WorkforceCommunityReach` | Workforce / community reach | Number / `number`, 0 decimals | Yes | `0` | — | No | No | Shown |
| `OtherReach` | Other reach | Number / `number`, 0 decimals | Yes | `0` | — | No | No | Shown |
| `EvidenceSummary` | Evidence summary | Multiple lines plain text / `text` | No | No default; runtime `""` | 5,000 | No | No | Shown |
| `EvidenceReferenceIdsJson` | Evidence reference IDs (JSON) | Multiple lines plain text / `text` | Yes | `[]` | 10,000 | No | No | Hidden |
| `WorkOutput` | Output | Multiple lines plain text / `text` | No | No default; runtime `""` | 5,000 | No | No | Shown |
| `WorkOutcome` | Outcome | Multiple lines plain text / `text` | No | No default; runtime `""` | 5,000 | No | No | Shown |
| `NextStep` | Next step | Multiple lines plain text / `text` | No | No default; runtime `""` | 1,000 | No | No | Shown |
| `FollowUpNeeded` | Follow-up needed | Yes/No / `boolean` | Yes | `false` | — | No | No | Shown |
| `FollowUpDate` | Follow-up date | Date and Time / `dateTime`, Date Only | No | No default; runtime `null` | — | Yes | No | Shown |
| `OrbitReportable` | ORBIT reportable | Yes/No / `boolean` | Yes | `false` | — | Compound secondary | No | Shown |
| `OrbitPrimaryDeliverableCode` | ORBIT primary deliverable code | Single line text / `text` | No | No default; runtime `null` | 255 | No | No | Shown |
| `OrbitSupportingCodesJson` | ORBIT supporting codes (JSON) | Multiple lines plain text / `text` | Yes | `[]` | 10,000 | No | No | Hidden |
| `StemPocMinutes` | STEM PoC minutes | Number / `number`, 0 decimals | Yes | `0` | — | No | No | Shown |
| `TacMinutes` | TaC minutes | Number / `number`, 0 decimals | Yes | `0` | — | No | No | Shown |
| `OrbitEvidence` | ORBIT evidence | Multiple lines plain text / `text` | No | No default; runtime `""` | 5,000 | No | No | Shown |
| `SchemaVersion` | Schema version | Number / `number`, 0 decimals | Yes | `2` | — | No | No | Hidden |
| `RecordVersion` | Record version | Number / `number`, 0 decimals | Yes | `1` | — | No | No | Hidden |
| `IsSample` | Development/sample record | Yes/No / `boolean` | Yes | `false` | — | No | No | Shown |

Choice definitions:

| Column | Ordered choices | Fill-in |
| --- | --- | --- |
| `RecordStatus` | `complete`, `draft` | Off |
| `EngagementScope` | `none`, `specific`, `regional`, `allDistricts` | Off |

Indexes:

- Simple unique index: `AppId` (`indexed = true`, `enforceUniqueValues = true`).
- Simple index: `ActivityDate` (`indexed = true`).
- Simple index: `FollowUpDate` (`indexed = true`).
- Compound index in SharePoint List Settings: primary `ActivityDate`, secondary `OrbitReportable`.
- Do not index multiline text/JSON, long text, reach, or version columns.

### 2.2 `IU_Organizations`

| Internal name | Display name | SharePoint / Graph type | Required | Default | App maximum | Index | Unique |
| --- | --- | --- | --- | --- | ---: | --- | --- |
| `Title` | Organization name | Single line text / `text` | Yes | No default | 255 | No | No |
| `AppId` | Application ID | Single line text / `text` | Yes | No default | 255 | Yes | Yes |
| `OrganizationType` | Organization type | Choice / `choice` | Yes | No default | — | No | No |

`OrganizationType` choices: `district`, `partner`, `iu`; fill-in off. Do not create regional or all-district pseudo-organizations.

### 2.3 `IU_Projects`

| Internal name | Display name | SharePoint / Graph type | Required | Default | App maximum | Index | Unique |
| --- | --- | --- | --- | --- | ---: | --- | --- |
| `Title` | Project name | Single line text / `text` | Yes | No default | 255 | No | No |
| `AppId` | Application ID | Single line text / `text` | Yes | No default | 255 | Yes | Yes |
| `ProjectDescription` | Project description | Multiple lines plain text / `text` | No | No default; runtime `""` | 1,000 | No | No |
| `ProjectStatus` | Project status | Choice / `choice` | Yes | No default | — | No | No |
| `Color` | Display color token | Single line text / `text` | Yes | No default | 255 | No | No |

`ProjectStatus` choices: `active`, `planning`, `complete`; fill-in off.

### 2.4 `IU_Contacts`

| Internal name | Display name | SharePoint / Graph type | Required | Default | App maximum | Index | Unique |
| --- | --- | --- | --- | --- | ---: | --- | --- |
| `Title` | Contact display name | Single line text / `text` | Yes | No default | 255 | No | No |
| `AppId` | Application ID | Single line text / `text` | Yes | No default | 255 | Yes | Yes |
| `Role` | Role | Single line text / `text` | Yes | No default | 255 | No | No |
| `OrganizationAppId` | Organization application ID | Single line text / `text` | No | No default; runtime `null` | 255 | No | No |

`OrganizationAppId` is ordinary text and must resolve to `IU_Organizations.AppId` when non-null. It is not a SharePoint Lookup.

### 2.5 `IU_Work_Categories`

| Internal name | Display name | SharePoint / Graph type | Required | Default | App maximum | Index | Unique |
| --- | --- | --- | --- | --- | ---: | --- | --- |
| `Title` | Category name | Single line text / `text` | Yes | No default | 255 | No | No |
| `AppId` | Application ID | Single line text / `text` | Yes | No default | 255 | Yes | Yes |
| `CategoryGroup` | Category group | Choice / `choice` | Yes | No default | — | No | No |

`CategoryGroup` choices: `work-area`, `topic`; fill-in off.

### 2.6 `IU_ORBIT_Deliverables`

| Internal name | Display name | SharePoint / Graph type | Required | Default | App maximum | Index | Unique |
| --- | --- | --- | --- | --- | ---: | --- | --- |
| `Title` | Deliverable label | Single line text / `text` | Yes | No default | 255 | No | No |
| `Code` | Deliverable code | Single line text / `text` | Yes | No default | 255 | Yes | Yes |

The initial codes are `A` through `G` with the labels already defined in `docs/ORBIT_MAPPING.md`. Codes, not labels or SharePoint item IDs, are stored on Work Records.

### 2.7 `IU_Settings`

| Internal name | Display name | SharePoint / Graph type | Required | Default | App maximum | Index | Unique |
| --- | --- | --- | --- | --- | ---: | --- | --- |
| `Title` | Setting key | Single line text / `text` | Yes | No default | 255 | Yes | Yes |
| `ValueJson` | Setting value (JSON) | Multiple lines plain text / `text` | Yes | No default | 10,000 | No | No |

The only baseline keys are `ReportingConfig` and `SystemSettings`. `Title` is the stable key; the SharePoint item ID is provider metadata only.

## 3. Every column

The list matrices in section 2 contain every custom column in the baseline schema. SharePoint system fields used by the integration are listed separately in section 11. No other business, lookup, attachment, evidence, reporting-period, user, or legacy-audit column is created by default.

## 4. SharePoint field types

- Single-line values use `text` with `allowMultipleLines = false` and `maxLength = 255`.
- Short/long/JSON values use `text` with `allowMultipleLines = true`, `textType = plain`, `appendChangesToExistingText = false`, and the exact app maximum shown.
- Business dates use `dateTime` with `format = dateOnly`.
- Counts, minutes, `SchemaVersion`, and `RecordVersion` use `number` with zero decimal places.
- Flags use `boolean`.
- Controlled finite runtime enums use `choice` with exact ordered choices and fill-in disabled.
- Built-in audit/identity fields retain their SharePoint system types.

## 5. Required versus optional

The `Required` column in each matrix is exact. A runtime string that permits `""` is optional in SharePoint and normalizes to `""` in the provider. Nullable runtime dates/IDs are optional and normalize to `null`. JSON arrays are required because the runtime requires arrays and uses `[]`, not null.

## 6. Default values

The `Default` column in each matrix is exact. Defaults are storage safety only; the application still writes complete explicit field values. No-default optional fields normalize through the provider as documented. No date, identity, title/name, activity type, duration, reference status/type/group, code, color, role, or settings JSON receives an invented default.

## 7. Index recommendations

- Create every simple index shown in the `Index` column during provisioning.
- Unique AppId/Code/Settings-key columns are indexed and use `enforceUniqueValues = true`.
- Create the Work Records compound index with primary `ActivityDate` and secondary `OrbitReportable` in SharePoint List Settings.
- Do not index multiline text/JSON columns or add speculative indexes.
- Verify returned simple-index state through Graph and the compound index through SharePoint List Settings.

## 8. Unique constraints

- `IU_Work_Records.AppId`: unique.
- `IU_Organizations.AppId`: unique.
- `IU_Projects.AppId`: unique.
- `IU_Contacts.AppId`: unique.
- `IU_Work_Categories.AppId`: unique.
- `IU_ORBIT_Deliverables.Code`: unique.
- `IU_Settings.Title`: unique.

No other column is unique. Enforce unique values before seeding data; do not attempt to retrofit uniqueness after duplicates exist.

## 9. JSON validation expectations

All JSON fields use the output of `JSON.stringify(value)` with no indentation or nonessential whitespace. The application validates before every write and strictly validates every read.

Array JSON columns must satisfy all of the following:

- Top-level value is an array, never null or an object.
- Every element is a non-empty trimmed string of at most 255 characters.
- No duplicate values.
- Serialized value is at most 10,000 characters.
- Empty collection is exactly `[]`.
- `ProjectIdsJson` values exist in `IU_Projects.AppId`.
- `OrganizationIdsJson` values exist in `IU_Organizations.AppId` and satisfy `EngagementScope` rules.
- `ContactIdsJson` values exist in `IU_Contacts.AppId`.
- `CategoryIdsJson` values exist in `IU_Work_Categories.AppId`.
- `OrbitSupportingCodesJson` values exist in `IU_ORBIT_Deliverables.Code`, exclude the primary code, and are empty when not reportable.
- `EvidenceReferenceIdsJson` values are arbitrary stable IDs; no evidence list is inferred.

`IU_Settings.ValueJson` must be a JSON object, not an array or scalar:

- `ReportingConfig` has exactly `minutesPerReportingDay`, `schoolYearStartMonth`, and four quarter objects with the current runtime shape and values valid for the reporting utilities.
- `SystemSettings` has exactly `activityTypes`, a non-empty array of unique, non-empty strings, each at most 255 characters.
- Unknown setting keys or invalid/oversized JSON fail verification; they are not silently ignored.

SharePoint column validation formulas are not the authority for JSON shape. The provider codec and shared application validation are authoritative.

## 10. Consolidated application-level text limits

| Field class | Exact application cap |
| --- | ---: |
| All single-line IDs, names, titles, labels, roles, types, and codes | 255 characters |
| `ShortDescription`, `ProjectDescription`, `NextStep` | 1,000 characters |
| `EvidenceSummary`, `WorkOutput`, `WorkOutcome`, `OrbitEvidence` | 5,000 characters |
| `DetailedNotes` | 10,000 characters |
| Each compact JSON field, including Settings | 10,000 serialized characters |
| Each string inside a JSON array | 255 characters |

Validation rejects an oversized value. Neither the provisioning layer nor provider may truncate it.

## 11. SharePoint system fields used

| SharePoint/Graph system value | Use |
| --- | --- |
| List GUID (`list.id`) | Environment configuration key for each list |
| List Item `ID` / Graph `listItem.id` | `metadata.providerId` for a Work Record |
| `Created` / `listItem.createdDateTime` | Canonical `metadata.createdAt` |
| `Modified` / `listItem.lastModifiedDateTime` | Canonical `metadata.modifiedAt` |
| `Created By` (`Author`) | Audit only; normally the application service principal |
| `Modified By` (`Editor`) | Audit only; application service principal or exceptional administrator |
| List-item ETag (`listItem.eTag`) | Opaque conditional-write token |
| Built-in `Title` | Runtime title/name/label/key as specified for each list |

Do not map SharePoint `GUID`, content type IDs, attachment metadata, version-history labels, or user fields into the runtime model.

## 12. `AppId` and List Item ID mapping

- `AppId` is the stable, portable application key and is stored in a custom unique text column.
- SharePoint assigns a numeric list-item `ID`; Graph exposes it as `listItem.id` string.
- For Work Records, `metadata.providerId = listItem.id` after create/read.
- Item IDs are only unique within one list and environment. They never replace `AppId` in relationships, exports, migrations, or URLs owned by the business model.
- Updates normally address the item by `metadata.providerId`, then verify the stored `AppId` matches. An indexed AppId query is the recovery path when provider metadata is unavailable.
- `AppId` is immutable after creation. The application refuses an identity change; the column is hidden on the Work Records UI.

## 13. `RecordVersion` and ETag behavior

`RecordVersion` is the positive whole-number concurrency value exposed as `metadata.version`. The ETag remains opaque and provider-private.

1. Create writes `RecordVersion = 1`.
2. Update reads the current item to obtain `RecordVersion` and ETag.
3. If stored version differs from `expectedVersion`, return the current record as a structured conflict without writing.
4. Otherwise PATCH fields with `RecordVersion = expectedVersion + 1` and `If-Match: {current ETag}`.
5. Graph `412 Precondition Failed` is a race conflict; re-read and return the current record.
6. Read the successful update back to obtain canonical Modified time and the new item state.

Exceptional manual correction must increment `RecordVersion` exactly once along with the corrected field values. The correction must never change `AppId`, `SchemaVersion`, or compact JSON without running the same validation rules.

## 14. `schemaVersion` storage

- Runtime `schemaVersion` maps to `IU_Work_Records.SchemaVersion`.
- Type: required Number, zero decimal places.
- Default and current supported value: `2`.
- The application writes `2` explicitly on create/update and rejects any unsupported value on read or write.
- No schema-version fields are added to reference/configuration lists because their current runtime types have none.

## 15. Created and Modified treatment

- SharePoint/Graph `Created` and `Modified` are canonical provider timestamps.
- The client cannot supply or override them.
- `Created` remains unchanged on updates; `Modified` comes from SharePoint after the write.
- Store/read them as UTC ISO timestamps in nested provider metadata.
- `ActivityDate` and `FollowUpDate` are date-only business values and are never derived from or timezone-shifted with system timestamps.
- A manual administrator edit legitimately changes Modified/Modified By and must also follow the RecordVersion rule above.

## 16. Legacy timestamp fields

No legacy timestamp columns are part of the baseline schema. Prototype timestamps are not authoritative and are not migrated into SharePoint system fields.

If a migration rehearsal demonstrates a real operational need, the only approved optional extension is:

| Internal name | Display name | Type | Required | Default | Indexed | Runtime mapping |
| --- | --- | --- | --- | --- | --- | --- |
| `LegacyCreatedAtUtc` | Legacy prototype created (UTC) | Date and Time, Date & Time | No | No default | No | None; migration metadata only |
| `LegacyModifiedAtUtc` | Legacy prototype modified (UTC) | Date and Time, Date & Time | No | No default | No | None; migration metadata only |

These fields may be added only before production list creation with a recorded steward decision. Their absence does not block provisioning.

## 17. Required site and list permissions

### Runtime

- Entra application permission: Microsoft Graph `Sites.Selected` (application).
- SharePoint selected-site grant: `write` on the dedicated IU Work Tracker site.
- No runtime `Sites.ReadWrite.All`, `Sites.Manage.All`, or `Sites.FullControl.All`.
- No item-level permissions and no unique list permissions in the baseline.

### Provisioning and permission administration

- A temporary administrative provisioning context needs `Sites.Manage.All` to create lists/columns through Graph.
- Granting/verifying/removing the selected-site application permission requires an IU SharePoint/tenant administrative context with `Sites.FullControl.All` as documented by Microsoft.
- Elevated provisioning permissions must not remain assigned to the runtime application.

### Human access

- Primary and backup steward: members of the SharePoint Site Owners group with Full Control.
- Normal IU Work Tracker users: access the web application through the Entra assignment group; no SharePoint list Edit permission is required.
- Direct editing of `IU_Work_Records`: exceptional steward/admin operation only.
- Reference/configuration maintenance: steward operation, subject to validation/reconciliation before the application consumes changes.

## 18. Expected Entra security group model

Use one dedicated user-assignment group per environment, following IU naming standards:

| Logical group | Purpose | SharePoint rights |
| --- | --- | --- |
| `IU Work Tracker Users — DEV/TEST` | Users allowed to sign into the development/test enterprise application | None required |
| `IU Work Tracker Users — PROD` | Users allowed to sign into the production enterprise application | None required |

- Assign each group to its matching Entra enterprise application; do not cross-assign production and development groups by default.
- The primary and backup steward must be named individuals or IU-approved managed roles and must remain in the appropriate SharePoint Site Owners group.
- Do not use the application user group as a SharePoint Edit group. Application-only Graph access separates app authorization from direct list access.
- Existing IU break-glass/global/SharePoint administrator policy remains unchanged.
- The exact tenant-approved group display names/object IDs are execution inputs and must be recorded in the provisioning manifest.

## 19. Required Microsoft Graph endpoints

Use only `https://graph.microsoft.com/v1.0`; no beta endpoint is required.

### Site, list, and column provisioning/verification

| Purpose | Method and endpoint |
| --- | --- |
| Resolve approved site once by path | `GET /sites/{hostname}:/{relative-path}` |
| Enumerate lists before/after | `GET /sites/{site-id}/lists` |
| Create a generic list | `POST /sites/{site-id}/lists` |
| Read list metadata/GUID | `GET /sites/{site-id}/lists/{list-id}` |
| Enumerate returned column definitions | `GET /sites/{site-id}/lists/{list-id}/columns` |
| Create one custom column | `POST /sites/{site-id}/lists/{list-id}/columns` |
| Configure/verify built-in Title or correct a column property | `PATCH /sites/{site-id}/lists/{list-id}/columns/{column-id}` |

Graph `columnDefinition` supplies `name`, `displayName`, `required`, `defaultValue`, `indexed`, `enforceUniqueValues`, `hidden`, and type facets. Create columns one at a time and verify the returned definition so a partial failure is identifiable.

The ActivityDate/OrbitReportable compound index must be created and verified in SharePoint List Settings; Microsoft Graph's column definition exposes single-column `indexed` state but not this compound-index declaration.

### Selected-site permission grant and rollback

| Purpose | Method and endpoint |
| --- | --- |
| Grant runtime application the selected-site `write` role | `POST /sites/{site-id}/permissions` |
| Enumerate/verify selected-site grants | `GET /sites/{site-id}/permissions` |
| Read the exact grant by ID | `GET /sites/{site-id}/permissions/{permission-id}` |
| Remove an incorrect grant | `DELETE /sites/{site-id}/permissions/{permission-id}` |

### Item smoke verification and later runtime readiness

| Purpose | Method and endpoint |
| --- | --- |
| Create a test/reference item | `POST /sites/{site-id}/lists/{list-id}/items` |
| List items and selected fields | `GET /sites/{site-id}/lists/{list-id}/items?expand=fields(select=...)` |
| Read item, fields, ID, timestamps, and ETag | `GET /sites/{site-id}/lists/{list-id}/items/{item-id}?expand=fields(select=...)` |
| Conditional field update | `PATCH /sites/{site-id}/lists/{list-id}/items/{item-id}/fields` with `If-Match` |
| Delete only a disposable test item | `DELETE /sites/{site-id}/lists/{list-id}/items/{item-id}` |

Microsoft Graph v1.0 does not expose deletion of a list resource in its list methods. If an empty provisioned list is wrong, delete it through SharePoint List Settings so it enters the site recycle bin, or use an IU-approved SharePoint administrative REST process. Do not improvise a Graph list-delete endpoint.

## 20. Provisioning order

1. Record the approved tenant ID, environment, site URL, site ID, site owners, security-group object ID, runtime app ID, change ticket, and retention confirmation in a provisioning manifest.
2. Verify the target is the dedicated non-production site for the first run. Enumerate existing lists and stop if any target title already exists unexpectedly.
3. Establish the temporary administrative provisioning context. Do not enable runtime authentication or change application configuration.
4. Create `IU_Organizations`; patch built-in Title; add custom columns; add AppId unique index; verify empty schema.
5. Create `IU_Projects`; configure Title and columns/index; verify.
6. Create `IU_Contacts`; configure Title and columns/index; verify.
7. Create `IU_Work_Categories`; configure Title and columns/index; verify.
8. Create `IU_ORBIT_Deliverables`; configure Title and Code unique index; verify.
9. Create `IU_Settings`; configure Title as unique/indexed and add ValueJson; verify.
10. Create `IU_Work_Records` last; configure built-in Title, then custom columns, simple indexes, unique AppId, hidden technical columns, and the compound ActivityDate/OrbitReportable index; verify.
11. Confirm attachments/content approval are off and inherited permissions/retention are correct on every list.
12. Capture all list and custom-column GUIDs plus the exact returned internal/display names.
13. Grant the runtime application `Sites.Selected` write access to the dedicated site; capture and verify permission ID. Do not add user list permissions.
14. Run the post-creation verification checklist. Any test items must be development-only and removed afterward.
15. Remove temporary elevated provisioning access that is not part of IU's normal administrator model.
16. Store the signed provisioning manifest. Do not seed production data, authenticate the application, or switch persistence in this phase.

## 21. Rollback/delete plan if provisioning is wrong

### Before any data is loaded

1. Stop immediately on the first mismatch. Do not continue creating dependent lists.
2. Save the current manifest, returned list/column IDs, screenshots or exports of List Settings, and the exact mismatch.
3. If only a mutable property is wrong—display name, required flag, default, hidden state, or simple index—correct it only after comparing the returned internal name/type to this specification.
4. If an internal name or field type is wrong, do not work around or rename it. Confirm the affected list is the newly created target and contains zero items, then delete and recreate that list.
5. Delete affected empty lists in reverse dependency order: Work Records, Settings, Deliverables, Categories, Contacts, Projects, Organizations.
6. Use SharePoint **List Settings → Delete this list**. SharePoint sends deleted lists to the recycle bin under tenant policy; do not permanently purge during rollback.
7. If the selected-site application grant is wrong, remove the exact recorded permission ID with `DELETE /sites/{site-id}/permissions/{permission-id}`.
8. Never delete the SharePoint site, unrelated lists, production data, Entra tenant groups, or app registration as part of list-schema rollback.

### If any list contains non-disposable data

- Stop and escalate to the primary/backup steward.
- Export and reconcile data before any schema replacement.
- Do not delete or bulk-copy based only on a display title; verify site ID, list GUID, item count, and manifest ownership.
- Prefer creating a corrected replacement list with a temporary clearly marked title, validating migration, and switching configuration only after approval.

## 22. Verification checklist after list creation

The operator-facing version is [SHAREPOINT_PROVISIONING_CHECKLIST.md](./SHAREPOINT_PROVISIONING_CHECKLIST.md). A provisioning run is not complete until all of these are evidenced:

- The site URL/ID, environment, owners, user security group, runtime app ID, permission ID, retention basis, and change ticket are recorded.
- Exactly seven target generic lists exist with the exact display titles and captured GUIDs.
- Every custom column has the exact API/internal name, display name, type facet, required flag, default, maximum, hidden state, indexed state, and uniqueness setting specified here.
- All closed Choice values match exact runtime casing and allow no fill-in value.
- AppId/Code/Settings-key duplicate tests fail as expected.
- JSON fields accept compact valid JSON and application validation rejects malformed, duplicate, unknown, non-string, or oversized values.
- Date-only fields round-trip without timezone movement.
- Number fields display zero decimal places; application validation covers whole-number/cross-field rules.
- Work Record item ID, Created, Modified, Created By, Modified By, and ETag are observable through Graph; no custom audit timestamp is required.
- `RecordVersion` starts at `1`; a matching ETag update increments it to `2`; stale ETag returns `412`.
- Simple and compound indexes match section 2.1; no JSON/multiline lookup index exists.
- Attachments and content approval are off; versioning and retention follow IU policy.
- Runtime application has only the approved selected-site grant; normal users have no direct list Edit permission.
- Disposable test items are removed and all seven production lists are empty before authorized seeding/migration.
- No application provider setting, authentication configuration, or production persistence has changed.

## Microsoft references

- [Create a SharePoint list with Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/list-create?view=graph-rest-1.0)
- [Create a list column definition](https://learn.microsoft.com/en-us/graph/api/list-post-columns?view=graph-rest-1.0)
- [Column definition properties](https://learn.microsoft.com/en-us/graph/api/resources/columndefinition?view=graph-rest-1.0)
- [Update a column definition](https://learn.microsoft.com/en-us/graph/api/columndefinition-update?view=graph-rest-1.0)
- [List item conditional update and `412` behavior](https://learn.microsoft.com/en-us/graph/api/listitem-update?view=graph-rest-1.0)
- [Create a selected-site permission](https://learn.microsoft.com/en-us/graph/api/site-post-permissions?view=graph-rest-1.0)
- [Delete an incorrect selected-site permission](https://learn.microsoft.com/en-us/graph/api/site-delete-permission?view=graph-rest-1.0)
- [Delete and restore a SharePoint list](https://support.microsoft.com/en-us/sharepoint/lists/data-and-lists/delete-a-list)
