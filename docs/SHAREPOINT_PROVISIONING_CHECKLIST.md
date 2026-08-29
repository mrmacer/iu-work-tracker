# IU Work Tracker — SharePoint Provisioning Checklist

Use this checklist only during an approved SharePoint provisioning or verification window. It does not authorize list creation, Entra changes, authentication, data migration, or persistence cutover.

Detailed field definitions are in [SHAREPOINT_PROVISIONING_SPEC.md](./SHAREPOINT_PROVISIONING_SPEC.md).

## A. Record the approved target

- [ ] Change/request ticket: `____________________________`
- [ ] Environment: `DEV/TEST` or `PROD` (circle one)
- [ ] Microsoft 365 tenant ID: `____________________________`
- [ ] SharePoint hostname: `____________________________`
- [ ] SharePoint site URL: `____________________________`
- [ ] Microsoft Graph site ID: `____________________________`
- [ ] Primary steward: `____________________________`
- [ ] Backup steward: `____________________________`
- [ ] IU SharePoint/tenant administrator: `____________________________`
- [ ] Entra user security-group name: `____________________________`
- [ ] Entra user security-group object ID: `____________________________`
- [ ] Runtime app registration display name: `____________________________`
- [ ] Runtime app/client ID: `____________________________`
- [ ] Existing IU retention policy/label confirmed: `____________________________`
- [ ] Confirm this is a dedicated IU-owned Work Tracker site.
- [ ] Confirm ordinary users will use the application and will not directly edit Work Records.
- [ ] Confirm primary and backup stewards are in the SharePoint Site Owners group.
- [ ] Confirm no production persistence or application provider setting will change during provisioning.

## B. Preflight before creating anything

- [ ] Verify the recorded site URL resolves to the recorded Graph site ID.
- [ ] Verify the site is the intended environment; first provisioning run is DEV/TEST.
- [ ] Enumerate existing site lists and record the result.
- [ ] Stop if any of the seven exact target titles already exists unexpectedly.
- [ ] Confirm no unrelated list/site will be modified.
- [ ] Confirm the provisioning identity is temporary and administratively approved.
- [ ] Confirm the runtime application does not have tenant-wide `Sites.ReadWrite.All`, `Sites.Manage.All`, or `Sites.FullControl.All`.
- [ ] Start a provisioning manifest and record every created list, column, and permission ID.

## C. Create the seven generic lists in order

For every list: template `genericList`; attachments off; content approval off; no folders; inherited site permissions; IU retention/default versioning retained.

### 1. `IU_Organizations`

- [ ] Create exact list title `IU_Organizations`.
- [ ] Capture list GUID: `____________________________`
- [ ] `Title` → display `Organization name`; single line; required; 255; no default.
- [ ] `AppId` → display `Application ID`; single line; required; 255; indexed; unique.
- [ ] `OrganizationType` → display `Organization type`; required closed Choice: `district`, `partner`, `iu`; no default; fill-in off.
- [ ] Confirm there is no Lookup column and no regional/all-district pseudo-organization field.
- [ ] Confirm the list is empty.

### 2. `IU_Projects`

- [ ] Create exact list title `IU_Projects`.
- [ ] Capture list GUID: `____________________________`
- [ ] `Title` → display `Project name`; single line; required; 255; no default.
- [ ] `AppId` → display `Application ID`; single line; required; 255; indexed; unique.
- [ ] `ProjectDescription` → display `Project description`; multiline plain; optional; maximum 1,000; no default.
- [ ] `ProjectStatus` → display `Project status`; required closed Choice: `active`, `planning`, `complete`; no default; fill-in off.
- [ ] `Color` → display `Display color token`; single line; required; maximum 255; no default.
- [ ] Confirm the list is empty.

### 3. `IU_Contacts`

- [ ] Create exact list title `IU_Contacts`.
- [ ] Capture list GUID: `____________________________`
- [ ] `Title` → display `Contact display name`; single line; required; 255; no default.
- [ ] `AppId` → display `Application ID`; single line; required; 255; indexed; unique.
- [ ] `Role` → display `Role`; single line; required; maximum 255; no default.
- [ ] `OrganizationAppId` → display `Organization application ID`; single line; optional; maximum 255; no default.
- [ ] Confirm `OrganizationAppId` is text, not a Lookup.
- [ ] Confirm the list is empty.

### 4. `IU_Work_Categories`

- [ ] Create exact list title `IU_Work_Categories`.
- [ ] Capture list GUID: `____________________________`
- [ ] `Title` → display `Category name`; single line; required; 255; no default.
- [ ] `AppId` → display `Application ID`; single line; required; 255; indexed; unique.
- [ ] `CategoryGroup` → display `Category group`; required closed Choice: `work-area`, `topic`; no default; fill-in off.
- [ ] Confirm the list is empty.

### 5. `IU_ORBIT_Deliverables`

- [ ] Create exact list title `IU_ORBIT_Deliverables`.
- [ ] Capture list GUID: `____________________________`
- [ ] `Title` → display `Deliverable label`; single line; required; 255; no default.
- [ ] `Code` → display `Deliverable code`; single line; required; 255; indexed; unique.
- [ ] Confirm no Lookup or multi-choice deliverable column exists.
- [ ] Confirm the list is empty.

### 6. `IU_Settings`

- [ ] Create exact list title `IU_Settings`.
- [ ] Capture list GUID: `____________________________`
- [ ] `Title` → display `Setting key`; single line; required; 255; indexed; unique.
- [ ] `ValueJson` → display `Setting value (JSON)`; multiline plain; required; maximum 10,000; no default.
- [ ] Confirm rich text and append-changes are off.
- [ ] Confirm baseline keys will be only `ReportingConfig` and `SystemSettings` when seeding is later authorized.
- [ ] Confirm the list is empty.

### 7. `IU_Work_Records`

- [ ] Create exact list title `IU_Work_Records` last.
- [ ] Capture list GUID: `____________________________`

Identity and activity:

- [ ] `Title` → `Work title`; single line; required; maximum 255; no default; shown.
- [ ] `AppId` → `Application ID`; single line; required; maximum 255; indexed; unique; hidden.
- [ ] `ActivityDate` → `Activity date`; Date Only; required; indexed; no default.
- [ ] `ActivityType` → `Activity type`; single line; required; maximum 255; no default.
- [ ] `ShortDescription` → `Short description`; multiline plain; optional; maximum 1,000; no default.
- [ ] `DetailedNotes` → `Detailed notes`; multiline plain; optional; maximum 10,000; no default.
- [ ] `DurationMinutes` → `Duration minutes`; Number; zero decimals; required; no default.
- [ ] `RecordStatus` → `Record status`; required closed Choice `complete`, `draft`; default `complete`; fill-in off.
- [ ] `EngagementScope` → `Engagement scope`; required closed Choice `none`, `specific`, `regional`, `allDistricts`; default `none`; fill-in off.

Relationship JSON:

- [ ] `ProjectIdsJson` → `Project IDs (JSON)`; multiline plain; required; default `[]`; max 10,000; hidden.
- [ ] `OrganizationIdsJson` → `Organization IDs (JSON)`; multiline plain; required; default `[]`; max 10,000; hidden.
- [ ] `ContactIdsJson` → `Contact IDs (JSON)`; multiline plain; required; default `[]`; max 10,000; hidden.
- [ ] `CategoryIdsJson` → `Category IDs (JSON)`; multiline plain; required; default `[]`; max 10,000; hidden.
- [ ] Confirm all JSON columns have rich text and append-changes off and are not indexed.

Reach and general evidence:

- [ ] `EducatorsLeadersReach` → `Educators / leaders reach`; Number; required; zero decimals; default `0`.
- [ ] `StudentsFamiliesReach` → `Students / families reach`; Number; required; zero decimals; default `0`.
- [ ] `WorkforceCommunityReach` → `Workforce / community reach`; Number; required; zero decimals; default `0`.
- [ ] `OtherReach` → `Other reach`; Number; required; zero decimals; default `0`.
- [ ] `EvidenceSummary` → `Evidence summary`; multiline plain; optional; maximum 5,000; no default.
- [ ] `EvidenceReferenceIdsJson` → `Evidence reference IDs (JSON)`; multiline plain; required; default `[]`; max 10,000; hidden.

Output, outcome, and follow-up:

- [ ] `WorkOutput` → `Output`; multiline plain; optional; maximum 5,000; no default.
- [ ] `WorkOutcome` → `Outcome`; multiline plain; optional; maximum 5,000; no default.
- [ ] `NextStep` → `Next step`; multiline plain; optional; maximum 1,000; no default.
- [ ] `FollowUpNeeded` → `Follow-up needed`; Yes/No; required; default `false`.
- [ ] `FollowUpDate` → `Follow-up date`; Date Only; optional; indexed; no default.

ORBIT:

- [ ] `OrbitReportable` → `ORBIT reportable`; Yes/No; required; default `false`.
- [ ] `OrbitPrimaryDeliverableCode` → `ORBIT primary deliverable code`; single line; optional; max 255; no default.
- [ ] `OrbitSupportingCodesJson` → `ORBIT supporting codes (JSON)`; multiline plain; required; default `[]`; max 10,000; hidden.
- [ ] `StemPocMinutes` → `STEM PoC minutes`; Number; zero decimals; required; default `0`.
- [ ] `TacMinutes` → `TaC minutes`; Number; zero decimals; required; default `0`.
- [ ] `OrbitEvidence` → `ORBIT evidence`; multiline plain; optional; maximum 5,000; no default.

Provider/evolution fields:

- [ ] `SchemaVersion` → `Schema version`; Number; zero decimals; required; default `2`; hidden.
- [ ] `RecordVersion` → `Record version`; Number; zero decimals; required; default `1`; hidden.
- [ ] `IsSample` → `Development/sample record`; Yes/No; required; default `false`.

Indexes:

- [ ] Verify simple unique index on `AppId`.
- [ ] Verify simple index on `ActivityDate`.
- [ ] Verify simple index on `FollowUpDate`.
- [ ] In List Settings, create/verify compound index with primary `ActivityDate`, secondary `OrbitReportable`.
- [ ] Confirm no JSON or multiline-text field is indexed.
- [ ] Confirm the list is empty.

## D. System fields and list behavior

For each list:

- [ ] Built-in item `ID`, `Created`, `Modified`, `Created By`, and `Modified By` exist.
- [ ] Built-in `Title` has the exact display name, required flag, uniqueness, and index setting specified for that list.
- [ ] Attachments are disabled.
- [ ] Content approval/moderation is disabled.
- [ ] Folders are not used.
- [ ] Version-history behavior follows the IU/site default.
- [ ] Retention follows the recorded IU Microsoft 365 policy/label.
- [ ] Permissions inherit from the dedicated site; no per-item permission exists.

Legacy timestamps:

- [ ] Confirm baseline provisioning has no `LegacyCreatedAtUtc` or `LegacyModifiedAtUtc` columns.
- [ ] If optional legacy metadata was explicitly approved before production provisioning, record the decision and verify both fields are optional Date & Time, unindexed, and unmapped to runtime metadata.

## E. Permissions and Entra checks

- [ ] Primary steward is in Site Owners.
- [ ] Backup steward is in Site Owners.
- [ ] DEV/TEST and PROD use separate Entra user-assignment groups.
- [ ] Correct group is assigned to the matching enterprise application.
- [ ] Normal users do not have SharePoint list Edit rights.
- [ ] Runtime application has Microsoft Graph application permission `Sites.Selected`.
- [ ] Runtime application has exactly one approved selected-site `write` grant for this environment.
- [ ] Capture selected-site permission ID: `____________________________`
- [ ] Runtime application does not retain provisioning-wide permissions.
- [ ] No direct user/list-item permission was created by the provisioning process.

## F. Schema and constraint verification

Run destructive/negative tests only in DEV/TEST with disposable items.

- [ ] Enumerate all seven list GUIDs and columns through Graph; compare internal and display names to the specification.
- [ ] Compare every returned type facet, required flag, default, `indexed`, `enforceUniqueValues`, `hidden`, and text maximum.
- [ ] Verify all Choice values use exact casing and reject fill-in values.
- [ ] Verify duplicate `AppId` fails in each AppId reference list and Work Records.
- [ ] Verify duplicate Deliverable `Code` fails.
- [ ] Verify duplicate Settings `Title` fails.
- [ ] Verify a valid compact JSON array round-trips exactly.
- [ ] Verify malformed JSON, non-array JSON, non-string elements, duplicates, unknown canonical IDs, and JSON over 10,000 characters are rejected by application validation.
- [ ] Verify Date Only values round-trip as the same `YYYY-MM-DD` without timezone movement.
- [ ] Verify Number fields use zero decimal display; confirm application validation enforces integer/cross-field rules.
- [ ] Verify single-line values over 255 and app-capped multiline values are rejected without truncation.

## G. Work Record concurrency smoke test

Complete only after a development Graph test identity is separately authorized; do not add authentication as part of this checklist.

- [ ] Create a disposable valid Work Record with `RecordVersion = 1` and `SchemaVersion = 2`.
- [ ] Capture AppId: `____________________________`
- [ ] Capture SharePoint item ID: `____________________________`
- [ ] Confirm AppId differs from and remains independent of item ID.
- [ ] Capture initial ETag: `____________________________`
- [ ] Confirm `Created` and `Modified` are supplied by SharePoint/Graph.
- [ ] Update with the matching ETag and increment `RecordVersion` from `1` to `2`.
- [ ] Confirm Created is unchanged and Modified advances.
- [ ] Retry with the stale ETag; confirm Graph returns `412 Precondition Failed` and no overwrite occurs.
- [ ] Delete the disposable item and confirm the list is empty.

## H. Rollback checks

- [ ] Stop on the first internal-name or type mismatch.
- [ ] Confirm the exact site ID, list GUID, title, manifest ownership, and item count before deleting anything.
- [ ] If the wrong list is empty, delete it through SharePoint List Settings so it enters the recycle bin; do not purge it.
- [ ] Delete affected empty lists only in reverse order: Work Records, Settings, Deliverables, Categories, Contacts, Projects, Organizations.
- [ ] If data is present, stop and obtain steward approval; export/reconcile before any replacement.
- [ ] If the selected-site grant is wrong, remove only the recorded permission ID.
- [ ] Never delete the site, unrelated lists, production data, tenant groups, or app registration as schema rollback.

## I. Final handoff

- [ ] Exactly seven target lists exist.
- [ ] All list and custom-column GUIDs are captured in the signed provisioning manifest.
- [ ] All schema, indexes, defaults, uniqueness, hidden states, permissions, retention, and behavior checks pass.
- [ ] All disposable test data is removed.
- [ ] Production lists remain empty until authorized seeding/migration.
- [ ] Temporary elevated provisioning access is removed or returned to IU's normal administrator baseline.
- [ ] D1 remains the active production persistence provider.
- [ ] No application authentication or Graph runtime integration has been enabled.
- [ ] No production data has been migrated.
- [ ] Primary steward sign-off: `____________________________  Date: __________`
- [ ] Backup steward sign-off: `____________________________  Date: __________`
- [ ] Administrator sign-off: `____________________________  Date: __________`
