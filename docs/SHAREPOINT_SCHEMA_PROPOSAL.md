# SharePoint Schema Proposal

No lists are created in V1. This proposal should be validated against the IU tenant, permissions, naming rules, retention policy, and expected reporting volume before provisioning.

## Recommended lists

### IU_Work_Records

Core fields: Title (text), AppId (text, unique), ActivityDate (date/time), StartTime and EndTime (date/time, optional), DurationMinutes (number), Status (choice), ActivityTypeId (text), ShortDescription (multiline plain text), DetailedNotes (multiline plain text), Output, Outcome, SuccessEvidence, Barrier, Opportunity, NextStep (multiline plain text), FollowUpNeeded (yes/no), FollowUpDate (date), ProjectIds, OrganizationIds, ContactIds, CategoryIds (multiline text containing a versioned JSON array of stable AppIds), Reach counts (numbers), RegionalScope (choice), OrbitReportable (yes/no), OrbitPrimaryDeliverable (choice), OrbitSupportingDeliverables (multiline JSON), StemPocMinutes and TacMinutes (numbers), OrbitEvidenceSummary (multiline text), AppCreatedAt and AppModifiedAt (date/time), SchemaVersion (number), and SyncFingerprint (text).

Indexes: AppId unique; ActivityDate; Status; FollowUpDate; OrbitReportable plus ActivityDate if tenant list tooling supports the compound access pattern. Query by date ranges first and filter relationship ID arrays client-side for V1-scale volumes.

### IU_Projects

AppId (unique text), Title, Description, Status, StartDate, EndDate, OrganizationIds and ContactIds (versioned JSON), CreatedAt, ModifiedAt, Active. Index AppId, Status, and date fields.

### IU_Organizations

AppId (unique text), Title/canonical name, OrganizationType, LeaType, ExternalId, ParentOrganizationAppId, Region, Active, CreatedAt, ModifiedAt. Index AppId, OrganizationType, Active, and ExternalId where populated.

### IU_Contacts

AppId (unique text), Title/display name, Role, OrganizationAppId, Email, Phone, Tags (JSON), Active, CreatedAt, ModifiedAt. Index AppId, OrganizationAppId, Active.

### IU_Work_Categories

AppId (unique text), Title, Group, Description, ColorToken, SortOrder, Active. Index AppId, Group, Active.

### IU_Evidence

AppId (unique text), WorkRecordAppId, EvidenceType, Title, Url, DriveId, DriveItemId, Description, CapturedDate, CreatedAt, ModifiedAt. Store file bytes in a document library, not a list attachment. Index AppId, WorkRecordAppId, CapturedDate.

### IU_Followups

AppId (unique text), WorkRecordAppId, Summary, DueDate, Status, OwnerContactAppId, CompletedDate, ReminderState, CreatedAt, ModifiedAt. Index AppId, WorkRecordAppId, DueDate, Status.

### IU_Settings

Key (unique text), ValueJson (multiline plain text), Version (number), EffectiveFrom (date), EffectiveTo (date), Description. Stores quarter boundaries, minutes per reporting day, capacity targets, and controlled-vocabulary configuration.

## Lookup strategy

Use stable AppId text references for high-cardinality or multi-valued relationships. Avoid SharePoint multi-lookup columns for projects, organizations, contacts, categories, deliverables, and evidence because Graph expansion, list thresholds, batching, offline writes, and migration become unnecessarily fragile. The adapter resolves AppIds against cached canonical entities.

A single-value SharePoint lookup may be reasonable later for a primary project or primary organization only if tenant testing proves a clear reporting benefit. Do not make correctness depend on display names.

## ID and synchronization rules

- Generate UUID AppIds before the network write; SharePoint item IDs remain provider metadata.
- Enforce duplicate prevention through an indexed, unique AppId field and idempotent upsert behavior.
- Store a schema version and a sync fingerprint/ETag. Use `If-Match` for safe updates.
- Preserve unsaved payloads on failures and expose Saved, Saving, Offline, Sync pending, and Sync error states.
- Discover and validate internal field names from list metadata; never infer them from display labels.
- Batch relationship resolution and cache canonical entities to minimize Graph calls.
