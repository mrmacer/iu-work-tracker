# SharePoint Schema Proposal — V1.1-Aligned

No SharePoint list, Microsoft Graph call, authentication flow, or synchronization process is implemented in V1.1. This proposal maps the runtime meaning that a future provider must preserve.

## IU_Work_Records

| Runtime concern | Proposed SharePoint storage |
| --- | --- |
| Stable identity | `AppId` unique indexed text |
| Core activity | `Title`, `ActivityDate`, `ActivityType`, `DurationMinutes`, `Status`, descriptions, output/outcome/next step |
| Engagement | `EngagementScope` choice: none/specific/regional/allDistricts |
| Relationships | versioned JSON text arrays for `ProjectIds`, `OrganizationIds`, `ContactIds`, `CategoryIds` |
| Reach | four number fields |
| Evidence | `EvidenceSummary`, `EvidenceReferenceIdsJson`; references point to future stable evidence identities, not uploaded bytes in V1.1 |
| Follow-up | `FollowUpNeeded`, `FollowUpDate` |
| ORBIT | reportable yes/no, primary choice, supporting JSON, PoC minutes, TaC minutes, evidence summary |
| Data evolution | `SchemaVersion` number |
| Concurrency | SharePoint item ID and ETag map into nested provider metadata; ETag participates in expected-version updates |
| Audit time | provider-created and modified timestamps map into nested metadata; the server/provider owns them |

The proposed adapter must not store scope concepts such as “regional” as organization IDs. `OrganizationIds` always contains canonical real entities.

## Reference lists

- `IU_Projects`: AppId, name/title, description, status, display metadata.
- `IU_Organizations`: AppId, canonical name, type (`district`, `partner`, `iu`), external IDs and active state when later authorized.
- `IU_Contacts`: AppId, display name, role, optional organization ID. V1.1 does not implement CRM behavior.
- `IU_Work_Categories`: AppId, name, group, description/sort metadata.
- `IU_ORBIT_Deliverables`: stable code and label, active/effective metadata if institutional policy requires it.
- `IU_Settings`: versioned reporting configuration and controlled activity types.

Use stable application IDs in multi-value JSON fields rather than display names or fragile multi-lookup columns. A future SharePoint provider resolves those IDs against canonical lists and returns the same domain types now consumed by the UI.

## Provider contract implications

The future adapter implements work-record list/create/update and all reference/configuration getters. Create uses an application-generated AppId but provider-owned item ID, timestamps, and initial version. Update includes the expected ETag/version and returns a structured conflict instead of last-write-wins behavior. Validation remains below the UI and applies before persistence.

Tenant permissions, retention, indexing thresholds, internal field names, and list provisioning remain intentionally unimplemented and must be validated before integration begins.
