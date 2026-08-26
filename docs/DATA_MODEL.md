# Data Model

## Relationship overview

```text
Project 0..* <——> 0..* WorkRecord 0..* <——> 0..* Organization
                              | 0..* <——> Contact
                              | 0..* <——> Category
                              | 1 ——> 0..* Evidence
                              | 1 ——> 0..* FollowUp
                              └ 0..1 ——> OrbitClassification
Configuration ——> derived quarter, reporting-day conversion, vocabulary
```

## WorkRecord

The canonical activity record.

| Group | Fields |
| --- | --- |
| Identity | `appId` UUID, provider item ID, title, activity date, start/end timestamp, `durationMinutes`, status, created/modified timestamps, sync state, sample flag |
| Description | activity type, short description, detailed notes, output, outcome, success/evidence summary, barrier, opportunity, next step |
| Relationships | project IDs, organization IDs, contact IDs, category IDs; regional/all-district scope flag |
| Reach | educator/leader, student/family, workforce/community, other participant counts; total derived where appropriate |
| Follow-up | needed flag, due date, next-step summary; normalized FollowUp records may be added for multiple actions |
| Reporting | optional OrbitClassification, precise STEM PoC minutes, TaC minutes, qualitative evidence references; school year and quarter derived from date |

Rules: title, activity date, duration, and activity type are the minimal quick-entry core. Counts are non-negative. End time may derive duration, but `durationMinutes` is authoritative. Multi-valued relationships use stable application IDs.

## Project

`appId`, name, description, status, start/end dates, organization IDs, contact IDs, milestone summaries, created/modified timestamps. Time, activities, reach, outcomes, and evidence are calculated from related Work Records.

## Organization

`appId`, canonical name, organization type, subtype/LEA classification, PDE/NCES or other external ID, parent organization ID, active flag, region, created/modified timestamps. District names are selected from this entity rather than stored as uncontrolled record text.

## Contact

`appId`, display name, title/role, organization ID, email, phone, tags, active flag, created/modified timestamps. V1 treats contacts as reusable references rather than a CRM.

## Category

`appId`, name, group, description, color token, sort order, active flag. Categories are configurable data, not conditionals embedded in the UI.

## Evidence

`appId`, work record ID, evidence type, title, URL or future SharePoint drive/item reference, description, captured date, created/modified timestamps. V1 can capture links and notes; future storage can attach files without changing the Work Record.

## FollowUp

`appId`, work record ID, summary, due date, status, owner contact ID, completed date, reminder state, created/modified timestamps.

## OrbitClassification

`reportable`, primary deliverable, supporting deliverables, STEM PoC minutes, TaC minutes, qualitative evidence summary. Quarter and reporting days are derived rather than entered. The model permits one activity to support several deliverables while maintaining one official primary classification.

## Configuration

Named, versioned values for work categories, activity types, deliverables, school-year start, quarter boundaries, minutes per reporting day (default 420), reach labels, and capacity targets. Historical calculation settings should be versioned so rule changes do not rewrite source duration.

## Provider boundary

UI code consumes a `DataProvider` contract (`list`, `get`, `create`, `update`, configuration/entity lookup). V1 uses a prototype provider; a future `SharePointDataProvider` implements the same domain operations and maps storage IDs, field names, retries, and sync states internally.
