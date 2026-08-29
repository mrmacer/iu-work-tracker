# IU Work Tracker — AI Handoff

## Purpose

This file is the canonical short-form handoff between Claude Code, Codex, and future coding agents working on IU Work Tracker.

READ THIS FILE FIRST before performing development work.

Git is the source of truth for code state.

Authoritative project documents remain authoritative for detailed architecture and specifications. This file summarizes current state; it does not replace them.

---

## Current State

IU Work Tracker V1.1 hardening is complete.

The application has:

- universal Work Record architecture
- optional/subordinate ORBIT classification
- DataProvider abstraction
- stable application IDs
- runtime validation
- optimistic concurrency architecture
- accessibility hardening
- 26 original V1.1 tests plus subsequent auth tests
- Microsoft delegated authentication implemented for DEV
- Microsoft Graph connectivity verified
- SharePoint DEV site connectivity verified

DEV SharePoint target:

https://siu29.sharepoint.com/sites/IUWorkTrackerDEV

Microsoft sign-in has been successfully tested with an IU29 account.

Graph `/me` succeeded.

The DEV SharePoint site resolved successfully.

Existing SharePoint resources were readable.

DEV SharePoint infrastructure has been provisioned according to the authoritative provisioning specification.

Graph smoke testing successfully demonstrated:

CREATE
READ
UPDATE
STALE-WRITE REJECTION
DELETE

Specific concurrency verification:

- an update changed the SharePoint ETag
- reuse of the stale ETag returned HTTP 412
- JSON arrays round-tripped successfully

The synthetic smoke-test Work Record was permanently deleted.

Verification of that exact deleted item returned HTTP 404.

No synthetic smoke-test Work Record should remain.

---

## CRITICAL CURRENT PERSISTENCE STATE

`DelegatedSharePointDataProvider` IS NOW IMPLEMENTED AND LIVE-VERIFIED IN DEV.

See docs/SHAREPOINT_PROVIDER_INTEGRATION_REPORT.md for full detail. Summary:

- Normal application actions (Log Work: create and update) now persist through the existing `DataProvider` boundary into live DEV SharePoint `IU_Work_Records` when a Microsoft account is signed in under DEV configuration.
- Provider selection (`selectDataProvider()` in `lib/data-provider.ts`) is explicit and non-interactive: SharePoint activates only when DEV config is present AND a Microsoft account is already signed in; it never forces a sign-in prompt.
- The prototype `ApiDataProvider` remains the default for every user who is not signed in, and remains the automatic fallback if the SharePoint provider fails to load for any reason. It has not been removed or bypassed, and was not asked to be.
- Live end-to-end verification (create → update including ORBIT presence → confirm via Graph → delete → confirm 404) was completed successfully against `https://siu29.sharepoint.com/sites/IUWorkTrackerDEV` through the normal Log Work UI. No synthetic data was left behind.
- Reference/configuration data (Projects, Organizations, Contacts, Categories, Deliverables, ReportingConfig, SystemSettings) still comes from static seed data even when SharePoint is the active Work Records provider — reading those six lists from SharePoint was explicitly out of scope for this phase.
- This is DEV-only and single-user-verified. It does not itself authorize production use; production still requires the separately-planned `Sites.Selected` application-permission architecture (see "DEV vs PRODUCTION" below).

Do not confuse:

"SharePoint works and is wired into the app in DEV"

with:

"The application is in production-ready cutover to SharePoint."

Those are not the same thing.

---

## Next Development Goal

WAIT FOR EXPLICIT USER INSTRUCTION.

Candidate next phases (none authorized to begin automatically):

- Broader DEV verification: multiple concurrent users/records, Graph throttling/retry behavior, pagination beyond one page.
- Reading the six reference/configuration lists from SharePoint instead of static seed data.
- Evaluating production readiness: `Sites.Selected` application-permission architecture, confidential server-side app registration, migration planning per docs/SHAREPOINT_INTEGRATION_PLAN.md §18-20.

DO NOT begin any of these unless explicitly instructed.

---

## Future Candidate Feature — NOT CURRENT SCOPE

A future feature under consideration is:

INBOX INTELLIGENCE

Concept:

Paste an email or email thread into IU Work Tracker.

AI analyzes it and may extract:

- summary
- priority
- action items
- follow-up
- timing/due date
- people
- organizations
- district/LEA
- related project
- related work category
- potential Work Record connection

Potential workflow:

PASTE EMAIL
→ ANALYZE
→ HUMAN REVIEW
→ SAVE TO INBOX
→ optionally CREATE WORK ITEM

A standalone prototype named:

email-kb-capture.html

was created as an exploratory concept.

DO NOT treat that standalone prototype as the desired production architecture.

Preferred direction is eventual integration into IU Work Tracker rather than creating a second dashboard/database/authentication/storage system.

Initial implementation should likely remain manual copy/paste.

Possible future Microsoft Graph Outlook integration may be evaluated separately.

DO NOT request Mail.Read or other mailbox permissions without explicit authorization.

---

## Explicitly Out of Scope Unless Requested

Do NOT independently begin:

- UI redesign
- dashboard redesign
- R10–R15
- production SharePoint migration
- Power Automate
- Outlook mailbox integration
- Mail.Read permission changes
- schema expansion
- SharePoint lookup fields
- ORBIT redesign
- new authentication architecture
- unrelated refactoring
- package upgrades
- deployment changes

Do not "clean up" working architecture merely because another implementation appears preferable.

---

## DEV vs PRODUCTION

Current SharePoint integration is DEV ONLY.

DEV uses delegated Microsoft Graph authentication.

The delegated Graph permission architecture is intentionally a development solution.

Production security architecture should be evaluated separately.

The preferred production direction remains narrower SharePoint access such as Sites.Selected if IU29 tenant administration permits it.

Do NOT normalize broad delegated permissions into production architecture without explicit approval.

---

## Authoritative Documents

Before modifying the relevant subsystem, read the appropriate documents.

Primary:

- docs/SHAREPOINT_INTEGRATION_PLAN.md
- docs/SHAREPOINT_PROVISIONING_SPEC.md
- docs/SHAREPOINT_PROVISIONING_CHECKLIST.md
- docs/SHAREPOINT_DEV_PROVISIONING_REPORT.md
- docs/DELEGATED_AUTH_SETUP.md
- docs/DELEGATED_AUTH_IMPLEMENTATION_REPORT.md
- docs/V1_1_IMPLEMENTATION_REPORT.md

Also consult:

- docs/PRODUCT_VISION.md
- docs/DATA_MODEL.md
- docs/ORBIT_MAPPING.md
- docs/V1_PLAN.md

If this handoff conflicts with a detailed authoritative specification, STOP and identify the conflict rather than silently choosing one.

---

## Agent Start Protocol

Every coding agent must begin by:

1. Read docs/AI_HANDOFF.md.
2. Run `git status`.
3. Inspect recent Git history.
4. Read only the authoritative documents relevant to the requested task.
5. Inspect the existing implementation before proposing replacement architecture.
6. State the smallest reasonable implementation plan.
7. Modify only what the requested task requires.

Do not begin by scanning or rewriting the entire repository unless necessary.

---

## Cost / Scope Discipline

This project is intentionally optimizing coding-agent usage.

Prefer:

- targeted file inspection
- targeted searches
- existing tests
- small patches
- existing architecture
- incremental verification

Avoid:

- repository-wide analysis when unnecessary
- speculative refactors
- rebuilding established context
- duplicate documentation
- unnecessary dependency research
- broad test generation unrelated to the change
- implementing "nice to have" improvements

If the requested task can be completed safely with a small patch, make the small patch.

---

## Agent Stop Protocol

At the end of a meaningful development phase:

1. Run appropriate existing tests.
2. Run type/lint/build checks when relevant.
3. Run `git status`.
4. Summarize files changed.
5. Update this handoff ONLY if project state materially changed.
6. Update "Next Development Goal" if the completed phase changes it.
7. Do not silently start the next phase.
8. Recommend a Git checkpoint.

If a task fails or remains incomplete, record that clearly here rather than describing it as complete.

---

## Git Handoff Rule

Claude Code and Codex must NEVER intentionally edit this repository simultaneously.

Before changing agents:

1. finish or stop the current task
2. inspect `git status`
3. review changes
4. commit a checkpoint if the state is worth preserving
5. then open the other coding agent

Recommended checkpoint format:

git add .
git commit -m "checkpoint: <completed phase>"

Do not commit broken or unknown state merely to make the working tree clean.

---

## Last Verified Milestone

`DelegatedSharePointDataProvider` implemented, unit-tested (64/64 tests, build, typecheck, lint clean on all files touched), and live-verified end-to-end through the normal Log Work UI against DEV SharePoint: create, update (including ORBIT presence), Graph-side confirmation of every field/timestamp/version rule, delete, and confirmed HTTP 404 afterward. No synthetic data left behind. See docs/SHAREPOINT_PROVIDER_INTEGRATION_REPORT.md.

The prototype `ApiDataProvider` remains the default/fallback provider by design.

---

## Next Agent

WAIT FOR EXPLICIT USER INSTRUCTION.

Do not begin any further SharePoint integration phase (reference-data lists, production `Sites.Selected` planning, broader concurrency/multi-user testing) automatically.
