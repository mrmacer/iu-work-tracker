# V1.1 Plan and Implemented Scope

## Goal

Make the universal Work Record trustworthy enough to enter, classify, persist, update, project, and later map to SharePoint without changing its meaning.

## Implemented in V1.1 (R1–R9)

- Semantic ORBIT checkbox/label control with mouse, touch-compatible pointer semantics, keyboard control, and accessible state.
- Provider-loaded Work Records, projects, organizations, contacts, categories, deliverables, reporting configuration, and system settings.
- Explicit engagement scope separated from real canonical organizations.
- General evidence references, contacts, schema version, and nested provider metadata.
- Optional ORBIT classification with primary/supporting deliverables, pure calculations, and conservative PoC/TaC allocation.
- Explicit activity type plus progressive classification/reach/evidence/contact controls.
- Shared runtime validation, structured results, server timestamps, immutable creation time, and optimistic version conflicts.
- Drizzle schema plus generated migrations as the only D1 schema authority; sample seeding remains separate.
- Modal focus entry, focus trap, Escape protection, focus restoration, and accessible step names/state.

## Migration workflow

`db/schema.ts` is authoritative. Run `npm run db:generate` after an approved schema change and inspect the generated SQL in `drizzle/`. Deployment applies those ordered migrations through the Sites workflow. The API assumes migration is complete and never creates or alters tables during a request.

For an existing local V1 database, apply `drizzle/0001_free_corsair.sql` through the local D1 migration runner before starting V1.1. A fresh database applies `0000` then `0001`. Migration `0001` preserves development data, adds V1.1 columns/default versions, converts the exact old regional sample relationship, carries ORBIT evidence into general evidence summary, and corrects the one legacy sample allocation that exceeded duration.

Sample seeding uses `INSERT OR IGNORE` after the schema is ready. It does not perform schema migration and does not overwrite user records.

## Acceptance checks

- Routine non-ORBIT entry keeps optional classifications behind disclosure.
- The ORBIT pointer and keyboard paths persist `reportable: true` with a valid primary deliverable.
- Invalid scope, arrays, counts, time, and nested ORBIT shapes return structured validation errors.
- Create and update result types are explicit; failed saves keep the open draft.
- Stale updates return a conflict with the current version.
- Ten clearly marked development scenarios exercise Today, History, Project, organization/LEA, scope, and ORBIT projections from the same records.

## Explicitly deferred

R10 Today enhancements, R11 project-specific history, R12 expanded History search, R13 production fallback redesign, R14 Home cleanup, and R15 sample reset/hide controls remain deferred. Microsoft authentication, Graph, SharePoint lists, Power Automate, Outlook/Teams integration, uploads, CRM, AI summaries, and advanced dashboards are also deferred.
