# IU Work Tracker V1 Audit

**Audit date:** 2026-08-26  
**Audit scope:** Product, Quick Log UX, universal record, ORBIT reporting layer, LEA and project relationships, kiosk experience, Today, sample data, and SharePoint/Microsoft Graph readiness.  
**Change boundary:** Audit only. No product recommendations were implemented.

## Overall assessment

V1 proves the core product idea. The application feels like a focused work tool, the universal record covers the most common activity/time/outcome relationships, non-ORBIT work fits naturally, sample data is honestly labeled, and summaries are calculated from shared records instead of copied into parallel modules.

V1 does **not yet fully satisfy “Log it once. Use it everywhere.”** The most immediate blocker is that the visible ORBIT switch cannot be activated in Quick Log. More broadly, the data-provider boundary currently covers Work Records only while projects, organizations, categories, deliverables, and configuration are imported directly into the UI as hard-coded arrays. The proposed data model and SharePoint schema also contain entities and metadata that do not exist in the runtime model. Those gaps would force avoidable UI changes during a future SharePoint transition.

**Verdict:** Promising V1 concept and shell; suitable for continued prototype work, but not yet trustworthy for real production reporting data. Complete a focused V1.1 hardening pass before Graph/SharePoint integration.

**Priority count:** 1 Critical, 8 High, 4 Medium, 2 Low recommendations.

## Audit method and observed evidence

- Reviewed the runtime TypeScript model, D1 schema, API mapping, provider contract, sample records, product documents, ORBIT mapping, and SharePoint proposal.
- Exercised all five Quick Log steps without saving an audit record.
- Reviewed Home, Today, History, Projects, and STEM/ORBIT at 1440×900 and 639×734.
- Verified desktop screen overflow: Home, Today, Projects, and STEM/ORBIT each fit within the 800-pixel application workspace.
- Verified mobile screen overflow: Home, Today, and History fit; Projects required 1,462 pixels and STEM/ORBIT 1,195 pixels inside a 674-pixel workspace.
- Verified Quick Log step 2 required internal scrolling on mobile (636 pixels of content in a 504-pixel panel) and slight scrolling on desktop.
- Verified four visible “Log work” actions on the desktop Home screen.
- Verified the Quick Log dialog opens with keyboard focus remaining on a background navigation button rather than entering the dialog.
- Verified the ORBIT checkbox remains unchecked when the visible switch is activated; the visible switch is not functionally connected to the hidden checkbox.

---

# 1. What Works Well

## Product and kiosk model

- Home provides immediate orientation, a clear purpose statement, a compact daily snapshot, and task-oriented navigation.
- The desktop application is finite: the principal screens fit within the workspace rather than behaving like an endless marketing page.
- The navigation vocabulary is understandable within five seconds: Home, Today, History, Projects, and STEM/ORBIT.
- Quick Log is clearly positioned as the primary workflow and uses five mental steps rather than one giant form.
- Optional outcome, reach, follow-up, project, organization, category, and ORBIT fields do not block a minimal record.
- Mobile uses a bottom task bar and floating Log Work action, which keeps the primary workflow reachable.

## Universal record fundamentals

The runtime Work Record supports:

- stable client-generated `appId`
- title, activity date, activity type, description, and detailed notes
- precise duration in minutes
- zero, one, or multiple project, organization, and category IDs
- educator/leader, student/family, workforce/community, and other reach counts
- output, outcome, next step, follow-up flag, and follow-up date
- optional nested ORBIT data
- created/modified timestamps, sample flag, status, and UI sync state

The same Work Record collection drives Home, Today, History, Projects, and ORBIT summaries. Project time and record counts are calculated from relationships rather than duplicated.

## ORBIT remains subordinate

- `orbit.reportable` defaults to false.
- Non-ORBIT work saves without an ORBIT deliverable or ORBIT-specific time.
- ORBIT data is nested under the Work Record rather than controlling the general activity taxonomy.
- Minutes remain authoritative and STEM days are derived using the configured 420-minute display rule.
- Primary and supporting deliverable fields exist in the runtime record.
- Reach and organizations live in the universal record, allowing future ORBIT projections to reuse normal work data.

## SharePoint-positive choices

- No component calls `localStorage` or `sessionStorage`.
- Stable application IDs are independent from D1 or future SharePoint item IDs.
- Multi-valued relationships use arrays of stable IDs rather than repeated display-name text.
- `app_id` has a uniqueness constraint, providing a foundation for idempotent create retries.
- The SharePoint proposal correctly warns against excessive multi-lookup fields and favors AppId references for Graph reliability.
- A save error leaves the Quick Log modal open, preserving the current in-memory draft.

## Sample data

All six seed records are marked `isSample: true` and visibly labeled in Home, Today, History, and the ORBIT workspace. The set includes:

- five ORBIT-reportable records and one non-ORBIT record
- two named LEAs plus one regional pseudo-organization
- a partner interaction
- student reach
- five project-linked records and one record without a project
- district planning, AI professional learning, internal planning, competition planning, ecosystem partnership, and makerspace work

The ORBIT workspace explicitly states that the totals are development samples and not IU29 performance data.

## Quick Log scenario audit

The timing assessment below assumes the ORBIT switch defect is repaired. “Routine” means the necessary project and organization already exist.

| Scenario | Natural V1 path | 30–60 seconds? | Audit finding |
| --- | --- | --- | --- |
| District planning meeting | District meeting + LEA + optional project/category | Yes | Strong fit; avoid the silent District Meeting default causing accidental classification. |
| Professional development | Professional learning + organization + educator reach | Yes | Strong fit. |
| Classroom / educator support | Classroom support + LEA + educator reach | Yes | Strong fit; outcome can remain optional. |
| STEM competition planning | Project planning + competition project | Yes if project exists | Good universal-record fit; static project list is the constraint. |
| Student competition/event | Student program + project + student reach | Borderline | Natural but reach and relationship decisions make step 2 dense. |
| Makerspace visit | Student program + Makerspace + LEA | Yes if project exists | Strong fit. |
| STEM Ecosystem partner meeting | Partner meeting + partner + ecosystem project | Borderline | Natural when the partner is already present; current list has only one partner. |
| Workforce outreach | Partner meeting or Other | Borderline/no | No workforce classification or broad organization catalog; requires a workaround. |
| State STEM PoC meeting | Other + future state organization + ORBIT time | No in current UI | No state-meeting type/organization and the ORBIT switch is blocked. |
| STEELS support | District meeting or classroom support + STEELS | Yes except ORBIT blocker | Strong taxonomy fit. |
| AI professional learning | Professional learning + AI project/category | Yes except ORBIT blocker | Exact V1 fit. |
| Resource development | Resource development + optional project | Yes | Strong non-ORBIT fit; general evidence is missing. |
| IU internal meeting | Internal planning + IU | Yes | Exact V1 fit and demonstrates non-ORBIT work. |
| Grant planning | Project planning or Other | No without workaround | No grant category and no project creation/reference-data provider. |
| Technology / application development | Other | No without workaround | No technology/development activity or category in the runtime reference set. |

**Quick Log conclusion:** The five-step structure is obvious and kiosk-like. A simple non-ORBIT entry can plausibly fit 30–60 seconds. The experience becomes slower for relationship-heavy activities because step 2 exposes every organization, project, category, and reach input at once. Several common scenarios fall into “Other” because V1 reference data is both incomplete and hard-coded.

---

# 2. Critical Issues

## C1 — ORBIT cannot be enabled in Quick Log

The styled switch uses a `div` wrapper around a transparent checkbox and a decorative span. The visible switch is not a label or button connected to the checkbox. Browser testing confirmed that activating the visible switch leaves `orbit.reportable` false.

This blocks creation of new ORBIT-reportable records through the primary workflow. It also makes the displayed ORBIT sample dashboard misleading because users cannot produce equivalent records themselves.

Priority and fix details are in recommendation R1.

---

# 3. Important Improvements

## I1 — The provider abstraction stops at Work Records

`DataProvider` exposes only `listWorkRecords` and `saveWorkRecord`. Components import `PROJECTS`, `ORGANIZATIONS`, `CATEGORIES`, `DELIVERABLES`, and the reporting-day value directly from `lib/models.ts`. A `SharePointDataProvider` cannot replace those sources without changing screen and form components.

## I2 — The documented and runtime models disagree

The documents describe contacts, evidence records, configuration, provider item IDs, regional scope, start/end time, schema version, and sync fingerprint. The runtime Work Record and D1 table omit most of them. Conversely, runtime fields such as `isSample` and UI-only `syncState` are not clearly separated from institutional fields.

## I3 — Regional scope is modeled as an organization

“Regional / Multiple LEAs” is a pseudo-organization. This supports a human label but cannot identify which LEAs were served, and it conflates engagement scope with a canonical organization. Accurate distinct-LEA counts by quarter are impossible for those records.

## I4 — ORBIT fields are only partially operable

Supporting deliverables exist in the record and samples but have no Quick Log control. A reportable record is not required to have a primary deliverable. STEM PoC and TaC minutes can exceed or contradict total activity duration without review. TaC days and reporting quarter are described but not derived by shared domain utilities.

## I5 — Quick Log makes high-cost defaults and too many step-2 decisions

Every new record silently starts as “District meeting” and 60 minutes. The duration default is useful; the activity-type default creates a data-quality risk. Step 2 places organizations, projects, nine categories, and four reach counters in one scrolling surface. The user must also visit the ORBIT step for every activity, even when the work is clearly non-STEM.

## I6 — Server validation and concurrency are insufficient for institutional data

The API validates only appId, title, date, activity type, and positive duration. It trusts client timestamps, sample status, arrays, counts, ORBIT time, and supporting values. Updates are blind last-write-wins operations with no ETag, version, sync fingerprint, or stale-edit detection.

## I7 — The database schema has multiple authorities

The table definition exists in `db/schema.ts`, a Drizzle migration, a manual migration, and a runtime `CREATE TABLE` string inside the API route. These can drift. Future maintainers may update one path while deployment or runtime initialization uses another.

## I8 — The modal is not keyboard-safe

When Quick Log opens, focus remains on the background Log Work button. The dialog does not establish initial focus, trap focus, restore focus intentionally, or expose descriptive names for the numbered step buttons. This makes a keyboard or screen-reader user navigate hidden background controls while a modal is open.

---

# 4. Nice-to-Have Improvements

- Reduce four visible desktop Log Work entry points to one persistent primary action plus one contextual Home action.
- Offer recently used projects and organizations before the full reference list; this is a convenience layer, not a new module.
- Separate development samples from user-entered records with a reset/hide-samples control in prototype environments.
- Let activity type/category reference data include inactive and ordered states so future vocabulary changes do not rewrite historical records.
- Consider a compact record review before save only for unusually complex ORBIT records; do not add review friction to routine work.

---

# 5. Data Model Risks

| Risk | Severity | Finding |
| --- | --- | --- |
| Evidence is ORBIT-only | High | `orbit.evidence` cannot capture evidence for non-ORBIT work, resource development, grants, internal projects, or district outcomes. A universal evidence summary/reference belongs at Work Record level; ORBIT should project it. |
| Contacts absent from runtime | High | Documents promise contact relationships, but `WorkRecord`, D1, and Quick Log have no `contactIds`. Partner history cannot later be derived without remapping old records. |
| Scope conflated with organization | High | The regional pseudo-organization cannot support exact LEA counts. Use an explicit scope enum plus zero/many real organization IDs. |
| Reference values are code constants | High | Activity types, categories, projects, organizations, deliverables, and reporting-day rules are not configurable records. Historical meaning and SharePoint mapping will be brittle. |
| Supporting deliverables are write-inaccessible | High | The field exists, but only seed data can populate it. The UI and data model present different capabilities. |
| `status` is inert | Medium | The model supports `draft` but Quick Log always creates `complete`; Today cannot show unfinished records. Either implement a minimal draft lifecycle or remove the promise from V1-facing copy. |
| `detailedNotes` is inert | Medium | The field persists but has no UI. It is not redundant if reserved for advanced detail, but the boundary with `description` should be defined. |
| Timestamps are client-authored | High | `createdAt` and `modifiedAt` depend on browser time. Provider-managed timestamps should be authoritative, with optional app timestamps clearly distinguished. |
| Reach lacks counting semantics | Medium | Counts can be reused across views, but the model does not define whether repeat participants, families, or cross-category totals may be summed. Reporting could overstate reach. |
| ORBIT time can contradict duration | High | STEM PoC plus TaC minutes have no relation/validation rule against activity duration. Explicit overlap rules are required before official aggregation. |
| Sample state is mixed with records | Low | Samples are well labeled but live in the same production-shaped table and can be edited. Prototype seeding should be isolated or resettable before real use. |

### Fields that should be derived

- total participants
- STEM PoC days and TaC days
- school year and reporting quarter
- distinct LEAs served
- project/organization/category time totals
- project and district engagement counts
- created/modified provider timestamps where the provider supplies them

### Fields that should remain optional

- description and detailed notes
- project, organization, contact, and category relationships
- reach counts
- output, outcome, evidence, next step, and follow-up
- ORBIT classification and all ORBIT detail for non-reportable work
- supporting deliverables

### Fields that should become configurable reference data

- activity types
- work categories and topics
- projects, organizations, and contacts
- deliverables and active reporting year
- school-year/quarter boundaries
- minutes per reporting day and role-specific accounting rules
- reach labels and counting guidance

---

# 6. SharePoint Integration Risks

## Positive readiness

- No direct browser-storage dependency exists in components.
- Client-generated UUIDs support offline creation and retry identity.
- Relationships are ID-based arrays.
- The SharePoint proposal anticipates internal-field discovery and avoids lookup-heavy designs.

## Risks requiring resolution before Graph work

1. **Reference data bypasses the provider.** Projects, organizations, categories, deliverables, and settings must be loaded through provider/domain services before a SharePoint provider can be substituted cleanly.
2. **The contract is too small.** It lacks entity lookup, record get, paged/range queries, configuration retrieval, sync result metadata, and explicit create/update semantics.
3. **JSON arrays trade lookup pain for search pain.** The proposed AppId-array strategy is reasonable for Graph writes, but SharePoint cannot efficiently filter inside multiline JSON. Quarter reports by district/project would require date-bounded Work Record reads plus client-side filtering or a deliberately maintained reporting index. This tradeoff must be documented and volume-tested.
4. **No provider item ID or ETag is represented.** A future adapter needs provider metadata outside the business fields to support `If-Match`, duplicate resolution, and safe retries.
5. **Blind upsert is not conflict-safe.** AppId uniqueness prevents duplicate creates but not stale overwrites.
6. **Timestamp ownership is unclear.** SharePoint `Created`/`Modified`, app timestamps, and offline draft timestamps need separate semantics.
7. **Schema/version metadata is missing at runtime.** The proposal includes `SchemaVersion` and `SyncFingerprint`, but V1 records do not.
8. **Fallback semantics differ from durable storage.** If API loading fails, the app swaps to an in-memory provider. It labels the session as prototype mode, but a save can still look successful while disappearing on reload. Production must never present that as institutional persistence.
9. **Canonical entity coverage is incomplete.** The actual organization set is a seed constant, not an entity catalog. Existing records using a pseudo-regional organization will need migration.
10. **There are multiple schema sources.** Graph field mapping will become fragile if runtime SQL, Drizzle, documentation, and SharePoint provisioning evolve independently.

The current frontend can preserve its screen structure during a SharePoint move, but it cannot make a clean provider swap yet. Without V1.1 contract work, Quick Log, Projects, History labels, and ORBIT calculations would all require edits.

---

# 7. UX / Kiosk Risks

## Home

**Works:** Strong orientation, no desktop scrolling, compact situational awareness, clear tasks.  
**Risk:** Four visible Log Work actions on desktop dilute hierarchy rather than strengthen it. The header, side rail, hero, and command card repeat the same action.

## Quick Log

**Works:** One mental task per step; minimal required core; outcomes and ORBIT use progressive disclosure.  
**Risks:** Broken ORBIT switch; dense step 2; silent District Meeting default; no modal focus management; numeric step buttons lack descriptive accessible names; mobile step 2 scrolls; Save & Log Another is hidden on mobile; all work must pass through the ORBIT step.

## Today

**Works:** It answers the basic daily question with activities, total time, and open-follow-up count. It fits both desktop and the tested mobile viewport.  
**Risks:** It does not show districts touched, projects touched, ORBIT-reportable time, or drafts. Follow-ups are only a count; on mobile the panel containing follow-up details is hidden. Today therefore reports activity but is not yet a complete operational surface.

## History

**Works:** Compact searchable record list; samples are individually labeled; edit access is obvious.  
**Risks:** Search covers only title, description, and activity type—not projects, organizations, categories, outcomes, or dates. Mobile hides relationship and ORBIT columns, reducing “use everywhere” visibility. Multi-relationships are represented by only the first project and organization.

## Projects

**Works:** Counts and time are derived from Work Records. Projects are containers rather than separate activity databases.  
**Risks:** “View connected work” opens unfiltered History, so it does not actually show the selected project. Projects are hard-coded, cannot load through the provider, and lack runtime organization/contact/milestone metadata. The mobile page is more than twice the workspace height.

## STEM / ORBIT

**Works:** Clearly labeled reporting layer; explicit sample disclaimer; precise minutes-to-days calculation; primary-deliverable distribution.  
**Risks:** It does not show quarter, LEAs served, supporting alignment, TaC days, evidence coverage, or data-quality review. Mobile requires substantial scrolling. The screen claims quarter is derived, but no shared derivation exists in code.

## Kiosk-loop conclusion

The desktop loop is recognizable: choose a task, complete it, and return to a workspace. Quick Log currently returns to Today after save rather than the command center; that is defensible for immediate confirmation, but it should be an intentional product decision. Secondary workspaces behave more like compact dashboards than single-purpose tools, which is acceptable at desktop scale but starts to break the kiosk model on mobile Projects and ORBIT.

---

# 8. Recommended V1.1 Changes

These recommendations intentionally exclude major features and external integrations.

| ID | Problem | Impact | Recommendation | Priority | Files/components affected |
| --- | --- | --- | --- | --- | --- |
| R1 | The visible ORBIT switch cannot change the checkbox. | Users cannot create ORBIT-reportable work through Quick Log; core reporting data is blocked. | Restore a semantic label/button connection for the switch, verify mouse/touch/keyboard behavior, and add an interaction test that enables ORBIT and saves a reportable record. | **Critical** | `app/IUWorkTracker.tsx` (`LogWizard`), `app/globals.css` (`.switch`), tests |
| R2 | Projects, organizations, categories, deliverables, and settings bypass `DataProvider`. | A SharePoint adapter would require changes across Quick Log and every summary screen. | Expand the domain/provider contract to retrieve reference entities and reporting configuration; pass provider-loaded data into screens instead of importing constants. Keep seed implementations for V1.1. | **High** | `lib/data-provider.ts`, `lib/models.ts`, `app/IUWorkTracker.tsx`, provider tests |
| R3 | Regional/all-district scope is stored as a pseudo-organization. | Quarterly LEA counts and “districts not served” analysis will be inaccurate. | Add explicit `engagementScope` (`none`, `specific`, `regional`, `allDistricts`) and keep `organizationIds` for real canonical entities only. Define migration behavior for the current regional seed. | **High** | `lib/models.ts`, `db/schema.ts`, API mapping/migration, `LogWizard`, data-model and SharePoint docs |
| R4 | Runtime Work Record omits general evidence, contacts, provider metadata, and schema version while documents promise them. | Old records would need later remapping, and non-ORBIT impact evidence cannot be reused everywhere. | Reconcile the actual and documented models. Add only the minimal structural fields needed now: general evidence summary/reference IDs, contact IDs, schema version, and provider metadata kept outside business fields. Do not add a visible CRM or upload system. | **High** | `lib/models.ts`, `db/schema.ts`, API route/migration, `docs/DATA_MODEL.md`, `docs/SHAREPOINT_SCHEMA_PROPOSAL.md`, advanced-detail portion of `LogWizard` |
| R5 | ORBIT supporting alignment, primary-deliverable validity, quarter, and time rules are incomplete. | Official reporting projections can be incomplete or internally contradictory. | Repair R1, expose optional supporting deliverables behind ORBIT disclosure, require a primary deliverable only when reportable, define PoC/TaC overlap rules, and implement tested pure derivation utilities for school year, quarter, and reporting days. | **High** | `lib/models.ts`, new domain calculation module, `LogWizard`, `Orbit`, API validation, `docs/ORBIT_MAPPING.md`, tests |
| R6 | Quick Log silently defaults every record to District Meeting and step 2 presents all relationship decisions at once. | Routine records can be misclassified; complex entries exceed the intended 30–60 seconds, especially on mobile. | Use an explicit/recent activity-type choice rather than a silent district default. Keep organization and project prominent, collapse categories/reach behind “Add classification/reach,” and let a clear non-ORBIT path save without extra reporting controls. Do not add fields. | **High** | `emptyRecord`, `LogWizard`, `app/globals.css`, interaction tests |
| R7 | The API trusts client timestamps and nested values and performs last-write-wins updates. | Invalid counts/arrays/ORBIT values and stale edits can corrupt institutional data; retry behavior is unsafe. | Add shared runtime validation, non-negative/invariant checks, server-owned modified time, immutable created time, and a version/ETag-style expected value for updates. Return structured save/conflict results through `DataProvider`. | **High** | `app/api/records/route.ts`, `lib/data-provider.ts`, `lib/models.ts`, D1 schema/migration, API tests |
| R8 | D1 schema exists in runtime SQL, Drizzle schema/migration, and a manual migration. | Schema drift is likely and future SharePoint mappings may target the wrong shape. | Select Drizzle schema + generated migrations as the single D1 authority. Remove request-time schema creation after migration behavior is verified; keep seeding separate and explicit. | **High** | `db/schema.ts`, `drizzle/**`, `migrations/**`, `app/api/records/route.ts`, deployment documentation |
| R9 | Quick Log does not manage modal focus or accessible step names. | Keyboard and screen-reader users can remain in background navigation and cannot understand numbered steps. | Move focus into the dialog, trap and restore focus, support Escape with draft protection, label each step with its title/current state, and test keyboard-only completion. | **High** | `LogWizard`, `app/globals.css`, accessibility/interaction tests |
| R10 | Today shows follow-up count but not actionable mobile follow-ups, drafts, or core context touched today. | The screen only partially answers the daily operational question. | Add a compact context line for distinct districts, projects, and ORBIT time; show draft count only if drafts are supported; make the next one or two follow-ups actionable on mobile. Avoid charts and additional dashboard modules. | **Medium** | `Today`, shared derivation helpers, responsive CSS, tests |
| R11 | Project “View connected work” opens all History records. | The control promises a project-specific view but breaks user trust. | Pass a project filter into History or render a focused connected-record panel. Continue deriving all project summaries from Work Records. | **Medium** | `Projects`, `History`, view/filter state, tests |
| R12 | History search and labels use only a subset of the universal relationships. | Users cannot reliably reuse records by district, project, category, outcome, or reporting classification. | Extend the client-side query projection to resolved entity names, categories, outcome, and ORBIT code; show a compact multi-relationship indicator rather than only the first item. | **Medium** | `History`, entity-resolution utilities, tests |
| R13 | In-memory fallback saves disappear on reload while appearing successful in the current session. | Users may mistake a preview save for durable storage during outages. | In durable/production mode, treat fallback as read-only or mark every fallback-saved record “Not persisted” with an explicit retry path. Keep in-memory fallback only for clearly labeled development use. | **Medium** | `lib/data-provider.ts`, save state in `IUWorkTracker`, storage banner/status, tests |
| R14 | Four desktop Log Work actions compete on Home. | Repetition weakens hierarchy and adds visual density. | Keep the persistent side-rail action and one prominent Home action; remove redundant header/command duplication after checking mobile reachability. | **Low** | `Header`, `SideNav`, `Home`, responsive CSS |
| R15 | Development sample records share the production-shaped collection and can be edited. | Prototype exploration can distort the demonstration baseline and later contaminate real reports. | Add a development-only hide/reset mechanism or seed namespace before real data is introduced; retain visible sample labels. | **Low** | sample seed/provider initialization, sample banner, tests |

## Recommended V1.1 scope

Implement **R1–R9 only** as the focused V1.1 trust pass:

1. Repair and test ORBIT entry.
2. Complete the provider boundary for reference/config data.
3. Correct LEA scope semantics and reconcile the runtime/documented model.
4. Add minimal ORBIT derivation and validation utilities.
5. Reduce Quick Log decisions without adding visible fields.
6. Add validation, version-aware updates, and one schema authority.
7. Make the modal keyboard-safe.

Treat R10–R13 as the next refinement queue after those foundations are stable. R14–R15 are optional cleanup.

---

# 9. Items Explicitly Deferred

The audit does not recommend implementing these in V1.1:

- Microsoft Entra authentication
- Microsoft Graph or SharePoint list integration
- SharePoint list provisioning
- Power Automate flows
- Outlook, calendar, or Teams integration
- AI summaries or classification
- advanced dashboards or Power BI-style analytics
- CRM features
- document/file upload storage
- official ORBIT export/submission
- reminders/notifications beyond the existing follow-up representation
- multi-user permissions or leadership portals
- automated school/district directory import
- grant management, project management, or competition modules
- dozens of settings or an administrative configuration console

The next implementation should make the existing core trustworthy, not broaden the product surface.
