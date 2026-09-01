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

## Hosted Deployment — GREEN

The verified GREEN build at commit `7e22e5909d1218c88a3be4deec45d3c61418da0f` is deployed to the existing owner-only Sites environment:

https://iu-work-tracker.gmacer.chatgpt.site/

Full detail: `docs/HOSTED_DEPLOYMENT_REPORT.md`.

- Sites version 3 deployed successfully from a clean checkout of the exact commit; the two known local-only untracked exclusions and `.env.local` were not included.
- The hosted Sites environment securely supplies the required public Microsoft/SharePoint identifiers and a server-only `ANTHROPIC_API_KEY` secret. The key is absent from client assets, rendered output, logs, and source.
- Hosted delegated Microsoft sign-in succeeded for Macer, Gregory with the existing redirect URI and only `User.Read` plus `Sites.ReadWrite.All`.
- The correct `IUWorkTrackerDEV` site resolved, all 10 existing lists were readable, and the footer reported `SharePoint DEV connected` with no prototype-provider fallback.
- One normal-UI Work Record smoke test persisted to `IU_Work_Records` with AppId, RecordVersion 1, ETag, Created, and Modified; it survived reload.
- One normal-UI Inbox Intelligence smoke test made exactly one successful hosted Anthropic request, rendered structured review output, persisted to `IU_Inbox_Intelligence`, survived reload, and passed `open→waiting→resolved→open` concurrency/version/timestamp checks.
- Only the approved compact source excerpt and structured fields were stored. No raw email/body/thread, Anthropic payload, or credential was persisted.
- Both exact synthetic items were permanently removed; exact lookups returned item-not-found/HTTP 404 and both active lists returned to their pre-test empty state.

This is a hosted DEV milestone, not production SharePoint authorization or migration approval.

---

## CRITICAL CURRENT PERSISTENCE STATE

`DelegatedSharePointDataProvider` IS NOW IMPLEMENTED AND LIVE-VERIFIED IN DEV.

See docs/SHAREPOINT_PROVIDER_INTEGRATION_REPORT.md for full detail. Summary:

- Normal application actions (Log Work: create and update) now persist through the existing `DataProvider` boundary into live DEV SharePoint `IU_Work_Records` when a Microsoft account is signed in under DEV configuration.
- Provider selection (`selectDataProvider()` in `lib/data-provider.ts`) is explicit and non-interactive: SharePoint activates only when DEV config is present AND a Microsoft account is already signed in; it never forces a sign-in prompt.
- The prototype `ApiDataProvider` was, at the time this verification ran, the default for every user who is not signed in and the automatic fallback on SharePoint failure. **Superseded by "Framework Migration — Patch 3A" below**: `ApiDataProvider`/Cloudflare D1 has since been removed; the default/fallback for those exact same cases is now the non-durable in-memory `MemoryDataProvider`, never a database.
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

## Inbox Intelligence V1 — GREEN

Authorized and implemented; one live AI test completed successfully. Full detail: docs/INBOX_INTELLIGENCE_V1_REPORT.md.

Summary:

- One new kiosk screen (`app/InboxIntelligence.tsx`) plus one new server route (`app/api/inbox-intelligence/route.ts`) implement PASTE EMAIL → ANALYZE WITH AI → HUMAN REVIEW → SAVE (session-only) → optionally CREATE WORK RECORD, entirely inside IU Work Tracker.
- The Anthropic API key (`ANTHROPIC_API_KEY`) is read server-side only — at the time this verification ran, via the Cloudflare Workers `env` binding; **since "Framework Migration — Patch 3A" below, via standard `process.env`**. Either way it is never a `NEXT_PUBLIC_*` variable. Verified structurally (client bundle grepped, zero matches) and live (the browser's network log showed only calls to the app's own `/api/inbox-intelligence` route, never to `api.anthropic.com`).
- AI output is never automatically persisted. Structured extraction uses the Anthropic TypeScript SDK's `messages.parse()` + a strict Zod schema; a response that fails validation is surfaced as an error, never fabricated.
- Persistence is session-only (`SessionInboxIntelligenceProvider`, in-memory) — no SharePoint list has been provisioned for this record type. A proposed `IU_Inbox_Intelligence` schema is documented in the report for a future phase's review, not provisioned.
- Creating a Work Record from an analyzed email reuses the existing, unmodified `DataProvider`/Log Work wizard path (`openLog()` + `buildWorkRecordDraftFromAnalysis()`); the user must still review and explicitly save.
- 89/89 automated tests pass; every test mocks the Anthropic client — no automated test makes a real API call.
- **One live, paid Anthropic API call was made** with a non-sensitive synthetic email (a made-up grant-deadline message, no real IU data). It succeeded end-to-end: structured extraction validated against the schema, every field (summary, priority, needsAttention, two correctly-dated action items, follow-up, tags, people/organizations/districts/projects, suggested Work Record) rendered correctly for review, and model/token diagnostics were captured (`claude-opus-5 · 2,373 in / 732 out tokens`). Confirmed live that nothing was persisted automatically at any layer (raw email, Inbox Intelligence record, or Work Record) — everything requires an explicit user click.
- An earlier attempt with a placeholder (non-real) key failed safely at authentication with no credential exposed, before the real key was supplied; only one call reached the model.

DO NOT request Mail.Read or other mailbox permissions. DO NOT integrate Microsoft Graph/Outlook. DO NOT provision the proposed SharePoint list without separate explicit instruction. DO NOT treat session-only persistence as durable — it is not.

---

## Inbox Intelligence V1.1 — GREEN

Authorized, implemented, provisioned manually, and live-verified in DEV. Full detail: `docs/INBOX_INTELLIGENCE_SHAREPOINT_REPORT.md`.

- `IU_Inbox_Intelligence` exists at list ID `892dbe47-6fa2-42f0-b9c4-1ed7a3664737` with the exact approved §23 schema: canonical built-in internal field `Title`, 24 approved custom columns, unique/indexed `AppId`, exact defaults/choices/date settings, and no application Lookup fields.
- `.env.local` configures `NEXT_PUBLIC_SHAREPOINT_IU_INBOX_INTELLIGENCE_LIST_ID`, so a signed-in DEV session selects `DelegatedSharePointInboxIntelligenceProvider`; unsigned/non-DEV contexts retain the session fallback.
- Live normal-UI verification passed: **Save to Inbox → Graph read-back → full application reload → open→waiting → waiting→resolved → resolved→open (reopen)**.
- `RecordVersion` advanced `1→2→3→4`; the SharePoint ETag advanced on every update. `Created` remained unchanged and `Modified` advanced each time.
- Action-item JSON and all extracted/matched relationship JSON arrays round-tripped exactly.
- No raw email/body/thread/signature, Anthropic payload, or API credential was persisted. The live persistence run made zero Anthropic API calls.
- The exact synthetic item was removed; its subsequent item lookup returned `itemNotFound`/HTTP 404 and the list returned to zero items.
- Production persistence/authentication remains out of scope; this is DEV-only and single-user-verified.

Do not reprovision or alter this list schema, broaden Entra permissions, or begin production cutover without explicit instruction.

---

## Dashboard Action Center (Patch 1) — implemented

Home now includes a compact "Action Center" panel, entirely a read view over already-durable Inbox Intelligence records loaded through the existing `InboxIntelligenceProvider` — no new persistence path, no new list, no AI call.

- Needs attention: records where the stored `analysis.needsAttention` flag is true and `status !== "resolved"` — never inferred from `status` alone.
- Waiting on: records with `status === "waiting"`.
- Open/Waiting/Resolved counts, deterministic due-date labels (`Due tomorrow`, weekday name, `Overdue`) only when an action item already has a real stored `dueDate`, and a small deterministic display limit (3 items/section) with a link into the full Inbox Intelligence screen.
- Dashboard load makes zero Anthropic requests — guarded by an automated test.
- Deliberately does not show "waiting N days": no dedicated status-transition timestamp exists yet to support that accurately, and a plain label beats false precision.

Do not add a Voice Intelligence, People/Organization/CRM, or a second Inbox management surface on Home as a result of this patch — those remain out of scope.

---

## Email Noise Torture Test (Patch 2) — implemented

A testing + extraction-hardening patch answering: can the Inbox Intelligence extraction pipeline distinguish real work signal from normal institutional email garbage (signatures, disclaimers, meeting-join boilerplate, quoted threads, forwarded headers, marketing footers)?

- New `tests/fixtures/email-noise.ts`: composable synthetic fixtures for 15 noise categories (external-sender warning, signature block, confidentiality notice, Teams/Zoom join boilerplate, social links, unsubscribe footer, image placeholders, auto-reply, calendar block, forwarded header, legal/security footer, quoted-thread wrapper) plus a `buildEmail(...)` joiner.
- New `tests/email-noise-filter.test.ts` (6 tests) and `tests/email-noise-torture.test.ts` (24 tests, covering the 20 required numbered cases plus combined scenarios) — all use the existing dependency-injected fake-client pattern from `tests/inbox-intelligence-analysis.test.ts`. **Zero real Anthropic API calls.**
- Hardening made, smallest-change only, after auditing the original prompt/pipeline against all 15 categories:
  1. New `lib/email-noise-filter.ts` (`stripDeterministicEmailNoise`) — conservative, whole-line-match-only preprocessing removing only the external-sender warning banner, standalone image-placeholder lines, and standalone unsubscribe/footer lines before the email reaches the model. Signatures, quoted threads, forwarded headers, and calendar/meeting blocks are deliberately left untouched (too judgment-dependent to strip safely). Wired into `analyzeEmailWithClaude` in `lib/anthropic-email-analysis.ts`; the `MAX_EMAIL_LENGTH` check still runs against the original pasted text.
  2. `SYSTEM_PROMPT` (now exported for testability) gained one added "Thread and boilerplate handling" paragraph: prioritize the newest authored message over quoted/forwarded content (never resurrect an old, already-addressed quoted request as a new action unless the current message reaffirms it); ignore signatures/disclaimers/meeting-join boilerplate/marketing-social footers as sources of actions, entities, tags, or deadlines; don't list a technology/platform/vendor name (Microsoft, Teams, Zoom, Facebook, LinkedIn, YouTube) as an entity merely because it appears in boilerplate; never list the same entity more than once.
  3. `normalizeEmailAnalysis` in `lib/inbox-intelligence-models.ts` now also deterministically dedupes `people`/`organizations`/`districts`/`projects`/`tags` (previously only `actionItems[].dueDate` was normalized) — a mechanical safety net independent of model compliance.
- Persisted `InboxIntelligenceRecord` SharePoint schema, status lifecycle, and concurrency model are unchanged.
- Known, intentionally deferred limitation: these tests validate the deterministic preprocessor and the surrounding pipeline (schema validation, normalization, entity matching, Work Record mapping) against hand-authored "desired correct" model responses — they cannot, without a real (forbidden) paid API call, prove the live model actually follows the hardened prompt. A hosted live smoke test remains a separate future authorization.

---

## Framework Migration — Patch 3A: real Next.js, D1 retired — implemented

A pure infrastructure conversion, no product/UI redesign. The app moved off `vinext` (Cloudflare's Vite-based Next.js-compatible runtime, targeting a genuine Cloudflare Worker) onto **real Next.js 16** (App Router, unchanged), so it can later be deployed to Vercel like any standard Next.js app.

- **Runtime**: `worker/index.ts` (the Cloudflare Worker entry point), `vite.config.ts`, and the `@cloudflare/vite-plugin`/`vinext`/`wrangler`/`@cloudflare/workers-types` dependencies are removed. `npm run dev`/`build`/`start` now run `next dev`/`next build`/`next start` directly — no custom server.
- **D1 retired as a runtime dependency, no replacement database introduced.** `app/api/records/route.ts` (the Cloudflare D1-backed prototype store) and the unused `db/`/`drizzle/` scaffolding are removed entirely. **SharePoint remains the only durable production data store** (docs/PRODUCT_VISION.md "Treat SharePoint as the intended institutional source of truth") — this was already true in practice; D1 never backed the real Microsoft-authenticated SharePoint path, only the prototype fallback.
- **Provider selection changed to eliminate the silent durable fallback**: `selectDataProvider()` (`lib/data-provider.ts`) now returns exactly two kinds — `"sharepoint"` (DEV Microsoft config present AND already signed in — unchanged) or `"memory"` (every other case: no config, config present but not yet signed in, or an init failure). The `"memory"` kind is the same in-memory, session-only `MemoryDataProvider` used for local dev/tests — never Cloudflare D1, never any other database. The Home storage banner and footer now say so plainly ("Preview mode — changes in this session are not saved. Sign in with Microsoft (top right) to connect SharePoint and save durably." / "Preview session active") instead of the old, more durable-sounding "Prototype data store connected."
- **`ANTHROPIC_API_KEY`** in `app/api/inbox-intelligence/route.ts` now reads from `process.env.ANTHROPIC_API_KEY` (standard Next.js server-side env access) instead of the Cloudflare Workers `env` binding via `cloudflare:workers`. Same guarantees as before: server-only, never `NEXT_PUBLIC_*`, never logged, never rendered, never called in tests.
- **Account menu identity source changed**: the ChatGPT/OpenAI Sites identity wiring from Patch 2.6 (`app/chatgpt-auth.ts`, the `oai-authenticated-user-*` header reading, `chatGPTUser`/`chatGPTSignOutHref` props) is removed. `app/DevMicrosoftConnection.tsx` is now a single unified account menu driven only by the app's own Microsoft-authenticated identity (MSAL `AccountInfo.name`/`.username`, no new Graph request) — "Signed in as [name] [username/UPN]" plus "Sign out", reusing the existing `logoutRedirect` path exactly. When DEV Microsoft config isn't present, the menu honestly shows "Not signed in" with no sign-in/out action, rather than any hosting-platform identity.
- Also removed as dead template scaffolding, never used by the app: `cloudflare-env.d.ts`, `examples/d1/`, `.openai/hosting.json`, `@openai/sites-vite-plugin`.
- Microsoft auth (`@azure/msal-browser`, authorization-code+PKCE, `sessionStorage` cache, `loginRedirect`/`acquireTokenSilent`/`acquireTokenRedirect`/`logoutRedirect`, delegated `User.Read` + `Sites.ReadWrite.All` only), SharePoint persistence (list IDs, `RecordVersion`, ETag/`If-Match`, conflict handling), and the Inbox Intelligence extraction pipeline (Patch 2/2.5) are all **unchanged** — this was infrastructure surgery only.
- **CURRENT FALLBACK HOST vs NEXT DEPLOYMENT TARGET** (at the time this migration patch ran): `https://iu-work-tracker.gmacer.chatgpt.site/` (ChatGPT/OpenAI Sites, commit `80eee7f`) was the live fallback host; this branch targeted Vercel as the next deployment target, not yet deployed anywhere. **Superseded**: Vercel deployment has since completed and succeeded — production is now `https://iu-work-tracker.vercel.app/`, with `main` as the normal source/deployment branch. The Sites deployment's current status going forward is not tracked here; treat Vercel as canonical production unless told otherwise.
- Tests: `tests/account-menu.test.tsx` rewritten against Microsoft identity (was ChatGPT identity); `tests/rendered-html.test.mjs` rewritten to start a real `next start` server and fetch it (was importing the Cloudflare Worker build output directly); `tests/data-provider-selection.test.ts` added for the browser-context provider-selection behavior. Baseline test count unchanged: **180 total = 179 vitest + 1 node test**.

Do NOT reintroduce Cloudflare D1, `vinext`, or any other durable database as a "prototype" persistence layer. Do NOT deploy to Vercel, change Microsoft Entra configuration, or change SharePoint configuration as a result of this section alone — those remain separate, explicit future steps.

---

## Voice Intelligence V1 — implemented (no durable persistence yet)

**Purpose**: the user records long, rambling voice notes while driving — work completed, meetings, follow-ups, ideas, decisions, questions, things worth remembering. Voice Intelligence turns a pasted transcript into independent, reviewable candidates. Same philosophy as Inbox Intelligence: **AI proposes, human reviews — AI never silently persists.** V1 deliberately begins with **paste transcript**, not microphone recording, audio upload, or transcription infrastructure (the user already has transcription elsewhere).

- New destination: sidebar "Voice Intelligence" and a Home command card, alongside the existing Inbox Intelligence entry — Home is not otherwise restructured.
- **Two states**: paste a transcript (`app/VoiceIntelligence.tsx`, textarea + "Analyze transcript") → AI-suggested review workspace showing each candidate as its own compact, editable card. No AI request occurs until that button is explicitly clicked — loading the Voice screen, like loading Home, costs zero Anthropic requests.
- **Segmentation, not summarization**: the system prompt (`lib/anthropic-voice-analysis.ts`) instructs the model to treat one transcript as potentially many unrelated topics and propose each as an independent candidate — never collapsed into one summary.
- **Candidate types**: `COMPLETED_WORK`, `ACTION`, `PERSON`, `ORGANIZATION`, `DISTRICT`, `PROJECT`, `IDEA`, `DECISION`, `QUESTION`, `KNOWLEDGE` (`lib/voice-intelligence-models.ts`). Each candidate is `{ type, title, detail, sourceExcerpt, durationText }` — the smallest useful model, not premature enterprise modeling. These are proposals only: a `PERSON` candidate does not create a Person, a `KNOWLEDGE` candidate is not institutional knowledge, none of the ten types write to any database.
- **Completed work vs. future action**: the prompt explicitly requires distinguishing what already happened (`COMPLETED_WORK`) from what still needs to happen (`ACTION`) — never describing a future action in past tense or turning completed work into a task.
- **Duration is EXPLICIT ONLY, with a deterministic safeguard**: the model may only propose `durationText` when the transcript states an explicit approximate duration ("about an hour", "30 minutes"), never inferred or estimated. `isDurationSupportedByTranscript()` / `normalizeVoiceAnalysis()` then deterministically re-verify every returned `durationText` actually appears (normalized substring match) in the original transcript — an unsupported/invented duration is silently dropped to `null` before it ever reaches the review UI. No natural-language duration parser was built; the goal is explicit-only, not comprehensive.
- **Deterministic dedupe**: `normalizeVoiceAnalysis()` also drops an exact repeated `(type, normalized title)` candidate — the smallest guard against obvious duplication, not semantic deduplication.
- **Raw transcript is never persisted**: it lives only in React state while the screen is open (`useState`, nothing else) — never SharePoint, never a Work Record, never Inbox Intelligence, never `localStorage`/`sessionStorage`/`IndexedDB`/cookies, never logged server-side (the API route logs only the caught error object, matching `app/api/inbox-intelligence/route.ts`'s existing discipline), never sent to analytics.
- **Server boundary**: `POST /api/voice-intelligence` (`app/api/voice-intelligence/route.ts`) follows the exact Inbox Intelligence pattern — `ANTHROPIC_API_KEY` read via `process.env` (server-only, never `NEXT_PUBLIC_*`), Zod strict schema (`VoiceAnalysisSchema`) + `zodOutputFormat`, same dependency-injected `analyzeTranscriptWithClaude(rawTranscript, client)` shape as `analyzeEmailWithClaude` so tests never touch the real Anthropic SDK.
- **Review UI**: select/deselect (local-only — "Selected candidates are ready for review. Nothing has been saved yet."), edit title/detail, change type, remove a duration, remove a candidate entirely. No button named "Save" exists in V1 — "Analyze another transcript" is the only reset action.
- Transcript length: `MAX_TRANSCRIPT_LENGTH = 40000` characters (`lib/voice-intelligence-config.ts`), enforced client- and server-side, never silently truncated — an over-limit transcript is rejected with a message to split it.

**Next planned step (not built yet)**: approved `COMPLETED_WORK` candidates connect into the existing Universal Work Record pathway (`openLog()` / `createWorkRecord`), the same way Inbox Intelligence's "Create Work Record" already works — no new persistence architecture, no SharePoint schema change anticipated. Do not build this, People/Organization/Project/Knowledge persistence, or microphone/audio capture without separate explicit instruction.

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

Hosted deployment GREEN at `https://iu-work-tracker.gmacer.chatgpt.site/` for commit `7e22e5909d1218c88a3be4deec45d3c61418da0f`: Microsoft delegated sign-in, SharePoint Work Record persistence, Inbox Intelligence V1 analysis, and durable Inbox Intelligence V1.1 persistence were all live-verified through the hosted normal UI. Exactly one Anthropic request was made. Both synthetic records were permanently removed, exact lookups returned item-not-found/HTTP 404, and both lists are empty. See `docs/HOSTED_DEPLOYMENT_REPORT.md`.

The prototype `ApiDataProvider` remains the default/fallback provider by design.

---

## Next Agent

WAIT FOR EXPLICIT USER INSTRUCTION.

Do not begin any further SharePoint integration phase (reference-data lists, production `Sites.Selected` planning, broader concurrency/multi-user testing) automatically.
