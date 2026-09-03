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

**Next planned step at the time this section was written** (superseded — see "Voice → Universal Work Record Handoff (Patch 5)" below): approved `COMPLETED_WORK` candidates connect into the existing Universal Work Record pathway. That step is now implemented.

**VOICE INTELLIGENCE V1 LIVE TORTURE TEST: GREEN.** Three explicitly authorized live Anthropic calls were made against real (synthetic, non-sensitive) transcripts: a messy multi-topic ramble segmented correctly into independent candidates; completed work stayed distinct from future actions and an explicit duration ("about 45 minutes") survived; vague duration language ("really long", "took forever", "spent a while") correctly did not become a duration. Repeated actions deduplicated reasonably, an irrelevant personal aside was ignored, and useful project/entity context surfaced without obvious pollution. 3/3 PASS.

---

## Voice → Universal Work Record Handoff (Patch 5) — implemented

A `COMPLETED_WORK` Voice candidate can become the starting point for a **new** Universal Work Record — reusing the exact existing form and save path Inbox Intelligence's "Create Work Record" already uses. No second Work Record form, no second persistence pathway, no new SharePoint schema.

- **`COMPLETED_WORK` only**: the "Log as work" action appears solely when a candidate's *current* type (after any human edit) is `COMPLETED_WORK`. Changing a candidate's type away from `COMPLETED_WORK` removes the action immediately; changing another candidate's type *to* `COMPLETED_WORK` adds it. Eligibility is computed live from review state, not from the original AI output.
- **Human-reviewed state is authoritative**: `buildWorkRecordDraftFromVoiceCandidate()` (`lib/voice-intelligence-work-record.ts`) receives whatever candidate object the UI currently holds — if the user edited the title, detail, type, or duration before clicking, the edited values are what reach the Work Record form. The function has no notion of "original AI output" at all.
- **`Log as work` does not persist.** It only calls the existing `openLog(draft)` — the same function `app/InboxIntelligence.tsx` already calls for "Create Work Record" — which opens the existing Log Work wizard with `draft` as its starting state. No `createWorkRecord`/`updateWorkRecord` call, no SharePoint call, no `localStorage`/`sessionStorage` write, and no Anthropic call happens from this click. The Voice candidate's own review-state entry is untouched.
- **Existing Work Record form owns final review; existing save path owns persistence.** Every prefilled field remains fully editable in the wizard exactly as if a human had typed it — nothing is visually or functionally locked. Durable persistence only happens when the user reaches the wizard's own "Save & done"/"Save & log another", which goes through the same unmodified `DataProvider.createWorkRecord()` (SharePoint when signed in, the non-durable in-memory provider otherwise — Preview mode is untouched and not special-cased inside Voice Intelligence).
- **Prefill mapping is deliberately narrow**: only `title` → Work Record `title` and `detail` → Work Record `description`. `sourceExcerpt` is never copied into any Work Record field — it is review provenance, not work-log content.
- **No automatic category mapping**: `activityType` is left at the form's normal blank default; the user must choose one (the wizard's own "Continue" gate already requires this before advancing past step 1 — unchanged, existing behavior).
- **No automatic relationship mapping**: districts/organizations/projects/contacts are left empty. Even when the same transcript also produced separate `DISTRICT`/`PROJECT`/etc. candidates, Patch 5 does not attach them — candidate-to-candidate relationship inference is explicitly out of scope for this patch.
- **No automatic reporting/ORBIT mapping**: `reach`, `orbit.reportable`, and all ORBIT fields are left at the form's normal defaults.
- **Duration — two safety layers, both must pass**: (1) Patch 4's existing transcript-support check (unchanged) must already have kept `durationText` non-null; (2) `parseDeterministicDurationMinutes()` (`lib/voice-intelligence-work-record.ts`) then converts only unambiguous explicit forms — a plain number/decimal plus hour(s)/minute(s)/hr(s)/min(s), with a leading `about`/`around`/`approximately`/`roughly`/`maybe` ignored for the conversion, plus the fixed idiom "a/an hour" → 60. Word numbers ("two hours") and vague phrases ("a while", "most of the morning", "a couple hours") deliberately do not convert. If either layer fails, the Work Record's duration is left at its normal default (60 minutes, same as a manually opened blank entry) rather than guessed.
- **No new AI call**: the handoff is pure deterministic client-side mapping from an already-reviewed candidate.
- **Other candidate types** (`ACTION`, `PERSON`, `ORGANIZATION`, `DISTRICT`, `PROJECT`, `IDEA`, `DECISION`, `QUESTION`, `KNOWLEDGE`) remain transient review-only intelligence in this patch — no destination was added for them; their own pathways are a separate, future, explicit design.
- Tested end-to-end: a mocked-fetch Voice analysis → "Log as work" → the real Log Work wizard opens prefilled → a real `MemoryDataProvider.createWorkRecord()` (spied, not mocked) is called exactly once on explicit "Save & done" — proving the wizard and save path are the SAME ones every other Work Record creation already uses, not a parallel implementation.

**VOICE → WORK RECORD → SHAREPOINT PRODUCTION PIPELINE: GREEN.** Voice Intelligence extraction, the "Log as work" handoff, and durable SharePoint persistence via the existing Work Record save path together form one verified, working pipeline in production.

---

## Meeting Notes V1 — implemented (transient, no durable persistence yet)

A third intake/review workspace (`app/MeetingNotes.tsx`), alongside Inbox Intelligence and Voice Intelligence — same philosophy: **AI proposes, human reviews, nothing is silently persisted.** One continuous screen (not a paste/review phase swap like Voice): meeting details, agenda, and general notes stay visible and editable throughout; a Meeting Intelligence review section and Draft Minutes simply append below once an explicit analysis has run.

- **Meeting content is durable as of Patch 6B** — see "Meeting Notes durability (Patch 6B)" below. `MeetingDraft` (title, date, meetingType, attendeesText, agendaText, notesText — `lib/meeting-minutes.ts`) is still the live-editing shape in `app/MeetingNotes.tsx`'s React state, but an explicit "Save Meeting" now persists it via the same `DataProvider`-style boundary Work Records and Inbox Intelligence already use.
- **Explicit "Analyze Meeting" only**: no AI request on page load, navigation, typing, or field edits — loading Home and loading this screen both cost zero Anthropic requests, exactly like Voice Intelligence. `POST /api/meeting-intelligence` sends only the six analysis-relevant fields — no SharePoint context, no reference data, no RAG/vector lookup.
- **Candidate types**: `SUMMARY`, `DECISION`, `ACTION`, `COMPLETED_WORK`, `IDEA`, `QUESTION`, `KNOWLEDGE`, `FOLLOW_UP_AGENDA` (`lib/meeting-intelligence-models.ts`). Shape: `{ type, title, detail, sourceExcerpt, ownerText, dueText, durationText }` — `ownerText`/`dueText`/`durationText` are nullable and only populated when the system prompt judges them explicit; `sourceExcerpt` may be empty only for `SUMMARY`.
- **Owner safety**: `ownerText` on an `ACTION` candidate is prompt-instructed to stay null unless the notes explicitly assign a person ("Annie will call the district" → `"Annie"`; "we should send it" → `null`). This is prompt-level guidance (like Email Noise's entity-pollution guardrails), not a deterministic verifier — tests prove the pipeline preserves an explicit or null `ownerText` exactly as the model returned it, never fabricating one.
- **Due-date safety**: `dueText` preserves explicit phrases verbatim ("Friday", "before the next meeting") — never a fabricated ISO date. Same prompt-level-guidance caveat as owner safety.
- **Duration safety — reuses Voice Intelligence's exact rule**: `stripUnsupportedDurations()` in `lib/meeting-intelligence-models.ts` re-verifies every `durationText` against the combined agenda+notes text using the same `isDurationSupportedByTranscript()` function Patch 4 already built (imported directly, not reimplemented) — an unsupported/invented duration is dropped to `null` before reaching the review UI.
- **Deterministic dedupe**: same `(type, normalized title)` rule as Voice Intelligence, independently implemented (not shared code) — consistent with this codebase's one-file-per-resource precedent.
- **Draft Minutes — no second AI call**: `buildDraftMinutes()` (`lib/meeting-minutes.ts`) is a pure, synchronous, deterministic composition of meeting details + agenda + the currently-*selected* `SUMMARY`/`DECISION`/`ACTION` candidates (owner/due rendered only when present). Ignored (deselected) and removed candidates are excluded; `IDEA`/`QUESTION`/`KNOWLEDGE`/`FOLLOW_UP_AGENDA` are intentionally left out of the minutes text itself. Updates live as candidates are edited.
- **Copy Minutes**: copies the exact currently-rendered Draft Minutes text via the Clipboard API. No file generation, no PDF/DOCX, no SharePoint save, no email.
- **`COMPLETED_WORK` → "Log as work" reuses the Patch 5 mechanism exactly**: `lib/meeting-intelligence-work-record.ts`'s `buildWorkRecordDraftFromMeetingCandidate()` is a tiny adapter (not a generalized shared mapper, to avoid any risk to Patch 5's tested `buildWorkRecordDraftFromVoiceCandidate`) that reuses the same `parseDeterministicDurationMinutes()` Patch 5 already built. Same conservative rules: only `title`→title and `detail`→description map; no category/district/project/organization/ORBIT inference; the click opens the existing `openLog()`/Log Work wizard and performs zero persistence — proven end-to-end with a real `MemoryDataProvider.createWorkRecord()` spy, exactly one call on explicit "Save & done".
- **`ACTION`/`KNOWLEDGE`/other candidate types**: remain transient reviewed intelligence only in V1 — no Dashboard Action Center connection, no task persistence, no reminders, no Knowledge Base write pathway. Their destinations are separate, future, explicit design work.
- No SharePoint schema, no new Microsoft permissions, and no calendar/Outlook/Teams integration were added in V1 — see Patch 6B below for the durability layer added on top of this unchanged workflow.

---

## Meeting Notes durability (Patch 6B) — application/provider layer implemented; SharePoint list NOT yet provisioned

Adds a durable `MeetingRecord` on top of the unchanged Meeting Notes V1 workflow above, following the exact `DataProvider`-style boundary Work Records and Inbox Intelligence already use. **The SharePoint list itself has not been created** — this patch is application/provider work built and fully tested against an in-memory provider, awaiting the explicit SharePoint schema approval gate before any list is provisioned.

- **`MeetingRecord` model** (`lib/meeting-intelligence-models.ts`): `appId`, `schemaVersion` (currently `1`), `title`, `meetingDate`, `meetingType`, `attendeesText`, `agendaText`, `notesText`, `reviewedCandidates`, `minutesText`, `analysisModel`/`analyzedAt` (both null or both set — never independently), and the standard nested `metadata` (`ProviderMetadata` — provider ID, numeric version, SharePoint's own Created/Modified, sync state). No duplicate timestamp fields were invented alongside SharePoint's canonical Created/Modified.
- **Reviewed intelligence is one JSON blob, not child lists**: `reviewedCandidates: ReviewedMeetingCandidate[]` (`MeetingCandidateSchema.extend({ selected: z.boolean() })`) is stored as a single `IntelligenceJson` field in SharePoint (`lib/sharepoint-meeting-records.ts`). There is no per-candidate-type list or column (no `ActionItemsJson`, no `DecisionsJson`, etc.) — every candidate type (`ACTION`/`DECISION`/`IDEA`/`QUESTION`/`KNOWLEDGE`/`FOLLOW_UP_AGENDA`/`SUMMARY`/`COMPLETED_WORK`) lives in the same blob.
- **Human review is authoritative**: `reviewedCandidates` is always the CURRENT human-edited review state — never the original AI output. The save path only ever reads current React state (`stripReviewIds()` in `app/MeetingNotes.tsx`), and `tests/meeting-notes-durability.test.tsx` proves it end-to-end: a retyped title, a removed owner, and a deselected candidate all survive save-then-reopen exactly as edited.
- **`minutesText` is deterministic — no AI call on save**: `saveMeeting()` calls the unchanged `buildDraftMinutes()` (`lib/meeting-minutes.ts`, built in Patch 6) synchronously at save time. Saving a meeting, like Draft Minutes/Copy Minutes/Analyze itself, never triggers a second Anthropic request.
- **Pre-analysis save is supported**: `reviewedCandidates` may be `[]` and `minutesText` may be the metadata/agenda-only form — a meeting may be saved before "Analyze Meeting" has ever run (an agenda prepared ahead of time, or notes still in progress). Nothing is fabricated to fill these in.
- **Provider boundary** (`lib/meeting-record-provider.ts`, mirroring `lib/inbox-intelligence-provider.ts` exactly): a `MeetingRecordProvider` interface with `list`/`create`/`update` only — **no delete** in this patch. `MemoryMeetingRecordProvider` is the non-durable fallback (used whenever DEV SharePoint config is absent, no Microsoft account is signed in, or `NEXT_PUBLIC_SHAREPOINT_IU_MEETING_RECORDS_LIST_ID` is unset). `DelegatedSharePointMeetingRecordProvider` (`lib/sharepoint-meeting-records.ts`) follows the identical numeric `RecordVersion` + Graph ETag/If-Match optimistic-concurrency pattern as Work Records and Inbox Intelligence — a Graph `412` is mapped to a structured conflict, never auto-merged. `selectMeetingRecordProvider()` only activates SharePoint when a Microsoft account is already signed in (non-interactive) and the list ID env var is configured; that env var is intentionally unset until the list is actually provisioned, so this always resolves to the memory provider today.
- **UI durability copy is honest about mode**: the Meeting Notes screen shows "SharePoint DEV connected — meeting records are saved to your IU Work Tracker workspace when you choose Save Meeting." when `storageMode === "sharepoint"`, or "Preview mode — meeting records are not durable in this mode." otherwise — a deliberate reversal of the V1 "not saved yet" language now that agenda/notes/reviewed intelligence are durable.
- **No autosave, anywhere**: typing, Analyze/Re-analyze, editing/deselecting/removing a candidate, Copy Minutes, and opening the Log Work wizard never call `saveRecord`/`updateRecord`. Only the explicit "Save Meeting" click persists anything; `tests/meeting-notes-durability.test.tsx` asserts zero save calls from every other action.
- **Create vs. update**: `saveMeeting()` routes through `saveRecord` (create) when `identity.metadata.version === 0`, else `updateRecord` (update) with the current `expectedVersion` — the exact same convention `save()` already uses for Work Records.
- **Reopen — zero AI call**: clicking a row in the new "Saved Meetings" list (`SavedMeetingRow`) reconstructs the full editor state — details, agenda, notes, every candidate with its current type/title/detail/owner/due/duration/selected state, minutes, and analysis metadata — from the stored record alone. Candidate `id`s (the transient React list key) are regenerated, never persisted or reused.
- **Re-analyze requires confirmation**: once `candidates.length > 0`, clicking Analyze again first shows `window.confirm("Re-analyzing will replace the current Meeting Intelligence candidates with a new AI analysis. Your agenda and notes will remain.")` — declining performs zero AI calls and leaves the current candidates untouched. Re-analysis never auto-saves.
- **Dirty-state tracking**: a JSON-snapshot fingerprint (`snapshotOf()`, deliberately excluding the transient candidate `id`) drives an "Unsaved changes" indicator and gates "Save Meeting" (disabled when clean). "New Meeting" prompts `window.confirm("Discard the changes in this meeting?")` only when dirty; declining discards nothing.
- **`COMPLETED_WORK` → "Log as work" remains completely uncoupled from Meeting Record persistence**: saving a meeting never creates a Work Record, and logging a candidate as work never creates or updates a Meeting Record — proven by `tests/meeting-work-record-handoff.test.tsx`'s "Meeting Notes / Work Record persistence stay uncoupled" regression tests (`createSpy`/`updateSpy` call-count assertions on both provider mocks, in the same integration render).
- **Still out of scope in 6B**: no meeting deletion; no Dashboard Action Center wiring for `ACTION`/other candidate types; no Knowledge Base/Markdown/Obsidian/embeddings pathway; no relationship inference to districts/organizations/people/projects (attendees remain plain text); no Outlook/Teams/Calendar integration; no new Microsoft Graph permissions beyond the existing delegated `User.Read` + `Sites.ReadWrite.All`.
- **Test coverage**: `tests/sharepoint-meeting-records.test.ts` (field mapping, limits, Graph operations against mocked fetch), `tests/meeting-record-provider.test.ts` (Memory + Delegated-SharePoint provider contract, `selectMeetingRecordProvider()` fallback), `tests/meeting-notes-durability.test.tsx` (save create/update counts, no-autosave, reopen reconstruction, human-review-is-authoritative, pre-analysis save, dirty-state transitions, New Meeting/Re-analyze confirmation gates), plus the two new regression tests in `tests/meeting-work-record-handoff.test.tsx`. Zero real SharePoint writes and zero real Anthropic calls in any of them.

**Status: GREEN, committed (`cdcaab5`).** The `IU Meeting Records` SharePoint list was approved and provisioned (Patch 6B.2); `NEXT_PUBLIC_SHAREPOINT_IU_MEETING_RECORDS_LIST_ID` is set in DEV `.env.local`, and a live create → reload → update → reload smoke test passed against real DEV SharePoint. A synthetic record (`MEETING DURABILITY TEST — DELETE ME`, AppId `meeting-durability-test-6b2`) remains in that list for manual cleanup.

---

## Durable Projects (Patch 7 / 7B) — GREEN, live-verified against real DEV SharePoint, committed (`9c4b442`)

Turns the existing Projects screen from a five-card, code-only reference-data view into a place where the user can create and maintain real projects — see docs/PRODUCT_VISION.md "Log it once. Use it everywhere." A `Project` is an organizational object Work Records accumulate around; this patch deliberately does not add tasks, subtasks, Kanban, Gantt, milestones, comments, attachments, project members, or any other project-management surface.

- **`IU_Projects` is the one authoritative durable Projects list.** It predates this patch (original seven-list reference-data provisioning, `docs/SHAREPOINT_PROVISIONING_CHECKLIST.md`) and already had `Title`/`AppId`/`ProjectDescription`/`Color`. It was extended in place (Patch 7B, manually, after live schema verification) with `StartDate` (Date Only), `TargetDate` (Date Only), `StemOrbit` (Yes/No), and `RecordVersion` (Number, required) — no new list was created. `ProjectStatus` **remains the existing Choice column** (never converted to text); its allowed values are now `active`/`planning`/`paused`/`complete` (`"paused"` added to the pre-existing three). No Lookup/Person columns, no child lists. `AppId` indexing status is unconfirmed but not a blocker — the app resolves by an OData filter regardless.
- **The `Project` model gained an optional durable dimension, not a parallel one**: `lib/models.ts`'s existing `Project` type (`appId`, `name`, `description`, `status`, `color`) is extended in place with `startDate?`, `targetDate?`, `stemOrbit?`, and `metadata?` (`ProviderMetadata`) — all optional, so the five existing seeded projects in `lib/reference-data.ts` remain valid, unchanged `Project` values. A `Project` with `metadata` present was created/loaded through the provider; one without it is a static seeded project — this is exactly how the UI decides whether "Edit" is offered.
- **Provider boundary** (`lib/project-provider.ts`, mirroring `lib/meeting-record-provider.ts`): a `ProjectProvider` interface with `list`/`create`/`update` only — **no delete**, matching Meeting Records and Inbox Intelligence. `MemoryProjectProvider` is the non-durable fallback; `DelegatedSharePointProjectProvider` (`lib/sharepoint-projects.ts`) follows the identical numeric `RecordVersion` + Graph ETag/If-Match optimistic-concurrency pattern — **live-verified**: create → reload → update → reload against real DEV SharePoint, `RecordVersion` `1 → 2`, `AppId`/`providerId` stable throughout.
- **One canonical env var**: `NEXT_PUBLIC_SHAREPOINT_IU_PROJECTS_LIST_ID` — the same variable the original provisioning already used. Patch 7 briefly introduced a separate `NEXT_PUBLIC_SHAREPOINT_IU_DURABLE_PROJECTS_LIST_ID` as a safety rail while `IU_Projects`'s live schema was unverified; **that variable has been fully retired** (Patch 7B) now that the schema is confirmed and extended — there is no second Project-list concept anywhere in the codebase.
- **A second, deliberately separate provider-selection wiring keeps Work Record validation correct**: `createWorkRecord`/`updateWorkRecord` validate `projectIds` against `this.references.projects` (`lib/validation.ts`), which was always just the static seed list. `DataProvider` (`lib/data-provider.ts`) gained `setDurableProjects(projects)` (implemented once on the shared `ReferenceProvider` base class both `MemoryDataProvider` and `DelegatedSharePointDataProvider` extend) — the app calls this every time the durable Project list changes so a newly-created project is immediately valid to reference from a Work Record, not just visually offered in the picker.
- **Create/Edit Project UI**: the dashed "Project creation is planned…" placeholder on the Projects screen is replaced with a clickable "+ Create Project" card (same dashed visual style) opening a compact modal (name/description/status/start date/target date/STEM-ORBIT checkbox — no color picker; new projects cycle deterministically through the same five `project-mark` colors the seeded cards already use). The same modal handles Edit (only offered on durable projects) — create vs. update is decided the same way every other durable resource decides it: `metadata.version > 0`. No autosave anywhere; Cancel discards with zero writes. **Live-verified** against real SharePoint: Create, Edit (prefilled with live-updated values), and browser-refresh persistence all confirmed in the actual running app.
- **Totals are always derived, never stored**: a Project (seeded or durable) carries no record count or duration field — `Projects` filters `records` by `projectIds.includes(project.appId)` and sums `durationMinutes`, exactly as it always has. A brand-new project naturally shows "0 records / 0m invested" until a Work Record references it — confirmed live (the synthetic SharePoint project has never had a Work Record logged against it).
- **Completed-project picker behavior**: the Log Work project picker excludes `status === "complete"` projects by default, but never hides one already selected on the record being edited — a completed project remains a valid, visible historical connection, just not offered as a default new-work choice. Planning/Active/Paused all remain normally selectable. **Live-verified**: the synthetic project appeared in the real Log Work project picker alongside all five seeded projects.
- **Legacy compatibility**: the five seeded projects (`project-steels`, `project-ai`, `project-keystone`, `project-ecosystem`, `project-makerspace`) are untouched and continue to work exactly as before — they are not migrated into the durable store in this patch, remain read-only (no Edit control), and every existing sample Work Record's `projectIds` relationship is unaffected. **Live-verified**: all five still render, no duplicate cards, alongside the new durable project.
- **Seed migration remains deferred**, pending separate explicit approval. The long-term design goal: write those exact five `appId`s into `IU_Projects`, then retire the static `lib/reference-data.ts` seed source — since every existing Work Record already keys off `appId` (never name or SharePoint item ID), that migration requires no Work Record changes at all.
- **No AI anywhere in this workflow** — zero Anthropic calls for Create Project, Edit Project, viewing Projects, or the Log Work project picker.
- **Test coverage**: `tests/project-provider.test.ts` (model validation, Memory + Delegated-SharePoint provider contract), `tests/project-provider-selection.test.ts` (canonical env var resolves to `sharepoint`; the retired temporary env var is ignored entirely), `tests/sharepoint-projects.test.ts` (field mapping incl. every `PROJECT_STATUSES` Choice value round-tripping as an exact lowercase string, Graph operations against mocked fetch), `tests/projects-ui.test.tsx` (Create/Edit UI, create-vs-update routing, no-autosave, Work Record integration through the existing Log Work wizard, completed-project picker exclusion, legacy seeded-project regression, zero Meeting Record/Inbox Intelligence writes). Zero real SharePoint writes and zero real Anthropic calls in any automated test — the live verification above was a manual, one-time smoke test via a temporary local-only page (deleted after use, never committed), not part of the test suite.

**Status: GREEN, live-verified, committed.** A synthetic project (`PROJECT DURABILITY TEST — DELETE ME`, AppId `project-durability-test-7b`) remains in the live `IU_Projects` list for manual cleanup — no delete functionality exists in this codebase.

---

## Durable Contacts (Patch 8B) — GREEN, live-verified against real DEV SharePoint, uncommitted

Promotes the existing static `Contact` reference-data concept into a durable, user-editable domain, following the exact `Project` precedent — see `docs/PRODUCT_VISION.md` "Log it once. Use it everywhere." A `Contact` is a durable Person record everything else can safely connect to; this patch deliberately builds only identity/basic metadata — no Contact detail/connected-work view, no intelligence matching, no Organization durability, no import (those are Patches 8C–8F).

- **`IU_Contacts` is the one authoritative durable Contacts list.** It predates this patch (original seven-list reference-data provisioning, `docs/SHAREPOINT_PROVISIONING_CHECKLIST.md`) and already had `Title`/`AppId`/`Role`/`OrganizationAppId`. Live-inspected before any change (Patch 8B), then extended in place (manually, after approval) with `Email` (Single line text), `Status` (Single line text — deliberately kept as plain text, not a SharePoint Choice column), `Notes` (Multiple lines of text, plain), and `RecordVersion` (Number, required, default 1); `Role` was relaxed from required to optional. No new list was created, no Lookup/Person columns, no child lists. Every internal name was live-verified to exactly match what the codec assumed — no rename was required.
- **The `Contact` model gained an optional durable dimension, not a parallel one**: `lib/models.ts`'s existing `Contact` type (`appId`, `displayName`, `role`, `organizationId`) is extended in place with `status` (required, new `ContactStatus` union), `email?`, `notes?`, and `metadata?` (`ProviderMetadata`) — `role` relaxed from required to optional. The three existing seeded contacts in `lib/reference-data.ts` were given `status: "active"` and remain otherwise unchanged. `organizationId` stays singular and nullable — unchanged, no `organizationIds` array. A `Contact` with `metadata` present was created/loaded through the provider; one without it is a static seeded contact — this is exactly how the UI decides whether "Edit" is offered.
- **Provider boundary** (`lib/contact-provider.ts`, mirroring `lib/project-provider.ts`): a `ContactProvider` interface with `list`/`create`/`update` only — **no delete** (Archived is a status, not a deletion — matching the approved architecture decision). `MemoryContactProvider` is the non-durable fallback; `DelegatedSharePointContactProvider` (`lib/sharepoint-contacts.ts`) follows the identical numeric `RecordVersion` + Graph ETag/If-Match optimistic-concurrency pattern — **live-verified**: created a real Contact (Annie Milewiski) through the normal Contacts UI, confirmed `AppId` is a real UUID, survived a hard browser refresh, edited her through the normal UI (`RecordVersion` `1 → 2`, `Modified` advanced, `ETag` changed), and confirmed the update persisted through another reload — all against real DEV `IU_Contacts`.
- **One canonical env var, no temporary safety rail**: unlike Patch 7's Project list (which used a temporary env var while `IU_Projects`'s live schema was unverified), the existing `NEXT_PUBLIC_SHAREPOINT_IU_CONTACTS_LIST_ID` was reused directly — its live schema was inspected and approved *before* `lib/contact-provider.ts`/`lib/sharepoint-contacts.ts` were written, so there was never an unverified-list safety concern to guard against.
- **Work Record integration reuses the existing relationship unchanged**: `WorkRecord.contactIds: string[]` (already validated, SharePoint-serialized as `ContactIdsJson`, and UI-selectable before this patch) is the sole Contact relationship — no second mechanism was added. `DataProvider` (`lib/data-provider.ts`) gained `setDurableContacts(contacts)` (implemented once on the shared `ReferenceProvider` base class), mirroring `setDurableProjects` exactly, so `createWorkRecord`/`updateWorkRecord`'s `contactIds` validation (`lib/validation.ts`) sees the merged (seeded + durable) contact set. **Live-verified**: selected the durable Contact in the real Log Work Contacts picker, saved a Work Record, reloaded, and confirmed the relationship persisted.
- **Contacts directory UX**: a new "Contacts" nav item renders a compact card grid (same visual language as Projects — `project-card`/`project-grid` classes reused, not a new design system) with a search box (name/role/email/organization, case-insensitive substring match — the same idiom History's search already uses) and a "Show archived" toggle (archived contacts hidden by default). "+ Add Contact" opens a compact modal (name/role/organization/email/status/notes); the same modal handles Edit (only offered on durable contacts). No connected work, no timeline, no Last Interaction, no Waiting On — deliberately deferred to Patch 8C.
- **Conservative, deterministic duplicate detection**: trim + lowercase exact matching only, never fuzzy, never first-name-only (`normalizeContactName`/`normalizeContactEmail` in `lib/contact-provider.ts`). A same-name match is informational only and never blocks Save. A same-email match requires an explicit second click on Save ("Create Anyway") before it will create a second Contact — never silently created, never silently merged.
- **No AI anywhere in this workflow** — zero Anthropic calls for browsing, searching, creating, or editing Contacts; verified both by test (a `fetch` spy) and by static source-text checks (`lib/contact-provider.ts`/`lib/sharepoint-contacts.ts` contain no reference to `@anthropic-ai/sdk` or any `/api/*-intelligence` route).
- **No intelligence integration in this patch**: Inbox/Meeting/Voice Intelligence are untouched — no `matchedContactIds`, `attendeeContactIds`, or `ownerContactId` field exists anywhere. That is Patch 8D.
- **Test coverage**: `tests/contact-provider.test.ts` (model/status validation, normalization, Memory + Delegated-SharePoint provider contract), `tests/contact-provider-selection.test.ts` (canonical env var resolves to `sharepoint`), `tests/sharepoint-contacts.test.ts` (field mapping incl. every `CONTACT_STATUSES` value, missing-optional-field handling, Graph operations against mocked fetch), `tests/contacts-ui.test.tsx` (directory, search, archived filtering, duplicate detection for both name and email, Create/Edit UI, Work Record integration incl. an unknown-Contact-ID validation-failure test and a static-seeded-Contact validation-success test, legacy seeded-contact regression, zero Meeting/Inbox/Project writes, zero AI). Zero real SharePoint writes and zero real Anthropic calls in any automated test — the live verification above was a manual, one-time smoke test through the real application UI, not part of the test suite.

**Status: GREEN, live-verified, uncommitted.** Annie Milewiski (real dev data, not synthetic test data — left in place per instruction) remains in the live `IU_Contacts` list.

---

## Contact Detail + Connected Work (Patch 8C) — GREEN, live-verified against real DEV SharePoint

Turns the Patch 8B Contacts directory into a relationship-memory view: opening a Contact answers "what work connects to this person" — entirely derived from records already stored elsewhere. No schema change, no new SharePoint columns, no new Graph requests, no AI.

- **The relationship stays exactly `Contact.appId → WorkRecord.contactIds → WorkRecord.projectIds → Project`** — no `Contact.projectIds`, no `Project.contactIds`, no direct persisted Contact↔Project relationship. See `docs/DATA_MODEL.md` "Contact connected work (Patch 8C)" for the full derivation contract.
- **`lib/contact-relationships.ts`** — a new, self-contained pure derivation module (no I/O, no AI), mirroring `lib/inbox-action-center.ts`'s role for the Home Action Center. `buildContactRelationshipSummary(workRecords, projects, contactAppId)` returns `connectedWorkRecords`, `recentWorkRecords` (5 most recent), `connectedProjects`, `totalMinutes`, `workRecordCount`, and `lastInteractionDate` — computed fresh on every render from arrays the screen already has, never persisted anywhere.
- **Last Interaction uses `WorkRecord.activityDate`** (the same business-date field and sort direction History already uses) — deliberately never `ProviderMetadata.createdAt`/`modifiedAt` (SharePoint Created/Modified describe persistence timing, not interaction timing). A dedicated test proves this by setting a decoy record's metadata timestamps to the year 2099 and confirming they are never read.
- **Contact Detail UX**: Contacts directory cards are now openable (the card's title block is a real `<button>`; the sibling "Edit" button stays a separate control so click targets never overlap or nest) — clicking one opens `ContactDetail` as an internal state toggle (`openContactId`) inside the existing `Contacts` screen component, not a new top-level `View` or URL route. Information hierarchy: identity header (name, resolved Organization via `contact.organizationId`, status, role, email) → Relationship Snapshot (Last Interaction, Connected Time, Work Record count) → Connected Projects (chips) → Recent Work (up to 5 rows, newest first, each opening the existing Log Work edit dialog for that record — no new Work Record CRUD surface) → Notes (existing `contact.notes` field, read-only here; editing stays through the existing Edit Contact modal). "← Back to Contacts" returns to the directory. Works identically for durable and seed Contacts (Edit is offered only when `contact.metadata` is present, same rule as the directory card) and for archived Contacts when explicitly opened.
- **Empty states, never a hidden screen**: "No recorded interaction yet.", "No connected projects yet.", "No work has been logged with this contact yet.", "0m" connected time — a Contact with zero connected Work Records renders calmly, never crashes, never fabricates a date.
- **Zero AI, zero new persistence, zero N+1 fetching**: no AI summary/relevance scoring anywhere in Contact Detail; no `Waiting On`/follow-up/task system (deferred to Patch 8D+, once Inbox/Meeting Contact links exist); derivation runs entirely over the Work Records and Projects the app already loaded — no Contact-specific Graph requests, no background refresh loop, no new cache layer.
- **Test coverage**: `tests/contact-relationships.test.ts` (pure derivation — exact-appId-only matching, never inferring from name/email/org/free-text, newest-first sort, project union + dedup + graceful-skip-unresolved, duration sum, last-interaction using business date only, unrelated-newer-record isolation, clean empty summary, recent-work cap vs. true count) and `tests/contact-detail-ui.test.tsx` (opening detail, organization resolution, status display, Relationship Snapshot/Connected Projects/Recent Work rendering, clicking a Recent Work row into the existing Log Work dialog, Notes rendering, empty states, seed-Contact and archived-Contact detail access, zero-fetch boundary, and static guardrails confirming `Contact` still has no `projectIds` and `Project` still has no `contactIds`).

**Status: GREEN, live-verified.**

---

## Intelligence Contact matching (Patch 8D) — GREEN

Connects Intelligence-detected people to the durable Contact system from Patch 8B. Governing principle: **AI may propose. The human decides.** No silent Contact creation, no silent identity guessing, no automatic overwriting of Contact data — see `docs/DATA_MODEL.md` "Intelligence Contact matching (Patch 8D)" for the full matching-ladder and persistence contract.

- **Workflow audit first** (Step B/C of the patch instructions): `Inbox Intelligence` already detects people (`EmailAnalysis.people: string[]`) AND already persists reviewed entity links (`matchedOrganizationIds`/`matchedDistrictIds`/`matchedProjectIds` on the durable `InboxIntelligenceRecord`) — the clear primary implementation target. `Voice Intelligence` detects people (`PERSON` candidates) but has **zero durable persistence of any kind** — no `VoiceRecord` type, no provider, no SharePoint list; its own doc comment states nothing there is ever written anywhere. `Meeting Intelligence` (`MeetingRecord`, Patch 6B) is durable but has **no `PERSON` candidate type and no entity-matching precedent at all** — `attendeesText`/`ownerText` remain plain free text by design (Patch 6B: "no district/people entity resolution on attendeesText"). This audit is why the three workflows were NOT treated with equal parity.
- **`lib/contact-matching.ts`** — the one shared, deterministic matcher (`matchContactCandidates()`), reused unmodified by both Inbox and Voice. Zero AI calls, zero network calls, zero mutation. Reuses `normalizeContactName`/`normalizeContactEmail` from `lib/contact-provider.ts` rather than a second normalization implementation, adding only whitespace-collapsing on top for AI-extracted free text. Returns every plausible candidate (never picks one by array order) so a same-name ambiguity always reaches the human. Archived Contacts are matchable.
- **`app/ContactMatchPanel.tsx`** — the one shared review component behind every "possible Contact matches" section in both Inbox and Voice. Renders unreviewed / matched / ignored states; resolves a matched decision's display name against the FULL current Contact list (not the transient recomputed `candidates` array), so the display stays correct even if the underlying detected-name text is later edited. `resolveMatchedContacts()` (same file) is the pure lookup used to render a saved Inbox record's already-decided matches read-only.
- **Inbox Intelligence integration** (`app/InboxIntelligence.tsx`): a "People — possible Contact matches" section appears in the review phase (before Save) whenever `analysis.people` is non-empty, one `ContactMatchPanel` per detected name — organization context for the `useful` tier comes from that same email's already-resolved `resolveEmailAnalysisEntities()` organization/district IDs. Review decisions (`personDecisions: Record<string, ContactMatchDecision>`) are local component state, reset on each new Analyze — unreviewed and ignored people leave no persisted trace, exactly like Organization/Project matching leaves no "was this reviewed" record. At Save, only the human's accepted "matched" decisions (deduped) become `matchedContactIds`, passed into `buildInboxIntelligenceRecord()`. There is **no post-save re-review UI** — deliberately mirroring the existing Organization/District/Project precedent, which is also resolved once at save time and shown read-only in the saved list thereafter (`InboxRow` now also displays matched Contact names alongside the existing matched Project/Organization display).
- **Voice Intelligence integration** (`app/VoiceIntelligence.tsx`): `PERSON`-type `ReviewCandidate`s gain an optional, purely transient `contactDecision` field — reviewed the same way as every other Voice review field (never written to SharePoint, Work Records, Inbox Intelligence, Organizations, Projects, or any browser storage). No organization-context elevation is available for Voice (no per-transcript resolved-organization-IDs precedent exists there), so Voice matching only ever reaches `review` or `strong` tiers in practice, never `useful` — an intentional, documented simplification, not a bug.
- **`app/ContactFormModal.tsx`** — Patch 8B's Create/Edit Contact form, extracted out of `app/IUWorkTracker.tsx` into its own file specifically so Inbox/Voice's "Add Person" can reuse it without a circular import between `IUWorkTracker.tsx` and the Intelligence screens it renders. Behavior is byte-for-byte unchanged; `IUWorkTracker.tsx`'s own Contacts screen now imports it from there too. `onSaved` now receives the saved `Contact` (previously took no arguments) so "Add Person" can immediately record the newly created Contact's `appId` as the matched decision.
- **No automatic anything**: opening a review panel, computing candidates, and selecting Contacts perform zero AI calls (verified by `fetch`-spy tests) and zero Contact mutations beyond an explicit "Add Person" create. A `strong` (exact email) match is displayed exactly like any other candidate — the human still clicks "Match Existing". Matching an existing Contact never calls `updateContact`.
- **SharePoint schema dependency**: `matchedContactIds` requires a new `MatchedContactIdsJson` column (Multiple lines of text, plain) on the live `IU_Inbox_Intelligence` list, mirroring `MatchedProjectIdsJson` etc. exactly — added by the user before this patch's code was written, following the Patch 8B precedent (schema change first, code second). Reading an older item saved before the column existed degrades gracefully to `matchedContactIds: []` (never a parse error) via the same `parseJsonStringArray` fallback the sibling fields already use.
- **Test coverage**: `tests/contact-matching.test.ts` (the full matching ladder, normalization, multi-candidate ambiguity, archived-Contact matchability), `tests/contact-match-panel-ui.test.tsx` (no candidate — including `strong` — is ever pre-decided; Match Existing/Ignore/Add Person/Change/Reconsider button semantics), `tests/inbox-contact-matching-ui.test.tsx` (end-to-end through the real component: unreviewed start state, Match Existing → persisted `matchedContactIds` → displayed on the saved row with zero Contact mutation, Ignore → no Contact created, Add Person → real Contact form → association, existing Organization/Project auto-resolution regression, zero extra AI/network calls), `tests/voice-contact-matching-ui.test.tsx` (the same review actions for a `PERSON` candidate, transient-only), plus extensions to `tests/inbox-intelligence-provider.test.ts` and `tests/sharepoint-inbox-intelligence.test.ts` (default `[]`, dedup, full round-trip, graceful-missing-column read).
- **Explicitly deferred, not attempted**: Meeting Intelligence Contact linkage (no `PERSON` type, no entity-matching precedent — would require a larger Meeting domain migration); durable Voice persistence of any kind; post-save re-review of Inbox Contact matches; AI relationship/importance scoring; Contact profile enrichment from Intelligence evidence; `matchedContactIds` on any model other than `InboxIntelligenceRecord`.

**Status: GREEN.**

---

## Durable Organizations (Patch 8E) — GREEN, live-verified against real DEV SharePoint

Promotes `Organization` from static reference data to a durable domain, following the exact Project/Contact precedent from Patches 7 and 8B. District stays exactly `Organization.type === "district"` — no separate District model, provider, list, or codec exists or was added.

- **`IU_Organizations` predates this patch** (original seven-list reference-data provisioning, `docs/SHAREPOINT_PROVISIONING_CHECKLIST.md`) and already had `Title`/`AppId`/`OrganizationType`, read-only. It was live-inspected by the user before any code was written (Patch 8E live-inspection gate, mirroring the Patch 8B Contacts gate exactly): confirmed empty, and extended in place with `RecordVersion` (Number, required, default 1, zero decimals) — no other column added, no second list or env var created. `Title`→`name`, `AppId`→`appId`, `OrganizationType`→`type` (Choice, unchanged values `district`/`partner`/`iu`) were confirmed to match the provisioning docs exactly, so `lib/sharepoint-organizations.ts` was written directly against the live schema with no ambiguity to resolve.
- **The `Organization` model gained an optional durable dimension, not a parallel one**: `lib/models.ts`'s existing `Organization` type (`appId`, `name`, `type`) is extended in place with optional `metadata?: ProviderMetadata` — the four existing seeded organizations in `lib/reference-data.ts` are unchanged. An `Organization` with `metadata` present was created/loaded through the provider; one without it is a static seeded organization — the same rule the UI already uses for Project/Contact "Edit" eligibility.
- **Provider boundary** (`lib/organization-provider.ts`, mirroring `lib/project-provider.ts` line for line): an `OrganizationProvider` interface with `list`/`create`/`update` only — no delete. `MemoryOrganizationProvider` is the non-durable fallback; `DelegatedSharePointOrganizationProvider` (`lib/sharepoint-organizations.ts`) follows the identical numeric `RecordVersion` + Graph ETag/If-Match optimistic-concurrency pattern (create sends `RecordVersion: 1`; update increments it; a 412 surfaces as a structured conflict carrying the current record — never an auto-merge, never a silent overwrite).
- **One canonical env var, no temporary safety rail**: like Patch 8B's Contacts (and unlike Patch 7's Projects, which used a temporary env var while schema was unverified), the existing `NEXT_PUBLIC_SHAREPOINT_IU_ORGANIZATIONS_LIST_ID` is used directly — its live schema was inspected and the `RecordVersion` column added *before* `lib/organization-provider.ts`/`lib/sharepoint-organizations.ts` were written, so there was never an unverified-list safety concern to guard against.
- **Reference-merge architecture fixes the boundary once, not each consumer individually**: `DataProvider` (`lib/data-provider.ts`) gained `setDurableOrganizations(organizations)` (implemented once on the shared `ReferenceProvider` base class), mirroring `setDurableProjects`/`setDurableContacts` exactly. Because `references.organizations` is merged (seed + durable) at this single point, every existing consumer sees a durable Organization automatically with zero changes to that consumer: `createWorkRecord`/`updateWorkRecord`'s `organizationIds` validation (`lib/validation.ts`, including the district/`engagementScope` rules), the Log Work District(s) and Organization/partner pickers, the Contacts form's Organization select, Inbox Intelligence's `matchedOrganizationIds`/`matchedDistrictIds` exact-name matching, and `lib/contact-matching.ts`'s organization-context elevation for the `useful` match tier. **Live-verified in the browser**: created a durable district through the Organizations screen, confirmed it appeared in the Log Work District(s) picker (and only there, never in Organization/partner) and in the Add Contact form's Organization select.
- **Organizations directory UX**: a new "Organizations" nav item renders the same compact card grid as Projects/Contacts (`project-card`/`project-grid` classes reused) with a search box (name or type label, case-insensitive substring) and no archived-filter concept (Organization has no status field). "+ Add Organization" opens a compact modal with exactly **Name** and **Type** — no Notes, no expanded taxonomy; the same modal handles Edit (only offered on durable organizations).
- **Conservative, deterministic duplicate detection**: trim + collapse-whitespace + lowercase exact matching only (`normalizeOrganizationName`, mirroring `normalizeContactName`), never fuzzy. A same-name match is informational only and never blocks Save — Organization has no email field to serve as Contact's stronger "requires a second click" signal, so this is single-tier, unlike Contact's two-tier (name warns, email blocks) duplicate handling.
- **No AI anywhere in this workflow** — zero Anthropic calls for browsing, searching, creating, or editing Organizations; Inbox/Voice Intelligence never call `saveOrganization`/`updateOrganization` — Organization matching there stays exactly as read-only as it always was, now simply seeing a larger candidate set.
- **Test coverage**: `tests/organization-provider.test.ts` (model/type validation, normalization, Memory + Delegated-SharePoint provider contract including RecordVersion-1-on-create, 412-conflict, UUID appId), `tests/organization-provider-selection.test.ts` (canonical env var resolves to `sharepoint`, no other list-id var activates it), `tests/sharepoint-organizations.test.ts` (field mapping incl. every `ORGANIZATION_TYPES` Choice value, missing-column/missing-timestamp handling, Graph operations against mocked fetch), `tests/data-provider-organizations.test.ts` (the reference-merge boundary itself: seed-only before, seed+durable after, no duplicate accumulation across re-syncs, WorkRecord validation seeing a durable Organization including district/engagementScope rules), `tests/organizations-ui.test.tsx` (directory, search, seeded-organization read-only regression, duplicate-name warning that never blocks Save, Create/Edit UI, Work Record integration seeing a durable district/partner in the correct picker only, Contact form seeing a durable Organization, Inbox district matching seeing a durable district on the saved row), `tests/organization-architecture-guardrails.test.ts` (no District model/provider/list, no `Organization.contactIds`/`projectIds`, `Contact.organizationId`/`WorkRecord.organizationIds` shapes unchanged, `matchedOrganizationIds`/`matchedDistrictIds` preserved, taxonomy not expanded, single canonical env var, no AI/fuzzy-matching code, no new Microsoft permissions). Zero real SharePoint writes and zero real Anthropic calls in any automated test.
- **Explicitly out of scope, not attempted**: expanded `OrganizationType` taxonomy (manufacturer/nonprofit/business/community/other); `Notes` on Organization; imports/CSV/provenance; CRM pipeline/enrichment; AI/relationship/importance scoring; Organization connected-work detail screen; new Microsoft permissions.

**Status: GREEN, live-verified in Preview mode (directory, Create/Edit, Log Work pickers, Contact form). Live SharePoint create/update against the real DEV `IU_Organizations` list was not independently re-verified by this session — Microsoft sign-in in this environment requires the user's own real credentials, which were not entered; the codec/provider are exercised end-to-end by mocked-fetch tests instead.**

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
