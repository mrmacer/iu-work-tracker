# IU Work Tracker — Inbox Intelligence V1 Report

**Phase:** Inbox Intelligence V1 (authorized explicitly; original concept in docs/AI_HANDOFF.md's former "Future Candidate Feature" section).

**Status:** GREEN — implementation, automated tests, and one live AI test all complete.

**Scope:** Paste-email → AI-extraction → human review → optional session save → optional Create Work Record. No Outlook/Graph mail integration, no SharePoint list provisioned.

## Architecture

Inbox Intelligence lives entirely inside IU Work Tracker as one new kiosk screen plus one new server route — no second app, dashboard, auth system, or database.

```
Home (command card) ──▶ InboxIntelligence.tsx (client screen)
                             │
                             │ POST /api/inbox-intelligence  { rawEmail }
                             ▼
                     app/api/inbox-intelligence/route.ts (server route)
                             │  reads env.ANTHROPIC_API_KEY (cloudflare:workers)
                             ▼
                lib/anthropic-email-analysis.ts → @anthropic-ai/sdk → Anthropic Messages API
                             │  client.messages.parse() + Zod schema (structured output)
                             ▼
                strict validation + dueDate normalization → AnalyzeEmailResult JSON
                             │
                             ▼
        InboxIntelligence.tsx renders an editable review screen (nothing persisted yet)
                             │
              ┌──────────────┴───────────────┐
              ▼                               ▼
   Save → SessionInboxIntelligenceProvider   Create Work Record → buildWorkRecordDraftFromAnalysis()
   (in-memory, this session only)             → existing openLog()/LogWizard/DataProvider (unchanged)
```

The browser never talks to Anthropic directly and never sees the API key — it only calls the app's own `/api/inbox-intelligence` route, mirroring the same server-owns-the-secret shape as `/api/records` (D1) already established in this codebase.

## Files changed

| File | Change |
|---|---|
| `lib/anthropic-config.ts` | New. Centralizes `ANTHROPIC_MODEL`, `ANTHROPIC_EFFORT`, `ANTHROPIC_MAX_OUTPUT_TOKENS`, `MAX_EMAIL_LENGTH` — the only place a future model/limit change needs to happen. |
| `lib/inbox-intelligence-models.ts` | New. `EmailAnalysisSchema`/`ActionItemSchema` (Zod), `EmailAnalysis`/`ActionItem` types, `normalizeActionItemDueDate`/`normalizeEmailAnalysis`, and the `InboxIntelligenceRecord` domain type. |
| `lib/anthropic-email-analysis.ts` | New. `analyzeEmailWithClaude(rawEmail, client)` — the one function that talks to Claude. Takes an injectable client (or `null`) so it is fully unit-testable without touching the network. |
| `app/api/inbox-intelligence/route.ts` | New. `POST` only. Reads `env.ANTHROPIC_API_KEY` via `cloudflare:workers` (same pattern as `env.DB` in `app/api/records/route.ts`), builds the Anthropic client server-side, delegates to `analyzeEmailWithClaude`, maps the result to an HTTP status. |
| `lib/inbox-intelligence-provider.ts` | New. `InboxIntelligenceProvider` interface + `SessionInboxIntelligenceProvider` (in-memory, one instance per screen mount). |
| `lib/inbox-intelligence-work-record.ts` | New. `buildWorkRecordDraftFromAnalysis()` — pure mapping from an `EmailAnalysis` to a `WorkRecord` draft, using the existing `DataProvider`/`WorkRecord` types and exact-name reference matching only. |
| `app/InboxIntelligence.tsx` | New. The screen: paste → review → saved states, reusing existing global CSS classes (`.panel`, `.form-stack`, `.metric-strip`, `.toggle-line`, `.log-footer`, etc.) rather than introducing new styles. |
| `app/IUWorkTracker.tsx` | Adds `"inbox"` to the `View` union and `navItems`, one `Command` card on Home, and one render branch — no other Home/kiosk changes. |
| `cloudflare-env.d.ts` | Adds `ANTHROPIC_API_KEY?: string` to `Cloudflare.Env`. |
| `.env.example` | Adds `ANTHROPIC_API_KEY=` (empty) with a comment forbidding a real value and forbidding a `NEXT_PUBLIC_` prefix. |
| `package.json` / `package-lock.json` | Adds `@anthropic-ai/sdk` (0.122.0, exact) and `zod` (4.5.2, exact) as `dependencies`. |
| `tests/inbox-intelligence-analysis.test.ts` | New. Schema validation, due-date normalization, and `analyzeEmailWithClaude` against a fake client (valid extraction, null `parsed_output`, rate-limit/auth/unknown errors, missing-key, length limit). |
| `tests/inbox-intelligence-provider.test.ts` | New. `buildWorkRecordDraftFromAnalysis` (purity, exact-match-only entity linking, engagementScope invariant, follow-up prefill) and `SessionInboxIntelligenceProvider` (ordering, counts, per-instance isolation). |

No existing file's behavior changed outside the additive edits listed above. `app/dev-sharepoint-smoke/` and `docs/MAC_WALKTHROUGH_AUTH_REFERENCE.md` remain the same pre-existing untracked items flagged in earlier reviews — untouched.

## API security design

- The Anthropic API key is read once, server-side, in `app/api/inbox-intelligence/route.ts`, via `env.ANTHROPIC_API_KEY` (the Cloudflare Workers environment binding — the same mechanism `env.DB` already uses in this codebase, not `process.env` and never a `NEXT_PUBLIC_*` variable).
- The browser calls only `POST /api/inbox-intelligence` with `{ rawEmail }`. It never constructs an Anthropic request and never receives the key in any response.
- **Verified structurally, not just by convention:** after `npm run build`, `dist/client/**` was grepped for `ANTHROPIC_API_KEY`, `@anthropic-ai`, and `x-api-key` — zero matches. The `@anthropic-ai/sdk` import in `app/InboxIntelligence.tsx` is `import type { ... }` only (TypeScript types, erased at build time); the client component never imports the SDK's runtime code.
- If `ANTHROPIC_API_KEY` is unset or empty, the route never constructs an Anthropic client and returns a generic `"AI analysis is not configured."` message — confirmed live (see "Exact next live-test step").
- Errors from the Anthropic SDK (`RateLimitError`, `AuthenticationError`, other `APIError`s, or anything else thrown) are caught and mapped to one of a small set of safe, generic messages. The raw SDK error is never serialized into the HTTP response. A test asserts this directly: an `AuthenticationError` constructed with a fake key string embedded in its message is asserted to never appear in the returned message.
- Server-side `console.error` logging on unexpected failures matches the existing `app/api/records/route.ts` convention (log the caught error object server-side only; never return it to the client).

## Extraction schema

Defined with Zod in `lib/inbox-intelligence-models.ts`, enforced via the Anthropic TypeScript SDK's `client.messages.parse()` + `zodOutputFormat()` structured-output feature (not manual JSON parsing, not a hand-rolled tool call):

```ts
{
  summary: string (1-2000 chars),
  priority: "high" | "medium" | "low",
  needsAttention: boolean,
  actionItems: { action: string; dueDate: string | null; owner: "me"|"sender"|"other"|"unknown" }[] (max 20),
  followUp: string (0-1000 chars),
  people: string[] (max 50), organizations: string[] (max 50),
  districts: string[] (max 50), projects: string[] (max 50),
  tags: string[] (max 20),
  suggestedWorkType: string | null,
  suggestedWorkRecord: { title: string (1-255 chars); description: string (0-2000 chars) }
}
```

Both objects are `.strict()` — an unexpected extra field from the model fails validation rather than being silently accepted. If `response.parsed_output` is `null` (the model's output didn't validate against the schema), the route returns `invalid_model_output` with a plain message and fabricates nothing.

One additional, deliberately narrow normalization runs after schema validation: any `actionItems[].dueDate` that isn't an exact `YYYY-MM-DD` string is coerced to `null` rather than rejecting the whole analysis — a loosely-formatted date is far more likely than a structurally broken response, and the instruction is "never fabricate a date," not "reject the whole email over one field." This is the one exception to "fail the whole response on any invalid field," and it is unit-tested directly (`normalizeActionItemDueDate`).

## AI behavior (prompting)

The system prompt (in `lib/anthropic-email-analysis.ts`) explicitly instructs the model to:
- Separate "what the email says" from "what the user may need to do."
- Leave `actionItems` empty for routine/FYI emails — not manufacture one per email.
- Only set a `dueDate` when the email states or clearly implies one; otherwise `null`.
- Only list people/organizations/districts/projects that are actually named in the text — never invent or guess entities.
- Treat `suggestedWorkType` as an unvalidated free-text guess; the application (not the model) is responsible for matching it to the controlled activity-type vocabulary.

## Privacy behavior

- The raw pasted email exists only in React component state on the client and in the one `POST /api/inbox-intelligence` request/response cycle. It is never written to any persistent store — not SharePoint, not D1, not even the in-memory session provider.
- What survives past the review screen (if the user clicks Save) is an `InboxIntelligenceRecord` containing a **200-character excerpt** (`sourceExcerpt`, whitespace-collapsed, ellipsized) plus the structured `analysis` — never the full email text.
- Nothing is persisted automatically. AI output only ever reaches session memory after an explicit user click on **Save**, and only ever reaches the real `DataProvider` (D1/SharePoint) after the user explicitly reviews and saves through the existing, unchanged Log Work wizard.

## Cost controls

- `MAX_EMAIL_LENGTH = 20000` characters, enforced in both the client (button disabled, inline character counter) and the server function (`analyzeEmailWithClaude` rejects before ever touching the Anthropic client) — defense in depth, and unit-tested at the server layer.
- The Analyze button is disabled while a request is in flight (`analyzing` state) and the handler itself short-circuits (`if (analyzing) return;`) as a second guard against a double-click or rapid-Enter firing two requests.
- No `useEffect` triggers analysis automatically; analysis fires only from the Analyze button's `onClick`. Nothing runs on keystrokes.
- No automatic retries: a failed analysis shows the error and leaves the pasted text in place for the user to explicitly click Analyze again.
- A small DEV diagnostic line appears on the review screen after a successful analysis: `Model: claude-opus-5 · <N> in / <N> out tokens`, sourced directly from `response.usage` and `response.model` — no separate telemetry system.
- Model, effort, and token ceiling are centralized in `lib/anthropic-config.ts` (currently `claude-opus-5`, effort `medium`, `max_tokens: 4096`) so a future cost/quality tradeoff is a one-constant change, not a feature rewrite.

## Persistence status

**Session-only. Not durable. No SharePoint list has been provisioned for this feature.**

`SessionInboxIntelligenceProvider` holds saved records in a plain in-memory array scoped to one component instance; a page reload loses everything, exactly as the task specified for this phase ("these may remain session-only for this phase... do not pretend session-only data is durable"). This is proven by a unit test asserting two provider instances never share state.

## Work Record integration

`buildWorkRecordDraftFromAnalysis(analysis, references, baseRecord)` is a pure function (no I/O, no DataProvider call) that maps AI-suggested fields onto a fresh, unsaved `WorkRecord`:

- `title`/`description` ← `suggestedWorkRecord`.
- `activityType` ← `suggestedWorkType` **only on an exact, case-insensitive match** against `references.settings.activityTypes`; otherwise left blank for the user to choose. Never a fuzzy match.
- `organizationIds`/`engagementScope` ← `organizations`/`districts`, matched by exact case-insensitive name against `references.organizations`. A matched canonical district sets `engagementScope: "specific"`; an unmatched name is silently dropped, never invented as a new organization. This keeps the draft's scope/organization invariant valid on its own (unit-tested), so the wizard never opens on an already-broken record.
- `projectIds` ← `projects`, same exact-match rule against `references.projects`.
- `followUpNeeded`/`followUpDate`/`nextStep` ← the extracted action items/follow-up text, never a fabricated date.
- `metadata.version` stays `0`, so the existing `save()` logic in `app/IUWorkTracker.tsx` always routes this through `createWorkRecord`, never `updateWorkRecord`.

**Create Work Record** calls this function and then the existing `openLog(draft)` — the same function every other "edit/create a Work Record" entry point in the app already uses. It opens the normal Log Work wizard, prefilled. Nothing is saved until the user reaches the wizard's own Save button, through the existing `DataProvider` write path (D1 or SharePoint, whichever `selectDataProvider()` picked) — completely unmodified by this feature.

## Tests

89/89 passing (was 64 before this phase), all new coverage mocks the Anthropic client — **no automated test calls the real API**:

- Schema validation: valid extraction accepted; invalid `priority`/`owner` enum, unknown extra field, and a missing required array field all rejected.
- `normalizeActionItemDueDate`/`normalizeEmailAnalysis`: well-formed date kept, `null` kept, malformed date discarded (not fabricated).
- `analyzeEmailWithClaude`: empty input, over-length input, no client (no API key), a valid successful call, a `null` parsed_output, a `RateLimitError`, an `AuthenticationError` (and that its message text never leaks into the returned message), and an arbitrary thrown error — all via a fake `{ messages: { parse } }` object, never a real `Anthropic` client.
- `buildWorkRecordDraftFromAnalysis`: purity/validity of a fully-matched draft, activity-type exact-match-only, organization/project exact-match-only (no invention), the district → `specific` scope invariant, and follow-up/next-step prefill without fabricating a date.
- `SessionInboxIntelligenceProvider`: newest-first ordering, `needsAttentionCount`/`openActionCount` computation, and per-instance isolation (proving nothing is durable).
- SSR/build safety: `npm run build` succeeds (client, SSR, and RSC bundles) with the new route and dependencies; the existing `tests/rendered-html.test.mjs` SSR-shell smoke test still passes unchanged.

Live-verified, no-key path (not a paid call): with the dev server running and no `ANTHROPIC_API_KEY` set, the real UI flow (Home → Inbox Intelligence → paste a synthetic email → Analyze email) was driven end-to-end in a browser. It rendered "AI analysis is not configured." cleanly, the Analyze button re-enabled, and the built client bundle was grepped for the API key name, the SDK package name, and `x-api-key` with zero matches.

**Live-verified, real API call (one paid request, see below).** A placeholder key first produced a clean, safe `AuthenticationError → "AI analysis is not configured correctly."` result with no key exposed — diagnosed (18-character placeholder value, not a real `sk-ant-...` key) without ever printing the value, and no second call was made until the user supplied a real key. With the real key in place, one live call succeeded end-to-end; see "Live AI test result" below.

## Live AI test result

One live, paid Anthropic API call was made with a non-sensitive synthetic email (a made-up grant-report-deadline message between fictional people at a fictional district — no real IU data). Verified directly in the browser and via network/server logs:

| Check | Result |
|---|---|
| Request succeeded | Yes — `POST /api/inbox-intelligence` returned `200` (server log: `POST /api/inbox-intelligence 200 in 11.8s`) |
| Structured extraction passed schema validation | Yes — the review screen rendered (a schema failure would have returned `invalid_model_output` and no review screen) |
| Result renders correctly | Yes — summary, priority, needs-attention, two distinct action items (each with owner and a due date computed correctly from relative day names + the email's header date, e.g. "by Wednesday" → `08/26/2026`), follow-up, tags, people, organizations, districts, projects, and suggested Work Record title/description all appeared, editable, under the "AI-suggested — review before saving" label |
| Extraction did not over-generate action items | Yes — a line explicitly marked "no action needed" in the test email produced no action item |
| Extraction did not invent entities | Yes — people/organizations/districts/projects exactly matched the names present in the pasted text; nothing extra appeared |
| No raw email persisted automatically | Yes — the raw email exists only in the one request body and client state; no write to any store occurs before Save |
| No Inbox Intelligence record persisted automatically | Yes — confirmed live: after analysis, "Recently analyzed" still read "Nothing analyzed yet this session" because Save was never clicked |
| No Work Record created automatically | Yes — the "Create Work Record" action only appears after an explicit Save, which was never clicked |
| Token/model diagnostics captured | Yes — rendered as `Model: claude-opus-5 · 2,373 in / 732 out tokens`, sourced from `response.model`/`response.usage` |
| No API credential exposed client-side | Yes — network log for this session shows exactly one call to `/api/inbox-intelligence` and **no** request to `api.anthropic.com` from the browser; server logs contain only the route path/status/timing; the rebuilt client bundle was re-grepped for `ANTHROPIC_API_KEY`/`@anthropic-ai`/`x-api-key` with zero matches |
| Only one paid call made | Yes — the earlier placeholder-key attempt failed at authentication (before any model invocation) and is not billable; exactly one real model call was made this session |

## Known limitations

1. Only one email, one session, one user account has exercised the live path so far. Broader real-world variety (very long threads, multiple languages, ambiguous dates, adversarial/prompt-injection-style email content) has not been tested.
2. `linkedWorkRecordAppId` exists on `InboxIntelligenceRecord` for the proposed future SharePoint shape, but V1 does not actually wire it up after a Work Record is created — doing so would require tracking wizard save/cancel outcomes across screens, which was judged more complexity than this phase's scope justifies. Documented here rather than half-built.
3. Reference matching is exact-name, case-insensitive only, as instructed — a differently-worded organization/district/project name in the email will not connect automatically; the user can still fix it in the wizard.
4. `selectDataProvider()`'s SharePoint-vs-prototype choice is unaffected by this feature; Inbox Intelligence writes Work Records through whichever provider is already active, unchanged.
5. Action item editing in the review screen supports edit and remove, not adding a brand-new action item from scratch — kept minimal per the review requirements as stated.

## Proposed SharePoint persistence model (NOT provisioned)

No new SharePoint list has been created, and none of the seven already-provisioned lists (`IU_Work_Records`, `IU_Projects`, `IU_Organizations`, `IU_Contacts`, `IU_Work_Categories`, `IU_ORBIT_Deliverables`, `IU_Settings`) is "explicitly intended for this kind of record" per `docs/SHAREPOINT_INTEGRATION_PLAN.md`/`docs/SHAREPOINT_PROVISIONING_SPEC.md` — all seven are Work Record or reference/config data, not email-derived intelligence. Repurposing any of them was avoided as instructed.

If a future phase provisions persistence for this feature, the proposed shape (mirroring the existing provisioning spec's exact style) is:

| Internal name | Display name | Type | Required | Default | Notes |
|---|---|---|---|---|---|
| `Title` | Title | Single line text | Yes | No default | Built-in; suggested Work Record title |
| `AppId` | Application ID | Single line text | Yes, unique | No default | Stable ID, same pattern as `IU_Work_Records.AppId` |
| `SourceExcerpt` | Source excerpt | Multiple lines plain text | No | `""` | Short excerpt only — **never the full raw email** |
| `SummaryText` | Summary | Multiple lines plain text | Yes | No default | |
| `Priority` | Priority | Choice (`high`,`medium`,`low`) | Yes | `medium` | |
| `NeedsAttention` | Needs attention | Yes/No | Yes | `false` | |
| `ActionItemsJson` | Action items (JSON) | Multiple lines plain text | Yes | `[]` | JSON array, same encoding discipline as `IU_Work_Records.*Json` |
| `FollowUpText` | Follow-up | Multiple lines plain text | No | `""` | |
| `PeopleJson` / `OrganizationsJson` / `DistrictsJson` / `ProjectsJson` / `TagsJson` | (matching) | Multiple lines plain text | Yes | `[]` | JSON string arrays |
| `SuggestedWorkType` | Suggested work type | Single line text | No | null | Unvalidated free text |
| `SuggestedWorkRecordTitle` / `SuggestedWorkRecordDescription` | (matching) | Single line text / Multiple lines plain text | Yes | No default | |
| `LinkedWorkRecordAppId` | Linked Work Record | Single line text | No | null | Stable `IU_Work_Records.AppId`, not a Lookup column — same no-Lookup rule as every other list |
| `ModelUsed` | Model used | Single line text | No | No default | Diagnostic |
| `InputTokens` / `OutputTokens` | (matching) | Number, 0 decimals | No | `0` | Diagnostic |
| `Created` / `Modified` | (built-in) | System Date/Time | — | — | Same provider-metadata rule as `IU_Work_Records` |

No Lookup columns, no Choice-vocabulary duplication, JSON arrays as plain multiline text — consistent with every rule already established for the seven provisioned lists. This table is a proposal for a future phase's review, not an instruction to provision it.

## Live-test step (completed)

Completed as described above: real key in the git-ignored `.env.local`, one non-sensitive synthetic email, one click of Analyze email, full review-screen rendering confirmed, diagnostics confirmed, no automatic persistence confirmed, no client-side credential exposure confirmed. See "Live AI test result."

Suggested next verification (not required for GREEN, optional for a future session): a second, different synthetic email exercising a case with no action items at all (a pure FYI), to further confirm the model doesn't over-generate action items across a wider sample. Not performed in this phase to keep to exactly one paid call as instructed.

## Readiness

**GREEN.** Implementation is complete and passes all automated tests (89/89), build, typecheck, and lint (on every file touched this phase). API-key exposure was verified structurally (zero matches in the rebuilt client bundle) both before and after the live call. The "no API key" and "misconfigured key" failure paths were verified live and failed safely with no credential exposure. One live, paid Anthropic API call with a non-sensitive synthetic email succeeded end-to-end: structured extraction passed schema validation, every requested field rendered correctly in the review screen, model/token diagnostics were captured, and no raw email, no Inbox Intelligence record, and no Work Record was persisted automatically. Exactly one billable model call was made this phase.

Inbox Intelligence V1 is ready for normal use with real IU email content, subject to the known limitations above — most notably that persistence remains session-only (nothing survives a page reload) until a future phase provisions the proposed SharePoint list.
