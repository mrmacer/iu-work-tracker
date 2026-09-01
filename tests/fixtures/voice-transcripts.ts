// Reusable, composable synthetic transcript fixtures for Voice Intelligence tests
// (docs/AI_HANDOFF.md "Voice Intelligence V1"). Entirely fictional names/organizations — no
// real personal or confidential recordings. These exist to test the deterministic pipeline
// (schema, normalization, UI, server boundary, non-persistence) — not to prove live-model
// segmentation quality, which requires a real (never-in-tests) Anthropic call. See
// tests/voice-intelligence-analysis.test.ts and the final Patch 4 report's recommended live
// torture-test transcripts.

/**
 * Deliberately messy: 8-12 topic shifts covering every category the Patch 4 spec asked for —
 * a completed district meeting with an explicit approximate duration, a follow-up action
 * (stated twice, for dedup), a website idea, a person relationship note, an organization
 * mentioned only in passing (should NOT become a candidate), a decision about an event date,
 * an unanswered funding question, a project status update, a lesson learned, and an
 * unrelated random aside (should also NOT become a candidate).
 */
export const MESSY_TRANSCRIPT = `Okay, quick voice memo on the drive home. I met with North Schuylkill this morning about science resources — that went about an hour, really productive.

I need to send Kim the Discovery materials we talked about. Actually yeah, I definitely need to send Kim those Discovery materials, don't let me forget.

Random thought — I was thinking the STEM site should have a partner map, like an interactive one showing all our district partners. Could be a cool addition.

Annie mentioned Zoom was being weird on her end today, not a big deal, she got it sorted.

Annie said we should hold off on the fall network meeting until October, so that's decided — pushing it to October.

I wonder if DEP has workforce development money we could tap for the STEM competition. Need to look into that.

The Keystone STEM Competition planning is moving along — we've got about half the schools registered so far, on track for the deadline.

One thing I learned today: the district contacts really respond better when you lead with the impact numbers instead of the activity description. Good to remember for future outreach.

Kim's become the go-to person for anything Discovery-related — she always gets back to me the same day, which has been a huge help.

Random aside, traffic on 81 was actually not bad today for once.

Also — yeah, send Kim the Discovery stuff, that's still on my list.`;

/** A transcript with no genuinely useful content — exercises the "no useful candidates" path. */
export const EMPTY_SIGNAL_TRANSCRIPT = `Just testing this thing out. Checking if the recording is working. Okay yeah, seems fine. That's it, that's the whole memo.`;

/**
 * One completed-work statement with an explicit duration, one clearly future action in
 * correctly future-tense phrasing, and one decision — for the completed-work-vs-action and
 * duration-support tests.
 */
export const COMPLETED_VS_ACTION_TRANSCRIPT = `Spent about 30 minutes this afternoon helping the STEELS committee finalize their agenda. I still need to call Annie about the venue for next month. We decided the kickoff will be the second week of September.`;

/** States a vague, non-explicit duration ("a while") that must NOT survive the duration guard. */
export const VAGUE_DURATION_TRANSCRIPT = `Worked on the grant narrative for a while this morning, made good progress on the budget section.`;
