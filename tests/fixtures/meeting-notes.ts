// Reusable synthetic meeting fixtures for Meeting Intelligence tests (docs/AI_HANDOFF.md
// "Meeting Notes V1"). Entirely fictional names/organizations — no real personal or
// confidential meeting content. These exist to test the deterministic pipeline (schema,
// normalization, UI, server boundary, non-persistence) — not to prove live-model extraction
// quality, which requires a real (never-in-tests) Anthropic call.

export const MEETING_AGENDA = `1. STEELS grant status update
2. Fall network meeting date
3. Discovery materials follow-up
4. Open discussion`;

/**
 * Covers: an actual decision, an explicit-owner action, an ambiguous-owner statement that
 * must NOT get an owner, completed work with an explicit duration, an explicit due phrase, a
 * casual off-topic aside that should not become a candidate, and a clearly unresolved topic
 * for a future agenda.
 */
export const MEETING_NOTES = `We reviewed the STEELS grant status — on track, about half the schools registered. Spent about 30 minutes walking through the budget section together.

Annie will call the district about the venue by Friday.

We should send the updated agenda to the full group — someone needs to handle that, not assigned yet.

Decided to move the fall network meeting to October.

Kim's traffic on the way in was rough today, not a big deal.

Still need to figure out whether DEP has workforce development funding available — didn't resolve this, worth revisiting.

We didn't get to the STEM site partner map idea — that should go on next month's agenda.`;

/** A due-phrase test transcript with only vague, non-explicit duration language. */
export const VAGUE_DURATION_NOTES = `Worked through the grant narrative for a while this morning — made good progress on the budget section.`;
