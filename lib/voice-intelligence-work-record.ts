import type { WorkRecord } from "./models";
import type { VoiceCandidate } from "./voice-intelligence-models";

const DETERMINISTIC_DURATION_PATTERN =
  /^(?:about|around|approximately|roughly|maybe)?\s*(an?|\d+(?:\.\d+)?)\s*(hours?|hrs?|minutes?|mins?)$/i;

/**
 * Deterministic, intentionally narrow durationText → minutes conversion — the SECOND safety
 * layer beyond Patch 4's transcript-support check (isDurationSupportedByTranscript in
 * voice-intelligence-models.ts). Only explicit, unambiguous forms convert: a plain number (or
 * decimal) plus hour(s)/hr(s)/minute(s)/min(s), with an optional leading approximate modifier
 * (about/around/approximately/roughly/maybe) ignored for the conversion itself, plus the
 * fixed idiom "a/an hour" → 60. Word numbers ("two hours") and vague phrases ("a while",
 * "most of the morning", "a couple hours") deliberately do not match — this returns null
 * rather than guessing, per docs/AI_HANDOFF.md "Voice → Universal Work Record Handoff".
 */
export function parseDeterministicDurationMinutes(durationText: string): number | null {
  const match = DETERMINISTIC_DURATION_PATTERN.exec(durationText.trim());
  if (!match) return null;
  const [, quantityRaw, unitRaw] = match;
  const isHour = /^(hour|hr)/i.test(unitRaw);

  let quantity: number;
  if (/^an?$/i.test(quantityRaw)) {
    if (!isHour) return null; // "a/an" is only a safe idiom for "a/an hour", not "a minute"
    quantity = 1;
  } else {
    quantity = Number(quantityRaw);
  }
  if (!Number.isFinite(quantity) || quantity <= 0) return null;

  return Math.round(isHour ? quantity * 60 : quantity);
}

/**
 * Maps a reviewed COMPLETED_WORK Voice candidate onto a blank Work Record draft. Pure and
 * side-effect-free — never writes through the DataProvider; the caller still opens the
 * existing Log Work wizard with this draft, and the user must review and explicitly save
 * through the existing, unmodified save path.
 *
 * Deliberately conservative, unlike lib/inbox-intelligence-work-record.ts's email mapper:
 * maps ONLY title, description, and (only when deterministically safe) duration.
 * sourceExcerpt is never copied into notes — it is review provenance, not work-log content.
 * Category, districts/organizations/projects/contacts, reach, and ORBIT are left exactly as
 * baseRecord's normal human-controlled defaults — candidate-to-candidate relationship
 * inference and reporting/ORBIT inference are both explicitly out of scope for this handoff.
 */
export function buildWorkRecordDraftFromVoiceCandidate(candidate: VoiceCandidate, baseRecord: WorkRecord): WorkRecord {
  const minutes = candidate.durationText ? parseDeterministicDurationMinutes(candidate.durationText) : null;
  return {
    ...baseRecord,
    title: candidate.title,
    description: candidate.detail,
    durationMinutes: minutes ?? baseRecord.durationMinutes,
  };
}
