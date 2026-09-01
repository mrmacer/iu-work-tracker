import type { WorkRecord } from "./models";
import type { MeetingCandidate } from "./meeting-intelligence-models";
import { parseDeterministicDurationMinutes } from "./voice-intelligence-work-record";

/**
 * Maps a reviewed COMPLETED_WORK Meeting candidate onto a blank Work Record draft. Pure and
 * side-effect-free — never writes through the DataProvider; the caller still opens the
 * existing Log Work wizard with this draft, and the user must review and explicitly save
 * through the existing, unmodified save path.
 *
 * Deliberately a tiny adapter rather than a generalized shared mapper: it reuses
 * parseDeterministicDurationMinutes from lib/voice-intelligence-work-record.ts (a pure,
 * candidate-shape-agnostic string→minutes converter) so the exact same conservative duration
 * rule applies to both Voice and Meeting candidates, without touching or risking Patch 5's
 * tested buildWorkRecordDraftFromVoiceCandidate. Maps ONLY title, description, and (only when
 * deterministically safe) duration — category, districts/organizations/projects/contacts,
 * reach, and ORBIT are left exactly as baseRecord's normal human-controlled defaults. No
 * relationship or reporting inference happens here, matching Patch 5's handoff exactly.
 */
export function buildWorkRecordDraftFromMeetingCandidate(candidate: MeetingCandidate, baseRecord: WorkRecord): WorkRecord {
  const minutes = candidate.durationText ? parseDeterministicDurationMinutes(candidate.durationText) : null;
  return {
    ...baseRecord,
    title: candidate.title,
    description: candidate.detail,
    durationMinutes: minutes ?? baseRecord.durationMinutes,
  };
}
