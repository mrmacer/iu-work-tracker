import { z } from "zod";
import { isDurationSupportedByTranscript } from "./voice-intelligence-models";

// Runtime shape AI extraction must conform to. Candidates are proposals only, never
// authoritative Decisions/Actions/Knowledge records, and nothing here is persisted — see
// docs/AI_HANDOFF.md "Meeting Notes V1".

export const MEETING_CANDIDATE_TYPES = [
  "SUMMARY",
  "DECISION",
  "ACTION",
  "COMPLETED_WORK",
  "IDEA",
  "QUESTION",
  "KNOWLEDGE",
  "FOLLOW_UP_AGENDA",
] as const;

export const MeetingCandidateTypeSchema = z.enum(MEETING_CANDIDATE_TYPES);
export type MeetingCandidateType = z.infer<typeof MeetingCandidateTypeSchema>;

/** User-facing label with normal capitalization for a candidate type. */
export const MEETING_CANDIDATE_TYPE_LABELS: Record<MeetingCandidateType, string> = {
  SUMMARY: "Summary",
  DECISION: "Decision",
  ACTION: "Action",
  COMPLETED_WORK: "Completed work",
  IDEA: "Idea",
  QUESTION: "Question",
  KNOWLEDGE: "Knowledge",
  FOLLOW_UP_AGENDA: "Follow-up agenda",
};

export const MeetingCandidateSchema = z
  .object({
    type: MeetingCandidateTypeSchema,
    title: z.string().trim().min(1).max(200),
    detail: z.string().trim().max(1000),
    // Every candidate except SUMMARY is expected to carry one; not min-length-enforced here
    // since a SUMMARY candidate may reasonably have none — see docs/AI_HANDOFF.md.
    sourceExcerpt: z.string().trim().max(400),
    // ACTION only: populated ONLY when the notes explicitly assign ownership. Never guessed.
    ownerText: z.string().trim().min(1).max(120).nullable(),
    // ACTION only: preserves explicit due language verbatim ("Friday", "before the next
    // meeting") — never a fabricated ISO date. See docs/AI_HANDOFF.md "Due-date safety".
    dueText: z.string().trim().min(1).max(60).nullable(),
    // COMPLETED_WORK only: same EXPLICIT-ONLY philosophy as Voice Intelligence.
    durationText: z.string().trim().min(1).max(60).nullable(),
  })
  .strict();

export const MeetingAnalysisSchema = z
  .object({
    candidates: z.array(MeetingCandidateSchema).max(60),
  })
  .strict();

export type MeetingCandidate = z.infer<typeof MeetingCandidateSchema>;
export type MeetingAnalysis = z.infer<typeof MeetingAnalysisSchema>;

function normalizeForMatch(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function stripUnsupportedDurations(candidates: MeetingCandidate[], meetingText: string): MeetingCandidate[] {
  return candidates.map((candidate) => ({
    ...candidate,
    durationText:
      candidate.durationText && isDurationSupportedByTranscript(candidate.durationText, meetingText)
        ? candidate.durationText
        : null,
  }));
}

/**
 * Smallest deterministic duplicate guard (same rule as Voice Intelligence's dedupe, and the
 * same "avoid candidate pollution" spirit as the Email Noise Torture Test): drops a later
 * candidate only when both its type and normalized title exactly match an earlier one. Never
 * merges across types, never semantic clustering.
 */
function dedupeCandidates(candidates: MeetingCandidate[]): MeetingCandidate[] {
  const seen = new Set<string>();
  return candidates.filter((candidate) => {
    const key = `${candidate.type}:${normalizeForMatch(candidate.title)}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

/**
 * Deterministic post-processing applied to every model response before it reaches review UI.
 * `meetingText` should be the combined agenda + notes text (the only trusted source for the
 * duration safeguard — see lib/anthropic-meeting-analysis.ts).
 */
export function normalizeMeetingAnalysis(analysis: MeetingAnalysis, meetingText: string): MeetingAnalysis {
  return {
    candidates: dedupeCandidates(stripUnsupportedDurations(analysis.candidates, meetingText)),
  };
}
