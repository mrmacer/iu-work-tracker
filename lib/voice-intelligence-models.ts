import { z } from "zod";

// Runtime shape AI extraction must conform to. Deliberately the smallest useful model per
// docs/AI_HANDOFF.md "Voice Intelligence V1" — candidates are proposals only, never
// authoritative People/Organization/Project/Knowledge records, and nothing here is persisted.

export const VOICE_CANDIDATE_TYPES = [
  "COMPLETED_WORK",
  "ACTION",
  "PERSON",
  "ORGANIZATION",
  "DISTRICT",
  "PROJECT",
  "IDEA",
  "DECISION",
  "QUESTION",
  "KNOWLEDGE",
] as const;

export const VoiceCandidateTypeSchema = z.enum(VOICE_CANDIDATE_TYPES);
export type VoiceCandidateType = z.infer<typeof VoiceCandidateTypeSchema>;

/** User-facing label with normal capitalization for a candidate type. */
export const VOICE_CANDIDATE_TYPE_LABELS: Record<VoiceCandidateType, string> = {
  COMPLETED_WORK: "Completed work",
  ACTION: "Action",
  PERSON: "Person",
  ORGANIZATION: "Organization",
  DISTRICT: "District",
  PROJECT: "Project",
  IDEA: "Idea",
  DECISION: "Decision",
  QUESTION: "Question",
  KNOWLEDGE: "Knowledge",
};

export const VoiceCandidateSchema = z
  .object({
    type: VoiceCandidateTypeSchema,
    title: z.string().trim().min(1).max(200),
    detail: z.string().trim().max(1000),
    sourceExcerpt: z.string().trim().min(1).max(400),
    // Only ever present when the transcript states an explicit approximate duration — see
    // stripUnsupportedDurations() below. Never inferred, never normalized to minutes.
    durationText: z.string().trim().min(1).max(60).nullable(),
  })
  .strict();

export const VoiceAnalysisSchema = z
  .object({
    candidates: z.array(VoiceCandidateSchema).max(60),
  })
  .strict();

export type VoiceCandidate = z.infer<typeof VoiceCandidateSchema>;
export type VoiceAnalysis = z.infer<typeof VoiceAnalysisSchema>;

function normalizeForMatch(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

/**
 * EXPLICIT ONLY (docs/AI_HANDOFF.md "Voice Intelligence V1" duration rule). A durationText the
 * model returned is kept only when materially equivalent language actually appears in the
 * original transcript — checked as a normalized substring match, deliberately not a natural-
 * language duration parser. Anything unsupported is dropped to null rather than trusted.
 */
export function isDurationSupportedByTranscript(durationText: string, transcript: string): boolean {
  const normalizedDuration = normalizeForMatch(durationText);
  if (!normalizedDuration) return false;
  return normalizeForMatch(transcript).includes(normalizedDuration);
}

function stripUnsupportedDurations(candidates: VoiceCandidate[], transcript: string): VoiceCandidate[] {
  return candidates.map((candidate) => ({
    ...candidate,
    durationText:
      candidate.durationText && isDurationSupportedByTranscript(candidate.durationText, transcript)
        ? candidate.durationText
        : null,
  }));
}

/**
 * Smallest deterministic duplicate guard (docs/AI_HANDOFF.md "Voice Intelligence V1" —
 * mirrors the Email Noise Torture Test's "avoid candidate pollution" spirit): drops a later
 * candidate only when both its type and normalized title exactly match an earlier one. This
 * is not semantic deduplication — different ideas about the same person/topic are preserved.
 */
function dedupeCandidates(candidates: VoiceCandidate[]): VoiceCandidate[] {
  const seen = new Set<string>();
  return candidates.filter((candidate) => {
    const key = `${candidate.type}:${normalizeForMatch(candidate.title)}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

/** Deterministic post-processing applied to every model response before it reaches review UI. */
export function normalizeVoiceAnalysis(analysis: VoiceAnalysis, transcript: string): VoiceAnalysis {
  return {
    candidates: dedupeCandidates(stripUnsupportedDurations(analysis.candidates, transcript)),
  };
}
