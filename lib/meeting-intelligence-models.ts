import { z } from "zod";
import type { ProviderMetadata } from "./models";
import type { ValidationIssue } from "./validation";
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

// ---------------------------------------------------------------------------------------
// Durable Meeting Record (Patch 6B — docs/AI_HANDOFF.md "Meeting Notes durability"). A
// MeetingRecord is the one durable artifact this patch introduces: meeting details, agenda,
// notes, the CURRENT human-reviewed candidate state, and the deterministic minutes composed
// from that state. Nothing here triggers a second AI request — see
// lib/anthropic-meeting-analysis.ts, called only from the explicit Analyze/Re-analyze action.
// ---------------------------------------------------------------------------------------

export const MEETING_RECORD_SCHEMA_VERSION = 1 as const;

/**
 * A reviewed candidate exactly as it will be durably persisted — the AI candidate's fields
 * plus the human's selected/ignored decision. Deliberately does NOT include the client-only
 * React list key the review UI uses (`id`) — that is transient UI state, not domain data, and
 * is regenerated on load. Persisting exactly the CURRENT edited candidate state — never the
 * original AI output — is the human-authority guarantee this record type exists to preserve.
 */
export const ReviewedMeetingCandidateSchema = MeetingCandidateSchema.extend({
  selected: z.boolean(),
}).strict();

export const ReviewedMeetingCandidatesSchema = z.array(ReviewedMeetingCandidateSchema).max(60);

export type ReviewedMeetingCandidate = z.infer<typeof ReviewedMeetingCandidateSchema>;

/**
 * A durable Meeting Record. `reviewedCandidates` may be an empty array and `minutesText` may
 * be the metadata-only form of buildDraftMinutes() — a meeting may be saved before it is ever
 * analyzed (an agenda prepared ahead of time, or notes still in progress); nothing is
 * fabricated to fill these in. `analysisModel`/`analyzedAt` are only ever set together, and
 * only once at least one explicit analysis has completed.
 */
export type MeetingRecord = {
  appId: string;
  schemaVersion: typeof MEETING_RECORD_SCHEMA_VERSION;
  title: string;
  meetingDate: string;
  meetingType: string;
  attendeesText: string;
  agendaText: string;
  notesText: string;
  reviewedCandidates: ReviewedMeetingCandidate[];
  minutesText: string;
  analysisModel: string | null;
  analyzedAt: string | null;
  metadata: ProviderMetadata;
};

/**
 * Builds a fresh, unsaved durable record from the current editor state. `metadata.version`
 * stays `0` so the provider's create()/update() branch always routes a first save through
 * create() — the exact same convention WorkRecord and InboxIntelligenceRecord already use.
 * Pure: performs no I/O.
 */
export function buildMeetingRecord(input: {
  appId: string;
  title: string;
  meetingDate: string;
  meetingType: string;
  attendeesText: string;
  agendaText: string;
  notesText: string;
  reviewedCandidates: ReviewedMeetingCandidate[];
  minutesText: string;
  analysisModel: string | null;
  analyzedAt: string | null;
}): MeetingRecord {
  return {
    ...input,
    schemaVersion: MEETING_RECORD_SCHEMA_VERSION,
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
  };
}

/**
 * Runtime-shape validation for a record about to be written — mirrors
 * validateInboxIntelligenceRecord's discipline (lib/sharepoint-inbox-intelligence.ts): a
 * human edit in the review screen must satisfy the same shape the AI extraction pipeline
 * itself enforces, so this never accepts something the AI path would have rejected.
 */
export function validateMeetingRecordShape(record: MeetingRecord): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  if (record.schemaVersion !== MEETING_RECORD_SCHEMA_VERSION) {
    issues.push({ path: "schemaVersion", code: "unsupported_schema", message: `schemaVersion must be ${MEETING_RECORD_SCHEMA_VERSION}.` });
  }
  const candidatesResult = ReviewedMeetingCandidatesSchema.safeParse(record.reviewedCandidates);
  if (!candidatesResult.success) {
    issues.push({ path: "reviewedCandidates", code: "invalid_candidates", message: "The reviewed Meeting Intelligence candidates no longer match the expected shape." });
  }
  if ((record.analysisModel === null) !== (record.analyzedAt === null)) {
    issues.push({ path: "analysisModel", code: "inconsistent_analysis_metadata", message: "analysisModel and analyzedAt must both be set or both be null." });
  }
  return issues;
}
