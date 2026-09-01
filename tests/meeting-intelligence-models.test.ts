import { describe, expect, it } from "vitest";
import {
  MEETING_CANDIDATE_TYPES,
  MeetingAnalysisSchema,
  MeetingCandidateSchema,
  normalizeMeetingAnalysis,
  type MeetingCandidate,
} from "../lib/meeting-intelligence-models";
import { MEETING_NOTES, VAGUE_DURATION_NOTES } from "./fixtures/meeting-notes";

function candidate(overrides: Partial<MeetingCandidate> = {}): MeetingCandidate {
  return MeetingCandidateSchema.parse({
    type: "ACTION",
    title: "Call the district about the venue",
    detail: "Follow up before Friday.",
    sourceExcerpt: "Annie will call the district about the venue by Friday.",
    ownerText: "Annie",
    dueText: "Friday",
    durationText: null,
    ...overrides,
  });
}

describe("MeetingCandidateSchema / MeetingAnalysisSchema", () => {
  it("accepts every documented candidate type", () => {
    for (const type of MEETING_CANDIDATE_TYPES) {
      expect(() => candidate({ type })).not.toThrow();
    }
  });

  it("rejects an unknown candidate type", () => {
    expect(() => MeetingCandidateSchema.parse({ ...candidate(), type: "TODO" })).toThrow();
  });

  it("rejects an empty title", () => {
    expect(() => MeetingCandidateSchema.parse({ ...candidate(), title: "" })).toThrow();
  });

  it("rejects an unknown/extra field (strict schema)", () => {
    expect(() => MeetingCandidateSchema.parse({ ...candidate(), assignee: "Annie" })).toThrow();
  });

  it("tolerates missing optional fields via explicit null (SUMMARY-shaped candidate)", () => {
    expect(() =>
      MeetingCandidateSchema.parse({
        type: "SUMMARY",
        title: "STEELS grant on track, fall meeting moved to October",
        detail: "Reviewed grant status, decided to delay the fall network meeting.",
        sourceExcerpt: "",
        ownerText: null,
        dueText: null,
        durationText: null,
      }),
    ).not.toThrow();
  });

  it("accepts an empty candidate array (no useful candidates found)", () => {
    expect(() => MeetingAnalysisSchema.parse({ candidates: [] })).not.toThrow();
  });

  it("caps the candidate array at a bounded maximum", () => {
    const candidates = Array.from({ length: 61 }, () => candidate());
    expect(() => MeetingAnalysisSchema.parse({ candidates })).toThrow();
  });
});

describe("normalizeMeetingAnalysis — duration safeguard (reuses Voice's EXPLICIT ONLY rule)", () => {
  it("keeps a durationText supported by the meeting text", () => {
    const analysis = { candidates: [candidate({ type: "COMPLETED_WORK", durationText: "about 30 minutes" })] };
    const result = normalizeMeetingAnalysis(analysis, MEETING_NOTES);
    expect(result.candidates[0].durationText).toBe("about 30 minutes");
  });

  it("strips a durationText with no support in the meeting text", () => {
    const analysis = { candidates: [candidate({ type: "COMPLETED_WORK", durationText: "three hours" })] };
    const result = normalizeMeetingAnalysis(analysis, MEETING_NOTES);
    expect(result.candidates[0].durationText).toBeNull();
  });

  it("rejects a vague duration phrase (only 'for a while' present)", () => {
    const analysis = { candidates: [candidate({ type: "COMPLETED_WORK", durationText: "45 minutes" })] };
    const result = normalizeMeetingAnalysis(analysis, VAGUE_DURATION_NOTES);
    expect(result.candidates[0].durationText).toBeNull();
  });
});

describe("normalizeMeetingAnalysis — deterministic dedupe", () => {
  it("drops an exact repeated (type, title) candidate", () => {
    const analysis = {
      candidates: [
        candidate({ type: "ACTION", title: "Call the district about the venue" }),
        candidate({ type: "ACTION", title: "call the district about the venue" }),
      ],
    };
    const result = normalizeMeetingAnalysis(analysis, MEETING_NOTES);
    expect(result.candidates).toHaveLength(1);
  });

  it("does not merge different candidates that merely share a type", () => {
    const analysis = {
      candidates: [
        candidate({ type: "ACTION", title: "Call the district about the venue" }),
        candidate({ type: "ACTION", title: "Send the updated agenda", ownerText: null, dueText: null }),
      ],
    };
    const result = normalizeMeetingAnalysis(analysis, MEETING_NOTES);
    expect(result.candidates).toHaveLength(2);
  });

  it("does not merge the same title across different types", () => {
    const analysis = {
      candidates: [
        candidate({ type: "ACTION", title: "Fall network meeting" }),
        candidate({ type: "DECISION", title: "Fall network meeting", ownerText: null, dueText: null }),
      ],
    };
    const result = normalizeMeetingAnalysis(analysis, MEETING_NOTES);
    expect(result.candidates).toHaveLength(2);
  });
});
