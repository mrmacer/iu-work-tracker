import { describe, expect, it } from "vitest";
import {
  VoiceAnalysisSchema,
  VoiceCandidateSchema,
  VOICE_CANDIDATE_TYPES,
  isDurationSupportedByTranscript,
  normalizeVoiceAnalysis,
  type VoiceCandidate,
} from "../lib/voice-intelligence-models";
import { COMPLETED_VS_ACTION_TRANSCRIPT, VAGUE_DURATION_TRANSCRIPT } from "./fixtures/voice-transcripts";

function candidate(overrides: Partial<VoiceCandidate> = {}): VoiceCandidate {
  return VoiceCandidateSchema.parse({
    type: "ACTION",
    title: "Send Kim the Discovery materials",
    detail: "Follow up from this morning's conversation.",
    sourceExcerpt: "I need to send Kim the Discovery materials",
    durationText: null,
    ...overrides,
  });
}

describe("VoiceCandidateSchema / VoiceAnalysisSchema", () => {
  it("accepts every documented candidate type", () => {
    for (const type of VOICE_CANDIDATE_TYPES) {
      expect(() => candidate({ type })).not.toThrow();
    }
  });

  it("rejects an unknown candidate type", () => {
    expect(() => VoiceCandidateSchema.parse({ ...candidate(), type: "TODO" })).toThrow();
  });

  it("rejects an empty title", () => {
    expect(() => VoiceCandidateSchema.parse({ ...candidate(), title: "" })).toThrow();
  });

  it("rejects an unknown/extra field (strict schema)", () => {
    expect(() => VoiceCandidateSchema.parse({ ...candidate(), estimatedMinutes: 30 })).toThrow();
  });

  it("caps the candidate array at a bounded maximum", () => {
    const candidates = Array.from({ length: 61 }, () => candidate());
    expect(() => VoiceAnalysisSchema.parse({ candidates })).toThrow();
  });

  it("accepts an empty candidate array (no useful candidates found)", () => {
    expect(() => VoiceAnalysisSchema.parse({ candidates: [] })).not.toThrow();
  });
});

describe("isDurationSupportedByTranscript — EXPLICIT ONLY", () => {
  it("supports duration language that appears in the transcript", () => {
    expect(isDurationSupportedByTranscript("about 30 minutes", COMPLETED_VS_ACTION_TRANSCRIPT)).toBe(true);
  });

  it("is case- and punctuation-insensitive", () => {
    expect(isDurationSupportedByTranscript("ABOUT 30 MINUTES!", COMPLETED_VS_ACTION_TRANSCRIPT)).toBe(true);
  });

  it("rejects duration language that does not appear in the transcript", () => {
    expect(isDurationSupportedByTranscript("about two hours", COMPLETED_VS_ACTION_TRANSCRIPT)).toBe(false);
  });

  it("rejects a vague, non-quoted duration phrase", () => {
    // The transcript says "for a while" — a candidate inventing "45 minutes" must not pass.
    expect(isDurationSupportedByTranscript("45 minutes", VAGUE_DURATION_TRANSCRIPT)).toBe(false);
  });
});

describe("normalizeVoiceAnalysis — duration safeguard", () => {
  it("keeps a durationText that is supported by the transcript", () => {
    const analysis = { candidates: [candidate({ type: "COMPLETED_WORK", durationText: "about 30 minutes" })] };
    const result = normalizeVoiceAnalysis(analysis, COMPLETED_VS_ACTION_TRANSCRIPT);
    expect(result.candidates[0].durationText).toBe("about 30 minutes");
  });

  it("strips a durationText the model invented with no support in the transcript", () => {
    const analysis = { candidates: [candidate({ type: "COMPLETED_WORK", durationText: "three hours" })] };
    const result = normalizeVoiceAnalysis(analysis, COMPLETED_VS_ACTION_TRANSCRIPT);
    expect(result.candidates[0].durationText).toBeNull();
  });

  it("never fabricates a duration for a candidate that had none", () => {
    const analysis = { candidates: [candidate({ durationText: null })] };
    const result = normalizeVoiceAnalysis(analysis, COMPLETED_VS_ACTION_TRANSCRIPT);
    expect(result.candidates[0].durationText).toBeNull();
  });
});

describe("normalizeVoiceAnalysis — deterministic dedupe", () => {
  it("drops an exact repeated (type, title) candidate", () => {
    const analysis = {
      candidates: [
        candidate({ type: "ACTION", title: "Send Kim the Discovery materials" }),
        candidate({ type: "ACTION", title: "send kim the discovery materials" }),
      ],
    };
    const result = normalizeVoiceAnalysis(analysis, COMPLETED_VS_ACTION_TRANSCRIPT);
    expect(result.candidates).toHaveLength(1);
  });

  it("does not merge different candidates that merely share a type", () => {
    const analysis = {
      candidates: [
        candidate({ type: "ACTION", title: "Send Kim the Discovery materials" }),
        candidate({ type: "ACTION", title: "Call Annie about the venue" }),
      ],
    };
    const result = normalizeVoiceAnalysis(analysis, COMPLETED_VS_ACTION_TRANSCRIPT);
    expect(result.candidates).toHaveLength(2);
  });

  it("does not merge the same title across different types", () => {
    const analysis = {
      candidates: [
        candidate({ type: "ACTION", title: "STEELS agenda" }),
        candidate({ type: "COMPLETED_WORK", title: "STEELS agenda" }),
      ],
    };
    const result = normalizeVoiceAnalysis(analysis, COMPLETED_VS_ACTION_TRANSCRIPT);
    expect(result.candidates).toHaveLength(2);
  });
});
