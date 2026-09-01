import { describe, expect, it } from "vitest";
import Anthropic from "@anthropic-ai/sdk";
import { analyzeTranscriptWithClaude } from "../lib/anthropic-voice-analysis";
import { MAX_TRANSCRIPT_LENGTH } from "../lib/voice-intelligence-config";
import { VoiceAnalysisSchema, type VoiceCandidate } from "../lib/voice-intelligence-models";
import { COMPLETED_VS_ACTION_TRANSCRIPT, MESSY_TRANSCRIPT } from "./fixtures/voice-transcripts";

// No automated test in this file ever calls the real Anthropic API — every client here is a
// fake object shaped like { messages: { parse } }, never a constructed Anthropic instance.
// The "segmentation" tests below hand-author the model's response to prove the pipeline
// (schema validation, normalization, non-collapsing of independent candidates) behaves
// correctly GIVEN a well-formed response — they do not and cannot prove live-model
// segmentation quality. That requires a real, deliberately separate, manually authorized call.

function fakeClient(parse: (params: unknown) => Promise<unknown>) {
  return { messages: { parse } } as unknown as Pick<Anthropic, "messages">;
}

function candidate(overrides: Partial<VoiceCandidate> = {}): VoiceCandidate {
  return {
    type: "ACTION",
    title: "Send Kim the Discovery materials",
    detail: "Follow-up from this morning's conversation with North Schuylkill.",
    sourceExcerpt: "I need to send Kim the Discovery materials",
    durationText: null,
    ...overrides,
  };
}

describe("analyzeTranscriptWithClaude", () => {
  it("returns a validation_error for an empty transcript without ever building a client call", async () => {
    let called = false;
    const client = fakeClient(async () => {
      called = true;
      return {};
    });
    const result = await analyzeTranscriptWithClaude("   ", client);
    expect(result.status).toBe("validation_error");
    expect(called).toBe(false);
  });

  it("rejects a transcript over the configured length limit before calling the model", async () => {
    let called = false;
    const client = fakeClient(async () => {
      called = true;
      return {};
    });
    const result = await analyzeTranscriptWithClaude("x".repeat(MAX_TRANSCRIPT_LENGTH + 1), client);
    expect(result.status).toBe("validation_error");
    expect(called).toBe(false);
    if (result.status === "validation_error") {
      expect(result.message).toMatch(/split it/i);
    }
  });

  it("returns server_error when no client is available (no API key configured)", async () => {
    const result = await analyzeTranscriptWithClaude("Met with the team this morning.", null);
    expect(result.status).toBe("server_error");
  });

  it("returns a normalized structured result for a valid model response", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 900, output_tokens: 420 },
      parsed_output: VoiceAnalysisSchema.parse({ candidates: [candidate()] }),
    }));
    const result = await analyzeTranscriptWithClaude(COMPLETED_VS_ACTION_TRANSCRIPT, client);
    expect(result.status).toBe("success");
    if (result.status === "success") {
      expect(result.analysis.candidates).toHaveLength(1);
      expect(result.usage).toEqual({ model: "claude-opus-5", inputTokens: 900, outputTokens: 420 });
    }
  });

  it("applies the duration safeguard to the model's response before returning it", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 500, output_tokens: 200 },
      parsed_output: VoiceAnalysisSchema.parse({
        candidates: [candidate({ type: "COMPLETED_WORK", durationText: "about 30 minutes" }), candidate({ title: "invented duration", durationText: "three hours" })],
      }),
    }));
    const result = await analyzeTranscriptWithClaude(COMPLETED_VS_ACTION_TRANSCRIPT, client);
    expect(result.status).toBe("success");
    if (result.status === "success") {
      expect(result.analysis.candidates[0].durationText).toBe("about 30 minutes");
      expect(result.analysis.candidates[1].durationText).toBeNull();
    }
  });

  it("treats a null parsed_output as invalid model output rather than crashing", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 100, output_tokens: 5 },
      parsed_output: null,
    }));
    const result = await analyzeTranscriptWithClaude("Met with the team this morning.", client);
    expect(result.status).toBe("invalid_model_output");
  });

  it("maps a RateLimitError to a safe network_error message", async () => {
    const client = fakeClient(async () => {
      throw new Anthropic.RateLimitError(429, {}, "rate limited", new Headers());
    });
    const result = await analyzeTranscriptWithClaude("Met with the team this morning.", client);
    expect(result.status).toBe("network_error");
  });

  it("maps an AuthenticationError to a server_error without ever including a key", async () => {
    const client = fakeClient(async () => {
      throw new Anthropic.AuthenticationError(401, {}, "invalid x-api-key: sk-ant-super-secret-value", new Headers());
    });
    const result = await analyzeTranscriptWithClaude("Met with the team this morning.", client);
    expect(result.status).toBe("server_error");
    if (result.status === "server_error") expect(result.message).not.toContain("sk-ant-super-secret-value");
  });

  it("maps an unexpected thrown value to a safe network_error", async () => {
    const client = fakeClient(async () => {
      throw new Error("boom");
    });
    const result = await analyzeTranscriptWithClaude("Met with the team this morning.", client);
    expect(result.status).toBe("network_error");
  });
});

describe("analyzeTranscriptWithClaude — segmentation pipeline (hand-authored response, not live-model proof)", () => {
  it("keeps independently-segmented candidates separate rather than collapsing them", async () => {
    // Hand-authored to represent what GOOD segmentation of MESSY_TRANSCRIPT looks like — this
    // proves the pipeline preserves independent candidates and applies normalization
    // correctly; it does not prove the live model actually segments this well.
    const wellSegmented: VoiceCandidate[] = [
      candidate({
        type: "COMPLETED_WORK",
        title: "Met with North Schuylkill about science resources",
        sourceExcerpt: "I met with North Schuylkill this morning about science resources",
        durationText: "about an hour",
      }),
      candidate({
        type: "ACTION",
        title: "Send Kim the Discovery materials",
        sourceExcerpt: "I need to send Kim the Discovery materials",
      }),
      candidate({
        type: "IDEA",
        title: "Add a partner map to the STEM site",
        sourceExcerpt: "the STEM site should have a partner map",
        durationText: null,
      }),
      candidate({
        type: "DECISION",
        title: "Delay the fall network meeting until October",
        sourceExcerpt: "we should hold off on the fall network meeting until October",
      }),
      candidate({
        type: "QUESTION",
        title: "Could DEP workforce funding support the STEM competition?",
        sourceExcerpt: "I wonder if DEP has workforce development money",
      }),
      candidate({
        type: "PROJECT",
        title: "Keystone STEM Competition registration on track",
        sourceExcerpt: "we've got about half the schools registered so far, on track for the deadline",
      }),
      candidate({
        type: "KNOWLEDGE",
        title: "Lead with impact numbers when reaching district contacts",
        sourceExcerpt: "district contacts really respond better when you lead with the impact numbers",
      }),
      candidate({
        type: "PERSON",
        title: "Kim is the reliable Discovery-materials contact",
        sourceExcerpt: "Kim's become the go-to person for anything Discovery-related",
      }),
    ];
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 2100, output_tokens: 950 },
      parsed_output: VoiceAnalysisSchema.parse({ candidates: wellSegmented }),
    }));
    const result = await analyzeTranscriptWithClaude(MESSY_TRANSCRIPT, client);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;

    expect(result.analysis.candidates).toHaveLength(8);
    const types = result.analysis.candidates.map((c) => c.type);
    expect(new Set(types).size).toBeGreaterThan(1); // not collapsed into one type/summary

    // Casual, non-meaningful mentions (Zoom, traffic on 81) never became candidates because
    // this mock never proposed them — the deterministic pipeline has nothing to police here;
    // avoiding that pollution is the system prompt's job (docs/AI_HANDOFF.md "Voice
    // Intelligence V1" — same caveat as Email Noise Torture Test's live-model limitation).
    expect(types).not.toContain("ORGANIZATION");

    // Completed work keeps its explicit, transcript-supported duration.
    const completed = result.analysis.candidates.find((c) => c.type === "COMPLETED_WORK");
    expect(completed?.durationText).toBe("about an hour");

    // The repeated "send Kim the Discovery materials" action was only proposed once by this
    // mock (representing correct model-side dedup); normalizeVoiceAnalysis's own dedupe is
    // covered directly in tests/voice-intelligence-models.test.ts.
    expect(result.analysis.candidates.filter((c) => c.type === "ACTION")).toHaveLength(1);
  });
});
