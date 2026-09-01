import { describe, expect, it } from "vitest";
import Anthropic from "@anthropic-ai/sdk";
import { analyzeMeetingWithClaude, type MeetingAnalysisInput } from "../lib/anthropic-meeting-analysis";
import { MAX_MEETING_CONTENT_LENGTH } from "../lib/meeting-intelligence-config";
import { MeetingAnalysisSchema, type MeetingCandidate } from "../lib/meeting-intelligence-models";
import { MEETING_AGENDA, MEETING_NOTES } from "./fixtures/meeting-notes";

// No automated test in this file ever calls the real Anthropic API — every client here is a
// fake object shaped like { messages: { parse } }, never a constructed Anthropic instance.
// The "segmentation pipeline" test hand-authors the model's response to prove the pipeline
// (schema validation, normalization) behaves correctly GIVEN a well-formed response — it does
// not and cannot prove live-model extraction quality.

function fakeClient(parse: (params: unknown) => Promise<unknown>) {
  return { messages: { parse } } as unknown as Pick<Anthropic, "messages">;
}

function baseInput(overrides: Partial<MeetingAnalysisInput> = {}): MeetingAnalysisInput {
  return {
    title: "STEELS quarterly planning",
    date: "2026-09-01",
    meetingType: "District Meeting",
    attendeesText: "Greg, Annie, Kim",
    agendaText: MEETING_AGENDA,
    notesText: MEETING_NOTES,
    ...overrides,
  };
}

function candidate(overrides: Partial<MeetingCandidate> = {}): MeetingCandidate {
  return {
    type: "ACTION",
    title: "Call the district about the venue",
    detail: "Confirm venue details before Friday.",
    sourceExcerpt: "Annie will call the district about the venue by Friday.",
    ownerText: "Annie",
    dueText: "Friday",
    durationText: null,
    ...overrides,
  };
}

describe("analyzeMeetingWithClaude", () => {
  it("returns a validation_error when neither agenda nor notes have content, without calling the model", async () => {
    let called = false;
    const client = fakeClient(async () => {
      called = true;
      return {};
    });
    const result = await analyzeMeetingWithClaude(baseInput({ agendaText: "   ", notesText: "" }), client);
    expect(result.status).toBe("validation_error");
    expect(called).toBe(false);
  });

  it("proceeds when only the agenda has content (notes empty)", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 200, output_tokens: 80 },
      parsed_output: MeetingAnalysisSchema.parse({ candidates: [] }),
    }));
    const result = await analyzeMeetingWithClaude(baseInput({ notesText: "" }), client);
    expect(result.status).toBe("success");
  });

  it("proceeds when only notes have content (agenda empty)", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 200, output_tokens: 80 },
      parsed_output: MeetingAnalysisSchema.parse({ candidates: [] }),
    }));
    const result = await analyzeMeetingWithClaude(baseInput({ agendaText: "" }), client);
    expect(result.status).toBe("success");
  });

  it("does not require a meeting title to analyze", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 200, output_tokens: 80 },
      parsed_output: MeetingAnalysisSchema.parse({ candidates: [] }),
    }));
    const result = await analyzeMeetingWithClaude(baseInput({ title: "", meetingType: "", attendeesText: "" }), client);
    expect(result.status).toBe("success");
  });

  it("rejects agenda+notes over the configured combined length limit before calling the model", async () => {
    let called = false;
    const client = fakeClient(async () => {
      called = true;
      return {};
    });
    const result = await analyzeMeetingWithClaude(
      baseInput({ agendaText: "x".repeat(MAX_MEETING_CONTENT_LENGTH), notesText: "y".repeat(10) }),
      client,
    );
    expect(result.status).toBe("validation_error");
    expect(called).toBe(false);
  });

  it("returns server_error when no client is available (no API key configured)", async () => {
    const result = await analyzeMeetingWithClaude(baseInput(), null);
    expect(result.status).toBe("server_error");
  });

  it("returns a normalized structured result for a valid model response", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 1200, output_tokens: 500 },
      parsed_output: MeetingAnalysisSchema.parse({ candidates: [candidate()] }),
    }));
    const result = await analyzeMeetingWithClaude(baseInput(), client);
    expect(result.status).toBe("success");
    if (result.status === "success") {
      expect(result.analysis.candidates).toHaveLength(1);
      expect(result.usage).toEqual({ model: "claude-opus-5", inputTokens: 1200, outputTokens: 500 });
    }
  });

  it("treats a null parsed_output as invalid model output rather than crashing", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 100, output_tokens: 5 },
      parsed_output: null,
    }));
    const result = await analyzeMeetingWithClaude(baseInput(), client);
    expect(result.status).toBe("invalid_model_output");
  });

  it("maps a RateLimitError to a safe network_error message", async () => {
    const client = fakeClient(async () => {
      throw new Anthropic.RateLimitError(429, {}, "rate limited", new Headers());
    });
    const result = await analyzeMeetingWithClaude(baseInput(), client);
    expect(result.status).toBe("network_error");
  });

  it("maps an AuthenticationError to a server_error without ever including a key", async () => {
    const client = fakeClient(async () => {
      throw new Anthropic.AuthenticationError(401, {}, "invalid x-api-key: sk-ant-super-secret-value", new Headers());
    });
    const result = await analyzeMeetingWithClaude(baseInput(), client);
    expect(result.status).toBe("server_error");
    if (result.status === "server_error") expect(result.message).not.toContain("sk-ant-super-secret-value");
  });
});

describe("analyzeMeetingWithClaude — owner / due safety (pipeline preservation, not live-model proof)", () => {
  it("preserves an explicit ownerText through the pipeline unchanged", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 500, output_tokens: 200 },
      parsed_output: MeetingAnalysisSchema.parse({ candidates: [candidate({ ownerText: "Annie" })] }),
    }));
    const result = await analyzeMeetingWithClaude(baseInput(), client);
    expect(result.status).toBe("success");
    if (result.status === "success") expect(result.analysis.candidates[0].ownerText).toBe("Annie");
  });

  it("preserves an ambiguous (null) ownerText as null — never fabricates one", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 500, output_tokens: 200 },
      parsed_output: MeetingAnalysisSchema.parse({
        candidates: [candidate({ title: "Send the updated agenda", ownerText: null, dueText: null, sourceExcerpt: "someone needs to handle that, not assigned yet" })],
      }),
    }));
    const result = await analyzeMeetingWithClaude(baseInput(), client);
    expect(result.status).toBe("success");
    if (result.status === "success") expect(result.analysis.candidates[0].ownerText).toBeNull();
  });

  it("preserves an explicit dueText through the pipeline unchanged", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 500, output_tokens: 200 },
      parsed_output: MeetingAnalysisSchema.parse({ candidates: [candidate({ dueText: "Friday" })] }),
    }));
    const result = await analyzeMeetingWithClaude(baseInput(), client);
    expect(result.status).toBe("success");
    if (result.status === "success") expect(result.analysis.candidates[0].dueText).toBe("Friday");
  });

  it("preserves a null dueText as null when nothing explicit was stated", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 500, output_tokens: 200 },
      parsed_output: MeetingAnalysisSchema.parse({ candidates: [candidate({ dueText: null })] }),
    }));
    const result = await analyzeMeetingWithClaude(baseInput(), client);
    expect(result.status).toBe("success");
    if (result.status === "success") expect(result.analysis.candidates[0].dueText).toBeNull();
  });
});

describe("analyzeMeetingWithClaude — segmentation pipeline (hand-authored response, not live-model proof)", () => {
  it("keeps independently-segmented candidates across multiple types, applying the duration safeguard", async () => {
    const wellSegmented: MeetingCandidate[] = [
      candidate({
        type: "SUMMARY",
        title: "STEELS on track; fall meeting moved to October",
        detail: "Reviewed grant status, delayed the fall network meeting, several open items remain.",
        sourceExcerpt: "",
        ownerText: null,
        dueText: null,
      }),
      candidate({
        type: "COMPLETED_WORK",
        title: "Reviewed the STEELS grant budget section",
        sourceExcerpt: "Spent about 30 minutes walking through the budget section together.",
        ownerText: null,
        dueText: null,
        durationText: "about 30 minutes",
      }),
      candidate({ type: "ACTION", title: "Call the district about the venue", ownerText: "Annie", dueText: "Friday" }),
      candidate({
        type: "ACTION",
        title: "Send the updated agenda to the full group",
        sourceExcerpt: "someone needs to handle that, not assigned yet",
        ownerText: null,
        dueText: null,
      }),
      candidate({
        type: "DECISION",
        title: "Move the fall network meeting to October",
        sourceExcerpt: "Decided to move the fall network meeting to October.",
        ownerText: null,
        dueText: null,
      }),
      candidate({
        type: "QUESTION",
        title: "Is DEP workforce development funding available?",
        sourceExcerpt: "Still need to figure out whether DEP has workforce development funding available",
        ownerText: null,
        dueText: null,
      }),
      candidate({
        type: "FOLLOW_UP_AGENDA",
        title: "STEM site partner map idea",
        sourceExcerpt: "that should go on next month's agenda",
        ownerText: null,
        dueText: null,
      }),
    ];
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 2200, output_tokens: 900 },
      parsed_output: MeetingAnalysisSchema.parse({ candidates: wellSegmented }),
    }));
    const result = await analyzeMeetingWithClaude(baseInput(), client);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;

    expect(result.analysis.candidates).toHaveLength(7);
    const types = new Set(result.analysis.candidates.map((c) => c.type));
    expect(types.size).toBeGreaterThan(1);

    // Kim's traffic aside was never proposed by this mock — nothing in the deterministic
    // pipeline needs to police that; avoiding it is the system prompt's job.
    const completed = result.analysis.candidates.find((c) => c.type === "COMPLETED_WORK");
    expect(completed?.durationText).toBe("about 30 minutes");

    const actions = result.analysis.candidates.filter((c) => c.type === "ACTION");
    expect(actions.find((a) => a.title.includes("venue"))?.ownerText).toBe("Annie");
    expect(actions.find((a) => a.title.includes("agenda"))?.ownerText).toBeNull();
  });
});
