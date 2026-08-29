import { describe, expect, it } from "vitest";
import Anthropic from "@anthropic-ai/sdk";
import { analyzeEmailWithClaude } from "../lib/anthropic-email-analysis";
import { MAX_EMAIL_LENGTH } from "../lib/anthropic-config";
import { EmailAnalysisSchema, normalizeActionItemDueDate, normalizeEmailAnalysis } from "../lib/inbox-intelligence-models";

// No automated test in this file ever calls the real Anthropic API — every client here
// is a fake object shaped like { messages: { parse } }, never a constructed Anthropic instance.

const VALID_ANALYSIS = {
  summary: "The district asked to reschedule Thursday's planning meeting.",
  priority: "medium" as const,
  needsAttention: true,
  actionItems: [{ action: "Propose two new meeting times", dueDate: "2026-09-05", owner: "me" as const }],
  followUp: "Reply with availability by Friday.",
  people: ["Dana Reyes"],
  organizations: [],
  districts: ["Example District"],
  projects: [],
  tags: ["scheduling"],
  suggestedWorkType: "District meeting",
  suggestedWorkRecord: { title: "Reschedule district planning meeting", description: "Coordinate a new time with the district." },
};

function fakeClient(parse: (params: unknown) => Promise<unknown>) {
  return { messages: { parse } } as unknown as Pick<Anthropic, "messages">;
}

describe("EmailAnalysisSchema", () => {
  it("accepts a well-formed extraction", () => {
    expect(EmailAnalysisSchema.safeParse(VALID_ANALYSIS).success).toBe(true);
  });

  it("rejects an invalid priority value", () => {
    const result = EmailAnalysisSchema.safeParse({ ...VALID_ANALYSIS, priority: "urgent" });
    expect(result.success).toBe(false);
  });

  it("rejects an invalid actionItem owner", () => {
    const result = EmailAnalysisSchema.safeParse({
      ...VALID_ANALYSIS,
      actionItems: [{ action: "x", dueDate: null, owner: "boss" }],
    });
    expect(result.success).toBe(false);
  });

  it("rejects an unknown extra field (strict schema)", () => {
    const result = EmailAnalysisSchema.safeParse({ ...VALID_ANALYSIS, extraField: "should not be here" });
    expect(result.success).toBe(false);
  });

  it("requires every array field to be present", () => {
    const { people: _people, ...withoutPeople } = VALID_ANALYSIS;
    void _people;
    expect(EmailAnalysisSchema.safeParse(withoutPeople).success).toBe(false);
  });
});

describe("normalizeActionItemDueDate", () => {
  it("keeps a well-formed date", () => {
    expect(normalizeActionItemDueDate("2026-09-05")).toBe("2026-09-05");
  });
  it("discards null rather than fabricating a date", () => {
    expect(normalizeActionItemDueDate(null)).toBeNull();
  });
  it("discards a malformed date instead of guessing", () => {
    expect(normalizeActionItemDueDate("next Friday")).toBeNull();
    expect(normalizeActionItemDueDate("09/05/2026")).toBeNull();
  });
});

describe("normalizeEmailAnalysis", () => {
  it("normalizes every action item's due date", () => {
    const parsed = EmailAnalysisSchema.parse({
      ...VALID_ANALYSIS,
      actionItems: [
        { action: "a", dueDate: "2026-09-05", owner: "me" },
        { action: "b", dueDate: "sometime soon", owner: "unknown" },
      ],
    });
    const normalized = normalizeEmailAnalysis(parsed);
    expect(normalized.actionItems[0].dueDate).toBe("2026-09-05");
    expect(normalized.actionItems[1].dueDate).toBeNull();
  });
});

describe("analyzeEmailWithClaude", () => {
  it("returns a validation_error for an empty email without ever building a client call", async () => {
    let called = false;
    const client = fakeClient(async () => {
      called = true;
      return {};
    });
    const result = await analyzeEmailWithClaude("   ", client);
    expect(result.status).toBe("validation_error");
    expect(called).toBe(false);
  });

  it("rejects an email over the configured length limit before calling the model", async () => {
    let called = false;
    const client = fakeClient(async () => {
      called = true;
      return {};
    });
    const result = await analyzeEmailWithClaude("x".repeat(MAX_EMAIL_LENGTH + 1), client);
    expect(result.status).toBe("validation_error");
    expect(called).toBe(false);
  });

  it("returns server_error when no client is available (no API key configured)", async () => {
    const result = await analyzeEmailWithClaude("Subject: test\n\nHello.", null);
    expect(result.status).toBe("server_error");
  });

  it("returns a normalized structured result for a valid model response", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 812, output_tokens: 340 },
      parsed_output: EmailAnalysisSchema.parse(VALID_ANALYSIS),
    }));
    const result = await analyzeEmailWithClaude("Subject: Planning meeting\n\nCan we reschedule?", client);
    expect(result.status).toBe("success");
    if (result.status === "success") {
      expect(result.analysis.suggestedWorkRecord.title).toBe(VALID_ANALYSIS.suggestedWorkRecord.title);
      expect(result.usage).toEqual({ model: "claude-opus-5", inputTokens: 812, outputTokens: 340 });
    }
  });

  it("treats a null parsed_output as invalid model output rather than crashing", async () => {
    const client = fakeClient(async () => ({
      model: "claude-opus-5",
      usage: { input_tokens: 100, output_tokens: 5 },
      parsed_output: null,
    }));
    const result = await analyzeEmailWithClaude("Subject: test\n\nHello.", client);
    expect(result.status).toBe("invalid_model_output");
  });

  it("maps a RateLimitError to a safe network_error message", async () => {
    const client = fakeClient(async () => {
      throw new Anthropic.RateLimitError(429, {}, "rate limited", new Headers());
    });
    const result = await analyzeEmailWithClaude("Subject: test\n\nHello.", client);
    expect(result.status).toBe("network_error");
  });

  it("maps an AuthenticationError to a server_error without ever including a key", async () => {
    const client = fakeClient(async () => {
      throw new Anthropic.AuthenticationError(401, {}, "invalid x-api-key: sk-ant-super-secret-value", new Headers());
    });
    const result = await analyzeEmailWithClaude("Subject: test\n\nHello.", client);
    expect(result.status).toBe("server_error");
    if (result.status === "server_error") expect(result.message).not.toContain("sk-ant-super-secret-value");
  });

  it("maps an unexpected thrown value to a safe network_error", async () => {
    const client = fakeClient(async () => {
      throw new Error("boom");
    });
    const result = await analyzeEmailWithClaude("Subject: test\n\nHello.", client);
    expect(result.status).toBe("network_error");
  });
});
