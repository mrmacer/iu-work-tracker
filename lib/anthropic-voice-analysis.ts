import Anthropic from "@anthropic-ai/sdk";
import { zodOutputFormat } from "@anthropic-ai/sdk/helpers/zod";
import {
  MAX_TRANSCRIPT_LENGTH,
  VOICE_ANTHROPIC_EFFORT,
  VOICE_ANTHROPIC_MAX_OUTPUT_TOKENS,
  VOICE_ANTHROPIC_MODEL,
} from "./voice-intelligence-config";
import { VoiceAnalysisSchema, normalizeVoiceAnalysis, type VoiceAnalysis } from "./voice-intelligence-models";

export type AnalyzeTranscriptUsage = {
  model: string;
  inputTokens: number;
  outputTokens: number;
};

export type AnalyzeTranscriptResult =
  | { status: "success"; analysis: VoiceAnalysis; usage: AnalyzeTranscriptUsage }
  | { status: "validation_error"; message: string }
  | { status: "invalid_model_output"; message: string }
  | { status: "network_error"; message: string }
  | { status: "server_error"; message: string };

export const VOICE_SYSTEM_PROMPT = `You extract structured intelligence from one pasted voice-note transcript for a school Intermediate Unit professional. The transcript may ramble across many unrelated topics in one recording — segmentation is the point. Identify each distinct, useful thought as its own independent candidate. Do not collapse unrelated topics into one candidate merely because they occurred in the same transcript, and do not return only one summary candidate.

Candidate types (use exactly one per candidate): COMPLETED_WORK, ACTION, PERSON, ORGANIZATION, DISTRICT, PROJECT, IDEA, DECISION, QUESTION, KNOWLEDGE.

Rules:
- Distinguish work already done (COMPLETED_WORK) from work that still needs to happen (ACTION). Never describe a future action in past tense, and never turn completed work into a future task.
- Never infer or estimate a duration. Only set durationText when the transcript states an explicit approximate duration ("about an hour", "30 minutes", "maybe two hours", "most of the morning") — copy that language closely. If no explicit duration is stated, leave durationText null. Never guess duration from transcript length or from how long similar work normally takes.
- Only propose a PERSON, ORGANIZATION, DISTRICT, or PROJECT candidate when the transcript contains genuinely useful context about that entity — never merely because a name was mentioned in passing, and never a PERSON candidate for the speaker referring to themselves.
- Only propose IDEA, DECISION, QUESTION, or KNOWLEDGE candidates when they are genuinely worth remembering, not for every passing remark.
- Never propose the same candidate more than once, even if the speaker repeats themselves.
- Every candidate needs a short sourceExcerpt: a concise, close-to-verbatim quote from the transcript that supports it — this is the evidence a human reviewer will check, not a paraphrase.
- These are proposals only. Never imply a candidate has been approved, completed as a task, or already saved anywhere.
- Return structured output only.`;

/**
 * Calls Claude to segment one pasted voice-note transcript into independent structured
 * candidates and validates the result strictly server-side, following the exact same
 * dependency-injected pattern as lib/anthropic-email-analysis.ts. The caller supplies the
 * client so a missing API key never reaches this function, and tests can inject a fake
 * client without any network access.
 */
export async function analyzeTranscriptWithClaude(
  rawTranscript: string,
  client: Pick<Anthropic, "messages"> | null,
): Promise<AnalyzeTranscriptResult> {
  const trimmed = rawTranscript.trim();
  if (!trimmed) {
    return { status: "validation_error", message: "Paste a transcript before analyzing." };
  }
  if (rawTranscript.length > MAX_TRANSCRIPT_LENGTH) {
    return {
      status: "validation_error",
      message: `This transcript is too long (${rawTranscript.length.toLocaleString()} characters). The limit is ${MAX_TRANSCRIPT_LENGTH.toLocaleString()} characters — split it into smaller pieces and analyze each separately.`,
    };
  }
  if (!client) {
    return { status: "server_error", message: "AI analysis is not configured." };
  }

  let response;
  try {
    response = await client.messages.parse({
      model: VOICE_ANTHROPIC_MODEL,
      max_tokens: VOICE_ANTHROPIC_MAX_OUTPUT_TOKENS,
      system: VOICE_SYSTEM_PROMPT,
      output_config: {
        format: zodOutputFormat(VoiceAnalysisSchema),
        effort: VOICE_ANTHROPIC_EFFORT,
      },
      messages: [{ role: "user", content: `Segment this transcript into independent candidates:\n\n${trimmed}` }],
    });
  } catch (error) {
    if (error instanceof Anthropic.RateLimitError) {
      return { status: "network_error", message: "The AI service is rate-limited right now. Try again in a moment." };
    }
    if (error instanceof Anthropic.AuthenticationError) {
      return { status: "server_error", message: "AI analysis is not configured correctly." };
    }
    if (error instanceof Anthropic.APIError) {
      return { status: "network_error", message: "The AI service could not complete the analysis." };
    }
    return { status: "network_error", message: "The AI service could not be reached." };
  }

  if (!response.parsed_output) {
    return { status: "invalid_model_output", message: "The AI response did not match the expected structure. Nothing was saved." };
  }

  return {
    status: "success",
    analysis: normalizeVoiceAnalysis(response.parsed_output, trimmed),
    usage: {
      model: response.model,
      inputTokens: response.usage.input_tokens,
      outputTokens: response.usage.output_tokens,
    },
  };
}
