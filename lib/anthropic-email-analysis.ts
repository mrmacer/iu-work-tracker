import Anthropic from "@anthropic-ai/sdk";
import { zodOutputFormat } from "@anthropic-ai/sdk/helpers/zod";
import { ANTHROPIC_EFFORT, ANTHROPIC_MAX_OUTPUT_TOKENS, ANTHROPIC_MODEL, MAX_EMAIL_LENGTH } from "./anthropic-config";
import { EmailAnalysisSchema, normalizeEmailAnalysis, type EmailAnalysis } from "./inbox-intelligence-models";

export type AnalyzeEmailUsage = {
  model: string;
  inputTokens: number;
  outputTokens: number;
};

export type AnalyzeEmailResult =
  | { status: "success"; analysis: EmailAnalysis; usage: AnalyzeEmailUsage }
  | { status: "validation_error"; message: string }
  | { status: "invalid_model_output"; message: string }
  | { status: "network_error"; message: string }
  | { status: "server_error"; message: string };

const SYSTEM_PROMPT = `You extract structured intelligence from one pasted work email for a school Intermediate Unit professional. The pasted text may include headers, quoted thread history, and a signature — read all of it, but summarize the current message.

Be strict about the difference between two things:
1. WHAT THE EMAIL SAYS — facts, requests, and information actually present in the text.
2. WHAT THE USER MAY NEED TO DO — your inference about follow-up action, clearly separated from #1.

Rules:
- Do not turn every email into an action item. Routine information, FYIs, and confirmations often have zero action items.
- Only include a due date (YYYY-MM-DD) when the email states or clearly implies a specific date. If no real deadline exists, or you are inferring one loosely, set dueDate to null. Never fabricate a date.
- Only list people, organizations, districts, and projects that are actually named in the email text. Never invent or guess entities that are not present.
- "priority" and "needsAttention" reflect this one email, not a general assessment of the sender.
- "suggestedWorkType" is a short free-text guess at the kind of work activity this represents, or null if unclear. It is not validated against any controlled vocabulary — leave matching that up to the application.
- "suggestedWorkRecord" is a minimal starting point for a work log entry a human will review and edit before saving — keep the title short and concrete, and the description one or two sentences.
- Every array field must be present; use an empty array when nothing applies. Never omit a field or return null for an array.`;

/**
 * Calls Claude to extract structured intelligence from one pasted email and validates the
 * result strictly server-side (docs/INBOX_INTELLIGENCE_V1_REPORT.md "AI extraction"). The
 * caller supplies the client so a missing API key never reaches this function, and tests can
 * inject a fake client without any network access.
 */
export async function analyzeEmailWithClaude(
  rawEmail: string,
  client: Pick<Anthropic, "messages"> | null,
): Promise<AnalyzeEmailResult> {
  const trimmed = rawEmail.trim();
  if (!trimmed) {
    return { status: "validation_error", message: "Paste an email before analyzing." };
  }
  if (rawEmail.length > MAX_EMAIL_LENGTH) {
    return {
      status: "validation_error",
      message: `The pasted email is too long (${rawEmail.length} characters). Trim it to ${MAX_EMAIL_LENGTH} characters or fewer.`,
    };
  }
  if (!client) {
    return { status: "server_error", message: "AI analysis is not configured." };
  }

  let response;
  try {
    response = await client.messages.parse({
      model: ANTHROPIC_MODEL,
      max_tokens: ANTHROPIC_MAX_OUTPUT_TOKENS,
      system: SYSTEM_PROMPT,
      output_config: {
        format: zodOutputFormat(EmailAnalysisSchema),
        effort: ANTHROPIC_EFFORT,
      },
      messages: [{ role: "user", content: `Analyze this pasted email:\n\n${trimmed}` }],
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
    analysis: normalizeEmailAnalysis(response.parsed_output),
    usage: {
      model: response.model,
      inputTokens: response.usage.input_tokens,
      outputTokens: response.usage.output_tokens,
    },
  };
}
