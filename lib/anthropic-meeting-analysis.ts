import Anthropic from "@anthropic-ai/sdk";
import { zodOutputFormat } from "@anthropic-ai/sdk/helpers/zod";
import {
  MAX_MEETING_CONTENT_LENGTH,
  MEETING_ANTHROPIC_EFFORT,
  MEETING_ANTHROPIC_MAX_OUTPUT_TOKENS,
  MEETING_ANTHROPIC_MODEL,
} from "./meeting-intelligence-config";
import { MeetingAnalysisSchema, normalizeMeetingAnalysis, type MeetingAnalysis } from "./meeting-intelligence-models";

export type MeetingAnalysisInput = {
  title: string;
  date: string;
  meetingType: string;
  attendeesText: string;
  agendaText: string;
  notesText: string;
};

export type AnalyzeMeetingUsage = {
  model: string;
  inputTokens: number;
  outputTokens: number;
};

export type AnalyzeMeetingResult =
  | { status: "success"; analysis: MeetingAnalysis; usage: AnalyzeMeetingUsage }
  | { status: "validation_error"; message: string }
  | { status: "invalid_model_output"; message: string }
  | { status: "network_error"; message: string }
  | { status: "server_error"; message: string };

export const MEETING_SYSTEM_PROMPT = `You extract structured intelligence from meeting details, an agenda, and general notes for a school Intermediate Unit professional. AI proposes; a human reviews everything before any of it is used anywhere.

Candidate types (use exactly one per candidate): SUMMARY, DECISION, ACTION, COMPLETED_WORK, IDEA, QUESTION, KNOWLEDGE, FOLLOW_UP_AGENDA.

Rules:
- Return exactly one SUMMARY candidate: a concise, meeting-level summary.
- DECISION only when the notes indicate an actual decision was made — never turn ordinary discussion into a decision.
- ACTION only when something should happen after the meeting — a future task, not something already done. Set ownerText only when the notes explicitly assign a person ("Annie will call the district," "I will send the draft") — never guess or default to the speaker for a general statement like "we should send it." Set dueText only when the notes state an explicit due phrase ("tomorrow," "Friday," "September 15," "before the next meeting") — copy that language closely; never invent a calendar date.
- COMPLETED_WORK only for something actually completed or performed during or before the meeting — never future work described as if already done. Set durationText only when the notes state an explicit approximate duration ("30 minutes," "about an hour"); never infer duration from meeting length or notes length.
- IDEA is a potential concept or suggestion not yet decided. QUESTION is an unresolved question or issue. KNOWLEDGE is a durable factual or contextual insight worth remembering beyond this meeting. FOLLOW_UP_AGENDA is an unresolved topic that clearly belongs on a future meeting's agenda — not every open thread qualifies.
- Do not invent people, organizations, projects, or districts as entities.
- Do not turn every passing comment into a candidate — ignore personal or off-topic clutter.
- Never propose the same candidate more than once.
- Every candidate except SUMMARY needs a short sourceExcerpt: a concise, close-to-verbatim quote from the agenda or notes that supports it.
- Return structured output only.`;

function promptContentFor(input: MeetingAnalysisInput, agenda: string, notes: string): string {
  return `Meeting title: ${input.title.trim() || "(untitled)"}
Date: ${input.date.trim() || "(not set)"}
Meeting type: ${input.meetingType.trim() || "(not set)"}
Attendees: ${input.attendeesText.trim() || "(not listed)"}

Agenda:
${agenda || "(none provided)"}

Notes:
${notes || "(none provided)"}`;
}

/**
 * Calls Claude to segment meeting details/agenda/notes into independent structured
 * candidates and validates the result strictly server-side, following the exact same
 * dependency-injected pattern as lib/anthropic-voice-analysis.ts and
 * lib/anthropic-email-analysis.ts. The caller supplies the client so a missing API key never
 * reaches this function, and tests can inject a fake client without any network access. Only
 * the six analysis-relevant fields are ever sent — no SharePoint context, no reference data,
 * no RAG/vector lookup.
 */
export async function analyzeMeetingWithClaude(
  input: MeetingAnalysisInput,
  client: Pick<Anthropic, "messages"> | null,
): Promise<AnalyzeMeetingResult> {
  const agenda = input.agendaText.trim();
  const notes = input.notesText.trim();
  if (!agenda && !notes) {
    return { status: "validation_error", message: "Add agenda items or notes before analyzing." };
  }
  const combinedLength = input.agendaText.length + input.notesText.length;
  if (combinedLength > MAX_MEETING_CONTENT_LENGTH) {
    return {
      status: "validation_error",
      message: `This meeting's agenda and notes are too long combined (${combinedLength.toLocaleString()} characters). The limit is ${MAX_MEETING_CONTENT_LENGTH.toLocaleString()} characters — trim it and try again.`,
    };
  }
  if (!client) {
    return { status: "server_error", message: "AI analysis is not configured." };
  }

  // The only text ever checked against for the duration safeguard — never the meeting title,
  // type, or attendees, which could never contain a supportable duration phrase anyway.
  const meetingText = [agenda, notes].filter(Boolean).join("\n\n");

  let response;
  try {
    response = await client.messages.parse({
      model: MEETING_ANTHROPIC_MODEL,
      max_tokens: MEETING_ANTHROPIC_MAX_OUTPUT_TOKENS,
      system: MEETING_SYSTEM_PROMPT,
      output_config: {
        format: zodOutputFormat(MeetingAnalysisSchema),
        effort: MEETING_ANTHROPIC_EFFORT,
      },
      messages: [{ role: "user", content: `Analyze this meeting:\n\n${promptContentFor(input, agenda, notes)}` }],
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
    analysis: normalizeMeetingAnalysis(response.parsed_output, meetingText),
    usage: {
      model: response.model,
      inputTokens: response.usage.input_tokens,
      outputTokens: response.usage.output_tokens,
    },
  };
}
