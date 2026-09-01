import Anthropic from "@anthropic-ai/sdk";
import { analyzeMeetingWithClaude, type AnalyzeMeetingResult, type MeetingAnalysisInput } from "../../../lib/anthropic-meeting-analysis";

export const dynamic = "force-dynamic";

// The Anthropic API key is a server-only secret read from process.env (standard Next.js
// server-side env access, matching app/api/inbox-intelligence/route.ts and
// app/api/voice-intelligence/route.ts exactly). It is never read from a NEXT_PUBLIC_*
// variable, never sent to the browser, and never included in any response body or log line.
// Meeting details/agenda/notes are likewise never logged or retained here — only the caught
// error object is logged below, never the request body.
function buildClient(): Pick<Anthropic, "messages"> | null {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return null;
  return new Anthropic({ apiKey });
}

function statusCodeFor(result: AnalyzeMeetingResult): number {
  switch (result.status) {
    case "success":
      return 200;
    case "validation_error":
      return 400;
    case "invalid_model_output":
      return 502;
    case "network_error":
      return 502;
    case "server_error":
      return 500;
  }
}

function stringField(value: unknown): string {
  return typeof value === "string" ? value : "";
}

export async function POST(request: Request) {
  try {
    const body = (await request.json().catch(() => null)) as Partial<Record<keyof MeetingAnalysisInput, unknown>> | null;
    const input: MeetingAnalysisInput = {
      title: stringField(body?.title),
      date: stringField(body?.date),
      meetingType: stringField(body?.meetingType),
      attendeesText: stringField(body?.attendeesText),
      agendaText: stringField(body?.agendaText),
      notesText: stringField(body?.notesText),
    };
    const result = await analyzeMeetingWithClaude(input, buildClient());
    return Response.json(result, { status: statusCodeFor(result) });
  } catch (error) {
    console.error("Meeting Intelligence analysis failed", error);
    return Response.json(
      { status: "server_error", message: "The meeting could not be analyzed." } satisfies AnalyzeMeetingResult,
      { status: 500 },
    );
  }
}
