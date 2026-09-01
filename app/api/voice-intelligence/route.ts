import Anthropic from "@anthropic-ai/sdk";
import { analyzeTranscriptWithClaude, type AnalyzeTranscriptResult } from "../../../lib/anthropic-voice-analysis";

export const dynamic = "force-dynamic";

// The Anthropic API key is a server-only secret read from process.env (standard Next.js
// server-side env access, matching app/api/inbox-intelligence/route.ts exactly). It is never
// read from a NEXT_PUBLIC_* variable, never sent to the browser, and never included in any
// response body or log line. The raw transcript is likewise never logged or persisted here —
// only the caught error object is logged below, never the request body.
function buildClient(): Pick<Anthropic, "messages"> | null {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return null;
  return new Anthropic({ apiKey });
}

function statusCodeFor(result: AnalyzeTranscriptResult): number {
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

export async function POST(request: Request) {
  try {
    const body = (await request.json().catch(() => null)) as { rawTranscript?: unknown } | null;
    const rawTranscript = typeof body?.rawTranscript === "string" ? body.rawTranscript : "";
    const result = await analyzeTranscriptWithClaude(rawTranscript, buildClient());
    return Response.json(result, { status: statusCodeFor(result) });
  } catch (error) {
    console.error("Voice Intelligence analysis failed", error);
    return Response.json(
      { status: "server_error", message: "The transcript could not be analyzed." } satisfies AnalyzeTranscriptResult,
      { status: 500 },
    );
  }
}
