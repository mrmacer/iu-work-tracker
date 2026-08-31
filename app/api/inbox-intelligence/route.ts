import Anthropic from "@anthropic-ai/sdk";
import { analyzeEmailWithClaude, type AnalyzeEmailResult } from "../../../lib/anthropic-email-analysis";

export const dynamic = "force-dynamic";

// The Anthropic API key is a server-only secret read from process.env (standard Next.js
// server-side env access). It is never read from a NEXT_PUBLIC_* variable, never sent to
// the browser, and never included in any response body or log line.
function buildClient(): Pick<Anthropic, "messages"> | null {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return null;
  return new Anthropic({ apiKey });
}

function statusCodeFor(result: AnalyzeEmailResult): number {
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
    const body = (await request.json().catch(() => null)) as { rawEmail?: unknown } | null;
    const rawEmail = typeof body?.rawEmail === "string" ? body.rawEmail : "";
    const result = await analyzeEmailWithClaude(rawEmail, buildClient());
    return Response.json(result, { status: statusCodeFor(result) });
  } catch (error) {
    console.error("Inbox Intelligence analysis failed", error);
    return Response.json(
      { status: "server_error", message: "The email could not be analyzed." } satisfies AnalyzeEmailResult,
      { status: 500 },
    );
  }
}
