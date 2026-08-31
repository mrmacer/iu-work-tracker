// Conservative, deterministic removal of unambiguous non-semantic email boilerplate before
// the pasted text is sent to the model. Only strips whole lines that match one of a small set
// of known-safe patterns — the exact categories docs/SHAREPOINT_INTEGRATION_PLAN.md-adjacent
// guidance calls "clearly boilerplate and safe to remove": external-sender warning banners,
// image placeholder lines, and standalone unsubscribe/footer-link lines.
//
// Deliberately does NOT touch signatures, quoted threads, forwarded headers, or calendar
// blocks — those can legitimately contain sender identity or real dates, so removing them
// risks deleting real content. That distinction is handled by system-prompt instructions
// instead (lib/anthropic-email-analysis.ts) — see docs/AI_HANDOFF.md "Email Noise Torture
// Test" for the reasoning split between preprocessing and prompt hardening.

const NOISE_LINE_PATTERNS: RegExp[] = [
  // Microsoft 365 / common external-sender warning banner (two common line shapes).
  /^caution:?\s*this email originated from outside(?: of)? the organization\.?$/i,
  /^do not click links? or open attachments? unless you recognize the sender.*$/i,
  // Standalone image placeholder lines emitted by various email clients.
  /^\[(?:image[^\]]*|cid:[^\]]*|logo|external image)\]$/i,
  // Standalone marketing-footer action lines.
  /^(?:unsubscribe|manage preferences|view (?:this )?(?:email )?in (?:a )?browser)$/i,
];

/**
 * Strips only whole lines that exactly match a known-safe boilerplate pattern. Blank lines
 * and everything else — including signatures, quoted/forwarded content, and calendar fields —
 * pass through untouched.
 */
export function stripDeterministicEmailNoise(rawEmail: string): string {
  const kept = rawEmail.split("\n").filter((line) => {
    const trimmed = line.trim();
    if (!trimmed) return true;
    return !NOISE_LINE_PATTERNS.some((pattern) => pattern.test(trimmed));
  });
  return kept.join("\n").replace(/\n{3,}/g, "\n\n");
}
