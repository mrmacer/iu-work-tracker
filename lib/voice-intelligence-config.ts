// Centralized Voice Intelligence / Anthropic API configuration. Deliberately independent of
// lib/anthropic-config.ts (Inbox Intelligence's own config) — same one-file-per-resource
// precedent already used for the SharePoint providers, not a shared/coupled constant.

/** claude-opus-5 is the current default per house policy; change only this constant to swap models. */
export const VOICE_ANTHROPIC_MODEL = "claude-opus-5";

/** Reasoning depth for the analysis call. Segmentation is a bounded, non-agentic task. */
export const VOICE_ANTHROPIC_EFFORT = "medium" as const;

/** A transcript can segment into many more candidates than an email has action items. */
export const VOICE_ANTHROPIC_MAX_OUTPUT_TOKENS = 8192;

/**
 * Conservative, explicit pasted-transcript length limit (cost control, and V1 deliberately
 * does not optimize for arbitrarily long hour-plus transcripts). Enforced both client- and
 * server-side; never silently truncated — an over-limit transcript is rejected with a message
 * telling the user to split it, exactly like Inbox Intelligence's MAX_EMAIL_LENGTH.
 */
export const MAX_TRANSCRIPT_LENGTH = 40000;
