// Centralized Meeting Intelligence / Anthropic API configuration. Independent of
// lib/anthropic-config.ts (Inbox) and lib/voice-intelligence-config.ts (Voice) — same
// one-file-per-resource precedent already used across this codebase's providers.

/** claude-opus-5 is the current default per house policy; change only this constant to swap models. */
export const MEETING_ANTHROPIC_MODEL = "claude-opus-5";

/** Reasoning depth for the analysis call. Extraction/segmentation is a bounded, non-agentic task. */
export const MEETING_ANTHROPIC_EFFORT = "medium" as const;

/** A meeting can segment into many candidates across 8 types, more than a single email produces. */
export const MEETING_ANTHROPIC_MAX_OUTPUT_TOKENS = 8192;

/**
 * Conservative, explicit combined agenda+notes length limit (cost control). Enforced both
 * client- and server-side; never silently truncated — an over-limit meeting is rejected with
 * a message telling the user to trim it, mirroring MAX_TRANSCRIPT_LENGTH's discipline.
 */
export const MAX_MEETING_CONTENT_LENGTH = 40000;
