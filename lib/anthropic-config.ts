// Centralized Inbox Intelligence / Anthropic API configuration.
// Change the model, effort, or limits here — nothing else in the feature should
// need to change to pick up a different model or a different cost/quality tradeoff.

/** claude-opus-5 is the current default per house policy; change only this constant to swap models. */
export const ANTHROPIC_MODEL = "claude-opus-5";

/** Reasoning depth for the analysis call. Extraction/summarization is a bounded, non-agentic task. */
export const ANTHROPIC_EFFORT = "medium" as const;

/** The extracted JSON is small; this comfortably covers summary + action items without over-provisioning. */
export const ANTHROPIC_MAX_OUTPUT_TOKENS = 4096;

/** Reasonable pasted-email length limit (cost control). Enforced both client- and server-side. */
export const MAX_EMAIL_LENGTH = 20000;
