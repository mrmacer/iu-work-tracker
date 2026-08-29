import { z } from "zod";

// Runtime shape AI extraction must conform to. Field names match the schema agreed
// in docs/INBOX_INTELLIGENCE_V1_REPORT.md exactly — do not rename without updating
// both the prompt and that document.

const dateOnlyOrNull = z.string().nullable();

export const ActionItemSchema = z
  .object({
    action: z.string().trim().min(1).max(500),
    dueDate: dateOnlyOrNull,
    owner: z.enum(["me", "sender", "other", "unknown"]),
  })
  .strict();

export const EmailAnalysisSchema = z
  .object({
    summary: z.string().trim().min(1).max(2000),
    priority: z.enum(["high", "medium", "low"]),
    needsAttention: z.boolean(),
    actionItems: z.array(ActionItemSchema).max(20),
    followUp: z.string().trim().max(1000),
    people: z.array(z.string().trim().min(1).max(200)).max(50),
    organizations: z.array(z.string().trim().min(1).max(200)).max(50),
    districts: z.array(z.string().trim().min(1).max(200)).max(50),
    projects: z.array(z.string().trim().min(1).max(200)).max(50),
    tags: z.array(z.string().trim().min(1).max(60)).max(20),
    suggestedWorkType: z.string().trim().max(200).nullable(),
    suggestedWorkRecord: z
      .object({
        title: z.string().trim().min(1).max(255),
        description: z.string().trim().max(2000),
      })
      .strict(),
  })
  .strict();

export type ActionItem = z.infer<typeof ActionItemSchema>;
export type EmailAnalysis = z.infer<typeof EmailAnalysisSchema>;

const DATE_ONLY_PATTERN = /^\d{4}-\d{2}-\d{2}$/;

/**
 * A due date the model could not confidently ground in the email is treated as
 * "no deadline" rather than fabricated. This never invents a date; it only
 * discards one that isn't in the exact YYYY-MM-DD shape the app uses everywhere else.
 */
export function normalizeActionItemDueDate(dueDate: string | null): string | null {
  return dueDate && DATE_ONLY_PATTERN.test(dueDate) ? dueDate : null;
}

export function normalizeEmailAnalysis(analysis: EmailAnalysis): EmailAnalysis {
  return {
    ...analysis,
    actionItems: analysis.actionItems.map((item) => ({
      ...item,
      dueDate: normalizeActionItemDueDate(item.dueDate),
    })),
  };
}

/**
 * A first-class Inbox Intelligence record. Deliberately does NOT retain the raw
 * pasted email — only a short, user-visible excerpt survives past the review
 * screen (see docs/INBOX_INTELLIGENCE_V1_REPORT.md "Privacy behavior"). V1
 * persistence is session-only (see lib/inbox-intelligence-provider.ts); no
 * SharePoint list has been provisioned for this record type yet.
 */
export type InboxIntelligenceRecord = {
  appId: string;
  createdAt: string;
  sourceExcerpt: string;
  analysis: EmailAnalysis;
  linkedWorkRecordAppId: string | null;
};
