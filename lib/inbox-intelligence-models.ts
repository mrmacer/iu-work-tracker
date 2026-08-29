import { z } from "zod";
import type { Organization, ProviderMetadata, ReferenceData } from "./models";

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

export const INBOX_INTELLIGENCE_SCHEMA_VERSION = 1 as const;

/** Deliberately three states — see docs/INBOX_INTELLIGENCE_SHAREPOINT_REPORT.md "Status model". */
export type InboxIntelligenceStatus = "open" | "waiting" | "resolved";

/**
 * A first-class, durable Inbox Intelligence record — the reviewed/confirmed intelligence
 * derived from an email, never the email itself. Deliberately does NOT retain the raw
 * pasted email, the full thread, signatures, or any Anthropic request/response payload —
 * only a short, user-visible excerpt (`sourceExcerpt`) and the structured `analysis` the
 * user reviewed. See docs/INBOX_INTELLIGENCE_SHAREPOINT_REPORT.md "Privacy behavior".
 *
 * `matched*Ids` are resolved canonical reference IDs (exact-name match only, same rule as
 * lib/inbox-intelligence-work-record.ts) kept alongside the AI's raw extracted names in
 * `analysis` — the raw names remain the reviewable/editable source of truth; the matched
 * IDs are a best-effort connection to existing IU Work Tracker reference data.
 *
 * `metadata` reuses the exact same `ProviderMetadata` shape as `WorkRecord` (providerId,
 * version, createdAt, modifiedAt, syncState) — no parallel concurrency/timestamp model.
 */
export type InboxIntelligenceRecord = {
  appId: string;
  schemaVersion: typeof INBOX_INTELLIGENCE_SCHEMA_VERSION;
  sourceType: "pasted-email";
  analyzedAt: string;
  sourceExcerpt: string;
  analysis: EmailAnalysis;
  matchedOrganizationIds: string[];
  matchedDistrictIds: string[];
  matchedProjectIds: string[];
  status: InboxIntelligenceStatus;
  resolvedAt: string | null;
  linkedWorkRecordAppId: string | null;
  metadata: ProviderMetadata;
};

// Deliberately simple, non-fuzzy matching: an extracted name is connected to a canonical
// reference only on an exact, case-insensitive match. No scoring, no partial matches — see
// docs/INBOX_INTELLIGENCE_V1_REPORT.md "AI behavior" for why this stays small. Shared by
// lib/inbox-intelligence-work-record.ts (Work Record prefill) and the durable SharePoint
// provider (matched*Ids) so the two never compute this differently.
function matchExactName(name: string, candidates: { appId: string; name: string }[]): string | null {
  const normalized = name.trim().toLowerCase();
  return candidates.find((item) => item.name.trim().toLowerCase() === normalized)?.appId ?? null;
}

function dedupe(ids: (string | null)[]): string[] {
  return [...new Set(ids.filter((id): id is string => id !== null))];
}

export type ResolvedEntityMatches = {
  organizationIds: string[];
  districtIds: string[];
  projectIds: string[];
};

/** Exact-name-only matches against canonical reference data. Never fuzzy, never invents an entity. */
export function resolveEmailAnalysisEntities(analysis: EmailAnalysis, references: ReferenceData): ResolvedEntityMatches {
  const isDistrict = (organization: Organization) => organization.type === "district";
  return {
    districtIds: dedupe(analysis.districts.map((name) => matchExactName(name, references.organizations.filter(isDistrict)))),
    organizationIds: dedupe(
      analysis.organizations.map((name) => matchExactName(name, references.organizations.filter((org) => !isDistrict(org)))),
    ),
    projectIds: dedupe(analysis.projects.map((name) => matchExactName(name, references.projects))),
  };
}
