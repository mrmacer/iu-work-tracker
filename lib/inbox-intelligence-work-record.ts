import type { EngagementScope, Organization, ReferenceData, WorkRecord } from "./models";
import type { EmailAnalysis } from "./inbox-intelligence-models";

// Deliberately simple, non-fuzzy matching: an extracted name is connected to a canonical
// reference only on an exact, case-insensitive match. No scoring, no partial matches — see
// docs/INBOX_INTELLIGENCE_V1_REPORT.md "AI behavior" for why V1 stays this small.
function matchExactName(name: string, candidates: { appId: string; name: string }[]): string | null {
  const normalized = name.trim().toLowerCase();
  return candidates.find((item) => item.name.trim().toLowerCase() === normalized)?.appId ?? null;
}

function dedupe(ids: (string | null)[]): string[] {
  return [...new Set(ids.filter((id): id is string => id !== null))];
}

/**
 * Maps AI-suggested email intelligence onto a blank Work Record draft. Pure and
 * side-effect-free: it never writes through the DataProvider — the caller still opens the
 * existing Log Work wizard with this draft, and the user must review and explicitly save.
 *
 * Reference matching is exact-name-only (see matchExactName above). A matched canonical
 * district always sets engagementScope to "specific" so the returned draft never violates
 * validateWorkRecord's scope/organization invariants on its own.
 */
export function buildWorkRecordDraftFromAnalysis(
  analysis: EmailAnalysis,
  references: ReferenceData,
  baseRecord: WorkRecord,
): WorkRecord {
  const isDistrict = (organization: Organization) => organization.type === "district";
  const districtIds = dedupe(analysis.districts.map((name) => matchExactName(name, references.organizations.filter(isDistrict))));
  const otherOrgIds = dedupe(
    analysis.organizations.map((name) => matchExactName(name, references.organizations.filter((org) => !isDistrict(org)))),
  );
  const projectIds = dedupe(analysis.projects.map((name) => matchExactName(name, references.projects)));
  const engagementScope: EngagementScope = districtIds.length > 0 ? "specific" : baseRecord.engagementScope;

  const matchedActivityType = analysis.suggestedWorkType
    ? references.settings.activityTypes.find(
        (type) => type.trim().toLowerCase() === analysis.suggestedWorkType!.trim().toLowerCase(),
      )
    : undefined;

  const firstDueDate = analysis.actionItems.find((item) => item.dueDate)?.dueDate ?? null;

  return {
    ...baseRecord,
    title: analysis.suggestedWorkRecord.title,
    description: analysis.suggestedWorkRecord.description,
    activityType: matchedActivityType ?? baseRecord.activityType,
    engagementScope,
    organizationIds: [...districtIds, ...otherOrgIds],
    projectIds,
    nextStep: analysis.followUp || analysis.actionItems[0]?.action || baseRecord.nextStep,
    followUpNeeded: analysis.actionItems.length > 0,
    followUpDate: firstDueDate,
  };
}
