import type { EngagementScope, ReferenceData, WorkRecord } from "./models";
import { resolveEmailAnalysisEntities, type EmailAnalysis } from "./inbox-intelligence-models";

/**
 * Maps AI-suggested email intelligence onto a blank Work Record draft. Pure and
 * side-effect-free: it never writes through the DataProvider — the caller still opens the
 * existing Log Work wizard with this draft, and the user must review and explicitly save.
 *
 * Reference matching is exact-name-only (see resolveEmailAnalysisEntities in
 * lib/inbox-intelligence-models.ts, shared with the durable SharePoint provider's
 * matched*Ids). A matched canonical district always sets engagementScope to "specific" so
 * the returned draft never violates validateWorkRecord's scope/organization invariants on
 * its own.
 */
export function buildWorkRecordDraftFromAnalysis(
  analysis: EmailAnalysis,
  references: ReferenceData,
  baseRecord: WorkRecord,
): WorkRecord {
  const { organizationIds: otherOrgIds, districtIds, projectIds } = resolveEmailAnalysisEntities(analysis, references);
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
