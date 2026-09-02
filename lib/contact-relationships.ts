import type { Project, WorkRecord } from "./models";

// Pure view-derivation over already-loaded, already-durable Work Records and Projects for the
// Contact Detail relationship summary (Patch 8C). No I/O, no AI, no fabricated dates — every
// value here comes from a field already stored on WorkRecord. Mirrors
// lib/inbox-action-center.ts's role for the Home Action Center. See docs/AI_HANDOFF.md
// "Contact connected work (Patch 8C)".
//
// The relationship stays exactly: Contact.appId → WorkRecord.contactIds → WorkRecord.projectIds
// → Project. Nothing here is persisted, and nothing here infers a connection from a name,
// email, organization, or free-text mention — only the explicit, durable `contactIds` array.

const RECENT_WORK_LIMIT = 5;

export type ContactRelationshipSummary = {
  /** Every Work Record whose contactIds includes this Contact's appId, newest activityDate first. */
  connectedWorkRecords: WorkRecord[];
  /** The most recent RECENT_WORK_LIMIT of the above — for the Contact Detail "Recent Work" list. */
  recentWorkRecords: WorkRecord[];
  /** Projects referenced by any connected Work Record, deduped by appId, resolved against the
   * current (seeded + durable) project reference set. An unresolved historical project ID is
   * silently skipped — never fabricated, never crashes the page. */
  connectedProjects: Project[];
  /** Sum of durationMinutes across every connected Work Record — the same authoritative
   * duration field Projects' own derived totals already use. */
  totalMinutes: number;
  workRecordCount: number;
  /** The connected Work Record's own business/activity date (WorkRecord.activityDate) —
   * never SharePoint Created/Modified, never a fabricated "now". Null when there is no
   * connected Work Record yet. */
  lastInteractionDate: string | null;
};

/** Newest activityDate first — the same business-date sort History already uses
 * (`records.sort((a, b) => b.activityDate.localeCompare(a.activityDate))`). */
export function selectConnectedWorkRecords(workRecords: WorkRecord[], contactAppId: string): WorkRecord[] {
  return workRecords
    .filter((record) => record.contactIds.includes(contactAppId))
    .sort((a, b) => b.activityDate.localeCompare(a.activityDate));
}

/** Deduplicated by appId; an unresolvable historical project ID is skipped, never fabricated. */
export function selectConnectedProjects(connectedWorkRecords: WorkRecord[], projects: Project[]): Project[] {
  const projectIds = new Set<string>();
  for (const record of connectedWorkRecords) {
    for (const id of record.projectIds) projectIds.add(id);
  }
  const byAppId = new Map(projects.map((project) => [project.appId, project]));
  const resolved: Project[] = [];
  for (const id of projectIds) {
    const project = byAppId.get(id);
    if (project) resolved.push(project);
  }
  return resolved;
}

export function buildContactRelationshipSummary(
  workRecords: WorkRecord[],
  projects: Project[],
  contactAppId: string,
): ContactRelationshipSummary {
  const connectedWorkRecords = selectConnectedWorkRecords(workRecords, contactAppId);
  const connectedProjects = selectConnectedProjects(connectedWorkRecords, projects);
  const totalMinutes = connectedWorkRecords.reduce((sum, record) => sum + record.durationMinutes, 0);
  return {
    connectedWorkRecords,
    recentWorkRecords: connectedWorkRecords.slice(0, RECENT_WORK_LIMIT),
    connectedProjects,
    totalMinutes,
    workRecordCount: connectedWorkRecords.length,
    lastInteractionDate: connectedWorkRecords[0]?.activityDate ?? null,
  };
}
