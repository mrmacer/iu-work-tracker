import { WORK_RECORD_SCHEMA_VERSION, type EngagementScope, type WorkRecord } from "./models";

const capturedAt = "2026-08-26T12:00:00.000Z";

function sample(
  appId: string,
  input: Partial<WorkRecord> & Pick<WorkRecord, "title" | "activityDate" | "activityType" | "durationMinutes">,
): WorkRecord {
  const base: WorkRecord = {
    appId,
    title: input.title,
    activityDate: input.activityDate,
    activityType: input.activityType,
    description: "Clearly marked development test record.",
    detailedNotes: "",
    durationMinutes: input.durationMinutes,
    status: "complete",
    engagementScope: "none",
    projectIds: [],
    organizationIds: [],
    contactIds: [],
    categoryIds: [],
    reach: { educatorsLeaders: 0, studentsFamilies: 0, workforceCommunity: 0, other: 0 },
    evidenceSummary: "",
    evidenceReferenceIds: [],
    output: "",
    outcome: "",
    nextStep: "",
    followUpNeeded: false,
    followUpDate: null,
    orbit: { reportable: false, primaryDeliverable: null, supportingDeliverables: [], stemPocMinutes: 0, tacMinutes: 0, evidence: "" },
    schemaVersion: WORK_RECORD_SCHEMA_VERSION,
    metadata: { providerId: `development:${appId}`, version: 1, createdAt: capturedAt, modifiedAt: capturedAt, syncState: "saved" },
    isSample: true,
  };
  return { ...base, ...input, metadata: { ...base.metadata, ...input.metadata }, reach: { ...base.reach, ...input.reach }, orbit: { ...base.orbit, ...input.orbit } };
}

const scope = (engagementScope: EngagementScope) => engagementScope;

export const SAMPLE_RECORDS: WorkRecord[] = [
  sample("sample-steels", { title: "District STEELS planning meeting", activityDate: "2026-08-26", activityType: "District meeting", description: "Planned next steps for local STEELS implementation.", durationMinutes: 60, engagementScope: scope("specific"), projectIds: ["project-steels"], organizationIds: ["org-north-valley"], contactIds: ["contact-north-valley-lead"], categoryIds: ["cat-district", "cat-steels", "cat-planning"], reach: { educatorsLeaders: 5, studentsFamilies: 0, workforceCommunity: 0, other: 0 }, output: "Draft implementation roadmap", outcome: "District team identified two priority actions.", nextStep: "Send STEELS resources", followUpNeeded: true, followUpDate: "2026-08-28", evidenceSummary: "Planning notes and draft roadmap", evidenceReferenceIds: ["dev-evidence-steels-roadmap"], orbit: { reportable: true, primaryDeliverable: "B", supportingDeliverables: ["D"], stemPocMinutes: 60, tacMinutes: 0, evidence: "Planning notes and draft roadmap" } }),
  sample("sample-ai", { title: "AI professional learning session", activityDate: "2026-08-26", activityType: "Professional learning", description: "Facilitated a practical responsible-AI learning session.", durationMinutes: 120, engagementScope: scope("regional"), projectIds: ["project-ai"], categoryIds: ["cat-pl", "cat-ai"], reach: { educatorsLeaders: 24, studentsFamilies: 0, workforceCommunity: 0, other: 0 }, output: "Workshop and resource guide", outcome: "Participants drafted classroom use guidelines.", evidenceSummary: "Session materials and participant reflection", orbit: { reportable: true, primaryDeliverable: "C", supportingDeliverables: ["D"], stemPocMinutes: 120, tacMinutes: 0, evidence: "Session materials and participant reflection" } }),
  sample("sample-keystone", { title: "Keystone STEM Competition planning", activityDate: "2026-08-25", activityType: "Project planning", durationMinutes: 90, engagementScope: scope("none"), projectIds: ["project-keystone"], organizationIds: ["org-iu"], contactIds: ["contact-iu-colleague"], categoryIds: ["cat-stem", "cat-students", "cat-planning"], reach: { educatorsLeaders: 3, studentsFamilies: 0, workforceCommunity: 0, other: 0 }, orbit: { reportable: true, primaryDeliverable: "F", supportingDeliverables: [], stemPocMinutes: 90, tacMinutes: 0, evidence: "Planning checklist" } }),
  sample("sample-ecosystem", { title: "STEM Ecosystem partner meeting", activityDate: "2026-08-24", activityType: "Partner meeting", durationMinutes: 60, engagementScope: scope("regional"), projectIds: ["project-ecosystem"], organizationIds: ["org-futureworks"], contactIds: ["contact-futureworks"], categoryIds: ["cat-stem", "cat-partnerships"], reach: { educatorsLeaders: 2, studentsFamilies: 0, workforceCommunity: 7, other: 0 }, orbit: { reportable: true, primaryDeliverable: "E", supportingDeliverables: ["A"], stemPocMinutes: 30, tacMinutes: 30, evidence: "Partner meeting notes" } }),
  sample("sample-internal", { title: "Internal IU administrative planning meeting", activityDate: "2026-08-26", activityType: "Internal planning", durationMinutes: 75, engagementScope: scope("none"), organizationIds: ["org-iu"], contactIds: ["contact-iu-colleague"], categoryIds: ["cat-internal", "cat-planning"], reach: { educatorsLeaders: 0, studentsFamilies: 0, workforceCommunity: 0, other: 4 } }),
  sample("sample-makerspace", { title: "Makerspace student visit", activityDate: "2026-08-21", activityType: "Student program", durationMinutes: 180, engagementScope: scope("specific"), projectIds: ["project-makerspace"], organizationIds: ["org-riverbend"], categoryIds: ["cat-stem", "cat-students"], reach: { educatorsLeaders: 4, studentsFamilies: 48, workforceCommunity: 2, other: 0 }, orbit: { reportable: true, primaryDeliverable: "F", supportingDeliverables: ["G"], stemPocMinutes: 180, tacMinutes: 0, evidence: "Facilitator notes" } }),
  sample("sample-statewide", { title: "Statewide STEM PoC meeting", activityDate: "2026-08-20", activityType: "Partner meeting", durationMinutes: 90, engagementScope: scope("none"), categoryIds: ["cat-stem", "cat-partnerships"], orbit: { reportable: true, primaryDeliverable: "A", supportingDeliverables: [], stemPocMinutes: 90, tacMinutes: 0, evidence: "Statewide meeting notes" } }),
  sample("sample-workforce", { title: "Workforce outreach conversation", activityDate: "2026-08-19", activityType: "Follow-up communication", durationMinutes: 30, engagementScope: scope("regional"), projectIds: ["project-ecosystem"], organizationIds: ["org-futureworks"], contactIds: ["contact-futureworks"], categoryIds: ["cat-partnerships"], reach: { educatorsLeaders: 0, studentsFamilies: 0, workforceCommunity: 1, other: 0 }, orbit: { reportable: true, primaryDeliverable: "E", supportingDeliverables: [], stemPocMinutes: 0, tacMinutes: 30, evidence: "Conversation summary" } }),
  sample("sample-regional-student", { title: "Regional student experience", activityDate: "2026-08-18", activityType: "Student program", durationMinutes: 210, engagementScope: scope("allDistricts"), projectIds: ["project-makerspace"], categoryIds: ["cat-stem", "cat-students"], reach: { educatorsLeaders: 12, studentsFamilies: 80, workforceCommunity: 4, other: 0 }, orbit: { reportable: true, primaryDeliverable: "F", supportingDeliverables: [], stemPocMinutes: 210, tacMinutes: 0, evidence: "Facilitator log" } }),
  sample("sample-technology", { title: "Non-STEM IU technology development work", activityDate: "2026-08-17", activityType: "Resource development", durationMinutes: 150, engagementScope: scope("none"), organizationIds: ["org-iu"], categoryIds: ["cat-internal"] }),
];
