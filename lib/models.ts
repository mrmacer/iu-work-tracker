export const WORK_RECORD_SCHEMA_VERSION = 2 as const;

export type SyncState = "saved" | "saving" | "sync-error";
export type EngagementScope = "none" | "specific" | "regional" | "allDistricts";

export type ProviderMetadata = {
  providerId?: string;
  version: number;
  createdAt: string;
  modifiedAt: string;
  syncState?: SyncState;
};

export type WorkRecord = {
  appId: string;
  title: string;
  activityDate: string;
  activityType: string;
  description: string;
  detailedNotes: string;
  durationMinutes: number;
  status: "complete" | "draft";
  engagementScope: EngagementScope;
  projectIds: string[];
  organizationIds: string[];
  contactIds: string[];
  categoryIds: string[];
  reach: { educatorsLeaders: number; studentsFamilies: number; workforceCommunity: number; other: number };
  evidenceSummary: string;
  evidenceReferenceIds: string[];
  output: string;
  outcome: string;
  nextStep: string;
  followUpNeeded: boolean;
  followUpDate: string | null;
  orbit: {
    reportable: boolean;
    primaryDeliverable: string | null;
    supportingDeliverables: string[];
    stemPocMinutes: number;
    tacMinutes: number;
    evidence: string;
  };
  schemaVersion: typeof WORK_RECORD_SCHEMA_VERSION;
  metadata: ProviderMetadata;
  isSample: boolean;
};

export type Project = { appId: string; name: string; description: string; status: "active" | "planning" | "complete"; color: string };
export type Organization = { appId: string; name: string; type: "district" | "partner" | "iu" };
export type Contact = { appId: string; displayName: string; role: string; organizationId: string | null };
export type Category = { appId: string; name: string; group: "work-area" | "topic" };
export type Deliverable = { code: string; label: string };
export type ReportingConfig = {
  minutesPerReportingDay: number;
  schoolYearStartMonth: number;
  quarters: readonly { code: "Q1" | "Q2" | "Q3" | "Q4"; startMonth: number; endMonth: number }[];
};
export type SystemSettings = { activityTypes: readonly string[] };
export type ReferenceData = {
  projects: Project[];
  organizations: Organization[];
  contacts: Contact[];
  categories: Category[];
  deliverables: Deliverable[];
  reportingConfig: ReportingConfig;
  settings: SystemSettings;
};
