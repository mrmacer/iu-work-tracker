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

/**
 * Patch 7 — Projects (previously reference/configuration data only, see docs/DATA_MODEL.md
 * "Reference and configuration entities") gains an optional durable dimension. `startDate`,
 * `targetDate`, `stemOrbit`, and `metadata` are all optional so the five existing seeded
 * projects in lib/reference-data.ts remain valid Project values unchanged — they are simply
 * never durable (no `metadata`) and never dated/STEM-flagged. A Project with `metadata`
 * present was created/loaded through lib/project-provider.ts; one without it is a static
 * seeded project. `"paused"` is new in Patch 7 — additive to the existing status vocabulary.
 */
export type Project = {
  appId: string;
  name: string;
  description: string;
  status: "planning" | "active" | "paused" | "complete";
  color: string;
  startDate?: string | null;
  targetDate?: string | null;
  stemOrbit?: boolean;
  metadata?: ProviderMetadata;
};
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
