import { MICROSOFT_GRAPH_BASE_URL } from "./microsoft-auth-config";
import { WORK_RECORD_SCHEMA_VERSION, type EngagementScope, type WorkRecord } from "./models";
import type { ValidationIssue } from "./validation";

// Field mapping, timestamp, ETag, and pagination behavior follow
// docs/SHAREPOINT_INTEGRATION_PLAN.md sections 7, 9-13 and
// docs/SHAREPOINT_PROVISIONING_SPEC.md section 2.1.

export type SharePointWorkRecordConfig = {
  siteId: string;
  workRecordsListId: string;
};

export type ResolvedWorkRecordItem = {
  itemId: string;
  etag: string;
  record: WorkRecord;
};

export class SharePointWorkRecordsError extends Error {
  constructor(
    public readonly kind: "network" | "auth" | "not_found" | "conflict" | "malformed" | "unexpected",
    message: string,
    public readonly current?: WorkRecord,
  ) {
    super(message);
    this.name = "SharePointWorkRecordsError";
  }
}

type FetchLike = typeof fetch;

type GraphListItem = {
  id: string;
  eTag?: string;
  "@odata.etag"?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  fields?: Record<string, unknown>;
};

// --- SharePoint-compatible text limits (docs/SHAREPOINT_PROVISIONING_SPEC.md 2.1 "App maximum") ---
// Resolves the "SharePoint text limits" pre-provisioning gate from
// docs/SHAREPOINT_INTEGRATION_PLAN.md section "Remaining architectural concerns": silent
// truncation is prohibited, so oversized values are rejected as a validation error instead.
const TEXT_LIMITS: { path: string; max: number; value: (record: WorkRecord) => string }[] = [
  { path: "title", max: 255, value: (r) => r.title },
  { path: "appId", max: 255, value: (r) => r.appId },
  { path: "activityType", max: 255, value: (r) => r.activityType },
  { path: "description", max: 1000, value: (r) => r.description },
  { path: "detailedNotes", max: 10000, value: (r) => r.detailedNotes },
  { path: "evidenceSummary", max: 5000, value: (r) => r.evidenceSummary },
  { path: "output", max: 5000, value: (r) => r.output },
  { path: "outcome", max: 5000, value: (r) => r.outcome },
  { path: "nextStep", max: 1000, value: (r) => r.nextStep },
  { path: "orbit.evidence", max: 5000, value: (r) => r.orbit.evidence },
  { path: "orbit.primaryDeliverable", max: 255, value: (r) => r.orbit.primaryDeliverable ?? "" },
];
const JSON_ARRAY_LIMITS: { path: string; max: number; value: (record: WorkRecord) => string[] }[] = [
  { path: "projectIds", max: 10000, value: (r) => r.projectIds },
  { path: "organizationIds", max: 10000, value: (r) => r.organizationIds },
  { path: "contactIds", max: 10000, value: (r) => r.contactIds },
  { path: "categoryIds", max: 10000, value: (r) => r.categoryIds },
  { path: "evidenceReferenceIds", max: 10000, value: (r) => r.evidenceReferenceIds },
  { path: "orbit.supportingDeliverables", max: 10000, value: (r) => r.orbit.supportingDeliverables },
];

export function validateSharePointTextLimits(record: WorkRecord): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  for (const limit of TEXT_LIMITS) {
    if (limit.value(record).length > limit.max) {
      issues.push({
        path: limit.path,
        code: "sharepoint_text_limit",
        message: `${limit.path} exceeds the SharePoint-compatible limit of ${limit.max} characters.`,
      });
    }
  }
  for (const limit of JSON_ARRAY_LIMITS) {
    if (JSON.stringify(limit.value(record)).length > limit.max) {
      issues.push({
        path: limit.path,
        code: "sharepoint_text_limit",
        message: `${limit.path} exceeds the SharePoint-compatible JSON column limit of ${limit.max} characters.`,
      });
    }
  }
  return issues;
}

function toGraphDateOnly(date: string): string {
  return `${date}T00:00:00.000Z`;
}

function fromGraphDateOnly(value: unknown): string | null {
  return typeof value === "string" && value.length >= 10 ? value.slice(0, 10) : null;
}

function parseJsonStringArray(value: unknown, path: string): string[] {
  if (value === undefined || value === null || value === "") return [];
  let parsed: unknown;
  try {
    parsed = JSON.parse(String(value));
  } catch {
    throw new SharePointWorkRecordsError("malformed", `${path} contains malformed JSON.`);
  }
  if (!Array.isArray(parsed) || !parsed.every((item) => typeof item === "string")) {
    throw new SharePointWorkRecordsError("malformed", `${path} must be a JSON array of strings.`);
  }
  return parsed;
}

/** Maps a runtime WorkRecord to IU_Work_Records Graph fields. Never touches Created/Modified/ETag. */
export function toSharePointFields(record: WorkRecord, version: number): Record<string, unknown> {
  return {
    Title: record.title,
    AppId: record.appId,
    ActivityDate: toGraphDateOnly(record.activityDate),
    ActivityType: record.activityType,
    ShortDescription: record.description,
    DetailedNotes: record.detailedNotes,
    DurationMinutes: record.durationMinutes,
    RecordStatus: record.status,
    EngagementScope: record.engagementScope,
    ProjectIdsJson: JSON.stringify(record.projectIds),
    OrganizationIdsJson: JSON.stringify(record.organizationIds),
    ContactIdsJson: JSON.stringify(record.contactIds),
    CategoryIdsJson: JSON.stringify(record.categoryIds),
    EducatorsLeadersReach: record.reach.educatorsLeaders,
    StudentsFamiliesReach: record.reach.studentsFamilies,
    WorkforceCommunityReach: record.reach.workforceCommunity,
    OtherReach: record.reach.other,
    EvidenceSummary: record.evidenceSummary,
    EvidenceReferenceIdsJson: JSON.stringify(record.evidenceReferenceIds),
    WorkOutput: record.output,
    WorkOutcome: record.outcome,
    NextStep: record.nextStep,
    FollowUpNeeded: record.followUpNeeded,
    FollowUpDate: record.followUpDate ? toGraphDateOnly(record.followUpDate) : null,
    OrbitReportable: record.orbit.reportable,
    OrbitPrimaryDeliverableCode: record.orbit.primaryDeliverable,
    OrbitSupportingCodesJson: JSON.stringify(record.orbit.supportingDeliverables),
    StemPocMinutes: record.orbit.stemPocMinutes,
    TacMinutes: record.orbit.tacMinutes,
    OrbitEvidence: record.orbit.evidence,
    SchemaVersion: WORK_RECORD_SCHEMA_VERSION,
    RecordVersion: version,
    IsSample: record.isSample,
  };
}

/** Maps one Graph listItem (with $expand=fields) back to a runtime WorkRecord. Strict: throws rather than coerces. */
export function fromSharePointItem(item: GraphListItem): WorkRecord {
  const fields = item.fields;
  if (!fields) throw new SharePointWorkRecordsError("malformed", `Microsoft Graph did not return fields for item ${item.id}.`);

  const schemaVersion = Number(fields.SchemaVersion);
  if (schemaVersion !== WORK_RECORD_SCHEMA_VERSION) {
    throw new SharePointWorkRecordsError("malformed", `Item ${item.id} has unsupported SchemaVersion ${String(fields.SchemaVersion)}.`);
  }
  const version = Number(fields.RecordVersion);
  if (!Number.isInteger(version) || version < 1) {
    throw new SharePointWorkRecordsError("malformed", `Item ${item.id} has an invalid RecordVersion.`);
  }
  if (!item.createdDateTime || !item.lastModifiedDateTime) {
    throw new SharePointWorkRecordsError("malformed", `Item ${item.id} is missing SharePoint Created/Modified timestamps.`);
  }
  const activityDate = fromGraphDateOnly(fields.ActivityDate);
  if (!activityDate) throw new SharePointWorkRecordsError("malformed", `Item ${item.id} has an invalid ActivityDate.`);
  const appId = typeof fields.AppId === "string" ? fields.AppId : "";
  if (!appId) throw new SharePointWorkRecordsError("malformed", `Item ${item.id} is missing AppId.`);

  const scope = String(fields.EngagementScope);
  const engagementScope: EngagementScope =
    scope === "specific" || scope === "regional" || scope === "allDistricts" ? scope : "none";

  return {
    appId,
    title: String(fields.Title ?? ""),
    activityDate,
    activityType: String(fields.ActivityType ?? ""),
    description: String(fields.ShortDescription ?? ""),
    detailedNotes: String(fields.DetailedNotes ?? ""),
    durationMinutes: Number(fields.DurationMinutes ?? 0),
    status: fields.RecordStatus === "draft" ? "draft" : "complete",
    engagementScope,
    projectIds: parseJsonStringArray(fields.ProjectIdsJson, "ProjectIdsJson"),
    organizationIds: parseJsonStringArray(fields.OrganizationIdsJson, "OrganizationIdsJson"),
    contactIds: parseJsonStringArray(fields.ContactIdsJson, "ContactIdsJson"),
    categoryIds: parseJsonStringArray(fields.CategoryIdsJson, "CategoryIdsJson"),
    reach: {
      educatorsLeaders: Number(fields.EducatorsLeadersReach ?? 0),
      studentsFamilies: Number(fields.StudentsFamiliesReach ?? 0),
      workforceCommunity: Number(fields.WorkforceCommunityReach ?? 0),
      other: Number(fields.OtherReach ?? 0),
    },
    evidenceSummary: String(fields.EvidenceSummary ?? ""),
    evidenceReferenceIds: parseJsonStringArray(fields.EvidenceReferenceIdsJson, "EvidenceReferenceIdsJson"),
    output: String(fields.WorkOutput ?? ""),
    outcome: String(fields.WorkOutcome ?? ""),
    nextStep: String(fields.NextStep ?? ""),
    followUpNeeded: Boolean(fields.FollowUpNeeded),
    followUpDate: fromGraphDateOnly(fields.FollowUpDate),
    orbit: {
      reportable: Boolean(fields.OrbitReportable),
      primaryDeliverable: fields.OrbitPrimaryDeliverableCode ? String(fields.OrbitPrimaryDeliverableCode) : null,
      supportingDeliverables: parseJsonStringArray(fields.OrbitSupportingCodesJson, "OrbitSupportingCodesJson"),
      stemPocMinutes: Number(fields.StemPocMinutes ?? 0),
      tacMinutes: Number(fields.TacMinutes ?? 0),
      evidence: String(fields.OrbitEvidence ?? ""),
    },
    schemaVersion: WORK_RECORD_SCHEMA_VERSION,
    metadata: {
      providerId: item.id,
      version,
      createdAt: item.createdDateTime,
      modifiedAt: item.lastModifiedDateTime,
      syncState: "saved",
    },
    isSample: Boolean(fields.IsSample),
  };
}

function itemsPath(config: SharePointWorkRecordConfig): string {
  return `/sites/${encodeURIComponent(config.siteId)}/lists/${encodeURIComponent(config.workRecordsListId)}/items`;
}

function graphErrorFor(response: Response, operation: string): SharePointWorkRecordsError {
  if (response.status === 401) return new SharePointWorkRecordsError("auth", "Microsoft authentication has expired or is invalid.");
  if (response.status === 403) return new SharePointWorkRecordsError("auth", "The signed-in account is not authorized for the DEV SharePoint Work Records list.");
  if (response.status === 404) return new SharePointWorkRecordsError("not_found", `${operation} could not find the requested item.`);
  return new SharePointWorkRecordsError("unexpected", `${operation} failed with HTTP ${response.status}.`);
}

async function graphRequest(url: string, token: string, init: RequestInit, fetcher: FetchLike): Promise<Response> {
  try {
    return await fetcher(url, {
      ...init,
      headers: { Accept: "application/json", Authorization: `Bearer ${token}`, ...init.headers },
    });
  } catch {
    throw new SharePointWorkRecordsError("network", "Microsoft Graph could not be reached.");
  }
}

/** Follows every @odata.nextLink; never assumes one page contains all items (plan section 13). */
export async function listWorkRecords(
  config: SharePointWorkRecordConfig,
  token: string,
  fetcher: FetchLike = fetch,
): Promise<WorkRecord[]> {
  const records: WorkRecord[] = [];
  let url: string | undefined = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}?$expand=fields&$top=200`;
  while (url) {
    const response = await graphRequest(url, token, {}, fetcher);
    if (!response.ok) throw graphErrorFor(response, "Listing Work Records");
    const payload = (await response.json()) as { value?: GraphListItem[]; "@odata.nextLink"?: string };
    for (const item of payload.value ?? []) records.push(fromSharePointItem(item));
    url = payload["@odata.nextLink"];
  }
  return records;
}

export async function getWorkRecordItem(
  config: SharePointWorkRecordConfig,
  token: string,
  itemId: string,
  fetcher: FetchLike = fetch,
): Promise<ResolvedWorkRecordItem> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}/${encodeURIComponent(itemId)}?$expand=fields`;
  const response = await graphRequest(url, token, {}, fetcher);
  if (response.status === 404) throw new SharePointWorkRecordsError("not_found", `Work Record item ${itemId} was not found.`);
  if (!response.ok) throw graphErrorFor(response, "Reading the Work Record");
  const item = (await response.json()) as GraphListItem;
  return { itemId: item.id, etag: String(item.eTag ?? item["@odata.etag"] ?? ""), record: fromSharePointItem(item) };
}

/** Indexed AppId lookup: used for the create-time uniqueness check and as the update fallback (plan section 11). */
export async function findWorkRecordByAppId(
  config: SharePointWorkRecordConfig,
  token: string,
  appId: string,
  fetcher: FetchLike = fetch,
): Promise<ResolvedWorkRecordItem | null> {
  const escaped = appId.replace(/'/g, "''");
  const filter = encodeURIComponent(`fields/AppId eq '${escaped}'`);
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}?$expand=fields&$filter=${filter}`;
  const response = await graphRequest(url, token, {}, fetcher);
  if (!response.ok) throw graphErrorFor(response, "Looking up the Work Record by AppId");
  const payload = (await response.json()) as { value?: GraphListItem[] };
  const item = payload.value?.[0];
  if (!item) return null;
  return { itemId: item.id, etag: String(item.eTag ?? item["@odata.etag"] ?? ""), record: fromSharePointItem(item) };
}

/** Resolves the update target by metadata.providerId first, verifying AppId; falls back to indexed AppId lookup. */
export async function resolveWorkRecordItem(
  config: SharePointWorkRecordConfig,
  token: string,
  record: WorkRecord,
  fetcher: FetchLike = fetch,
): Promise<ResolvedWorkRecordItem | null> {
  if (record.metadata.providerId) {
    try {
      const resolved = await getWorkRecordItem(config, token, record.metadata.providerId, fetcher);
      if (resolved.record.appId === record.appId) return resolved;
    } catch (error) {
      if (!(error instanceof SharePointWorkRecordsError && error.kind === "not_found")) throw error;
    }
  }
  return findWorkRecordByAppId(config, token, record.appId, fetcher);
}

/** Create, then read the item back so the result carries SharePoint's id/timestamps (plan section 11, Create step 4). */
export async function createWorkRecordItem(
  config: SharePointWorkRecordConfig,
  token: string,
  record: WorkRecord,
  fetcher: FetchLike = fetch,
): Promise<WorkRecord> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}`;
  const response = await graphRequest(
    url,
    token,
    { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fields: toSharePointFields(record, 1) }) },
    fetcher,
  );
  if (!response.ok) throw graphErrorFor(response, "Creating the Work Record");
  const created = (await response.json()) as GraphListItem;
  const resolved = await getWorkRecordItem(config, token, created.id, fetcher);
  return resolved.record;
}

/** Conditional PATCH with If-Match; a 412 is re-read and surfaced as a conflict (plan section 11, Update steps 5-6). */
export async function updateWorkRecordItem(
  config: SharePointWorkRecordConfig,
  token: string,
  itemId: string,
  etag: string,
  record: WorkRecord,
  newVersion: number,
  fetcher: FetchLike = fetch,
): Promise<WorkRecord> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}/${encodeURIComponent(itemId)}/fields`;
  const response = await graphRequest(
    url,
    token,
    {
      method: "PATCH",
      headers: { "Content-Type": "application/json", "If-Match": etag },
      body: JSON.stringify(toSharePointFields(record, newVersion)),
    },
    fetcher,
  );
  if (response.status === 412) {
    const resolved = await getWorkRecordItem(config, token, itemId, fetcher);
    throw new SharePointWorkRecordsError("conflict", "The Work Record changed in SharePoint after it was loaded.", resolved.record);
  }
  if (!response.ok) throw graphErrorFor(response, "Updating the Work Record");
  const resolved = await getWorkRecordItem(config, token, itemId, fetcher);
  return resolved.record;
}
