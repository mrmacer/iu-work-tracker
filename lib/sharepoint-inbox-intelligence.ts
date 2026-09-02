import { MICROSOFT_GRAPH_BASE_URL } from "./microsoft-auth-config";
import {
  EmailAnalysisSchema,
  INBOX_INTELLIGENCE_SCHEMA_VERSION,
  type InboxIntelligenceRecord,
  type InboxIntelligenceStatus,
} from "./inbox-intelligence-models";
import type { ValidationIssue } from "./validation";

// Field mapping, timestamp, ETag, and pagination behavior follow the same conventions as
// lib/sharepoint-work-records.ts (see docs/SHAREPOINT_INTEGRATION_PLAN.md and
// docs/INBOX_INTELLIGENCE_SHAREPOINT_REPORT.md "SharePoint schema") applied to the new
// IU_Inbox_Intelligence list. This file is intentionally independent of
// sharepoint-work-records.ts rather than sharing low-level Graph plumbing — matching this
// codebase's existing one-file-per-resource pattern (lib/microsoft-graph.ts and
// lib/sharepoint-work-records.ts are likewise independent) rather than introducing a new
// shared abstraction as part of this phase.

export type SharePointInboxIntelligenceConfig = {
  siteId: string;
  inboxIntelligenceListId: string;
};

export type ResolvedInboxIntelligenceItem = {
  itemId: string;
  etag: string;
  record: InboxIntelligenceRecord;
};

export class SharePointInboxIntelligenceError extends Error {
  constructor(
    public readonly kind: "network" | "auth" | "not_found" | "conflict" | "malformed" | "unexpected",
    message: string,
    public readonly current?: InboxIntelligenceRecord,
  ) {
    super(message);
    this.name = "SharePointInboxIntelligenceError";
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

const STATUSES: InboxIntelligenceStatus[] = ["open", "waiting", "resolved"];

// SharePoint-compatible text limits (docs/INBOX_INTELLIGENCE_SHAREPOINT_REPORT.md "SharePoint
// schema"). Oversized values are rejected as a validation error, never truncated — the same
// rule already applied to Work Records.
const TEXT_LIMITS: { path: string; max: number; value: (record: InboxIntelligenceRecord) => string }[] = [
  { path: "appId", max: 255, value: (r) => r.appId },
  { path: "sourceExcerpt", max: 500, value: (r) => r.sourceExcerpt },
  { path: "analysis.suggestedWorkRecord.title", max: 255, value: (r) => r.analysis.suggestedWorkRecord.title },
  { path: "analysis.suggestedWorkType", max: 255, value: (r) => r.analysis.suggestedWorkType ?? "" },
  { path: "linkedWorkRecordAppId", max: 255, value: (r) => r.linkedWorkRecordAppId ?? "" },
];
const JSON_ARRAY_LIMITS: { path: string; max: number; value: (record: InboxIntelligenceRecord) => string[] }[] = [
  { path: "analysis.actionItems", max: 10000, value: (r) => r.analysis.actionItems.map((item) => item.action) },
  { path: "analysis.people", max: 10000, value: (r) => r.analysis.people },
  { path: "analysis.organizations", max: 10000, value: (r) => r.analysis.organizations },
  { path: "analysis.districts", max: 10000, value: (r) => r.analysis.districts },
  { path: "analysis.projects", max: 10000, value: (r) => r.analysis.projects },
  { path: "analysis.tags", max: 10000, value: (r) => r.analysis.tags },
  { path: "matchedOrganizationIds", max: 10000, value: (r) => r.matchedOrganizationIds },
  { path: "matchedDistrictIds", max: 10000, value: (r) => r.matchedDistrictIds },
  { path: "matchedProjectIds", max: 10000, value: (r) => r.matchedProjectIds },
  { path: "matchedContactIds", max: 10000, value: (r) => r.matchedContactIds },
];

export function validateInboxIntelligenceSharePointLimits(record: InboxIntelligenceRecord): ValidationIssue[] {
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

/**
 * Full runtime-shape validation for a record about to be written. Reuses the exact same
 * EmailAnalysisSchema the AI pipeline validates against (docs/INBOX_INTELLIGENCE_V1_REPORT.md
 * "Extraction schema") — a human edit in the review screen must satisfy the same shape as an
 * AI extraction did, so this never accepts something the AI path itself would have rejected.
 */
export function validateInboxIntelligenceRecord(record: InboxIntelligenceRecord): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  const analysisResult = EmailAnalysisSchema.safeParse(record.analysis);
  if (!analysisResult.success) {
    issues.push({ path: "analysis", code: "invalid_analysis", message: "The reviewed intelligence no longer matches the expected shape." });
  }
  if (record.schemaVersion !== INBOX_INTELLIGENCE_SCHEMA_VERSION) {
    issues.push({ path: "schemaVersion", code: "unsupported_schema", message: `schemaVersion must be ${INBOX_INTELLIGENCE_SCHEMA_VERSION}.` });
  }
  if (!STATUSES.includes(record.status)) {
    issues.push({ path: "status", code: "invalid_status", message: "status must be open, waiting, or resolved." });
  }
  if (record.status === "resolved" && !record.resolvedAt) {
    issues.push({ path: "resolvedAt", code: "required", message: "resolvedAt is required once status is resolved." });
  }
  if (record.status !== "resolved" && record.resolvedAt) {
    issues.push({ path: "resolvedAt", code: "unexpected_value", message: "resolvedAt must be null unless status is resolved." });
  }
  return [...issues, ...validateInboxIntelligenceSharePointLimits(record)];
}

function toGraphDateTime(value: string): string {
  return value; // already a full ISO 8601 timestamp; stored as-is, unlike ActivityDate's date-only truncation
}

/** Maps a durable InboxIntelligenceRecord to IU_Inbox_Intelligence Graph fields. */
export function toSharePointFields(record: InboxIntelligenceRecord, version: number): Record<string, unknown> {
  return {
    Title: record.analysis.suggestedWorkRecord.title,
    AppId: record.appId,
    SchemaVersion: INBOX_INTELLIGENCE_SCHEMA_VERSION,
    SourceType: record.sourceType,
    AnalyzedAt: toGraphDateTime(record.analyzedAt),
    SourceExcerpt: record.sourceExcerpt,
    SummaryText: record.analysis.summary,
    Priority: record.analysis.priority,
    NeedsAttention: record.analysis.needsAttention,
    ActionItemsJson: JSON.stringify(record.analysis.actionItems),
    FollowUpText: record.analysis.followUp,
    PeopleJson: JSON.stringify(record.analysis.people),
    OrganizationsJson: JSON.stringify(record.analysis.organizations),
    DistrictsJson: JSON.stringify(record.analysis.districts),
    ProjectsJson: JSON.stringify(record.analysis.projects),
    TagsJson: JSON.stringify(record.analysis.tags),
    MatchedOrganizationIdsJson: JSON.stringify(record.matchedOrganizationIds),
    MatchedDistrictIdsJson: JSON.stringify(record.matchedDistrictIds),
    MatchedProjectIdsJson: JSON.stringify(record.matchedProjectIds),
    // Patch 8D. REQUIRES a MatchedContactIdsJson column (Multiple lines of text, plain) on the
    // live IU_Inbox_Intelligence list — added the same way Patch 8B's IU_Contacts columns were
    // added, before this code writes to it. See docs/AI_HANDOFF.md "Intelligence Contact
    // matching (Patch 8D)".
    MatchedContactIdsJson: JSON.stringify(record.matchedContactIds),
    SuggestedWorkType: record.analysis.suggestedWorkType,
    SuggestedWorkRecordDescription: record.analysis.suggestedWorkRecord.description,
    LinkedWorkRecordAppId: record.linkedWorkRecordAppId,
    Status: record.status,
    ResolvedAt: record.resolvedAt ? toGraphDateTime(record.resolvedAt) : null,
    RecordVersion: version,
  };
}

function parseJsonStringArray(value: unknown, path: string): string[] {
  if (value === undefined || value === null || value === "") return [];
  let parsed: unknown;
  try {
    parsed = JSON.parse(String(value));
  } catch {
    throw new SharePointInboxIntelligenceError("malformed", `${path} contains malformed JSON.`);
  }
  if (!Array.isArray(parsed) || !parsed.every((item) => typeof item === "string")) {
    throw new SharePointInboxIntelligenceError("malformed", `${path} must be a JSON array of strings.`);
  }
  return parsed;
}

function parseActionItemsJson(value: unknown): InboxIntelligenceRecord["analysis"]["actionItems"] {
  if (value === undefined || value === null || value === "") return [];
  let parsed: unknown;
  try {
    parsed = JSON.parse(String(value));
  } catch {
    throw new SharePointInboxIntelligenceError("malformed", "ActionItemsJson contains malformed JSON.");
  }
  if (!Array.isArray(parsed)) throw new SharePointInboxIntelligenceError("malformed", "ActionItemsJson must be a JSON array.");
  return parsed.map((entry, index) => {
    if (
      !entry ||
      typeof entry !== "object" ||
      typeof (entry as Record<string, unknown>).action !== "string" ||
      typeof (entry as Record<string, unknown>).owner !== "string"
    ) {
      throw new SharePointInboxIntelligenceError("malformed", `ActionItemsJson[${index}] is not a valid action item.`);
    }
    const item = entry as Record<string, unknown>;
    return {
      action: String(item.action),
      dueDate: typeof item.dueDate === "string" ? item.dueDate : null,
      owner: (["me", "sender", "other", "unknown"].includes(String(item.owner)) ? item.owner : "unknown") as
        | "me"
        | "sender"
        | "other"
        | "unknown",
    };
  });
}

/** Maps one Graph listItem (with $expand=fields) back to a durable InboxIntelligenceRecord. Strict: throws rather than coerces. */
export function fromSharePointItem(item: GraphListItem): InboxIntelligenceRecord {
  const fields = item.fields;
  if (!fields) throw new SharePointInboxIntelligenceError("malformed", `Microsoft Graph did not return fields for item ${item.id}.`);

  const schemaVersion = Number(fields.SchemaVersion);
  if (schemaVersion !== INBOX_INTELLIGENCE_SCHEMA_VERSION) {
    throw new SharePointInboxIntelligenceError("malformed", `Item ${item.id} has unsupported SchemaVersion ${String(fields.SchemaVersion)}.`);
  }
  const version = Number(fields.RecordVersion);
  if (!Number.isInteger(version) || version < 1) {
    throw new SharePointInboxIntelligenceError("malformed", `Item ${item.id} has an invalid RecordVersion.`);
  }
  if (!item.createdDateTime || !item.lastModifiedDateTime) {
    throw new SharePointInboxIntelligenceError("malformed", `Item ${item.id} is missing SharePoint Created/Modified timestamps.`);
  }
  const appId = typeof fields.AppId === "string" ? fields.AppId : "";
  if (!appId) throw new SharePointInboxIntelligenceError("malformed", `Item ${item.id} is missing AppId.`);
  const status = String(fields.Status);
  if (!STATUSES.includes(status as InboxIntelligenceStatus)) {
    throw new SharePointInboxIntelligenceError("malformed", `Item ${item.id} has an invalid Status.`);
  }
  const priority = String(fields.Priority);
  if (priority !== "high" && priority !== "medium" && priority !== "low") {
    throw new SharePointInboxIntelligenceError("malformed", `Item ${item.id} has an invalid Priority.`);
  }
  const analyzedAt = typeof fields.AnalyzedAt === "string" ? fields.AnalyzedAt : null;
  if (!analyzedAt) throw new SharePointInboxIntelligenceError("malformed", `Item ${item.id} is missing AnalyzedAt.`);

  return {
    appId,
    schemaVersion: INBOX_INTELLIGENCE_SCHEMA_VERSION,
    sourceType: "pasted-email",
    analyzedAt,
    sourceExcerpt: String(fields.SourceExcerpt ?? ""),
    analysis: {
      summary: String(fields.SummaryText ?? ""),
      priority,
      needsAttention: Boolean(fields.NeedsAttention),
      actionItems: parseActionItemsJson(fields.ActionItemsJson),
      followUp: String(fields.FollowUpText ?? ""),
      people: parseJsonStringArray(fields.PeopleJson, "PeopleJson"),
      organizations: parseJsonStringArray(fields.OrganizationsJson, "OrganizationsJson"),
      districts: parseJsonStringArray(fields.DistrictsJson, "DistrictsJson"),
      projects: parseJsonStringArray(fields.ProjectsJson, "ProjectsJson"),
      tags: parseJsonStringArray(fields.TagsJson, "TagsJson"),
      suggestedWorkType: fields.SuggestedWorkType ? String(fields.SuggestedWorkType) : null,
      suggestedWorkRecord: {
        title: String(fields.Title ?? ""),
        description: String(fields.SuggestedWorkRecordDescription ?? ""),
      },
    },
    matchedOrganizationIds: parseJsonStringArray(fields.MatchedOrganizationIdsJson, "MatchedOrganizationIdsJson"),
    matchedDistrictIds: parseJsonStringArray(fields.MatchedDistrictIdsJson, "MatchedDistrictIdsJson"),
    matchedProjectIds: parseJsonStringArray(fields.MatchedProjectIdsJson, "MatchedProjectIdsJson"),
    // Missing field (a list item written before the column existed) reads back as [] via
    // parseJsonStringArray's undefined-handling — never a parse error on old items.
    matchedContactIds: parseJsonStringArray(fields.MatchedContactIdsJson, "MatchedContactIdsJson"),
    status: status as InboxIntelligenceStatus,
    resolvedAt: typeof fields.ResolvedAt === "string" ? fields.ResolvedAt : null,
    linkedWorkRecordAppId: fields.LinkedWorkRecordAppId ? String(fields.LinkedWorkRecordAppId) : null,
    metadata: {
      providerId: item.id,
      version,
      createdAt: item.createdDateTime,
      modifiedAt: item.lastModifiedDateTime,
      syncState: "saved",
    },
  };
}

function itemsPath(config: SharePointInboxIntelligenceConfig): string {
  return `/sites/${encodeURIComponent(config.siteId)}/lists/${encodeURIComponent(config.inboxIntelligenceListId)}/items`;
}

function graphErrorFor(response: Response, operation: string): SharePointInboxIntelligenceError {
  if (response.status === 401) return new SharePointInboxIntelligenceError("auth", "Microsoft authentication has expired or is invalid.");
  if (response.status === 403) return new SharePointInboxIntelligenceError("auth", "The signed-in account is not authorized for the DEV Inbox Intelligence list.");
  if (response.status === 404) return new SharePointInboxIntelligenceError("not_found", `${operation} could not find the requested item.`);
  return new SharePointInboxIntelligenceError("unexpected", `${operation} failed with HTTP ${response.status}.`);
}

async function graphRequest(url: string, token: string, init: RequestInit, fetcher: FetchLike): Promise<Response> {
  try {
    return await fetcher(url, {
      ...init,
      headers: { Accept: "application/json", Authorization: `Bearer ${token}`, ...init.headers },
    });
  } catch {
    throw new SharePointInboxIntelligenceError("network", "Microsoft Graph could not be reached.");
  }
}

/** Follows every @odata.nextLink; never assumes one page contains all items. */
export async function listInboxIntelligenceItems(
  config: SharePointInboxIntelligenceConfig,
  token: string,
  fetcher: FetchLike = fetch,
): Promise<InboxIntelligenceRecord[]> {
  const records: InboxIntelligenceRecord[] = [];
  let url: string | undefined = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}?$expand=fields&$top=200`;
  while (url) {
    const response = await graphRequest(url, token, {}, fetcher);
    if (!response.ok) throw graphErrorFor(response, "Listing Inbox Intelligence records");
    const payload = (await response.json()) as { value?: GraphListItem[]; "@odata.nextLink"?: string };
    for (const item of payload.value ?? []) records.push(fromSharePointItem(item));
    url = payload["@odata.nextLink"];
  }
  return records;
}

export async function getInboxIntelligenceItem(
  config: SharePointInboxIntelligenceConfig,
  token: string,
  itemId: string,
  fetcher: FetchLike = fetch,
): Promise<ResolvedInboxIntelligenceItem> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}/${encodeURIComponent(itemId)}?$expand=fields`;
  const response = await graphRequest(url, token, {}, fetcher);
  if (response.status === 404) throw new SharePointInboxIntelligenceError("not_found", `Inbox Intelligence item ${itemId} was not found.`);
  if (!response.ok) throw graphErrorFor(response, "Reading the Inbox Intelligence record");
  const item = (await response.json()) as GraphListItem;
  return { itemId: item.id, etag: String(item.eTag ?? item["@odata.etag"] ?? ""), record: fromSharePointItem(item) };
}

/** Indexed AppId lookup: create-time uniqueness check and the update fallback (mirrors Work Records). */
export async function findInboxIntelligenceByAppId(
  config: SharePointInboxIntelligenceConfig,
  token: string,
  appId: string,
  fetcher: FetchLike = fetch,
): Promise<ResolvedInboxIntelligenceItem | null> {
  const escaped = appId.replace(/'/g, "''");
  const filter = encodeURIComponent(`fields/AppId eq '${escaped}'`);
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}?$expand=fields&$filter=${filter}`;
  const response = await graphRequest(url, token, {}, fetcher);
  if (!response.ok) throw graphErrorFor(response, "Looking up the Inbox Intelligence record by AppId");
  const payload = (await response.json()) as { value?: GraphListItem[] };
  const item = payload.value?.[0];
  if (!item) return null;
  return { itemId: item.id, etag: String(item.eTag ?? item["@odata.etag"] ?? ""), record: fromSharePointItem(item) };
}

/** Resolves the update target by metadata.providerId first, verifying AppId; falls back to indexed AppId lookup. */
export async function resolveInboxIntelligenceItem(
  config: SharePointInboxIntelligenceConfig,
  token: string,
  record: InboxIntelligenceRecord,
  fetcher: FetchLike = fetch,
): Promise<ResolvedInboxIntelligenceItem | null> {
  if (record.metadata.providerId) {
    try {
      const resolved = await getInboxIntelligenceItem(config, token, record.metadata.providerId, fetcher);
      if (resolved.record.appId === record.appId) return resolved;
    } catch (error) {
      if (!(error instanceof SharePointInboxIntelligenceError && error.kind === "not_found")) throw error;
    }
  }
  return findInboxIntelligenceByAppId(config, token, record.appId, fetcher);
}

/** Create, then read the item back so the result carries SharePoint's id/timestamps. */
export async function createInboxIntelligenceItem(
  config: SharePointInboxIntelligenceConfig,
  token: string,
  record: InboxIntelligenceRecord,
  fetcher: FetchLike = fetch,
): Promise<InboxIntelligenceRecord> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}`;
  const response = await graphRequest(
    url,
    token,
    { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fields: toSharePointFields(record, 1) }) },
    fetcher,
  );
  if (!response.ok) throw graphErrorFor(response, "Creating the Inbox Intelligence record");
  const created = (await response.json()) as GraphListItem;
  const resolved = await getInboxIntelligenceItem(config, token, created.id, fetcher);
  return resolved.record;
}

/** Conditional PATCH with If-Match; a 412 is re-read and surfaced as a conflict. */
export async function updateInboxIntelligenceItem(
  config: SharePointInboxIntelligenceConfig,
  token: string,
  itemId: string,
  etag: string,
  record: InboxIntelligenceRecord,
  newVersion: number,
  fetcher: FetchLike = fetch,
): Promise<InboxIntelligenceRecord> {
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
    const resolved = await getInboxIntelligenceItem(config, token, itemId, fetcher);
    throw new SharePointInboxIntelligenceError("conflict", "The Inbox Intelligence record changed in SharePoint after it was loaded.", resolved.record);
  }
  if (!response.ok) throw graphErrorFor(response, "Updating the Inbox Intelligence record");
  const resolved = await getInboxIntelligenceItem(config, token, itemId, fetcher);
  return resolved.record;
}
