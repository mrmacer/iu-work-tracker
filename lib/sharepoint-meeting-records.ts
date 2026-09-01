import { MICROSOFT_GRAPH_BASE_URL } from "./microsoft-auth-config";
import {
  MEETING_RECORD_SCHEMA_VERSION,
  ReviewedMeetingCandidatesSchema,
  type MeetingRecord,
  type ReviewedMeetingCandidate,
} from "./meeting-intelligence-models";
import type { ValidationIssue } from "./validation";

// Field mapping, timestamp, ETag, and pagination behavior follow the exact same conventions
// as lib/sharepoint-inbox-intelligence.ts and lib/sharepoint-work-records.ts, applied to the
// PROPOSED (not yet provisioned) IU_Meeting_Records list — see the Patch 6B SharePoint schema
// approval gate in docs/AI_HANDOFF.md. This file is intentionally independent of those two
// files rather than sharing low-level Graph plumbing, matching this codebase's existing
// one-file-per-resource pattern. It compiles and is fully unit-testable against mocks, but is
// unreachable in production until NEXT_PUBLIC_SHAREPOINT_IU_MEETING_RECORDS_LIST_ID is
// configured — see selectMeetingRecordProvider() in lib/meeting-record-provider.ts.

export type SharePointMeetingRecordConfig = {
  siteId: string;
  meetingRecordsListId: string;
};

export type ResolvedMeetingRecordItem = {
  itemId: string;
  etag: string;
  record: MeetingRecord;
};

export class SharePointMeetingRecordsError extends Error {
  constructor(
    public readonly kind: "network" | "auth" | "not_found" | "conflict" | "malformed" | "unexpected",
    message: string,
    public readonly current?: MeetingRecord,
  ) {
    super(message);
    this.name = "SharePointMeetingRecordsError";
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

// Generous, explicit SharePoint-compatible text limits — agenda/notes/minutes/intelligence
// JSON for a real meeting can be long; these are deliberately not the small default a
// "Single line of text" column would impose. Oversized values are rejected as a validation
// error, never truncated, matching every other resource's discipline in this codebase.
const TEXT_LIMITS: { path: string; max: number; value: (record: MeetingRecord) => string }[] = [
  { path: "appId", max: 255, value: (r) => r.appId },
  { path: "title", max: 255, value: (r) => r.title },
  { path: "meetingType", max: 255, value: (r) => r.meetingType },
  { path: "attendeesText", max: 4000, value: (r) => r.attendeesText },
  { path: "agendaText", max: 20000, value: (r) => r.agendaText },
  { path: "notesText", max: 20000, value: (r) => r.notesText },
  { path: "minutesText", max: 20000, value: (r) => r.minutesText },
  { path: "analysisModel", max: 100, value: (r) => r.analysisModel ?? "" },
];
const JSON_LIMITS: { path: string; max: number; value: (record: MeetingRecord) => unknown }[] = [
  { path: "reviewedCandidates", max: 40000, value: (r) => r.reviewedCandidates },
];

export function validateMeetingRecordSharePointLimits(record: MeetingRecord): ValidationIssue[] {
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
  for (const limit of JSON_LIMITS) {
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

function toGraphDateTime(value: string): string {
  return value; // already a full ISO 8601 timestamp; stored as-is, matching AnalyzedAt's role elsewhere
}

/** Maps a durable MeetingRecord to the proposed IU_Meeting_Records Graph fields. */
export function toSharePointFields(record: MeetingRecord, version: number): Record<string, unknown> {
  return {
    Title: record.title || "Untitled meeting",
    AppId: record.appId,
    SchemaVersion: MEETING_RECORD_SCHEMA_VERSION,
    MeetingDate: toGraphDateOnly(record.meetingDate),
    MeetingType: record.meetingType,
    AttendeesText: record.attendeesText,
    AgendaText: record.agendaText,
    NotesText: record.notesText,
    IntelligenceJson: JSON.stringify(record.reviewedCandidates),
    MinutesText: record.minutesText,
    AnalysisModel: record.analysisModel,
    AnalyzedAt: record.analyzedAt ? toGraphDateTime(record.analyzedAt) : null,
    RecordVersion: version,
  };
}

function parseReviewedCandidatesJson(value: unknown): ReviewedMeetingCandidate[] {
  if (value === undefined || value === null || value === "") return [];
  let parsed: unknown;
  try {
    parsed = JSON.parse(String(value));
  } catch {
    throw new SharePointMeetingRecordsError("malformed", "IntelligenceJson contains malformed JSON.");
  }
  const result = ReviewedMeetingCandidatesSchema.safeParse(parsed);
  if (!result.success) {
    throw new SharePointMeetingRecordsError("malformed", "IntelligenceJson does not match the expected reviewed-candidate shape.");
  }
  return result.data;
}

/** Maps one Graph listItem (with $expand=fields) back to a durable MeetingRecord. Strict: throws rather than coerces. */
export function fromSharePointItem(item: GraphListItem): MeetingRecord {
  const fields = item.fields;
  if (!fields) throw new SharePointMeetingRecordsError("malformed", `Microsoft Graph did not return fields for item ${item.id}.`);

  const schemaVersion = Number(fields.SchemaVersion);
  if (schemaVersion !== MEETING_RECORD_SCHEMA_VERSION) {
    throw new SharePointMeetingRecordsError("malformed", `Item ${item.id} has unsupported SchemaVersion ${String(fields.SchemaVersion)}.`);
  }
  const version = Number(fields.RecordVersion);
  if (!Number.isInteger(version) || version < 1) {
    throw new SharePointMeetingRecordsError("malformed", `Item ${item.id} has an invalid RecordVersion.`);
  }
  if (!item.createdDateTime || !item.lastModifiedDateTime) {
    throw new SharePointMeetingRecordsError("malformed", `Item ${item.id} is missing SharePoint Created/Modified timestamps.`);
  }
  const appId = typeof fields.AppId === "string" ? fields.AppId : "";
  if (!appId) throw new SharePointMeetingRecordsError("malformed", `Item ${item.id} is missing AppId.`);
  const meetingDate = fromGraphDateOnly(fields.MeetingDate);
  if (!meetingDate) throw new SharePointMeetingRecordsError("malformed", `Item ${item.id} has an invalid MeetingDate.`);

  const analyzedAt = typeof fields.AnalyzedAt === "string" ? fields.AnalyzedAt : null;
  const analysisModel = typeof fields.AnalysisModel === "string" && fields.AnalysisModel ? fields.AnalysisModel : null;

  return {
    appId,
    schemaVersion: MEETING_RECORD_SCHEMA_VERSION,
    title: String(fields.Title ?? ""),
    meetingDate,
    meetingType: String(fields.MeetingType ?? ""),
    attendeesText: String(fields.AttendeesText ?? ""),
    agendaText: String(fields.AgendaText ?? ""),
    notesText: String(fields.NotesText ?? ""),
    reviewedCandidates: parseReviewedCandidatesJson(fields.IntelligenceJson),
    minutesText: String(fields.MinutesText ?? ""),
    analysisModel,
    analyzedAt,
    metadata: {
      providerId: item.id,
      version,
      createdAt: item.createdDateTime,
      modifiedAt: item.lastModifiedDateTime,
      syncState: "saved",
    },
  };
}

function itemsPath(config: SharePointMeetingRecordConfig): string {
  return `/sites/${encodeURIComponent(config.siteId)}/lists/${encodeURIComponent(config.meetingRecordsListId)}/items`;
}

function graphErrorFor(response: Response, operation: string): SharePointMeetingRecordsError {
  if (response.status === 401) return new SharePointMeetingRecordsError("auth", "Microsoft authentication has expired or is invalid.");
  if (response.status === 403) return new SharePointMeetingRecordsError("auth", "The signed-in account is not authorized for the DEV Meeting Records list.");
  if (response.status === 404) return new SharePointMeetingRecordsError("not_found", `${operation} could not find the requested item.`);
  return new SharePointMeetingRecordsError("unexpected", `${operation} failed with HTTP ${response.status}.`);
}

async function graphRequest(url: string, token: string, init: RequestInit, fetcher: FetchLike): Promise<Response> {
  try {
    return await fetcher(url, {
      ...init,
      headers: { Accept: "application/json", Authorization: `Bearer ${token}`, ...init.headers },
    });
  } catch {
    throw new SharePointMeetingRecordsError("network", "Microsoft Graph could not be reached.");
  }
}

/** Follows every @odata.nextLink; never assumes one page contains all items. */
export async function listMeetingRecordItems(
  config: SharePointMeetingRecordConfig,
  token: string,
  fetcher: FetchLike = fetch,
): Promise<MeetingRecord[]> {
  const records: MeetingRecord[] = [];
  let url: string | undefined = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}?$expand=fields&$top=200`;
  while (url) {
    const response = await graphRequest(url, token, {}, fetcher);
    if (!response.ok) throw graphErrorFor(response, "Listing Meeting Records");
    const payload = (await response.json()) as { value?: GraphListItem[]; "@odata.nextLink"?: string };
    for (const item of payload.value ?? []) records.push(fromSharePointItem(item));
    url = payload["@odata.nextLink"];
  }
  return records;
}

export async function getMeetingRecordItem(
  config: SharePointMeetingRecordConfig,
  token: string,
  itemId: string,
  fetcher: FetchLike = fetch,
): Promise<ResolvedMeetingRecordItem> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}/${encodeURIComponent(itemId)}?$expand=fields`;
  const response = await graphRequest(url, token, {}, fetcher);
  if (response.status === 404) throw new SharePointMeetingRecordsError("not_found", `Meeting Record item ${itemId} was not found.`);
  if (!response.ok) throw graphErrorFor(response, "Reading the Meeting Record");
  const item = (await response.json()) as GraphListItem;
  return { itemId: item.id, etag: String(item.eTag ?? item["@odata.etag"] ?? ""), record: fromSharePointItem(item) };
}

/** Indexed AppId lookup: create-time uniqueness check and the update fallback (mirrors Work Records / Inbox Intelligence). */
export async function findMeetingRecordByAppId(
  config: SharePointMeetingRecordConfig,
  token: string,
  appId: string,
  fetcher: FetchLike = fetch,
): Promise<ResolvedMeetingRecordItem | null> {
  const escaped = appId.replace(/'/g, "''");
  const filter = encodeURIComponent(`fields/AppId eq '${escaped}'`);
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}?$expand=fields&$filter=${filter}`;
  const response = await graphRequest(url, token, {}, fetcher);
  if (!response.ok) throw graphErrorFor(response, "Looking up the Meeting Record by AppId");
  const payload = (await response.json()) as { value?: GraphListItem[] };
  const item = payload.value?.[0];
  if (!item) return null;
  return { itemId: item.id, etag: String(item.eTag ?? item["@odata.etag"] ?? ""), record: fromSharePointItem(item) };
}

/** Resolves the update target by metadata.providerId first, verifying AppId; falls back to indexed AppId lookup. */
export async function resolveMeetingRecordItem(
  config: SharePointMeetingRecordConfig,
  token: string,
  record: MeetingRecord,
  fetcher: FetchLike = fetch,
): Promise<ResolvedMeetingRecordItem | null> {
  if (record.metadata.providerId) {
    try {
      const resolved = await getMeetingRecordItem(config, token, record.metadata.providerId, fetcher);
      if (resolved.record.appId === record.appId) return resolved;
    } catch (error) {
      if (!(error instanceof SharePointMeetingRecordsError && error.kind === "not_found")) throw error;
    }
  }
  return findMeetingRecordByAppId(config, token, record.appId, fetcher);
}

/** Create, then read the item back so the result carries SharePoint's id/timestamps. */
export async function createMeetingRecordItem(
  config: SharePointMeetingRecordConfig,
  token: string,
  record: MeetingRecord,
  fetcher: FetchLike = fetch,
): Promise<MeetingRecord> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}`;
  const response = await graphRequest(
    url,
    token,
    { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fields: toSharePointFields(record, 1) }) },
    fetcher,
  );
  if (!response.ok) throw graphErrorFor(response, "Creating the Meeting Record");
  const created = (await response.json()) as GraphListItem;
  const resolved = await getMeetingRecordItem(config, token, created.id, fetcher);
  return resolved.record;
}

/** Conditional PATCH with If-Match; a 412 is re-read and surfaced as a conflict. */
export async function updateMeetingRecordItem(
  config: SharePointMeetingRecordConfig,
  token: string,
  itemId: string,
  etag: string,
  record: MeetingRecord,
  newVersion: number,
  fetcher: FetchLike = fetch,
): Promise<MeetingRecord> {
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
    const resolved = await getMeetingRecordItem(config, token, itemId, fetcher);
    throw new SharePointMeetingRecordsError("conflict", "The Meeting Record changed in SharePoint after it was loaded.", resolved.record);
  }
  if (!response.ok) throw graphErrorFor(response, "Updating the Meeting Record");
  const resolved = await getMeetingRecordItem(config, token, itemId, fetcher);
  return resolved.record;
}
