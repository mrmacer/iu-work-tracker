import { MICROSOFT_GRAPH_BASE_URL } from "./microsoft-auth-config";
import type { Project } from "./models";
import { PROJECT_STATUSES, type ProjectStatus } from "./project-provider";
import type { ValidationIssue } from "./validation";

// Field mapping, timestamp, ETag, and pagination behavior follow the exact same conventions as
// lib/sharepoint-meeting-records.ts. Targets the real, live, approved IU_Projects list (Patch
// 7B) — see docs/AI_HANDOFF.md "Durable Projects (Patch 7)". IU_Projects predates this patch
// (original seven-list reference-data provisioning, docs/SHAREPOINT_PROVISIONING_CHECKLIST.md)
// and already had Title/AppId/ProjectDescription/Color; it was extended in place with
// StartDate, TargetDate, StemOrbit, and RecordVersion, and its existing ProjectStatus Choice
// column (kept as Choice, never converted to text) gained a fourth allowed value, "paused" —
// alongside the pre-existing "active"/"planning"/"complete". A Graph Choice field's value is
// written/read as a plain string identical to a text field's, so toSharePointFields()/
// fromSharePointItem() need no special handling for that — PROJECT_STATUSES' exact lowercase
// values (lib/project-provider.ts) already match the live Choice values one-to-one.

export type SharePointProjectConfig = {
  siteId: string;
  projectsListId: string;
};

export type ResolvedProjectItem = {
  itemId: string;
  etag: string;
  project: Project;
};

export class SharePointProjectsError extends Error {
  constructor(
    public readonly kind: "network" | "auth" | "not_found" | "conflict" | "malformed" | "unexpected",
    message: string,
    public readonly current?: Project,
  ) {
    super(message);
    this.name = "SharePointProjectsError";
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

const TEXT_LIMITS: { path: string; max: number; value: (project: Project) => string }[] = [
  { path: "appId", max: 255, value: (p) => p.appId },
  { path: "name", max: 255, value: (p) => p.name },
  { path: "description", max: 1000, value: (p) => p.description },
  { path: "status", max: 255, value: (p) => p.status },
  { path: "color", max: 255, value: (p) => p.color },
];

export function validateProjectSharePointLimits(project: Project): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  for (const limit of TEXT_LIMITS) {
    if (limit.value(project).length > limit.max) {
      issues.push({
        path: limit.path,
        code: "sharepoint_text_limit",
        message: `${limit.path} exceeds the SharePoint-compatible limit of ${limit.max} characters.`,
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

/** Maps a durable Project to the proposed IU_Projects Graph fields. */
export function toSharePointFields(project: Project, version: number): Record<string, unknown> {
  return {
    Title: project.name || "Untitled project",
    AppId: project.appId,
    ProjectDescription: project.description,
    ProjectStatus: project.status,
    Color: project.color,
    StartDate: project.startDate ? toGraphDateOnly(project.startDate) : null,
    TargetDate: project.targetDate ? toGraphDateOnly(project.targetDate) : null,
    StemOrbit: project.stemOrbit ?? false,
    RecordVersion: version,
  };
}

/** Maps one Graph listItem (with $expand=fields) back to a durable Project. Strict: throws rather than coerces. */
export function fromSharePointItem(item: GraphListItem): Project {
  const fields = item.fields;
  if (!fields) throw new SharePointProjectsError("malformed", `Microsoft Graph did not return fields for item ${item.id}.`);

  const version = Number(fields.RecordVersion);
  if (!Number.isInteger(version) || version < 1) {
    throw new SharePointProjectsError("malformed", `Item ${item.id} has an invalid RecordVersion.`);
  }
  if (!item.createdDateTime || !item.lastModifiedDateTime) {
    throw new SharePointProjectsError("malformed", `Item ${item.id} is missing SharePoint Created/Modified timestamps.`);
  }
  const appId = typeof fields.AppId === "string" ? fields.AppId : "";
  if (!appId) throw new SharePointProjectsError("malformed", `Item ${item.id} is missing AppId.`);
  const status = typeof fields.ProjectStatus === "string" ? fields.ProjectStatus : "";
  if (!PROJECT_STATUSES.includes(status as ProjectStatus)) {
    throw new SharePointProjectsError("malformed", `Item ${item.id} has an unrecognized ProjectStatus "${status}".`);
  }

  return {
    appId,
    name: String(fields.Title ?? ""),
    description: String(fields.ProjectDescription ?? ""),
    status: status as ProjectStatus,
    color: String(fields.Color ?? ""),
    startDate: fromGraphDateOnly(fields.StartDate),
    targetDate: fromGraphDateOnly(fields.TargetDate),
    stemOrbit: Boolean(fields.StemOrbit),
    metadata: {
      providerId: item.id,
      version,
      createdAt: item.createdDateTime,
      modifiedAt: item.lastModifiedDateTime,
      syncState: "saved",
    },
  };
}

function itemsPath(config: SharePointProjectConfig): string {
  return `/sites/${encodeURIComponent(config.siteId)}/lists/${encodeURIComponent(config.projectsListId)}/items`;
}

function graphErrorFor(response: Response, operation: string): SharePointProjectsError {
  if (response.status === 401) return new SharePointProjectsError("auth", "Microsoft authentication has expired or is invalid.");
  if (response.status === 403) return new SharePointProjectsError("auth", "The signed-in account is not authorized for the DEV Projects list.");
  if (response.status === 404) return new SharePointProjectsError("not_found", `${operation} could not find the requested item.`);
  return new SharePointProjectsError("unexpected", `${operation} failed with HTTP ${response.status}.`);
}

async function graphRequest(url: string, token: string, init: RequestInit, fetcher: FetchLike): Promise<Response> {
  try {
    return await fetcher(url, {
      ...init,
      headers: { Accept: "application/json", Authorization: `Bearer ${token}`, ...init.headers },
    });
  } catch {
    throw new SharePointProjectsError("network", "Microsoft Graph could not be reached.");
  }
}

/** Follows every @odata.nextLink; never assumes one page contains all items. */
export async function listProjectItems(
  config: SharePointProjectConfig,
  token: string,
  fetcher: FetchLike = fetch,
): Promise<Project[]> {
  const projects: Project[] = [];
  let url: string | undefined = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}?$expand=fields&$top=200`;
  while (url) {
    const response = await graphRequest(url, token, {}, fetcher);
    if (!response.ok) throw graphErrorFor(response, "Listing Projects");
    const payload = (await response.json()) as { value?: GraphListItem[]; "@odata.nextLink"?: string };
    for (const item of payload.value ?? []) projects.push(fromSharePointItem(item));
    url = payload["@odata.nextLink"];
  }
  return projects;
}

export async function getProjectItem(
  config: SharePointProjectConfig,
  token: string,
  itemId: string,
  fetcher: FetchLike = fetch,
): Promise<ResolvedProjectItem> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}/${encodeURIComponent(itemId)}?$expand=fields`;
  const response = await graphRequest(url, token, {}, fetcher);
  if (response.status === 404) throw new SharePointProjectsError("not_found", `Project item ${itemId} was not found.`);
  if (!response.ok) throw graphErrorFor(response, "Reading the Project");
  const item = (await response.json()) as GraphListItem;
  return { itemId: item.id, etag: String(item.eTag ?? item["@odata.etag"] ?? ""), project: fromSharePointItem(item) };
}

/** Indexed AppId lookup: create-time uniqueness check and the update fallback (mirrors every other resource). */
export async function findProjectByAppId(
  config: SharePointProjectConfig,
  token: string,
  appId: string,
  fetcher: FetchLike = fetch,
): Promise<ResolvedProjectItem | null> {
  const escaped = appId.replace(/'/g, "''");
  const filter = encodeURIComponent(`fields/AppId eq '${escaped}'`);
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}?$expand=fields&$filter=${filter}`;
  const response = await graphRequest(url, token, {}, fetcher);
  if (!response.ok) throw graphErrorFor(response, "Looking up the Project by AppId");
  const payload = (await response.json()) as { value?: GraphListItem[] };
  const item = payload.value?.[0];
  if (!item) return null;
  return { itemId: item.id, etag: String(item.eTag ?? item["@odata.etag"] ?? ""), project: fromSharePointItem(item) };
}

/** Resolves the update target by metadata.providerId first, verifying AppId; falls back to indexed AppId lookup. */
export async function resolveProjectItem(
  config: SharePointProjectConfig,
  token: string,
  project: Project,
  fetcher: FetchLike = fetch,
): Promise<ResolvedProjectItem | null> {
  const providerId = project.metadata?.providerId;
  if (providerId) {
    try {
      const resolved = await getProjectItem(config, token, providerId, fetcher);
      if (resolved.project.appId === project.appId) return resolved;
    } catch (error) {
      if (!(error instanceof SharePointProjectsError && error.kind === "not_found")) throw error;
    }
  }
  return findProjectByAppId(config, token, project.appId, fetcher);
}

/** Create, then read the item back so the result carries SharePoint's id/timestamps. */
export async function createProjectItem(
  config: SharePointProjectConfig,
  token: string,
  project: Project,
  fetcher: FetchLike = fetch,
): Promise<Project> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}`;
  const response = await graphRequest(
    url,
    token,
    { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fields: toSharePointFields(project, 1) }) },
    fetcher,
  );
  if (!response.ok) throw graphErrorFor(response, "Creating the Project");
  const created = (await response.json()) as GraphListItem;
  const resolved = await getProjectItem(config, token, created.id, fetcher);
  return resolved.project;
}

/** Conditional PATCH with If-Match; a 412 is re-read and surfaced as a conflict. */
export async function updateProjectItem(
  config: SharePointProjectConfig,
  token: string,
  itemId: string,
  etag: string,
  project: Project,
  newVersion: number,
  fetcher: FetchLike = fetch,
): Promise<Project> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}/${encodeURIComponent(itemId)}/fields`;
  const response = await graphRequest(
    url,
    token,
    {
      method: "PATCH",
      headers: { "Content-Type": "application/json", "If-Match": etag },
      body: JSON.stringify(toSharePointFields(project, newVersion)),
    },
    fetcher,
  );
  if (response.status === 412) {
    const resolved = await getProjectItem(config, token, itemId, fetcher);
    throw new SharePointProjectsError("conflict", "The Project changed in SharePoint after it was loaded.", resolved.project);
  }
  if (!response.ok) throw graphErrorFor(response, "Updating the Project");
  const resolved = await getProjectItem(config, token, itemId, fetcher);
  return resolved.project;
}
