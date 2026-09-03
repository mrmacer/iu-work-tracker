import { MICROSOFT_GRAPH_BASE_URL } from "./microsoft-auth-config";
import type { Organization } from "./models";
import { ORGANIZATION_TYPES, type OrganizationTypeValue } from "./organization-provider";
import type { ValidationIssue } from "./validation";

// Field mapping, timestamp, ETag, and pagination behavior follow the exact same conventions as
// lib/sharepoint-projects.ts. Targets the real, live IU_Organizations list (Patch 8E) — see
// docs/AI_HANDOFF.md "Durable Organizations (Patch 8E)". IU_Organizations predates this patch
// (original seven-list reference-data provisioning, docs/SHAREPOINT_PROVISIONING_CHECKLIST.md)
// and already had Title/AppId/OrganizationType; the user added a RecordVersion column (Number,
// required, default 1) before this file was written — live-verified, no other change made. A
// Graph Choice field's value is written/read as a plain string identical to a text field's, so
// toSharePointFields()/fromSharePointItem() need no special handling for that —
// ORGANIZATION_TYPES' exact lowercase values (lib/organization-provider.ts) already match the
// live Choice values ("district", "partner", "iu") one-to-one.

export type SharePointOrganizationConfig = {
  siteId: string;
  organizationsListId: string;
};

export type ResolvedOrganizationItem = {
  itemId: string;
  etag: string;
  organization: Organization;
};

export class SharePointOrganizationsError extends Error {
  constructor(
    public readonly kind: "network" | "auth" | "not_found" | "conflict" | "malformed" | "unexpected",
    message: string,
    public readonly current?: Organization,
  ) {
    super(message);
    this.name = "SharePointOrganizationsError";
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

const TEXT_LIMITS: { path: string; max: number; value: (organization: Organization) => string }[] = [
  { path: "appId", max: 255, value: (o) => o.appId },
  { path: "name", max: 255, value: (o) => o.name },
  { path: "type", max: 255, value: (o) => o.type },
];

export function validateOrganizationSharePointLimits(organization: Organization): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  for (const limit of TEXT_LIMITS) {
    if (limit.value(organization).length > limit.max) {
      issues.push({
        path: limit.path,
        code: "sharepoint_text_limit",
        message: `${limit.path} exceeds the SharePoint-compatible limit of ${limit.max} characters.`,
      });
    }
  }
  return issues;
}

/** Maps a durable Organization to the live IU_Organizations Graph fields. */
export function toSharePointFields(organization: Organization, version: number): Record<string, unknown> {
  return {
    Title: organization.name || "Untitled organization",
    AppId: organization.appId,
    OrganizationType: organization.type,
    RecordVersion: version,
  };
}

/** Maps one Graph listItem (with $expand=fields) back to a durable Organization. Strict: throws rather than coerces. */
export function fromSharePointItem(item: GraphListItem): Organization {
  const fields = item.fields;
  if (!fields) throw new SharePointOrganizationsError("malformed", `Microsoft Graph did not return fields for item ${item.id}.`);

  const version = Number(fields.RecordVersion);
  if (!Number.isInteger(version) || version < 1) {
    throw new SharePointOrganizationsError("malformed", `Item ${item.id} has an invalid RecordVersion.`);
  }
  if (!item.createdDateTime || !item.lastModifiedDateTime) {
    throw new SharePointOrganizationsError("malformed", `Item ${item.id} is missing SharePoint Created/Modified timestamps.`);
  }
  const appId = typeof fields.AppId === "string" ? fields.AppId : "";
  if (!appId) throw new SharePointOrganizationsError("malformed", `Item ${item.id} is missing AppId.`);
  const type = typeof fields.OrganizationType === "string" ? fields.OrganizationType : "";
  if (!ORGANIZATION_TYPES.includes(type as OrganizationTypeValue)) {
    throw new SharePointOrganizationsError("malformed", `Item ${item.id} has an unrecognized OrganizationType "${type}".`);
  }

  return {
    appId,
    name: String(fields.Title ?? ""),
    type: type as OrganizationTypeValue,
    metadata: {
      providerId: item.id,
      version,
      createdAt: item.createdDateTime,
      modifiedAt: item.lastModifiedDateTime,
      syncState: "saved",
    },
  };
}

function itemsPath(config: SharePointOrganizationConfig): string {
  return `/sites/${encodeURIComponent(config.siteId)}/lists/${encodeURIComponent(config.organizationsListId)}/items`;
}

function graphErrorFor(response: Response, operation: string): SharePointOrganizationsError {
  if (response.status === 401) return new SharePointOrganizationsError("auth", "Microsoft authentication has expired or is invalid.");
  if (response.status === 403) return new SharePointOrganizationsError("auth", "The signed-in account is not authorized for the DEV Organizations list.");
  if (response.status === 404) return new SharePointOrganizationsError("not_found", `${operation} could not find the requested item.`);
  return new SharePointOrganizationsError("unexpected", `${operation} failed with HTTP ${response.status}.`);
}

async function graphRequest(url: string, token: string, init: RequestInit, fetcher: FetchLike): Promise<Response> {
  try {
    return await fetcher(url, {
      ...init,
      headers: { Accept: "application/json", Authorization: `Bearer ${token}`, ...init.headers },
    });
  } catch {
    throw new SharePointOrganizationsError("network", "Microsoft Graph could not be reached.");
  }
}

/** Follows every @odata.nextLink; never assumes one page contains all items. */
export async function listOrganizationItems(
  config: SharePointOrganizationConfig,
  token: string,
  fetcher: FetchLike = fetch,
): Promise<Organization[]> {
  const organizations: Organization[] = [];
  let url: string | undefined = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}?$expand=fields&$top=200`;
  while (url) {
    const response = await graphRequest(url, token, {}, fetcher);
    if (!response.ok) throw graphErrorFor(response, "Listing Organizations");
    const payload = (await response.json()) as { value?: GraphListItem[]; "@odata.nextLink"?: string };
    for (const item of payload.value ?? []) organizations.push(fromSharePointItem(item));
    url = payload["@odata.nextLink"];
  }
  return organizations;
}

export async function getOrganizationItem(
  config: SharePointOrganizationConfig,
  token: string,
  itemId: string,
  fetcher: FetchLike = fetch,
): Promise<ResolvedOrganizationItem> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}/${encodeURIComponent(itemId)}?$expand=fields`;
  const response = await graphRequest(url, token, {}, fetcher);
  if (response.status === 404) throw new SharePointOrganizationsError("not_found", `Organization item ${itemId} was not found.`);
  if (!response.ok) throw graphErrorFor(response, "Reading the Organization");
  const item = (await response.json()) as GraphListItem;
  return { itemId: item.id, etag: String(item.eTag ?? item["@odata.etag"] ?? ""), organization: fromSharePointItem(item) };
}

/** Indexed AppId lookup: create-time uniqueness check and the update fallback (mirrors every other resource). */
export async function findOrganizationByAppId(
  config: SharePointOrganizationConfig,
  token: string,
  appId: string,
  fetcher: FetchLike = fetch,
): Promise<ResolvedOrganizationItem | null> {
  const escaped = appId.replace(/'/g, "''");
  const filter = encodeURIComponent(`fields/AppId eq '${escaped}'`);
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}?$expand=fields&$filter=${filter}`;
  const response = await graphRequest(url, token, {}, fetcher);
  if (!response.ok) throw graphErrorFor(response, "Looking up the Organization by AppId");
  const payload = (await response.json()) as { value?: GraphListItem[] };
  const item = payload.value?.[0];
  if (!item) return null;
  return { itemId: item.id, etag: String(item.eTag ?? item["@odata.etag"] ?? ""), organization: fromSharePointItem(item) };
}

/** Resolves the update target by metadata.providerId first, verifying AppId; falls back to indexed AppId lookup. */
export async function resolveOrganizationItem(
  config: SharePointOrganizationConfig,
  token: string,
  organization: Organization,
  fetcher: FetchLike = fetch,
): Promise<ResolvedOrganizationItem | null> {
  const providerId = organization.metadata?.providerId;
  if (providerId) {
    try {
      const resolved = await getOrganizationItem(config, token, providerId, fetcher);
      if (resolved.organization.appId === organization.appId) return resolved;
    } catch (error) {
      if (!(error instanceof SharePointOrganizationsError && error.kind === "not_found")) throw error;
    }
  }
  return findOrganizationByAppId(config, token, organization.appId, fetcher);
}

/** Create, then read the item back so the result carries SharePoint's id/timestamps. */
export async function createOrganizationItem(
  config: SharePointOrganizationConfig,
  token: string,
  organization: Organization,
  fetcher: FetchLike = fetch,
): Promise<Organization> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}`;
  const response = await graphRequest(
    url,
    token,
    { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fields: toSharePointFields(organization, 1) }) },
    fetcher,
  );
  if (!response.ok) throw graphErrorFor(response, "Creating the Organization");
  const created = (await response.json()) as GraphListItem;
  const resolved = await getOrganizationItem(config, token, created.id, fetcher);
  return resolved.organization;
}

/** Conditional PATCH with If-Match; a 412 is re-read and surfaced as a conflict. */
export async function updateOrganizationItem(
  config: SharePointOrganizationConfig,
  token: string,
  itemId: string,
  etag: string,
  organization: Organization,
  newVersion: number,
  fetcher: FetchLike = fetch,
): Promise<Organization> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}/${encodeURIComponent(itemId)}/fields`;
  const response = await graphRequest(
    url,
    token,
    {
      method: "PATCH",
      headers: { "Content-Type": "application/json", "If-Match": etag },
      body: JSON.stringify(toSharePointFields(organization, newVersion)),
    },
    fetcher,
  );
  if (response.status === 412) {
    const resolved = await getOrganizationItem(config, token, itemId, fetcher);
    throw new SharePointOrganizationsError("conflict", "The Organization changed in SharePoint after it was loaded.", resolved.organization);
  }
  if (!response.ok) throw graphErrorFor(response, "Updating the Organization");
  const resolved = await getOrganizationItem(config, token, itemId, fetcher);
  return resolved.organization;
}
