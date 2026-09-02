import { MICROSOFT_GRAPH_BASE_URL } from "./microsoft-auth-config";
import type { Contact, ContactStatus } from "./models";
import { CONTACT_STATUSES } from "./contact-provider";
import type { ValidationIssue } from "./validation";

// Field mapping, timestamp, ETag, and pagination behavior follow the exact same conventions as
// lib/sharepoint-projects.ts. Targets the real, live, approved IU_Contacts list (Patch 8B) —
// see docs/AI_HANDOFF.md "Durable Contacts (Patch 8B)". Title/AppId/Role/OrganizationAppId
// were live-verified before this file was written (Single line text; AppId required/indexed/
// unique; Role relaxed from required to optional; OrganizationAppId optional text, not a
// Lookup). Email/Status/Notes/RecordVersion are the four columns approved and added to the
// live list as part of Patch 8B — their internal names below are PENDING FINAL LIVE
// RE-VERIFICATION after creation; update this comment and the field names once confirmed.

export type SharePointContactConfig = {
  siteId: string;
  contactsListId: string;
};

export type ResolvedContactItem = {
  itemId: string;
  etag: string;
  contact: Contact;
};

export class SharePointContactsError extends Error {
  constructor(
    public readonly kind: "network" | "auth" | "not_found" | "conflict" | "malformed" | "unexpected",
    message: string,
    public readonly current?: Contact,
  ) {
    super(message);
    this.name = "SharePointContactsError";
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

const TEXT_LIMITS: { path: string; max: number; value: (contact: Contact) => string }[] = [
  { path: "appId", max: 255, value: (c) => c.appId },
  { path: "displayName", max: 255, value: (c) => c.displayName },
  { path: "role", max: 255, value: (c) => c.role ?? "" },
  { path: "email", max: 255, value: (c) => c.email ?? "" },
  { path: "notes", max: 2000, value: (c) => c.notes ?? "" },
];

export function validateContactSharePointLimits(contact: Contact): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  for (const limit of TEXT_LIMITS) {
    if (limit.value(contact).length > limit.max) {
      issues.push({
        path: limit.path,
        code: "sharepoint_text_limit",
        message: `${limit.path} exceeds the SharePoint-compatible limit of ${limit.max} characters.`,
      });
    }
  }
  return issues;
}

/** Maps a durable Contact to the live IU_Contacts Graph fields. */
export function toSharePointFields(contact: Contact, version: number): Record<string, unknown> {
  return {
    Title: contact.displayName || "Untitled contact",
    AppId: contact.appId,
    Role: contact.role ?? "",
    OrganizationAppId: contact.organizationId,
    Email: contact.email ?? "",
    Status: contact.status,
    Notes: contact.notes ?? "",
    RecordVersion: version,
  };
}

/** Maps one Graph listItem (with $expand=fields) back to a durable Contact. Strict: throws rather than coerces. */
export function fromSharePointItem(item: GraphListItem): Contact {
  const fields = item.fields;
  if (!fields) throw new SharePointContactsError("malformed", `Microsoft Graph did not return fields for item ${item.id}.`);

  const version = Number(fields.RecordVersion);
  if (!Number.isInteger(version) || version < 1) {
    throw new SharePointContactsError("malformed", `Item ${item.id} has an invalid RecordVersion.`);
  }
  if (!item.createdDateTime || !item.lastModifiedDateTime) {
    throw new SharePointContactsError("malformed", `Item ${item.id} is missing SharePoint Created/Modified timestamps.`);
  }
  const appId = typeof fields.AppId === "string" ? fields.AppId : "";
  if (!appId) throw new SharePointContactsError("malformed", `Item ${item.id} is missing AppId.`);
  const status = typeof fields.Status === "string" ? fields.Status : "";
  if (!CONTACT_STATUSES.includes(status as ContactStatus)) {
    throw new SharePointContactsError("malformed", `Item ${item.id} has an unrecognized Status "${status}".`);
  }
  const role = typeof fields.Role === "string" && fields.Role ? fields.Role : undefined;
  const email = typeof fields.Email === "string" && fields.Email ? fields.Email : undefined;
  const notes = typeof fields.Notes === "string" && fields.Notes ? fields.Notes : undefined;
  const organizationId = typeof fields.OrganizationAppId === "string" && fields.OrganizationAppId ? fields.OrganizationAppId : null;

  return {
    appId,
    displayName: String(fields.Title ?? ""),
    role,
    organizationId,
    email,
    status: status as ContactStatus,
    notes,
    metadata: {
      providerId: item.id,
      version,
      createdAt: item.createdDateTime,
      modifiedAt: item.lastModifiedDateTime,
      syncState: "saved",
    },
  };
}

function itemsPath(config: SharePointContactConfig): string {
  return `/sites/${encodeURIComponent(config.siteId)}/lists/${encodeURIComponent(config.contactsListId)}/items`;
}

function graphErrorFor(response: Response, operation: string): SharePointContactsError {
  if (response.status === 401) return new SharePointContactsError("auth", "Microsoft authentication has expired or is invalid.");
  if (response.status === 403) return new SharePointContactsError("auth", "The signed-in account is not authorized for the DEV Contacts list.");
  if (response.status === 404) return new SharePointContactsError("not_found", `${operation} could not find the requested item.`);
  return new SharePointContactsError("unexpected", `${operation} failed with HTTP ${response.status}.`);
}

async function graphRequest(url: string, token: string, init: RequestInit, fetcher: FetchLike): Promise<Response> {
  try {
    return await fetcher(url, {
      ...init,
      headers: { Accept: "application/json", Authorization: `Bearer ${token}`, ...init.headers },
    });
  } catch {
    throw new SharePointContactsError("network", "Microsoft Graph could not be reached.");
  }
}

/** Follows every @odata.nextLink; never assumes one page contains all items. */
export async function listContactItems(
  config: SharePointContactConfig,
  token: string,
  fetcher: FetchLike = fetch,
): Promise<Contact[]> {
  const contacts: Contact[] = [];
  let url: string | undefined = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}?$expand=fields&$top=200`;
  while (url) {
    const response = await graphRequest(url, token, {}, fetcher);
    if (!response.ok) throw graphErrorFor(response, "Listing Contacts");
    const payload = (await response.json()) as { value?: GraphListItem[]; "@odata.nextLink"?: string };
    for (const item of payload.value ?? []) contacts.push(fromSharePointItem(item));
    url = payload["@odata.nextLink"];
  }
  return contacts;
}

export async function getContactItem(
  config: SharePointContactConfig,
  token: string,
  itemId: string,
  fetcher: FetchLike = fetch,
): Promise<ResolvedContactItem> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}/${encodeURIComponent(itemId)}?$expand=fields`;
  const response = await graphRequest(url, token, {}, fetcher);
  if (response.status === 404) throw new SharePointContactsError("not_found", `Contact item ${itemId} was not found.`);
  if (!response.ok) throw graphErrorFor(response, "Reading the Contact");
  const item = (await response.json()) as GraphListItem;
  return { itemId: item.id, etag: String(item.eTag ?? item["@odata.etag"] ?? ""), contact: fromSharePointItem(item) };
}

/** Indexed AppId lookup: create-time uniqueness check and the update fallback (mirrors every other resource). */
export async function findContactByAppId(
  config: SharePointContactConfig,
  token: string,
  appId: string,
  fetcher: FetchLike = fetch,
): Promise<ResolvedContactItem | null> {
  const escaped = appId.replace(/'/g, "''");
  const filter = encodeURIComponent(`fields/AppId eq '${escaped}'`);
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}?$expand=fields&$filter=${filter}`;
  const response = await graphRequest(url, token, {}, fetcher);
  if (!response.ok) throw graphErrorFor(response, "Looking up the Contact by AppId");
  const payload = (await response.json()) as { value?: GraphListItem[] };
  const item = payload.value?.[0];
  if (!item) return null;
  return { itemId: item.id, etag: String(item.eTag ?? item["@odata.etag"] ?? ""), contact: fromSharePointItem(item) };
}

/** Resolves the update target by metadata.providerId first, verifying AppId; falls back to indexed AppId lookup. */
export async function resolveContactItem(
  config: SharePointContactConfig,
  token: string,
  contact: Contact,
  fetcher: FetchLike = fetch,
): Promise<ResolvedContactItem | null> {
  const providerId = contact.metadata?.providerId;
  if (providerId) {
    try {
      const resolved = await getContactItem(config, token, providerId, fetcher);
      if (resolved.contact.appId === contact.appId) return resolved;
    } catch (error) {
      if (!(error instanceof SharePointContactsError && error.kind === "not_found")) throw error;
    }
  }
  return findContactByAppId(config, token, contact.appId, fetcher);
}

/** Create, then read the item back so the result carries SharePoint's id/timestamps. */
export async function createContactItem(
  config: SharePointContactConfig,
  token: string,
  contact: Contact,
  fetcher: FetchLike = fetch,
): Promise<Contact> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}`;
  const response = await graphRequest(
    url,
    token,
    { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fields: toSharePointFields(contact, 1) }) },
    fetcher,
  );
  if (!response.ok) throw graphErrorFor(response, "Creating the Contact");
  const created = (await response.json()) as GraphListItem;
  const resolved = await getContactItem(config, token, created.id, fetcher);
  return resolved.contact;
}

/** Conditional PATCH with If-Match; a 412 is re-read and surfaced as a conflict. */
export async function updateContactItem(
  config: SharePointContactConfig,
  token: string,
  itemId: string,
  etag: string,
  contact: Contact,
  newVersion: number,
  fetcher: FetchLike = fetch,
): Promise<Contact> {
  const url = `${MICROSOFT_GRAPH_BASE_URL}${itemsPath(config)}/${encodeURIComponent(itemId)}/fields`;
  const response = await graphRequest(
    url,
    token,
    {
      method: "PATCH",
      headers: { "Content-Type": "application/json", "If-Match": etag },
      body: JSON.stringify(toSharePointFields(contact, newVersion)),
    },
    fetcher,
  );
  if (response.status === 412) {
    const resolved = await getContactItem(config, token, itemId, fetcher);
    throw new SharePointContactsError("conflict", "The Contact changed in SharePoint after it was loaded.", resolved.contact);
  }
  if (!response.ok) throw graphErrorFor(response, "Updating the Contact");
  const resolved = await getContactItem(config, token, itemId, fetcher);
  return resolved.contact;
}
