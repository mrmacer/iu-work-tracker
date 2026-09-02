import type { AccountInfo } from "@azure/msal-browser";
import {
  createBrowserMicrosoftAuthController,
  InteractiveRedirectStartedError,
  type MicrosoftAuthController,
} from "./microsoft-auth";
import { readDevMicrosoftConfig } from "./microsoft-auth-config";
import type { Contact, ContactStatus, ProviderMetadata } from "./models";
import {
  createContactItem,
  findContactByAppId,
  listContactItems,
  resolveContactItem,
  SharePointContactsError,
  updateContactItem,
  validateContactSharePointLimits,
  type SharePointContactConfig,
} from "./sharepoint-contacts";
import type { ValidationIssue } from "./validation";

// Patch 8B — Durable Contacts. Structurally identical to lib/project-provider.ts:
// list/create/update only (no delete — archived status is sufficient, matching the approved
// architecture decision), the same numeric-version + ETag/If-Match optimistic concurrency, the
// same memory-fallback-by-default provider selection. See docs/AI_HANDOFF.md "Durable Contacts
// (Patch 8B)". Unlike Patch 7's temporary DURABLE_PROJECTS safety rail, the existing
// NEXT_PUBLIC_SHAREPOINT_IU_CONTACTS_LIST_ID is reused directly — its live schema was
// inspected and approved before this file was written (Patch 8B live-inspection gate), so no
// separate temporary variable is warranted here.

export const CONTACT_STATUSES = ["active", "developing", "occasional", "dormant", "archived"] as const;

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export type ContactResult<T> =
  | { status: "success"; value: T }
  | { status: "validation_error"; errors: ValidationIssue[] }
  | { status: "conflict"; current: Contact; message: string }
  | { status: "network_error"; message: string }
  | { status: "persistence_error"; message: string };

/** Minimum operations only — no delete, matching ProjectProvider/MeetingRecordProvider/InboxIntelligenceProvider. */
export interface ContactProvider {
  list(): Promise<ContactResult<Contact[]>>;
  create(contact: Contact): Promise<ContactResult<Contact>>;
  update(contact: Contact, expectedVersion: number): Promise<ContactResult<Contact>>;
}

/**
 * Builds a fresh, unsaved durable Contact from the create-form state. `metadata.version` stays
 * `0` so the save path always routes a first save through create() — the same convention every
 * other durable resource in this codebase uses. Pure: performs no I/O.
 */
export function buildContactDraft(input: {
  appId: string;
  displayName: string;
  role: string | undefined;
  organizationId: string | null;
  email: string | undefined;
  status: ContactStatus;
  notes: string | undefined;
}): Contact {
  return {
    ...input,
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
  };
}

/** Normalizes a display name for exact-match duplicate detection: trim + lowercase, no fuzzy matching. */
export function normalizeContactName(value: string): string {
  return value.trim().toLowerCase();
}

/** Normalizes an email for exact-match duplicate detection and future strong-match evidence: trim + lowercase. */
export function normalizeContactEmail(value: string): string {
  return value.trim().toLowerCase();
}

/** Runtime-shape validation for a Contact about to be written. Self-contained — Contact has no
 * cross-entity relationships to check beyond organizationId, unlike WorkRecord (lib/validation.ts). */
export function validateContactShape(contact: Contact): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  if (typeof contact.displayName !== "string" || !contact.displayName.trim()) {
    issues.push({ path: "displayName", code: "required", message: "Contact name is required." });
  }
  if (!CONTACT_STATUSES.includes(contact.status)) {
    issues.push({ path: "status", code: "invalid_status", message: "status must be one of active, developing, occasional, dormant, archived." });
  }
  if (contact.role !== undefined && typeof contact.role !== "string") {
    issues.push({ path: "role", code: "invalid_string", message: "role must be a string." });
  }
  if (contact.notes !== undefined && typeof contact.notes !== "string") {
    issues.push({ path: "notes", code: "invalid_string", message: "notes must be a string." });
  }
  if (contact.email !== undefined && contact.email !== "" && !EMAIL_PATTERN.test(contact.email)) {
    issues.push({ path: "email", code: "invalid_email", message: "email must be a valid email address." });
  }
  if (contact.organizationId !== null && typeof contact.organizationId !== "string") {
    issues.push({ path: "organizationId", code: "invalid_id", message: "organizationId must be a string or null." });
  }
  return issues;
}

/**
 * Fallback persistence: in-memory only, scoped to one instance. Nothing here survives a page
 * reload — this is intentional, mirroring MemoryProjectProvider's role. Used whenever DEV
 * SharePoint configuration is absent, the user isn't signed in, or the SharePoint provider
 * fails to load for any reason. A contact "created" here is NOT durable.
 */
export class MemoryContactProvider implements ContactProvider {
  private contacts: Contact[] = [];
  private sequence = 0;

  async list(): Promise<ContactResult<Contact[]>> {
    return {
      status: "success",
      value: [...this.contacts].sort((a, b) => (b.metadata?.createdAt ?? "").localeCompare(a.metadata?.createdAt ?? "")),
    };
  }

  async create(contact: Contact): Promise<ContactResult<Contact>> {
    const shapeIssues = validateContactShape(contact);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const existing = this.contacts.find((item) => item.appId === contact.appId);
    if (existing) return { status: "conflict", current: existing, message: "A contact with this application ID already exists." };
    const now = new Date().toISOString();
    const saved: Contact = {
      ...contact,
      metadata: { providerId: `memory:${++this.sequence}`, version: 1, createdAt: now, modifiedAt: now, syncState: "saved" },
    };
    this.contacts = [saved, ...this.contacts];
    return { status: "success", value: saved };
  }

  async update(contact: Contact, expectedVersion: number): Promise<ContactResult<Contact>> {
    const shapeIssues = validateContactShape(contact);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const index = this.contacts.findIndex((item) => item.appId === contact.appId);
    if (index < 0) return { status: "persistence_error", message: "The contact no longer exists." };
    const current = this.contacts[index];
    if ((current.metadata as ProviderMetadata).version !== expectedVersion) {
      return { status: "conflict", current, message: "This contact was changed after you opened it. Your edits are still available." };
    }
    const saved: Contact = {
      ...contact,
      metadata: { ...(current.metadata as ProviderMetadata), version: (current.metadata as ProviderMetadata).version + 1, modifiedAt: new Date().toISOString() },
    };
    this.contacts[index] = saved;
    return { status: "success", value: saved };
  }
}

/**
 * Durable DEV SharePoint persistence for Contacts, following the exact same pattern as
 * DelegatedSharePointProjectProvider: the existing delegated Microsoft sign-in, the same
 * numeric-version + ETag/If-Match algorithm, the same re-read-before-write sequence.
 */
export class DelegatedSharePointContactProvider implements ContactProvider {
  constructor(
    private readonly controller: MicrosoftAuthController,
    private readonly account: AccountInfo,
    private readonly config: SharePointContactConfig,
  ) {}

  private token(): Promise<string> {
    return this.controller.acquireGraphToken(this.account);
  }

  async list(): Promise<ContactResult<Contact[]>> {
    try {
      const token = await this.token();
      const contacts = await listContactItems(this.config, token);
      contacts.sort((a, b) => (b.metadata?.createdAt ?? "").localeCompare(a.metadata?.createdAt ?? ""));
      return { status: "success", value: contacts };
    } catch (error) {
      return this.toErrorResult<Contact[]>(error);
    }
  }

  async create(contact: Contact): Promise<ContactResult<Contact>> {
    const shapeIssues = validateContactShape(contact);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const limitIssues = validateContactSharePointLimits(contact);
    if (limitIssues.length) return { status: "validation_error", errors: limitIssues };
    try {
      const token = await this.token();
      const existing = await findContactByAppId(this.config, token, contact.appId);
      if (existing) {
        return { status: "conflict", current: existing.contact, message: "A contact with this application ID already exists." };
      }
      const saved = await createContactItem(this.config, token, contact);
      return { status: "success", value: saved };
    } catch (error) {
      return this.toErrorResult<Contact>(error);
    }
  }

  async update(contact: Contact, expectedVersion: number): Promise<ContactResult<Contact>> {
    const shapeIssues = validateContactShape(contact);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const limitIssues = validateContactSharePointLimits(contact);
    if (limitIssues.length) return { status: "validation_error", errors: limitIssues };
    try {
      const token = await this.token();
      const resolved = await resolveContactItem(this.config, token, contact);
      if (!resolved) return { status: "persistence_error", message: "The contact no longer exists." };
      if ((resolved.contact.metadata as ProviderMetadata).version !== expectedVersion) {
        return { status: "conflict", current: resolved.contact, message: "This contact was changed after you opened it. Your edits are still available." };
      }
      const saved = await updateContactItem(this.config, token, resolved.itemId, resolved.etag, contact, expectedVersion + 1);
      return { status: "success", value: saved };
    } catch (error) {
      if (error instanceof SharePointContactsError && error.kind === "conflict" && error.current) {
        return { status: "conflict", current: error.current, message: "This contact was changed after you opened it. Your edits are still available." };
      }
      return this.toErrorResult<Contact>(error);
    }
  }

  private toErrorResult<T>(error: unknown): ContactResult<T> {
    if (error instanceof InteractiveRedirectStartedError) {
      return { status: "network_error", message: "Microsoft sign-in confirmation is required. Finish signing in, then try again." };
    }
    if (error instanceof SharePointContactsError) {
      return error.kind === "auth"
        ? { status: "network_error", message: error.message }
        : { status: "persistence_error", message: error.message };
    }
    return { status: "network_error", message: "The DEV SharePoint data store could not be reached." };
  }
}

export type ActiveContactProviderKind = "sharepoint" | "memory";

/**
 * Provider selection mirrors selectProjectProvider() exactly. Unlike Patch 7's Project list
 * (which reused an env var pointed at a different-purpose, unverified list), the existing
 * NEXT_PUBLIC_SHAREPOINT_IU_CONTACTS_LIST_ID is used directly here — its live schema was
 * inspected and its four missing columns (Email/Status/Notes/RecordVersion) approved and added
 * before this file was written (Patch 8B), so there is no unverified-list safety concern to
 * guard against with a second temporary variable.
 */
export async function selectContactProvider(): Promise<{
  provider: ContactProvider;
  kind: ActiveContactProviderKind;
}> {
  if (typeof window === "undefined") return { provider: new MemoryContactProvider(), kind: "memory" };

  const config = readDevMicrosoftConfig();
  const siteId = process.env.NEXT_PUBLIC_SHAREPOINT_SITE_ID?.trim();
  const listId = process.env.NEXT_PUBLIC_SHAREPOINT_IU_CONTACTS_LIST_ID?.trim();
  if (config.status !== "enabled" || !siteId || !listId) {
    return { provider: new MemoryContactProvider(), kind: "memory" };
  }

  try {
    const controller = createBrowserMicrosoftAuthController(config.value, window.location.origin);
    const account = await controller.initialize();
    if (!account) return { provider: new MemoryContactProvider(), kind: "memory" };
    return {
      provider: new DelegatedSharePointContactProvider(controller, account, { siteId, contactsListId: listId }),
      kind: "sharepoint",
    };
  } catch {
    return { provider: new MemoryContactProvider(), kind: "memory" };
  }
}
