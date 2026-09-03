import type { AccountInfo } from "@azure/msal-browser";
import {
  createBrowserMicrosoftAuthController,
  InteractiveRedirectStartedError,
  type MicrosoftAuthController,
} from "./microsoft-auth";
import { readDevMicrosoftConfig } from "./microsoft-auth-config";
import type { Organization, ProviderMetadata } from "./models";
import {
  createOrganizationItem,
  findOrganizationByAppId,
  listOrganizationItems,
  resolveOrganizationItem,
  SharePointOrganizationsError,
  updateOrganizationItem,
  validateOrganizationSharePointLimits,
  type SharePointOrganizationConfig,
} from "./sharepoint-organizations";
import type { ValidationIssue } from "./validation";

// Patch 8E — Durable Organizations. Structurally identical to lib/project-provider.ts and
// lib/contact-provider.ts: list/create/update only (no delete), the same numeric-version +
// ETag/If-Match optimistic concurrency, the same memory-fallback-by-default provider selection.
// See docs/AI_HANDOFF.md "Durable Organizations (Patch 8E)". Unlike Patch 7's temporary
// DURABLE_PROJECTS safety rail, the existing NEXT_PUBLIC_SHAREPOINT_IU_ORGANIZATIONS_LIST_ID is
// reused directly — its live schema was inspected and a RecordVersion column added by the user
// before this file was written (Patch 8E live-inspection gate), so no separate temporary
// variable is warranted here.

export const ORGANIZATION_TYPES = ["district", "partner", "iu"] as const;
export type OrganizationTypeValue = (typeof ORGANIZATION_TYPES)[number];

export type OrganizationResult<T> =
  | { status: "success"; value: T }
  | { status: "validation_error"; errors: ValidationIssue[] }
  | { status: "conflict"; current: Organization; message: string }
  | { status: "network_error"; message: string }
  | { status: "persistence_error"; message: string };

/** Minimum operations only — no delete, matching ProjectProvider/ContactProvider. */
export interface OrganizationProvider {
  list(): Promise<OrganizationResult<Organization[]>>;
  create(organization: Organization): Promise<OrganizationResult<Organization>>;
  update(organization: Organization, expectedVersion: number): Promise<OrganizationResult<Organization>>;
}

/**
 * Builds a fresh, unsaved durable Organization from the create-form state. `metadata.version`
 * stays `0` so the save path always routes a first save through create() — the same convention
 * every other durable resource in this codebase uses. Pure: performs no I/O.
 */
export function buildOrganizationDraft(input: { appId: string; name: string; type: OrganizationTypeValue }): Organization {
  return {
    ...input,
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
  };
}

/** Normalizes a name for exact-match duplicate detection: trim + collapse internal whitespace +
 * lowercase — no fuzzy matching. Mirrors normalizeContactName (lib/contact-provider.ts). */
export function normalizeOrganizationName(value: string): string {
  return value.trim().replace(/\s+/g, " ").toLowerCase();
}

/** Runtime-shape validation for an Organization about to be written. Self-contained — Organization
 * has no cross-entity relationships to check, unlike WorkRecord (lib/validation.ts). */
export function validateOrganizationShape(organization: Organization): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  if (typeof organization.name !== "string" || !organization.name.trim()) {
    issues.push({ path: "name", code: "required", message: "Organization name is required." });
  }
  if (!ORGANIZATION_TYPES.includes(organization.type)) {
    issues.push({ path: "type", code: "invalid_type", message: "type must be one of district, partner, iu." });
  }
  return issues;
}

/**
 * Fallback persistence: in-memory only, scoped to one instance. Nothing here survives a page
 * reload — this is intentional, mirroring MemoryProjectProvider's role. Used whenever DEV
 * SharePoint configuration is absent, the user isn't signed in, or the SharePoint provider
 * fails to load for any reason. An organization "created" here is NOT durable.
 */
export class MemoryOrganizationProvider implements OrganizationProvider {
  private organizations: Organization[] = [];
  private sequence = 0;

  async list(): Promise<OrganizationResult<Organization[]>> {
    return {
      status: "success",
      value: [...this.organizations].sort((a, b) => (b.metadata?.createdAt ?? "").localeCompare(a.metadata?.createdAt ?? "")),
    };
  }

  async create(organization: Organization): Promise<OrganizationResult<Organization>> {
    const shapeIssues = validateOrganizationShape(organization);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const existing = this.organizations.find((item) => item.appId === organization.appId);
    if (existing) return { status: "conflict", current: existing, message: "An organization with this application ID already exists." };
    const now = new Date().toISOString();
    const saved: Organization = {
      ...organization,
      metadata: { providerId: `memory:${++this.sequence}`, version: 1, createdAt: now, modifiedAt: now, syncState: "saved" },
    };
    this.organizations = [saved, ...this.organizations];
    return { status: "success", value: saved };
  }

  async update(organization: Organization, expectedVersion: number): Promise<OrganizationResult<Organization>> {
    const shapeIssues = validateOrganizationShape(organization);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const index = this.organizations.findIndex((item) => item.appId === organization.appId);
    if (index < 0) return { status: "persistence_error", message: "The organization no longer exists." };
    const current = this.organizations[index];
    if ((current.metadata as ProviderMetadata).version !== expectedVersion) {
      return { status: "conflict", current, message: "This organization was changed after you opened it. Your edits are still available." };
    }
    const saved: Organization = {
      ...organization,
      metadata: { ...(current.metadata as ProviderMetadata), version: (current.metadata as ProviderMetadata).version + 1, modifiedAt: new Date().toISOString() },
    };
    this.organizations[index] = saved;
    return { status: "success", value: saved };
  }
}

/**
 * Durable DEV SharePoint persistence for Organizations, following the exact same pattern as
 * DelegatedSharePointProjectProvider: the existing delegated Microsoft sign-in, the same
 * numeric-version + ETag/If-Match algorithm, the same re-read-before-write sequence.
 */
export class DelegatedSharePointOrganizationProvider implements OrganizationProvider {
  constructor(
    private readonly controller: MicrosoftAuthController,
    private readonly account: AccountInfo,
    private readonly config: SharePointOrganizationConfig,
  ) {}

  private token(): Promise<string> {
    return this.controller.acquireGraphToken(this.account);
  }

  async list(): Promise<OrganizationResult<Organization[]>> {
    try {
      const token = await this.token();
      const organizations = await listOrganizationItems(this.config, token);
      organizations.sort((a, b) => (b.metadata?.createdAt ?? "").localeCompare(a.metadata?.createdAt ?? ""));
      return { status: "success", value: organizations };
    } catch (error) {
      return this.toErrorResult<Organization[]>(error);
    }
  }

  async create(organization: Organization): Promise<OrganizationResult<Organization>> {
    const shapeIssues = validateOrganizationShape(organization);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const limitIssues = validateOrganizationSharePointLimits(organization);
    if (limitIssues.length) return { status: "validation_error", errors: limitIssues };
    try {
      const token = await this.token();
      const existing = await findOrganizationByAppId(this.config, token, organization.appId);
      if (existing) {
        return { status: "conflict", current: existing.organization, message: "An organization with this application ID already exists." };
      }
      const saved = await createOrganizationItem(this.config, token, organization);
      return { status: "success", value: saved };
    } catch (error) {
      return this.toErrorResult<Organization>(error);
    }
  }

  async update(organization: Organization, expectedVersion: number): Promise<OrganizationResult<Organization>> {
    const shapeIssues = validateOrganizationShape(organization);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const limitIssues = validateOrganizationSharePointLimits(organization);
    if (limitIssues.length) return { status: "validation_error", errors: limitIssues };
    try {
      const token = await this.token();
      const resolved = await resolveOrganizationItem(this.config, token, organization);
      if (!resolved) return { status: "persistence_error", message: "The organization no longer exists." };
      if ((resolved.organization.metadata as ProviderMetadata).version !== expectedVersion) {
        return { status: "conflict", current: resolved.organization, message: "This organization was changed after you opened it. Your edits are still available." };
      }
      const saved = await updateOrganizationItem(this.config, token, resolved.itemId, resolved.etag, organization, expectedVersion + 1);
      return { status: "success", value: saved };
    } catch (error) {
      if (error instanceof SharePointOrganizationsError && error.kind === "conflict" && error.current) {
        return { status: "conflict", current: error.current, message: "This organization was changed after you opened it. Your edits are still available." };
      }
      return this.toErrorResult<Organization>(error);
    }
  }

  private toErrorResult<T>(error: unknown): OrganizationResult<T> {
    if (error instanceof InteractiveRedirectStartedError) {
      return { status: "network_error", message: "Microsoft sign-in confirmation is required. Finish signing in, then try again." };
    }
    if (error instanceof SharePointOrganizationsError) {
      return error.kind === "auth"
        ? { status: "network_error", message: error.message }
        : { status: "persistence_error", message: error.message };
    }
    return { status: "network_error", message: "The DEV SharePoint data store could not be reached." };
  }
}

export type ActiveOrganizationProviderKind = "sharepoint" | "memory";

/**
 * Provider selection mirrors selectProjectProvider()/selectContactProvider() exactly.
 * NEXT_PUBLIC_SHAREPOINT_IU_ORGANIZATIONS_LIST_ID — the existing IU_Organizations list from the
 * original seven-list reference-data provisioning (docs/SHAREPOINT_PROVISIONING_CHECKLIST.md) —
 * is the single authoritative durable-Organizations list configuration. It was empty and
 * extended in place with a RecordVersion column (Patch 8E live-inspection gate) before this
 * file was written. There is deliberately no second env var.
 */
export async function selectOrganizationProvider(): Promise<{
  provider: OrganizationProvider;
  kind: ActiveOrganizationProviderKind;
}> {
  if (typeof window === "undefined") return { provider: new MemoryOrganizationProvider(), kind: "memory" };

  const config = readDevMicrosoftConfig();
  const siteId = process.env.NEXT_PUBLIC_SHAREPOINT_SITE_ID?.trim();
  const listId = process.env.NEXT_PUBLIC_SHAREPOINT_IU_ORGANIZATIONS_LIST_ID?.trim();
  if (config.status !== "enabled" || !siteId || !listId) {
    return { provider: new MemoryOrganizationProvider(), kind: "memory" };
  }

  try {
    const controller = createBrowserMicrosoftAuthController(config.value, window.location.origin);
    const account = await controller.initialize();
    if (!account) return { provider: new MemoryOrganizationProvider(), kind: "memory" };
    return {
      provider: new DelegatedSharePointOrganizationProvider(controller, account, { siteId, organizationsListId: listId }),
      kind: "sharepoint",
    };
  } catch {
    return { provider: new MemoryOrganizationProvider(), kind: "memory" };
  }
}
