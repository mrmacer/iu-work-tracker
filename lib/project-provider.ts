import type { AccountInfo } from "@azure/msal-browser";
import {
  createBrowserMicrosoftAuthController,
  InteractiveRedirectStartedError,
  type MicrosoftAuthController,
} from "./microsoft-auth";
import { readDevMicrosoftConfig } from "./microsoft-auth-config";
import type { Project, ProviderMetadata } from "./models";
import {
  createProjectItem,
  findProjectByAppId,
  listProjectItems,
  resolveProjectItem,
  SharePointProjectsError,
  updateProjectItem,
  validateProjectSharePointLimits,
  type SharePointProjectConfig,
} from "./sharepoint-projects";
import type { ValidationIssue } from "./validation";

// Patch 7 — Durable Projects. Structurally identical to lib/meeting-record-provider.ts:
// list/create/update only (no delete), the same numeric-version + ETag/If-Match optimistic
// concurrency, the same memory-fallback-by-default provider selection. See
// docs/AI_HANDOFF.md "Durable Projects (Patch 7)".

export const PROJECT_STATUSES = ["planning", "active", "paused", "complete"] as const;
export type ProjectStatus = (typeof PROJECT_STATUSES)[number];

const DATE_ONLY_PATTERN = /^\d{4}-\d{2}-\d{2}$/;

export type ProjectResult<T> =
  | { status: "success"; value: T }
  | { status: "validation_error"; errors: ValidationIssue[] }
  | { status: "conflict"; current: Project; message: string }
  | { status: "network_error"; message: string }
  | { status: "persistence_error"; message: string };

/** Minimum operations only — no delete, matching MeetingRecordProvider/InboxIntelligenceProvider. */
export interface ProjectProvider {
  list(): Promise<ProjectResult<Project[]>>;
  create(project: Project): Promise<ProjectResult<Project>>;
  update(project: Project, expectedVersion: number): Promise<ProjectResult<Project>>;
}

/**
 * Builds a fresh, unsaved durable Project from the create-form state. `metadata.version` stays
 * `0` so the save path always routes a first save through create() — the same convention every
 * other durable resource in this codebase uses. Pure: performs no I/O.
 */
export function buildProjectDraft(input: {
  appId: string;
  name: string;
  description: string;
  status: ProjectStatus;
  color: string;
  startDate: string | null;
  targetDate: string | null;
  stemOrbit: boolean;
}): Project {
  return {
    ...input,
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
  };
}

/** Runtime-shape validation for a Project about to be written. Self-contained — Project has no
 * cross-entity relationships to check, unlike WorkRecord (lib/validation.ts). */
export function validateProjectShape(project: Project): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  if (typeof project.name !== "string" || !project.name.trim()) {
    issues.push({ path: "name", code: "required", message: "Project name is required." });
  }
  if (!PROJECT_STATUSES.includes(project.status)) {
    issues.push({ path: "status", code: "invalid_status", message: "status must be one of planning, active, paused, complete." });
  }
  if (typeof project.color !== "string" || !project.color.trim()) {
    issues.push({ path: "color", code: "required", message: "color is required." });
  }
  for (const path of ["startDate", "targetDate"] as const) {
    const value = project[path];
    if (value !== null && value !== undefined && (typeof value !== "string" || !DATE_ONLY_PATTERN.test(value))) {
      issues.push({ path, code: "invalid_date", message: `${path} must be null or YYYY-MM-DD.` });
    }
  }
  if (project.stemOrbit !== undefined && typeof project.stemOrbit !== "boolean") {
    issues.push({ path: "stemOrbit", code: "invalid_boolean", message: "stemOrbit must be boolean." });
  }
  return issues;
}

/**
 * Fallback persistence: in-memory only, scoped to one instance. Nothing here survives a page
 * reload — this is intentional, mirroring MemoryMeetingRecordProvider's role. Used whenever DEV
 * SharePoint configuration is absent, the user isn't signed in, the durable Projects list has
 * not yet been provisioned/approved, or the SharePoint provider fails to load for any reason. A
 * project "created" here is NOT durable.
 */
export class MemoryProjectProvider implements ProjectProvider {
  private projects: Project[] = [];
  private sequence = 0;

  async list(): Promise<ProjectResult<Project[]>> {
    return {
      status: "success",
      value: [...this.projects].sort((a, b) => (b.metadata?.createdAt ?? "").localeCompare(a.metadata?.createdAt ?? "")),
    };
  }

  async create(project: Project): Promise<ProjectResult<Project>> {
    const shapeIssues = validateProjectShape(project);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const existing = this.projects.find((item) => item.appId === project.appId);
    if (existing) return { status: "conflict", current: existing, message: "A project with this application ID already exists." };
    const now = new Date().toISOString();
    const saved: Project = {
      ...project,
      metadata: { providerId: `memory:${++this.sequence}`, version: 1, createdAt: now, modifiedAt: now, syncState: "saved" },
    };
    this.projects = [saved, ...this.projects];
    return { status: "success", value: saved };
  }

  async update(project: Project, expectedVersion: number): Promise<ProjectResult<Project>> {
    const shapeIssues = validateProjectShape(project);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const index = this.projects.findIndex((item) => item.appId === project.appId);
    if (index < 0) return { status: "persistence_error", message: "The project no longer exists." };
    const current = this.projects[index];
    if ((current.metadata as ProviderMetadata).version !== expectedVersion) {
      return { status: "conflict", current, message: "This project was changed after you opened it. Your edits are still available." };
    }
    const saved: Project = {
      ...project,
      metadata: { ...(current.metadata as ProviderMetadata), version: (current.metadata as ProviderMetadata).version + 1, modifiedAt: new Date().toISOString() },
    };
    this.projects[index] = saved;
    return { status: "success", value: saved };
  }
}

/**
 * Durable DEV SharePoint persistence for Projects, following the exact same pattern as
 * DelegatedSharePointMeetingRecordProvider: the existing delegated Microsoft sign-in, the same
 * numeric-version + ETag/If-Match algorithm, the same re-read-before-write sequence.
 */
export class DelegatedSharePointProjectProvider implements ProjectProvider {
  constructor(
    private readonly controller: MicrosoftAuthController,
    private readonly account: AccountInfo,
    private readonly config: SharePointProjectConfig,
  ) {}

  private token(): Promise<string> {
    return this.controller.acquireGraphToken(this.account);
  }

  async list(): Promise<ProjectResult<Project[]>> {
    try {
      const token = await this.token();
      const projects = await listProjectItems(this.config, token);
      projects.sort((a, b) => (b.metadata?.createdAt ?? "").localeCompare(a.metadata?.createdAt ?? ""));
      return { status: "success", value: projects };
    } catch (error) {
      return this.toErrorResult<Project[]>(error);
    }
  }

  async create(project: Project): Promise<ProjectResult<Project>> {
    const shapeIssues = validateProjectShape(project);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const limitIssues = validateProjectSharePointLimits(project);
    if (limitIssues.length) return { status: "validation_error", errors: limitIssues };
    try {
      const token = await this.token();
      const existing = await findProjectByAppId(this.config, token, project.appId);
      if (existing) {
        return { status: "conflict", current: existing.project, message: "A project with this application ID already exists." };
      }
      const saved = await createProjectItem(this.config, token, project);
      return { status: "success", value: saved };
    } catch (error) {
      return this.toErrorResult<Project>(error);
    }
  }

  async update(project: Project, expectedVersion: number): Promise<ProjectResult<Project>> {
    const shapeIssues = validateProjectShape(project);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const limitIssues = validateProjectSharePointLimits(project);
    if (limitIssues.length) return { status: "validation_error", errors: limitIssues };
    try {
      const token = await this.token();
      const resolved = await resolveProjectItem(this.config, token, project);
      if (!resolved) return { status: "persistence_error", message: "The project no longer exists." };
      if ((resolved.project.metadata as ProviderMetadata).version !== expectedVersion) {
        return { status: "conflict", current: resolved.project, message: "This project was changed after you opened it. Your edits are still available." };
      }
      const saved = await updateProjectItem(this.config, token, resolved.itemId, resolved.etag, project, expectedVersion + 1);
      return { status: "success", value: saved };
    } catch (error) {
      if (error instanceof SharePointProjectsError && error.kind === "conflict" && error.current) {
        return { status: "conflict", current: error.current, message: "This project was changed after you opened it. Your edits are still available." };
      }
      return this.toErrorResult<Project>(error);
    }
  }

  private toErrorResult<T>(error: unknown): ProjectResult<T> {
    if (error instanceof InteractiveRedirectStartedError) {
      return { status: "network_error", message: "Microsoft sign-in confirmation is required. Finish signing in, then try again." };
    }
    if (error instanceof SharePointProjectsError) {
      return error.kind === "auth"
        ? { status: "network_error", message: error.message }
        : { status: "persistence_error", message: error.message };
    }
    return { status: "network_error", message: "The DEV SharePoint data store could not be reached." };
  }
}

export type ActiveProjectProviderKind = "sharepoint" | "memory";

/**
 * Provider selection mirrors selectMeetingRecordProvider() exactly. NEXT_PUBLIC_SHAREPOINT_IU_
 * PROJECTS_LIST_ID — the existing IU_Projects list from the original seven-list reference-data
 * provisioning (docs/SHAREPOINT_PROVISIONING_CHECKLIST.md) — is the single authoritative
 * durable-Projects list configuration (Patch 7B, after the Patch 7 schema approval gate). It
 * was extended in place with StartDate/TargetDate/StemOrbit/RecordVersion columns and a new
 * "paused" ProjectStatus Choice value; see docs/AI_HANDOFF.md "Durable Projects (Patch 7)".
 * There is deliberately no second env var: Patch 7 briefly used a separate
 * NEXT_PUBLIC_SHAREPOINT_IU_DURABLE_PROJECTS_LIST_ID while this list's live schema was
 * unverified, and that temporary safety rail has been fully retired now that it's confirmed
 * and extended — reusing it here permanently would create exactly the duplicate-list-concept
 * this patch exists to avoid.
 */
export async function selectProjectProvider(): Promise<{
  provider: ProjectProvider;
  kind: ActiveProjectProviderKind;
}> {
  if (typeof window === "undefined") return { provider: new MemoryProjectProvider(), kind: "memory" };

  const config = readDevMicrosoftConfig();
  const siteId = process.env.NEXT_PUBLIC_SHAREPOINT_SITE_ID?.trim();
  const listId = process.env.NEXT_PUBLIC_SHAREPOINT_IU_PROJECTS_LIST_ID?.trim();
  if (config.status !== "enabled" || !siteId || !listId) {
    return { provider: new MemoryProjectProvider(), kind: "memory" };
  }

  try {
    const controller = createBrowserMicrosoftAuthController(config.value, window.location.origin);
    const account = await controller.initialize();
    if (!account) return { provider: new MemoryProjectProvider(), kind: "memory" };
    return {
      provider: new DelegatedSharePointProjectProvider(controller, account, { siteId, projectsListId: listId }),
      kind: "sharepoint",
    };
  } catch {
    return { provider: new MemoryProjectProvider(), kind: "memory" };
  }
}
