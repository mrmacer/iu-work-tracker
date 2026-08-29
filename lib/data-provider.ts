import type { AccountInfo } from "@azure/msal-browser";
import type { Category, Contact, Deliverable, Organization, Project, ReferenceData, ReportingConfig, SystemSettings, WorkRecord } from "./models";
import {
  createBrowserMicrosoftAuthController,
  InteractiveRedirectStartedError,
  type MicrosoftAuthController,
} from "./microsoft-auth";
import { readDevMicrosoftConfig } from "./microsoft-auth-config";
import { REFERENCE_DATA } from "./reference-data";
import { SAMPLE_RECORDS } from "./sample-data";
import {
  createWorkRecordItem,
  findWorkRecordByAppId,
  listWorkRecords,
  resolveWorkRecordItem,
  SharePointWorkRecordsError,
  updateWorkRecordItem,
  validateSharePointTextLimits,
  type SharePointWorkRecordConfig,
} from "./sharepoint-work-records";
import { validateWorkRecord, type ValidationIssue } from "./validation";

export type ProviderResult<T> =
  | { status: "success"; value: T }
  | { status: "validation_error"; errors: ValidationIssue[] }
  | { status: "conflict"; current: WorkRecord; message: string }
  | { status: "network_error"; message: string }
  | { status: "persistence_error"; message: string };

export interface DataProvider {
  getWorkRecords(): Promise<ProviderResult<WorkRecord[]>>;
  createWorkRecord(record: WorkRecord): Promise<ProviderResult<WorkRecord>>;
  updateWorkRecord(record: WorkRecord, expectedVersion: number): Promise<ProviderResult<WorkRecord>>;
  getProjects(): Promise<Project[]>;
  getOrganizations(): Promise<Organization[]>;
  getContacts(): Promise<Contact[]>;
  getCategories(): Promise<Category[]>;
  getDeliverables(): Promise<Deliverable[]>;
  getReportingConfig(): Promise<ReportingConfig>;
  getSystemSettings(): Promise<SystemSettings>;
}

abstract class ReferenceProvider {
  protected references: ReferenceData = REFERENCE_DATA;
  async getProjects() { return structuredClone(this.references.projects); }
  async getOrganizations() { return structuredClone(this.references.organizations); }
  async getContacts() { return structuredClone(this.references.contacts); }
  async getCategories() { return structuredClone(this.references.categories); }
  async getDeliverables() { return structuredClone(this.references.deliverables); }
  async getReportingConfig() { return structuredClone(this.references.reportingConfig); }
  async getSystemSettings() { return structuredClone(this.references.settings); }
}

function parseResult<T>(payload: unknown, response: Response): ProviderResult<T> {
  if (payload && typeof payload === "object" && "status" in payload) return payload as ProviderResult<T>;
  return response.ok
    ? { status: "success", value: payload as T }
    : { status: "persistence_error", message: "The data store returned an unexpected response." };
}

export class ApiDataProvider extends ReferenceProvider implements DataProvider {
  private async request<T>(input: string, init?: RequestInit): Promise<ProviderResult<T>> {
    try {
      const response = await fetch(input, { cache: "no-store", ...init });
      const payload = await response.json().catch(() => null);
      return parseResult<T>(payload, response);
    } catch {
      return { status: "network_error", message: "The connected data store could not be reached." };
    }
  }
  getWorkRecords() { return this.request<WorkRecord[]>("/api/records"); }
  createWorkRecord(record: WorkRecord) {
    return this.request<WorkRecord>("/api/records", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(record) });
  }
  updateWorkRecord(record: WorkRecord, expectedVersion: number) {
    return this.request<WorkRecord>("/api/records", { method: "PUT", headers: { "content-type": "application/json" }, body: JSON.stringify({ record, expectedVersion }) });
  }
}

export class MemoryDataProvider extends ReferenceProvider implements DataProvider {
  private records: WorkRecord[];
  private sequence = 0;
  constructor(records: WorkRecord[] = SAMPLE_RECORDS, private readonly clock: () => string = () => new Date().toISOString()) {
    super();
    this.records = structuredClone(records);
  }
  async getWorkRecords(): Promise<ProviderResult<WorkRecord[]>> {
    return { status: "success", value: structuredClone(this.records) };
  }
  async createWorkRecord(record: WorkRecord): Promise<ProviderResult<WorkRecord>> {
    const validated = validateWorkRecord(record, this.references);
    if (!validated.valid) return { status: "validation_error", errors: validated.issues };
    if (this.records.some((item) => item.appId === record.appId)) return { status: "conflict", current: structuredClone(this.records.find((item) => item.appId === record.appId)!), message: "A record with this application ID already exists." };
    const now = this.clock();
    const saved: WorkRecord = { ...structuredClone(record), metadata: { providerId: `memory:${++this.sequence}`, version: 1, createdAt: now, modifiedAt: now, syncState: "saved" } };
    this.records.unshift(saved);
    return { status: "success", value: structuredClone(saved) };
  }
  async updateWorkRecord(record: WorkRecord, expectedVersion: number): Promise<ProviderResult<WorkRecord>> {
    const validated = validateWorkRecord(record, this.references);
    if (!validated.valid) return { status: "validation_error", errors: validated.issues };
    const index = this.records.findIndex((item) => item.appId === record.appId);
    if (index < 0) return { status: "persistence_error", message: "The record no longer exists." };
    const current = this.records[index];
    if (current.metadata.version !== expectedVersion) return { status: "conflict", current: structuredClone(current), message: "This record was changed after you opened it. Your draft is still available." };
    const saved: WorkRecord = {
      ...structuredClone(record),
      metadata: {
        providerId: current.metadata.providerId,
        version: current.metadata.version + 1,
        createdAt: current.metadata.createdAt,
        modifiedAt: this.clock(),
        syncState: "saved",
      },
    };
    this.records[index] = saved;
    return { status: "success", value: structuredClone(saved) };
  }
}

export class PrototypeFallbackProvider extends MemoryDataProvider {}

/**
 * DEV-only Work Record persistence against the verified DEV SharePoint site, using the
 * existing delegated Microsoft sign-in (see docs/DELEGATED_AUTH_SETUP.md and
 * docs/AI_HANDOFF.md "DEV vs PRODUCTION"). Reference/configuration data intentionally
 * continues to come from the same static seed data every other provider uses; those six
 * lists are steward-owned reference data, not part of this integration phase's scope.
 */
export class DelegatedSharePointDataProvider extends ReferenceProvider implements DataProvider {
  constructor(
    private readonly controller: MicrosoftAuthController,
    private readonly account: AccountInfo,
    private readonly config: SharePointWorkRecordConfig,
  ) {
    super();
  }

  private token(): Promise<string> {
    return this.controller.acquireGraphToken(this.account);
  }

  async getWorkRecords(): Promise<ProviderResult<WorkRecord[]>> {
    try {
      const token = await this.token();
      const records = await listWorkRecords(this.config, token);
      records.sort(
        (a, b) =>
          b.activityDate.localeCompare(a.activityDate) ||
          b.metadata.createdAt.localeCompare(a.metadata.createdAt),
      );
      return { status: "success", value: records };
    } catch (error) {
      return this.toErrorResult<WorkRecord[]>(error);
    }
  }

  async createWorkRecord(record: WorkRecord): Promise<ProviderResult<WorkRecord>> {
    const validated = validateWorkRecord(record, this.references);
    if (!validated.valid) return { status: "validation_error", errors: validated.issues };
    const limitIssues = validateSharePointTextLimits(record);
    if (limitIssues.length) return { status: "validation_error", errors: limitIssues };
    try {
      const token = await this.token();
      const existing = await findWorkRecordByAppId(this.config, token, record.appId);
      if (existing) {
        return { status: "conflict", current: existing.record, message: "A record with this application ID already exists." };
      }
      const saved = await createWorkRecordItem(this.config, token, record);
      return { status: "success", value: saved };
    } catch (error) {
      return this.toErrorResult<WorkRecord>(error);
    }
  }

  async updateWorkRecord(record: WorkRecord, expectedVersion: number): Promise<ProviderResult<WorkRecord>> {
    const validated = validateWorkRecord(record, this.references);
    if (!validated.valid) return { status: "validation_error", errors: validated.issues };
    const limitIssues = validateSharePointTextLimits(record);
    if (limitIssues.length) return { status: "validation_error", errors: limitIssues };
    try {
      const token = await this.token();
      const resolved = await resolveWorkRecordItem(this.config, token, record);
      if (!resolved) return { status: "persistence_error", message: "The record no longer exists." };
      if (resolved.record.metadata.version !== expectedVersion) {
        return {
          status: "conflict",
          current: resolved.record,
          message: "This record was changed after you opened it. Your draft is still available.",
        };
      }
      const saved = await updateWorkRecordItem(
        this.config,
        token,
        resolved.itemId,
        resolved.etag,
        record,
        expectedVersion + 1,
      );
      return { status: "success", value: saved };
    } catch (error) {
      if (error instanceof SharePointWorkRecordsError && error.kind === "conflict" && error.current) {
        return {
          status: "conflict",
          current: error.current,
          message: "This record was changed after you opened it. Your draft is still available.",
        };
      }
      return this.toErrorResult<WorkRecord>(error);
    }
  }

  private toErrorResult<T>(error: unknown): ProviderResult<T> {
    if (error instanceof InteractiveRedirectStartedError) {
      return { status: "network_error", message: "Microsoft sign-in confirmation is required. Finish signing in, then try again." };
    }
    if (error instanceof SharePointWorkRecordsError) {
      return error.kind === "auth"
        ? { status: "network_error", message: error.message }
        : { status: "persistence_error", message: error.message };
    }
    return { status: "network_error", message: "The DEV SharePoint data store could not be reached." };
  }
}

export type ActiveProviderKind = "sharepoint" | "api";

/**
 * DEV SharePoint provider selection. The prototype ApiDataProvider (and its existing
 * PrototypeFallbackProvider safety net in app/IUWorkTracker.tsx) remains the default and the
 * safe fallback. The SharePoint provider activates only when DEV Microsoft/SharePoint
 * configuration is present AND a Microsoft account is already signed in — this never
 * triggers an interactive sign-in prompt on its own.
 */
export async function selectDataProvider(): Promise<{ provider: DataProvider; kind: ActiveProviderKind }> {
  if (typeof window === "undefined") return { provider: new ApiDataProvider(), kind: "api" };

  const config = readDevMicrosoftConfig();
  const siteId = process.env.NEXT_PUBLIC_SHAREPOINT_SITE_ID?.trim();
  const workRecordsListId = process.env.NEXT_PUBLIC_SHAREPOINT_IU_WORK_RECORDS_LIST_ID?.trim();
  if (config.status !== "enabled" || !siteId || !workRecordsListId) {
    return { provider: new ApiDataProvider(), kind: "api" };
  }

  try {
    const controller = createBrowserMicrosoftAuthController(config.value, window.location.origin);
    const account = await controller.initialize();
    if (!account) return { provider: new ApiDataProvider(), kind: "api" };
    return {
      provider: new DelegatedSharePointDataProvider(controller, account, { siteId, workRecordsListId }),
      kind: "sharepoint",
    };
  } catch {
    return { provider: new ApiDataProvider(), kind: "api" };
  }
}
