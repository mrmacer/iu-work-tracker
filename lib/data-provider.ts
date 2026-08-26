import type { Category, Contact, Deliverable, Organization, Project, ReferenceData, ReportingConfig, SystemSettings, WorkRecord } from "./models";
import { REFERENCE_DATA } from "./reference-data";
import { SAMPLE_RECORDS } from "./sample-data";
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
