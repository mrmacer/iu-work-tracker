import { describe, expect, it } from "vitest";
import { MemoryDataProvider } from "../lib/data-provider";
import { WORK_RECORD_SCHEMA_VERSION, type EngagementScope, type WorkRecord } from "../lib/models";
import { REFERENCE_DATA } from "../lib/reference-data";
import { deriveReportingDays, deriveReportingQuarter, deriveSchoolYear, deriveStemPocMinutes, deriveTacMinutes } from "../lib/reporting";
import { SAMPLE_RECORDS } from "../lib/sample-data";
import { validateWorkRecord } from "../lib/validation";

function newRecord(overrides: Partial<WorkRecord> = {}): WorkRecord {
  const source = structuredClone(SAMPLE_RECORDS.find((record) => !record.orbit.reportable)!);
  return {
    ...source,
    appId: `test-${crypto.randomUUID()}`,
    title: "Development test work record",
    activityDate: "2026-08-26",
    activityType: "Resource development",
    engagementScope: "none",
    organizationIds: [],
    projectIds: [],
    contactIds: [],
    categoryIds: [],
    evidenceReferenceIds: [],
    isSample: false,
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
    schemaVersion: WORK_RECORD_SCHEMA_VERSION,
    ...overrides,
  };
}

function orbitRecord(overrides: Partial<WorkRecord> = {}) {
  return newRecord({
    durationMinutes: 90,
    orbit: { reportable: true, primaryDeliverable: "B", supportingDeliverables: ["D"], stemPocMinutes: 60, tacMinutes: 30, evidence: "Development evidence" },
    ...overrides,
  });
}

describe("reporting calculations", () => {
  it("derives school year around the July boundary", () => {
    expect(deriveSchoolYear("2026-06-30", REFERENCE_DATA.reportingConfig)).toBe("2025-2026");
    expect(deriveSchoolYear("2026-07-01", REFERENCE_DATA.reportingConfig)).toBe("2026-2027");
  });
  it("derives configured quarters", () => {
    expect(deriveReportingQuarter("2026-08-26", REFERENCE_DATA.reportingConfig)).toBe("Q1");
    expect(deriveReportingQuarter("2027-02-01", REFERENCE_DATA.reportingConfig)).toBe("Q3");
  });
  it("converts precise minutes into reporting days", () => {
    expect(deriveReportingDays(210, REFERENCE_DATA.reportingConfig)).toBe(0.5);
  });
  it("projects PoC and TaC only from reportable records", () => {
    expect(deriveStemPocMinutes(orbitRecord())).toBe(60);
    expect(deriveTacMinutes(orbitRecord())).toBe(30);
    expect(deriveStemPocMinutes(newRecord())).toBe(0);
  });
});

describe("work record validation", () => {
  it.each<[EngagementScope, string[]]>([
    ["none", []],
    ["specific", ["org-north-valley"]],
    ["regional", []],
    ["allDistricts", []],
  ])("accepts %s engagement scope", (engagementScope, organizationIds) => {
    expect(validateWorkRecord(newRecord({ engagementScope, organizationIds }), REFERENCE_DATA).valid).toBe(true);
  });
  it("rejects scope and district mismatches", () => {
    expect(validateWorkRecord(newRecord({ engagementScope: "specific" }), REFERENCE_DATA).valid).toBe(false);
    expect(validateWorkRecord(newRecord({ engagementScope: "regional", organizationIds: ["org-riverbend"] }), REFERENCE_DATA).valid).toBe(false);
  });
  it("requires a primary deliverable only for reportable work", () => {
    expect(validateWorkRecord(newRecord(), REFERENCE_DATA).valid).toBe(true);
    expect(validateWorkRecord(orbitRecord({ orbit: { reportable: true, primaryDeliverable: null, supportingDeliverables: [], stemPocMinutes: 0, tacMinutes: 0, evidence: "" } }), REFERENCE_DATA).valid).toBe(false);
  });
  it("accepts distinct supporting deliverables", () => {
    expect(validateWorkRecord(orbitRecord(), REFERENCE_DATA).valid).toBe(true);
  });
  it("rejects duplicate, primary, and invalid supporting deliverables", () => {
    for (const supportingDeliverables of [["D", "D"], ["B"], ["Z"]]) {
      expect(validateWorkRecord(orbitRecord({ orbit: { ...orbitRecord().orbit, supportingDeliverables } }), REFERENCE_DATA).valid).toBe(false);
    }
  });
  it("rejects negative counts and reporting time", () => {
    expect(validateWorkRecord(newRecord({ reach: { educatorsLeaders: -1, studentsFamilies: 0, workforceCommunity: 0, other: 0 } }), REFERENCE_DATA).valid).toBe(false);
    expect(validateWorkRecord(orbitRecord({ orbit: { ...orbitRecord().orbit, stemPocMinutes: -1 } }), REFERENCE_DATA).valid).toBe(false);
  });
  it("rejects hidden double-counting above activity duration", () => {
    expect(validateWorkRecord(orbitRecord({ durationMinutes: 60 }), REFERENCE_DATA).valid).toBe(false);
  });
  it("rejects malformed arrays and nested ORBIT values", () => {
    expect(validateWorkRecord({ ...newRecord(), projectIds: "project-steels" }, REFERENCE_DATA).valid).toBe(false);
    expect(validateWorkRecord({ ...newRecord(), orbit: [] }, REFERENCE_DATA).valid).toBe(false);
  });
});

describe("provider boundary and save safety", () => {
  it("retrieves every reference family through the provider", async () => {
    const provider = new MemoryDataProvider([]);
    const values = await Promise.all([provider.getProjects(), provider.getOrganizations(), provider.getContacts(), provider.getCategories(), provider.getDeliverables(), provider.getReportingConfig(), provider.getSystemSettings()]);
    expect(values.every((value) => Array.isArray(value) ? value.length > 0 : Boolean(value))).toBe(true);
  });
  it("creates non-ORBIT and ORBIT records with structured success", async () => {
    const provider = new MemoryDataProvider([], () => "2026-08-26T13:00:00.000Z");
    const plain = await provider.createWorkRecord(newRecord());
    const orbit = await provider.createWorkRecord(orbitRecord());
    expect(plain.status).toBe("success");
    expect(orbit.status).toBe("success");
    if (orbit.status === "success") expect(orbit.value.orbit.reportable).toBe(true);
  });
  it("updates safely, preserves created time, and owns modified time", async () => {
    const times = ["2026-08-26T13:00:00.000Z", "2026-08-26T13:05:00.000Z"];
    const provider = new MemoryDataProvider([], () => times.shift()!);
    const created = await provider.createWorkRecord(newRecord());
    expect(created.status).toBe("success");
    if (created.status !== "success") return;
    const updated = await provider.updateWorkRecord({ ...created.value, outcome: "Updated outcome", metadata: { ...created.value.metadata, createdAt: "client-change", modifiedAt: "client-change" } }, 1);
    expect(updated.status).toBe("success");
    if (updated.status !== "success") return;
    expect(updated.value.metadata.version).toBe(2);
    expect(updated.value.metadata.createdAt).toBe("2026-08-26T13:00:00.000Z");
    expect(updated.value.metadata.modifiedAt).toBe("2026-08-26T13:05:00.000Z");
  });
  it("returns a structured conflict for stale updates", async () => {
    const provider = new MemoryDataProvider([], () => new Date().toISOString());
    const created = await provider.createWorkRecord(newRecord());
    if (created.status !== "success") throw new Error("Setup failed");
    const first = await provider.updateWorkRecord({ ...created.value, outcome: "First" }, 1);
    expect(first.status).toBe("success");
    const stale = await provider.updateWorkRecord({ ...created.value, outcome: "Stale" }, 1);
    expect(stale.status).toBe("conflict");
  });
  it("returns structured validation errors", async () => {
    const provider = new MemoryDataProvider([]);
    const result = await provider.createWorkRecord(newRecord({ durationMinutes: -1 }));
    expect(result.status).toBe("validation_error");
  });
});

describe("log it once projections", () => {
  it("reuses one record across history, today, project, LEA, and ORBIT derivations", () => {
    const record = SAMPLE_RECORDS.find((item) => item.appId === "sample-steels")!;
    const history = SAMPLE_RECORDS.includes(record);
    const todayMinutes = SAMPLE_RECORDS.filter((item) => item.activityDate === record.activityDate).reduce((sum, item) => sum + item.durationMinutes, 0);
    const projectRecords = SAMPLE_RECORDS.filter((item) => item.projectIds.includes("project-steels"));
    const organizationRecords = SAMPLE_RECORDS.filter((item) => item.organizationIds.includes("org-north-valley"));
    const orbitRecords = SAMPLE_RECORDS.filter((item) => item.orbit.reportable);
    expect(history).toBe(true);
    expect(todayMinutes).toBeGreaterThanOrEqual(record.durationMinutes);
    expect(projectRecords).toContain(record);
    expect(organizationRecords).toContain(record);
    expect(orbitRecords).toContain(record);
    expect(deriveReportingDays(deriveStemPocMinutes(record), REFERENCE_DATA.reportingConfig)).toBeCloseTo(1 / 7);
    expect(SAMPLE_RECORDS.some((item) => item.organizationIds.includes("org-regional"))).toBe(false);
  });
});
