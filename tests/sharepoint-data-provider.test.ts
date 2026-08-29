import { afterEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { DelegatedSharePointDataProvider, selectDataProvider } from "../lib/data-provider";
import type { MicrosoftAuthController } from "../lib/microsoft-auth";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import { SAMPLE_RECORDS } from "../lib/sample-data";

const account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
const config = { siteId: "site-id", workRecordsListId: "list-id" };

function fakeController(acquireGraphToken: () => Promise<string>): MicrosoftAuthController {
  return { acquireGraphToken } as unknown as MicrosoftAuthController;
}

function testRecord(overrides: Partial<WorkRecord> = {}): WorkRecord {
  const source = structuredClone(SAMPLE_RECORDS.find((record) => !record.orbit.reportable)!);
  return {
    ...source,
    appId: "app-under-test",
    title: "DEV APP PROVIDER TEST — DELETE ME",
    activityDate: "2026-08-28",
    organizationIds: [],
    projectIds: [],
    contactIds: [],
    categoryIds: [],
    evidenceReferenceIds: [],
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
    schemaVersion: WORK_RECORD_SCHEMA_VERSION,
    ...overrides,
  };
}

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

function graphItem(record: WorkRecord, overrides: Partial<{ id: string; eTag: string; version: number }> = {}) {
  return {
    id: overrides.id ?? "1",
    eTag: overrides.eTag ?? '"1"',
    createdDateTime: "2026-08-28T12:00:00Z",
    lastModifiedDateTime: "2026-08-28T12:00:00Z",
    fields: {
      Title: record.title,
      AppId: record.appId,
      ActivityDate: `${record.activityDate}T00:00:00.000Z`,
      ActivityType: record.activityType,
      ShortDescription: record.description,
      DetailedNotes: record.detailedNotes,
      DurationMinutes: record.durationMinutes,
      RecordStatus: record.status,
      EngagementScope: record.engagementScope,
      ProjectIdsJson: "[]",
      OrganizationIdsJson: "[]",
      ContactIdsJson: "[]",
      CategoryIdsJson: "[]",
      EducatorsLeadersReach: 0,
      StudentsFamiliesReach: 0,
      WorkforceCommunityReach: 0,
      OtherReach: 0,
      EvidenceSummary: "",
      EvidenceReferenceIdsJson: "[]",
      WorkOutput: "",
      WorkOutcome: "",
      NextStep: "",
      FollowUpNeeded: false,
      FollowUpDate: null,
      OrbitReportable: false,
      OrbitPrimaryDeliverableCode: null,
      OrbitSupportingCodesJson: "[]",
      StemPocMinutes: 0,
      TacMinutes: 0,
      OrbitEvidence: "",
      SchemaVersion: WORK_RECORD_SCHEMA_VERSION,
      RecordVersion: overrides.version ?? 1,
      IsSample: false,
    },
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("DelegatedSharePointDataProvider", () => {
  it("rejects an invalid record before ever acquiring a Graph token", async () => {
    const acquireGraphToken = vi.fn();
    const provider = new DelegatedSharePointDataProvider(fakeController(acquireGraphToken), account, config);
    const result = await provider.createWorkRecord(testRecord({ title: "" }));
    expect(result.status).toBe("validation_error");
    expect(acquireGraphToken).not.toHaveBeenCalled();
  });

  it("rejects a title over the SharePoint-compatible limit before writing", async () => {
    const acquireGraphToken = vi.fn();
    const provider = new DelegatedSharePointDataProvider(fakeController(acquireGraphToken), account, config);
    const result = await provider.createWorkRecord(testRecord({ title: "x".repeat(300) }));
    expect(result.status).toBe("validation_error");
    expect(acquireGraphToken).not.toHaveBeenCalled();
  });

  it("returns a structured conflict when AppId already exists on create", async () => {
    const record = testRecord();
    const fetcher = vi.fn(async () => jsonResponse({ value: [graphItem(record)] })) as unknown as typeof fetch;
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointDataProvider(fakeController(async () => "token"), account, config);
    const result = await provider.createWorkRecord(record);
    expect(result.status).toBe("conflict");
  });

  it("creates successfully and normalizes the read-back record", async () => {
    const record = testRecord();
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ value: [] })) // AppId uniqueness check
      .mockResolvedValueOnce(jsonResponse({ id: "10" })) // create
      .mockResolvedValueOnce(jsonResponse(graphItem(record, { id: "10" }))) as unknown as typeof fetch; // read-back
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointDataProvider(fakeController(async () => "token"), account, config);
    const result = await provider.createWorkRecord(record);
    expect(result.status).toBe("success");
    if (result.status === "success") {
      expect(result.value.metadata.providerId).toBe("10");
      expect(result.value.metadata.version).toBe(1);
    }
  });

  it("returns conflict without writing when RecordVersion no longer matches expectedVersion", async () => {
    const record = testRecord({ metadata: { providerId: "10", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" } });
    const fetcher = vi.fn(async () => jsonResponse(graphItem(record, { id: "10", version: 2 }))) as unknown as typeof fetch;
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointDataProvider(fakeController(async () => "token"), account, config);
    const result = await provider.updateWorkRecord(record, 1);
    expect(result.status).toBe("conflict");
    expect(fetcher).toHaveBeenCalledTimes(1); // resolved and compared; no PATCH attempted
  });

  it("maps an expired Graph token to a network_error result instead of throwing", async () => {
    const provider = new DelegatedSharePointDataProvider(
      fakeController(async () => {
        throw new Error("boom");
      }),
      account,
      config,
    );
    const result = await provider.getWorkRecords();
    expect(result.status).toBe("network_error");
  });
});

describe("selectDataProvider", () => {
  it("falls back to the prototype ApiDataProvider outside a browser context", async () => {
    const { kind } = await selectDataProvider();
    expect(kind).toBe("api");
  });
});
