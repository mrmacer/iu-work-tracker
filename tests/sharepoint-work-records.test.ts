import { describe, expect, it, vi } from "vitest";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import { SAMPLE_RECORDS } from "../lib/sample-data";
import {
  createWorkRecordItem,
  findWorkRecordByAppId,
  fromSharePointItem,
  listWorkRecords,
  resolveWorkRecordItem,
  SharePointWorkRecordsError,
  toSharePointFields,
  updateWorkRecordItem,
  validateSharePointTextLimits,
  type SharePointWorkRecordConfig,
} from "../lib/sharepoint-work-records";

const config: SharePointWorkRecordConfig = {
  siteId: "siu29.sharepoint.com,site,web",
  workRecordsListId: "work-records-list-id",
};

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

function testRecord(overrides: Partial<WorkRecord> = {}): WorkRecord {
  const source = structuredClone(SAMPLE_RECORDS.find((record) => !record.orbit.reportable)!);
  return {
    ...source,
    appId: "app-under-test",
    title: "DEV APP PROVIDER TEST — DELETE ME",
    activityDate: "2026-08-28",
    projectIds: ["project-steels"],
    organizationIds: [],
    contactIds: [],
    categoryIds: [],
    evidenceReferenceIds: ["dev-evidence-1"],
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
    ...overrides,
  };
}

function graphItem(fields: Record<string, unknown>, overrides: Partial<{ id: string; eTag: string; createdDateTime: string; lastModifiedDateTime: string }> = {}) {
  return {
    id: overrides.id ?? "1",
    eTag: overrides.eTag ?? '"1"',
    createdDateTime: overrides.createdDateTime ?? "2026-08-28T12:00:00Z",
    lastModifiedDateTime: overrides.lastModifiedDateTime ?? "2026-08-28T12:00:00Z",
    fields,
  };
}

function fieldsFor(record: WorkRecord, version = 1) {
  return toSharePointFields(record, version);
}

describe("SharePoint Work Records field mapping", () => {
  it("round-trips a canonical Work Record without mutating relationship JSON", () => {
    const record = testRecord();
    const fields = fieldsFor(record, 1);
    expect(JSON.parse(String(fields.ProjectIdsJson))).toEqual(["project-steels"]);
    const mapped = fromSharePointItem(graphItem(fields));
    expect(mapped.appId).toBe(record.appId);
    expect(mapped.title).toBe(record.title);
    expect(mapped.durationMinutes).toBe(record.durationMinutes);
    expect(mapped.projectIds).toEqual(["project-steels"]);
    expect(mapped.evidenceReferenceIds).toEqual(["dev-evidence-1"]);
    expect(mapped.schemaVersion).toBe(WORK_RECORD_SCHEMA_VERSION);
    expect(mapped.metadata).toEqual({
      providerId: "1",
      version: 1,
      createdAt: "2026-08-28T12:00:00Z",
      modifiedAt: "2026-08-28T12:00:00Z",
      syncState: "saved",
    });
  });

  it("preserves ORBIT as optional/subordinate: absent when non-reportable, present when set", () => {
    const nonReportable = fromSharePointItem(graphItem(fieldsFor(testRecord())));
    expect(nonReportable.orbit.reportable).toBe(false);
    expect(nonReportable.orbit.primaryDeliverable).toBeNull();

    const reportable = testRecord({
      durationMinutes: 90,
      orbit: { reportable: true, primaryDeliverable: "B", supportingDeliverables: ["D"], stemPocMinutes: 60, tacMinutes: 30, evidence: "note" },
    });
    const mapped = fromSharePointItem(graphItem(fieldsFor(reportable)));
    expect(mapped.orbit).toEqual(reportable.orbit);
  });

  it("rejects an unsupported SchemaVersion instead of silently coercing it", () => {
    const fields = fieldsFor(testRecord());
    fields.SchemaVersion = 1;
    expect(() => fromSharePointItem(graphItem(fields))).toThrow(SharePointWorkRecordsError);
  });

  it("rejects malformed relationship JSON instead of silently discarding it", () => {
    const fields = fieldsFor(testRecord());
    fields.ProjectIdsJson = "{not json";
    expect(() => fromSharePointItem(graphItem(fields))).toThrow(/malformed JSON/);
  });

  it("flags SharePoint-compatible text-limit violations without truncating", () => {
    const oversizedTitle = testRecord({ title: "x".repeat(256) });
    const issues = validateSharePointTextLimits(oversizedTitle);
    expect(issues).toContainEqual(expect.objectContaining({ path: "title", code: "sharepoint_text_limit" }));
  });

  it("accepts values within the SharePoint-compatible limits", () => {
    expect(validateSharePointTextLimits(testRecord())).toEqual([]);
  });
});

describe("SharePoint Work Records Graph operations", () => {
  it("follows every @odata.nextLink when listing", async () => {
    const record = testRecord();
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({ value: [graphItem(fieldsFor(record))], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next-page" }),
      )
      .mockResolvedValueOnce(jsonResponse({ value: [graphItem(fieldsFor(testRecord({ appId: "second" })))] })) as unknown as typeof fetch;

    const records = await listWorkRecords(config, "token", fetcher);
    expect(records).toHaveLength(2);
    expect(fetcher).toHaveBeenCalledTimes(2);
    expect(fetcher).toHaveBeenNthCalledWith(2, "https://graph.microsoft.com/v1.0/next-page", expect.objectContaining({}));
  });

  it("creates an item and reads it back for SharePoint-owned id/timestamps", async () => {
    const record = testRecord();
    const fetcherMock = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ id: "42" }))
      .mockResolvedValueOnce(jsonResponse(graphItem(fieldsFor(record), { id: "42" })));
    const fetcher = fetcherMock as unknown as typeof fetch;

    const saved = await createWorkRecordItem(config, "token", record, fetcher);
    expect(saved.metadata.providerId).toBe("42");
    expect(saved.metadata.version).toBe(1);
    const [, createInit] = fetcherMock.mock.calls[0] as [string, RequestInit];
    expect(createInit.method).toBe("POST");
    expect(JSON.parse(String(createInit.body)).fields.RecordVersion).toBe(1);
  });

  it("looks up an existing item by indexed AppId", async () => {
    const record = testRecord();
    const fetcher = vi.fn(async (url: string) => {
      expect(url).toContain("$filter=");
      expect(url).toContain(encodeURIComponent("fields/AppId eq 'app-under-test'"));
      return jsonResponse({ value: [graphItem(fieldsFor(record), { id: "7" })] });
    }) as unknown as typeof fetch;

    const found = await findWorkRecordByAppId(config, "token", "app-under-test", fetcher);
    expect(found?.itemId).toBe("7");
    expect(found?.record.appId).toBe("app-under-test");
  });

  it("returns null when no item matches the AppId", async () => {
    const fetcher = vi.fn(async () => jsonResponse({ value: [] })) as unknown as typeof fetch;
    expect(await findWorkRecordByAppId(config, "token", "missing", fetcher)).toBeNull();
  });

  it("resolves an update target by providerId and verifies the stored AppId matches", async () => {
    const record = testRecord({ metadata: { providerId: "9", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" } });
    const fetcher = vi.fn(async () => jsonResponse(graphItem(fieldsFor(record), { id: "9" }))) as unknown as typeof fetch;
    const resolved = await resolveWorkRecordItem(config, "token", record, fetcher);
    expect(resolved?.itemId).toBe("9");
  });

  it("falls back to AppId lookup when the providerId item is gone", async () => {
    const record = testRecord({ metadata: { providerId: "stale-id", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" } });
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ error: "not found" }, 404))
      .mockResolvedValueOnce(jsonResponse({ value: [graphItem(fieldsFor(record), { id: "fresh-id" })] })) as unknown as typeof fetch;
    const resolved = await resolveWorkRecordItem(config, "token", record, fetcher);
    expect(resolved?.itemId).toBe("fresh-id");
  });

  it("sends If-Match on update and surfaces a 412 as a conflict carrying the current record", async () => {
    const record = testRecord();
    const fetcher = vi.fn(async (_url: string, init: RequestInit) => {
      expect((init.headers as Record<string, string>)["If-Match"]).toBe('"stale"');
      return jsonResponse({ error: "precondition failed" }, 412);
    }) as unknown as typeof fetch;
    const conflictFetcher = vi
      .fn()
      .mockImplementationOnce(fetcher)
      .mockResolvedValueOnce(jsonResponse(graphItem(fieldsFor(record), { id: "5", eTag: '"2"' }))) as unknown as typeof fetch;

    await expect(updateWorkRecordItem(config, "token", "5", '"stale"', record, 2, conflictFetcher)).rejects.toMatchObject({
      kind: "conflict",
      current: expect.objectContaining({ appId: record.appId }),
    });
  });

  it("updates and reads back the incremented version on success", async () => {
    const record = testRecord();
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(new Response(null, { status: 204 }))
      .mockResolvedValueOnce(jsonResponse(graphItem(fieldsFor(record, 2), { id: "5" }))) as unknown as typeof fetch;

    const saved = await updateWorkRecordItem(config, "token", "5", '"1"', record, 2, fetcher);
    expect(saved.metadata.version).toBe(2);
  });

  it("maps a Graph 401 to an auth error", async () => {
    const fetcher = vi.fn(async () => jsonResponse({ error: "expired" }, 401)) as unknown as typeof fetch;
    await expect(listWorkRecords(config, "expired-token", fetcher)).rejects.toMatchObject({ kind: "auth" });
  });
});
