import { describe, expect, it, vi } from "vitest";
import { MEETING_RECORD_SCHEMA_VERSION, type MeetingRecord, type ReviewedMeetingCandidate } from "../lib/meeting-intelligence-models";
import {
  createMeetingRecordItem,
  findMeetingRecordByAppId,
  fromSharePointItem,
  listMeetingRecordItems,
  resolveMeetingRecordItem,
  SharePointMeetingRecordsError,
  toSharePointFields,
  updateMeetingRecordItem,
  validateMeetingRecordSharePointLimits,
  type SharePointMeetingRecordConfig,
} from "../lib/sharepoint-meeting-records";

const config: SharePointMeetingRecordConfig = {
  siteId: "siu29.sharepoint.com,site,web",
  meetingRecordsListId: "meeting-records-list-id",
};

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

function candidate(overrides: Partial<ReviewedMeetingCandidate> = {}): ReviewedMeetingCandidate {
  return {
    type: "ACTION",
    title: "Call the district about the venue",
    detail: "Confirm venue details.",
    sourceExcerpt: "Annie will call the district about the venue by Friday.",
    ownerText: "Annie",
    dueText: "Friday",
    durationText: null,
    selected: true,
    ...overrides,
  };
}

function testRecord(overrides: Partial<MeetingRecord> = {}): MeetingRecord {
  return {
    appId: "meeting-under-test",
    schemaVersion: MEETING_RECORD_SCHEMA_VERSION,
    title: "STEELS quarterly planning",
    meetingDate: "2026-08-29",
    meetingType: "District Meeting",
    attendeesText: "Greg, Annie, Kim",
    agendaText: "Review grant budget. Discuss fall meeting timing.",
    notesText: "Walked through the budget together. Annie will call the district.",
    reviewedCandidates: [candidate()],
    minutesText: "STEELS quarterly planning\n2026-08-29 · District Meeting",
    analysisModel: "claude-opus-5",
    analyzedAt: "2026-08-29T12:00:00.000Z",
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
    ...overrides,
  };
}

function graphItem(record: MeetingRecord, version: number, overrides: Partial<{ id: string; eTag: string }> = {}) {
  return {
    id: overrides.id ?? "1",
    eTag: overrides.eTag ?? '"1"',
    createdDateTime: "2026-08-29T12:00:00Z",
    lastModifiedDateTime: "2026-08-29T12:00:00Z",
    fields: toSharePointFields(record, version),
  };
}

describe("Meeting Record field mapping", () => {
  it("round-trips a record, storing reviewedCandidates as a single JSON blob field", () => {
    const record = testRecord();
    const fields = toSharePointFields(record, 1);
    expect(JSON.parse(String(fields.IntelligenceJson))).toEqual(record.reviewedCandidates);
    // No per-candidate-type columns — one JSON field carries every candidate type.
    expect(Object.keys(fields)).not.toContain("ActionItemsJson");
    expect(Object.keys(fields)).not.toContain("DecisionsJson");

    const mapped = fromSharePointItem(graphItem(record, 1));
    expect(mapped.appId).toBe(record.appId);
    expect(mapped.title).toBe(record.title);
    expect(mapped.meetingDate).toBe(record.meetingDate);
    expect(mapped.reviewedCandidates).toEqual(record.reviewedCandidates);
    expect(mapped.minutesText).toBe(record.minutesText);
    expect(mapped.analysisModel).toBe(record.analysisModel);
    expect(mapped.analyzedAt).toBe(record.analyzedAt);
    expect(mapped.metadata).toEqual({ providerId: "1", version: 1, createdAt: "2026-08-29T12:00:00Z", modifiedAt: "2026-08-29T12:00:00Z", syncState: "saved" });
  });

  it("supports saving before any analysis has run — empty candidates, null analysis metadata", () => {
    const record = testRecord({ reviewedCandidates: [], analysisModel: null, analyzedAt: null, minutesText: "Untitled meeting" });
    const mapped = fromSharePointItem(graphItem(record, 1));
    expect(mapped.reviewedCandidates).toEqual([]);
    expect(mapped.analysisModel).toBeNull();
    expect(mapped.analyzedAt).toBeNull();
  });

  it("rejects an unsupported SchemaVersion instead of silently coercing it", () => {
    const item = graphItem(testRecord(), 1);
    item.fields.SchemaVersion = 99;
    expect(() => fromSharePointItem(item)).toThrow(SharePointMeetingRecordsError);
  });

  it("rejects malformed IntelligenceJson instead of silently discarding it", () => {
    const item = graphItem(testRecord(), 1);
    item.fields.IntelligenceJson = "{not json";
    expect(() => fromSharePointItem(item)).toThrow(/malformed JSON/);
  });

  it("rejects IntelligenceJson that no longer matches the reviewed-candidate shape", () => {
    const item = graphItem(testRecord(), 1);
    item.fields.IntelligenceJson = JSON.stringify([{ type: "ACTION" }]); // missing required fields
    expect(() => fromSharePointItem(item)).toThrow(/expected reviewed-candidate shape/);
  });

  it("flags an oversized text field via the SharePoint-compatible limit check, without truncating", () => {
    const oversized = testRecord({ notesText: "x".repeat(20001) });
    const issues = validateMeetingRecordSharePointLimits(oversized);
    expect(issues).toContainEqual(expect.objectContaining({ path: "notesText", code: "sharepoint_text_limit" }));
  });

  it("flags an oversized reviewedCandidates JSON blob via the limit check", () => {
    const oversized = testRecord({ reviewedCandidates: Array.from({ length: 60 }, () => candidate({ detail: "x".repeat(1000) })) });
    const issues = validateMeetingRecordSharePointLimits(oversized);
    expect(issues).toContainEqual(expect.objectContaining({ path: "reviewedCandidates", code: "sharepoint_text_limit" }));
  });

  it("accepts a record within every limit", () => {
    expect(validateMeetingRecordSharePointLimits(testRecord())).toEqual([]);
  });
});

describe("Meeting Record Graph operations", () => {
  it("follows every @odata.nextLink when listing", async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({ value: [graphItem(testRecord(), 1)], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next-page" }),
      )
      .mockResolvedValueOnce(jsonResponse({ value: [graphItem(testRecord({ appId: "second" }), 1, { id: "2" })] })) as unknown as typeof fetch;

    const records = await listMeetingRecordItems(config, "token", fetcher);
    expect(records).toHaveLength(2);
    expect(fetcher).toHaveBeenCalledTimes(2);
  });

  it("creates an item, sends RecordVersion 1, and reads it back for SharePoint-owned id/timestamps", async () => {
    const record = testRecord();
    const fetcherMock = vi.fn().mockResolvedValueOnce(jsonResponse({ id: "42" })).mockResolvedValueOnce(jsonResponse(graphItem(record, 1, { id: "42" })));
    const created = await createMeetingRecordItem(config, "token", record, fetcherMock as unknown as typeof fetch);
    expect(created.metadata.providerId).toBe("42");
    expect(created.metadata.version).toBe(1);
    const [, createInit] = fetcherMock.mock.calls[0] as [string, RequestInit];
    expect(createInit.method).toBe("POST");
    expect(JSON.parse(String(createInit.body)).fields.RecordVersion).toBe(1);
  });

  it("looks up an existing item by indexed AppId", async () => {
    const record = testRecord();
    const fetcher = vi.fn(async (url: string) => {
      expect(url).toContain(encodeURIComponent(`fields/AppId eq '${record.appId}'`));
      return jsonResponse({ value: [graphItem(record, 1, { id: "7" })] });
    }) as unknown as typeof fetch;
    const found = await findMeetingRecordByAppId(config, "token", record.appId, fetcher);
    expect(found?.itemId).toBe("7");
  });

  it("resolves an update target by providerId and falls back to AppId lookup when it's gone", async () => {
    const record = testRecord({ metadata: { providerId: "stale-id", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" } });
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ error: "not found" }, 404))
      .mockResolvedValueOnce(jsonResponse({ value: [graphItem(record, 1, { id: "fresh-id" })] })) as unknown as typeof fetch;
    const resolved = await resolveMeetingRecordItem(config, "token", record, fetcher);
    expect(resolved?.itemId).toBe("fresh-id");
  });

  it("sends If-Match on update and surfaces a 412 as a conflict carrying the current record", async () => {
    const record = testRecord();
    const fetcher = vi
      .fn()
      .mockImplementationOnce(async (_url: string, init: RequestInit) => {
        expect((init.headers as Record<string, string>)["If-Match"]).toBe('"stale"');
        return jsonResponse({ error: "precondition failed" }, 412);
      })
      .mockResolvedValueOnce(jsonResponse(graphItem(record, 2, { id: "5", eTag: '"2"' }))) as unknown as typeof fetch;

    await expect(updateMeetingRecordItem(config, "token", "5", '"stale"', record, 2, fetcher)).rejects.toMatchObject({
      kind: "conflict",
      current: expect.objectContaining({ appId: record.appId }),
    });
  });

  it("updates and reads back the incremented version on success", async () => {
    const record = testRecord();
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(new Response(null, { status: 204 }))
      .mockResolvedValueOnce(jsonResponse(graphItem(record, 2, { id: "5" }))) as unknown as typeof fetch;
    const saved = await updateMeetingRecordItem(config, "token", "5", '"1"', record, 2, fetcher);
    expect(saved.metadata.version).toBe(2);
  });

  it("maps a Graph 401 to an auth error", async () => {
    const fetcher = vi.fn(async () => jsonResponse({ error: "expired" }, 401)) as unknown as typeof fetch;
    await expect(listMeetingRecordItems(config, "expired-token", fetcher)).rejects.toMatchObject({ kind: "auth" });
  });
});
