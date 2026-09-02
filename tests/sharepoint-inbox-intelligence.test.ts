import { describe, expect, it, vi } from "vitest";
import {
  buildInboxIntelligenceRecord,
  EmailAnalysisSchema,
  type EmailAnalysis,
  type InboxIntelligenceRecord,
} from "../lib/inbox-intelligence-models";
import { REFERENCE_DATA } from "../lib/reference-data";
import {
  createInboxIntelligenceItem,
  findInboxIntelligenceByAppId,
  fromSharePointItem,
  listInboxIntelligenceItems,
  resolveInboxIntelligenceItem,
  SharePointInboxIntelligenceError,
  toSharePointFields,
  updateInboxIntelligenceItem,
  validateInboxIntelligenceRecord,
  validateInboxIntelligenceSharePointLimits,
  type SharePointInboxIntelligenceConfig,
} from "../lib/sharepoint-inbox-intelligence";

const config: SharePointInboxIntelligenceConfig = {
  siteId: "siu29.sharepoint.com,site,web",
  inboxIntelligenceListId: "inbox-intelligence-list-id",
};

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

function analysis(overrides: Partial<EmailAnalysis> = {}): EmailAnalysis {
  return EmailAnalysisSchema.parse({
    summary: "Test summary",
    priority: "high",
    needsAttention: true,
    actionItems: [{ action: "Reply", dueDate: "2026-09-05", owner: "me" }],
    followUp: "Check back Friday",
    people: ["Pat Alvarez"],
    organizations: [],
    districts: [],
    projects: ["Robotics Club Expansion"],
    tags: ["grant"],
    suggestedWorkType: null,
    suggestedWorkRecord: { title: "DEV INBOX TEST", description: "Synthetic record for tests" },
    ...overrides,
  });
}

function testRecord(overrides: Partial<InboxIntelligenceRecord> = {}): InboxIntelligenceRecord {
  return { ...buildInboxIntelligenceRecord(analysis(), "short excerpt", REFERENCE_DATA, "2026-08-29T12:00:00.000Z"), ...overrides };
}

function graphItem(record: InboxIntelligenceRecord, version: number, overrides: Partial<{ id: string; eTag: string }> = {}) {
  return {
    id: overrides.id ?? "1",
    eTag: overrides.eTag ?? '"1"',
    createdDateTime: "2026-08-29T12:00:00Z",
    lastModifiedDateTime: "2026-08-29T12:00:00Z",
    fields: toSharePointFields(record, version),
  };
}

describe("Inbox Intelligence field mapping", () => {
  it("round-trips a record without mutating relationship JSON, and carries no raw-email field", () => {
    const record = testRecord();
    const fields = toSharePointFields(record, 1);
    expect(JSON.parse(String(fields.ProjectsJson))).toEqual(["Robotics Club Expansion"]);
    expect(Object.keys(fields)).not.toContain("RawEmail");
    expect(Object.keys(fields)).not.toContain("EmailBody");

    const mapped = fromSharePointItem(graphItem(record, 1));
    expect(mapped.appId).toBe(record.appId);
    expect(mapped.analysis.suggestedWorkRecord.title).toBe(record.analysis.suggestedWorkRecord.title);
    expect(mapped.analysis.actionItems).toEqual(record.analysis.actionItems);
    expect(mapped.sourceExcerpt).toBe("short excerpt");
    expect(mapped.metadata).toEqual({ providerId: "1", version: 1, createdAt: "2026-08-29T12:00:00Z", modifiedAt: "2026-08-29T12:00:00Z", syncState: "saved" });
  });

  it("round-trips matchedContactIds through MatchedContactIdsJson (Patch 8D)", () => {
    const record = testRecord({ matchedContactIds: ["contact-north-valley-lead", "contact-futureworks"] });
    const fields = toSharePointFields(record, 1);
    expect(JSON.parse(String(fields.MatchedContactIdsJson))).toEqual(["contact-north-valley-lead", "contact-futureworks"]);

    const mapped = fromSharePointItem(graphItem(record, 1));
    expect(mapped.matchedContactIds).toEqual(["contact-north-valley-lead", "contact-futureworks"]);
  });

  it("reads matchedContactIds as [] for an item saved before the MatchedContactIdsJson column existed, never a parse error", () => {
    const item = graphItem(testRecord(), 1);
    delete (item.fields as Record<string, unknown>).MatchedContactIdsJson;
    expect(() => fromSharePointItem(item)).not.toThrow();
    expect(fromSharePointItem(item).matchedContactIds).toEqual([]);
  });

  it("rejects an unsupported SchemaVersion instead of silently coercing it", () => {
    const fields = toSharePointFields(testRecord(), 1);
    fields.SchemaVersion = 99;
    expect(() => fromSharePointItem(graphItem(testRecord(), 1) && { ...graphItem(testRecord(), 1), fields })).toThrow(
      SharePointInboxIntelligenceError,
    );
  });

  it("rejects malformed action-item JSON instead of silently discarding it", () => {
    const item = graphItem(testRecord(), 1);
    item.fields.ActionItemsJson = "{not json";
    expect(() => fromSharePointItem(item)).toThrow(/malformed JSON/);
  });

  it("flags an oversized field via the SharePoint-compatible limit check without truncating", () => {
    const oversized = testRecord({ sourceExcerpt: "x".repeat(501) });
    const issues = validateInboxIntelligenceSharePointLimits(oversized);
    expect(issues).toContainEqual(expect.objectContaining({ path: "sourceExcerpt", code: "sharepoint_text_limit" }));
  });

  it("accepts a record within limits and with a consistent status/resolvedAt pairing", () => {
    expect(validateInboxIntelligenceRecord(testRecord())).toEqual([]);
  });

  it("rejects resolved without resolvedAt, and non-resolved with a resolvedAt set", () => {
    expect(validateInboxIntelligenceRecord(testRecord({ status: "resolved", resolvedAt: null }))).toContainEqual(
      expect.objectContaining({ path: "resolvedAt", code: "required" }),
    );
    expect(validateInboxIntelligenceRecord(testRecord({ status: "open", resolvedAt: "2026-08-29T12:00:00.000Z" }))).toContainEqual(
      expect.objectContaining({ path: "resolvedAt", code: "unexpected_value" }),
    );
  });

  it("re-validates the analysis against the same schema the AI pipeline enforces", () => {
    const broken = testRecord({ analysis: { ...analysis(), priority: "urgent" as unknown as EmailAnalysis["priority"] } });
    expect(validateInboxIntelligenceRecord(broken)).toContainEqual(expect.objectContaining({ path: "analysis" }));
  });
});

describe("Inbox Intelligence Graph operations", () => {
  it("follows every @odata.nextLink when listing", async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({ value: [graphItem(testRecord(), 1)], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next-page" }),
      )
      .mockResolvedValueOnce(jsonResponse({ value: [graphItem(testRecord({ appId: "second" }), 1, { id: "2" })] })) as unknown as typeof fetch;

    const records = await listInboxIntelligenceItems(config, "token", fetcher);
    expect(records).toHaveLength(2);
    expect(fetcher).toHaveBeenCalledTimes(2);
  });

  it("creates an item and reads it back for SharePoint-owned id/timestamps", async () => {
    const record = testRecord();
    const fetcherMock = vi.fn().mockResolvedValueOnce(jsonResponse({ id: "42" })).mockResolvedValueOnce(jsonResponse(graphItem(record, 1, { id: "42" })));
    const created = await createInboxIntelligenceItem(config, "token", record, fetcherMock as unknown as typeof fetch);
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
    const found = await findInboxIntelligenceByAppId(config, "token", record.appId, fetcher);
    expect(found?.itemId).toBe("7");
  });

  it("resolves an update target by providerId and falls back to AppId lookup when it's gone", async () => {
    const record = testRecord({ metadata: { providerId: "stale-id", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" } });
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ error: "not found" }, 404))
      .mockResolvedValueOnce(jsonResponse({ value: [graphItem(record, 1, { id: "fresh-id" })] })) as unknown as typeof fetch;
    const resolved = await resolveInboxIntelligenceItem(config, "token", record, fetcher);
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

    await expect(updateInboxIntelligenceItem(config, "token", "5", '"stale"', record, 2, fetcher)).rejects.toMatchObject({
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
    const saved = await updateInboxIntelligenceItem(config, "token", "5", '"1"', record, 2, fetcher);
    expect(saved.metadata.version).toBe(2);
  });

  it("maps a Graph 401 to an auth error", async () => {
    const fetcher = vi.fn(async () => jsonResponse({ error: "expired" }, 401)) as unknown as typeof fetch;
    await expect(listInboxIntelligenceItems(config, "expired-token", fetcher)).rejects.toMatchObject({ kind: "auth" });
  });
});
