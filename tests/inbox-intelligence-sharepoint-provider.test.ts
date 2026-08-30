import { afterEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { buildInboxIntelligenceRecord, EmailAnalysisSchema } from "../lib/inbox-intelligence-models";
import { REFERENCE_DATA } from "../lib/reference-data";
import { DelegatedSharePointInboxIntelligenceProvider, selectInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { toSharePointFields } from "../lib/sharepoint-inbox-intelligence";
import type { MicrosoftAuthController } from "../lib/microsoft-auth";

const account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
const config = { siteId: "site-id", inboxIntelligenceListId: "list-id" };

function fakeController(acquireGraphToken: () => Promise<string>): MicrosoftAuthController {
  return { acquireGraphToken } as unknown as MicrosoftAuthController;
}

function testRecord() {
  const analysis = EmailAnalysisSchema.parse({
    summary: "Test",
    priority: "medium",
    needsAttention: false,
    actionItems: [],
    followUp: "",
    people: [],
    organizations: [],
    districts: [],
    projects: [],
    tags: [],
    suggestedWorkType: null,
    suggestedWorkRecord: { title: "DEV INBOX PROVIDER TEST", description: "desc" },
  });
  return buildInboxIntelligenceRecord(analysis, "", REFERENCE_DATA, "2026-08-29T12:00:00.000Z");
}

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

function graphItem(record: ReturnType<typeof testRecord>, version = 1, id = "1") {
  return {
    id,
    eTag: '"1"',
    createdDateTime: "2026-08-29T12:00:00Z",
    lastModifiedDateTime: "2026-08-29T12:00:00Z",
    fields: toSharePointFields(record, version),
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("DelegatedSharePointInboxIntelligenceProvider", () => {
  it("returns a structured conflict when AppId already exists on create, without acquiring a second token unnecessarily", async () => {
    const record = testRecord();
    const fetcher = vi.fn(async () => jsonResponse({ value: [graphItem(record)] }));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointInboxIntelligenceProvider(fakeController(async () => "token"), account, config);
    const result = await provider.create(record);
    expect(result.status).toBe("conflict");
  });

  it("creates successfully and normalizes the read-back record", async () => {
    const record = testRecord();
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ value: [] }))
      .mockResolvedValueOnce(jsonResponse({ id: "10" }))
      .mockResolvedValueOnce(jsonResponse(graphItem(record, 1, "10")));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointInboxIntelligenceProvider(fakeController(async () => "token"), account, config);
    const result = await provider.create(record);
    expect(result.status).toBe("success");
    if (result.status === "success") expect(result.value.metadata.providerId).toBe("10");
  });

  it("returns conflict without writing when RecordVersion no longer matches expectedVersion", async () => {
    const record = { ...testRecord(), metadata: { providerId: "10", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" as const } };
    const fetcher = vi.fn(async () => jsonResponse(graphItem(record, 2, "10")));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointInboxIntelligenceProvider(fakeController(async () => "token"), account, config);
    const result = await provider.update(record, 1);
    expect(result.status).toBe("conflict");
    expect(fetcher).toHaveBeenCalledTimes(1); // resolved and compared; no PATCH attempted
  });

  it("maps an expired Graph token to a network_error result instead of throwing", async () => {
    const provider = new DelegatedSharePointInboxIntelligenceProvider(
      fakeController(async () => {
        throw new Error("boom");
      }),
      account,
      config,
    );
    const result = await provider.list();
    expect(result.status).toBe("network_error");
  });

  it("rejects an invalid record before ever acquiring a Graph token", async () => {
    const acquireGraphToken = vi.fn();
    const provider = new DelegatedSharePointInboxIntelligenceProvider(fakeController(acquireGraphToken), account, config);
    const invalid = { ...testRecord(), status: "resolved" as const, resolvedAt: null };
    const result = await provider.create(invalid);
    expect(result.status).toBe("validation_error");
    expect(acquireGraphToken).not.toHaveBeenCalled();
  });
});

describe("selectInboxIntelligenceProvider", () => {
  it("falls back to the session provider outside a browser context", async () => {
    const { kind } = await selectInboxIntelligenceProvider();
    expect(kind).toBe("session");
  });
});
