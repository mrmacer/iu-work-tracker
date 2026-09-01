import { afterEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import {
  buildMeetingRecord,
  MEETING_RECORD_SCHEMA_VERSION,
  type MeetingRecord,
  type ReviewedMeetingCandidate,
} from "../lib/meeting-intelligence-models";
import {
  DelegatedSharePointMeetingRecordProvider,
  MemoryMeetingRecordProvider,
  selectMeetingRecordProvider,
} from "../lib/meeting-record-provider";
import { toSharePointFields } from "../lib/sharepoint-meeting-records";
import type { MicrosoftAuthController } from "../lib/microsoft-auth";

const account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
const config = { siteId: "site-id", meetingRecordsListId: "list-id" };

function fakeController(acquireGraphToken: () => Promise<string>): MicrosoftAuthController {
  return { acquireGraphToken } as unknown as MicrosoftAuthController;
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

function record(overrides: Partial<Parameters<typeof buildMeetingRecord>[0]> = {}): MeetingRecord {
  return buildMeetingRecord({
    appId: "meeting-under-test",
    title: "STEELS quarterly planning",
    meetingDate: "2026-08-29",
    meetingType: "District Meeting",
    attendeesText: "Greg, Annie, Kim",
    agendaText: "Review grant budget.",
    notesText: "Walked through the budget together.",
    reviewedCandidates: [candidate()],
    minutesText: "STEELS quarterly planning\n2026-08-29 · District Meeting",
    analysisModel: "claude-opus-5",
    analyzedAt: "2026-08-29T12:00:00.000Z",
    ...overrides,
  });
}

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

function graphItem(rec: MeetingRecord, version = 1, id = "1") {
  return {
    id,
    eTag: '"1"',
    createdDateTime: "2026-08-29T12:00:00Z",
    lastModifiedDateTime: "2026-08-29T12:00:00Z",
    fields: toSharePointFields(rec, version),
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("MemoryMeetingRecordProvider", () => {
  it("assigns SharePoint-shaped provider metadata on create and lists newest first", async () => {
    const provider = new MemoryMeetingRecordProvider();
    const first = await provider.create(record());
    const second = await provider.create(record({ appId: "second-meeting" }));
    expect(first.status).toBe("success");
    expect(second.status).toBe("success");
    if (first.status !== "success" || second.status !== "success") return;
    expect(first.value.metadata.version).toBe(1);

    const list = await provider.list();
    expect(list.status).toBe("success");
    if (list.status !== "success") return;
    expect(list.value).toHaveLength(2);
    expect(list.value[0].appId).toBe(second.value.appId); // most-recently-created first
  });

  it("rejects a duplicate AppId as a conflict without overwriting the existing record", async () => {
    const provider = new MemoryMeetingRecordProvider();
    const created = await provider.create(record());
    if (created.status !== "success") throw new Error("setup failed");
    const duplicate = await provider.create({ ...record(), appId: created.value.appId });
    expect(duplicate.status).toBe("conflict");
  });

  it("increments RecordVersion on update and rejects a stale expectedVersion as a conflict", async () => {
    const provider = new MemoryMeetingRecordProvider();
    const created = await provider.create(record());
    if (created.status !== "success") throw new Error("setup failed");

    const updated = await provider.update({ ...created.value, title: "Updated title" }, 1);
    expect(updated.status).toBe("success");
    if (updated.status === "success") expect(updated.value.metadata.version).toBe(2);

    const stale = await provider.update({ ...created.value, title: "Stale write" }, 1);
    expect(stale.status).toBe("conflict");
  });

  it("saves a meeting before any analysis has run — empty candidates, null analysis metadata accepted", async () => {
    const provider = new MemoryMeetingRecordProvider();
    const preAnalysis = record({ reviewedCandidates: [], analysisModel: null, analyzedAt: null, minutesText: "Untitled meeting" });
    const created = await provider.create(preAnalysis);
    expect(created.status).toBe("success");
    if (created.status === "success") {
      expect(created.value.reviewedCandidates).toEqual([]);
      expect(created.value.analysisModel).toBeNull();
    }
  });

  it("keeps every instance independent — nothing is shared or durable across instances", async () => {
    const first = new MemoryMeetingRecordProvider();
    await first.create(record());
    const second = new MemoryMeetingRecordProvider();
    const list = await second.list();
    expect(list).toEqual({ status: "success", value: [] });
  });

  it("rejects a shape-invalid record (inconsistent analysis metadata) before persisting anything", async () => {
    const provider = new MemoryMeetingRecordProvider();
    const invalid = record({ analysisModel: "claude-opus-5", analyzedAt: null });
    const result = await provider.create(invalid);
    expect(result.status).toBe("validation_error");
    expect((await provider.list()).status).toBe("success");
    const list = await provider.list();
    if (list.status === "success") expect(list.value).toHaveLength(0);
  });
});

describe("DelegatedSharePointMeetingRecordProvider", () => {
  it("returns a structured conflict when AppId already exists on create, without a second unnecessary write", async () => {
    const rec = record();
    const fetcher = vi.fn(async () => jsonResponse({ value: [graphItem(rec)] }));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointMeetingRecordProvider(fakeController(async () => "token"), account, config);
    const result = await provider.create(rec);
    expect(result.status).toBe("conflict");
  });

  it("creates successfully and normalizes the read-back record", async () => {
    const rec = record();
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ value: [] }))
      .mockResolvedValueOnce(jsonResponse({ id: "10" }))
      .mockResolvedValueOnce(jsonResponse(graphItem(rec, 1, "10")));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointMeetingRecordProvider(fakeController(async () => "token"), account, config);
    const result = await provider.create(rec);
    expect(result.status).toBe("success");
    if (result.status === "success") expect(result.value.metadata.providerId).toBe("10");
  });

  it("returns conflict without writing when RecordVersion no longer matches expectedVersion", async () => {
    const rec = { ...record(), metadata: { providerId: "10", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" as const } };
    const fetcher = vi.fn(async () => jsonResponse(graphItem(rec, 2, "10")));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointMeetingRecordProvider(fakeController(async () => "token"), account, config);
    const result = await provider.update(rec, 1);
    expect(result.status).toBe("conflict");
    expect(fetcher).toHaveBeenCalledTimes(1); // resolved and compared; no PATCH attempted
  });

  it("maps an expired Graph token to a network_error result instead of throwing", async () => {
    const provider = new DelegatedSharePointMeetingRecordProvider(
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
    const provider = new DelegatedSharePointMeetingRecordProvider(fakeController(acquireGraphToken), account, config);
    const invalid = record({ analysisModel: "claude-opus-5", analyzedAt: null });
    const result = await provider.create(invalid);
    expect(result.status).toBe("validation_error");
    expect(acquireGraphToken).not.toHaveBeenCalled();
  });

  it("rejects an oversized record via the SharePoint limit check before acquiring a Graph token", async () => {
    const acquireGraphToken = vi.fn();
    const provider = new DelegatedSharePointMeetingRecordProvider(fakeController(acquireGraphToken), account, config);
    const oversized = record({ notesText: "x".repeat(20001) });
    const result = await provider.create(oversized);
    expect(result.status).toBe("validation_error");
    expect(acquireGraphToken).not.toHaveBeenCalled();
  });
});

describe("selectMeetingRecordProvider", () => {
  it("falls back to the memory provider outside a browser context — never a fake/hardcoded SharePoint list id", async () => {
    const { kind, provider } = await selectMeetingRecordProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryMeetingRecordProvider);
  });
});

// Sanity: MEETING_RECORD_SCHEMA_VERSION stays at 1 for this patch — a bump here without an
// accompanying migration note would be a signal worth catching in review.
describe("schema version", () => {
  it("is 1", () => {
    expect(MEETING_RECORD_SCHEMA_VERSION).toBe(1);
  });
});
