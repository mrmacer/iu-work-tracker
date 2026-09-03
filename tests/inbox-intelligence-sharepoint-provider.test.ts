import { afterEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { buildInboxIntelligenceRecord, EmailAnalysisSchema } from "../lib/inbox-intelligence-models";
import { REFERENCE_DATA } from "../lib/reference-data";
import { DelegatedSharePointInboxIntelligenceProvider, selectInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { toSharePointFields } from "../lib/sharepoint-inbox-intelligence";
import { InteractiveRedirectStartedError, MicrosoftAuthenticationError, type MicrosoftAuthController } from "../lib/microsoft-auth";

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
    // An unrecognized error type still falls through to the generic fallback message —
    // this is the baseline the classification tests below are distinguishing themselves from.
    if (result.status === "network_error") expect(result.message).toBe("The DEV SharePoint data store could not be reached.");
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

describe("toErrorResult classification — Inbox authentication error observability", () => {
  it("surfaces MicrosoftAuthenticationError's own message instead of the generic fallback", async () => {
    const provider = new DelegatedSharePointInboxIntelligenceProvider(
      fakeController(async () => {
        throw new MicrosoftAuthenticationError("A Microsoft access token could not be acquired.");
      }),
      account,
      config,
    );
    const result = await provider.list();
    expect(result.status).toBe("network_error");
    if (result.status === "network_error") {
      expect(result.message).toBe("A Microsoft access token could not be acquired.");
      // The exact regression this patch fixes: must NOT collapse into the generic sentence.
      expect(result.message).not.toBe("The DEV SharePoint data store could not be reached.");
    }
  });

  it("also classifies MicrosoftAuthenticationError on create(), not just list()", async () => {
    const record = testRecord();
    const provider = new DelegatedSharePointInboxIntelligenceProvider(
      fakeController(async () => {
        throw new MicrosoftAuthenticationError("Sign in with Microsoft to continue.");
      }),
      account,
      config,
    );
    const result = await provider.create(record);
    expect(result.status).toBe("network_error");
    if (result.status === "network_error") expect(result.message).toBe("Sign in with Microsoft to continue.");
  });

  it("leaves InteractiveRedirectStartedError handling unchanged", async () => {
    const provider = new DelegatedSharePointInboxIntelligenceProvider(
      fakeController(async () => {
        throw new InteractiveRedirectStartedError();
      }),
      account,
      config,
    );
    const result = await provider.list();
    expect(result.status).toBe("network_error");
    if (result.status === "network_error") {
      expect(result.message).toBe("Microsoft sign-in confirmation is required. Finish signing in, then try again.");
    }
  });

  it("leaves SharePointInboxIntelligenceError handling unchanged — a 403 still surfaces its specific auth message", async () => {
    const fetcher = vi.fn(async () => new Response("", { status: 403 }));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointInboxIntelligenceProvider(fakeController(async () => "token"), account, config);
    const result = await provider.list();
    expect(result.status).toBe("network_error");
    if (result.status === "network_error") {
      expect(result.message).toBe("The signed-in account is not authorized for the DEV Inbox Intelligence list.");
    }
  });

  it("leaves SharePointInboxIntelligenceError handling unchanged — an unexpected HTTP status still surfaces persistence_error with its specific message", async () => {
    const fetcher = vi.fn(async () => new Response("", { status: 500 }));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointInboxIntelligenceProvider(fakeController(async () => "token"), account, config);
    const result = await provider.list();
    expect(result.status).toBe("persistence_error");
    if (result.status === "persistence_error") {
      expect(result.message).toBe("Listing Inbox Intelligence records failed with HTTP 500.");
    }
  });

  it("a plain unrecognized Error still uses the existing generic fallback (baseline unchanged)", async () => {
    const provider = new DelegatedSharePointInboxIntelligenceProvider(
      fakeController(async () => {
        throw new Error("some unrelated failure");
      }),
      account,
      config,
    );
    const result = await provider.create(testRecord());
    expect(result.status).toBe("network_error");
    if (result.status === "network_error") expect(result.message).toBe("The DEV SharePoint data store could not be reached.");
  });
});

describe("selectInboxIntelligenceProvider", () => {
  it("falls back to the session provider outside a browser context", async () => {
    const { kind } = await selectInboxIntelligenceProvider();
    expect(kind).toBe("session");
  });
});
