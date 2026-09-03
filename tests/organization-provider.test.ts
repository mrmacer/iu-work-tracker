import { afterEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import type { Organization } from "../lib/models";
import {
  buildOrganizationDraft,
  DelegatedSharePointOrganizationProvider,
  MemoryOrganizationProvider,
  normalizeOrganizationName,
  ORGANIZATION_TYPES,
  selectOrganizationProvider,
  validateOrganizationShape,
} from "../lib/organization-provider";
import { toSharePointFields } from "../lib/sharepoint-organizations";
import type { MicrosoftAuthController } from "../lib/microsoft-auth";

const account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
const config = { siteId: "site-id", organizationsListId: "list-id" };

function fakeController(acquireGraphToken: () => Promise<string>): MicrosoftAuthController {
  return { acquireGraphToken } as unknown as MicrosoftAuthController;
}

function organization(overrides: Partial<Parameters<typeof buildOrganizationDraft>[0]> = {}): Organization {
  return buildOrganizationDraft({
    appId: "org-under-test",
    name: "Test Valley SD",
    type: "district",
    ...overrides,
  });
}

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

function graphItem(o: Organization, version = 1, id = "1") {
  return {
    id,
    eTag: '"1"',
    createdDateTime: "2026-08-29T12:00:00Z",
    lastModifiedDateTime: "2026-08-29T12:00:00Z",
    fields: toSharePointFields(o, version),
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("Organization model", () => {
  it("builds a fresh, unsaved draft with metadata.version 0 — routes create() first", () => {
    const o = organization();
    expect(o.metadata?.version).toBe(0);
  });

  it("requires a non-empty name", () => {
    const issues = validateOrganizationShape(organization({ name: "  " }));
    expect(issues).toContainEqual(expect.objectContaining({ path: "name", code: "required" }));
  });

  it("accepts exactly the three existing types: district, partner, iu — no expanded taxonomy", () => {
    expect(ORGANIZATION_TYPES).toEqual(["district", "partner", "iu"]);
    for (const type of ORGANIZATION_TYPES) {
      expect(validateOrganizationShape(organization({ type }))).toEqual([]);
    }
  });

  it("rejects a type outside the allowed set", () => {
    const issues = validateOrganizationShape(organization({ type: "nonprofit" as Organization["type"] }));
    expect(issues).toContainEqual(expect.objectContaining({ path: "type", code: "invalid_type" }));
  });
});

describe("normalizeOrganizationName", () => {
  it("trims, collapses internal whitespace, and lowercases — no fuzzy matching", () => {
    expect(normalizeOrganizationName("  North   Valley SD  ")).toBe("north valley sd");
    expect(normalizeOrganizationName("North Valley SD")).toBe(normalizeOrganizationName("  NORTH  VALLEY   SD "));
  });

  it("does not treat similar-but-different names as equal", () => {
    expect(normalizeOrganizationName("North Valley SD")).not.toBe(normalizeOrganizationName("North Valley School District"));
  });
});

describe("MemoryOrganizationProvider", () => {
  it("assigns SharePoint-shaped provider metadata on create and lists newest first", async () => {
    const provider = new MemoryOrganizationProvider();
    const first = await provider.create(organization());
    const second = await provider.create(organization({ appId: "second-org" }));
    expect(first.status).toBe("success");
    expect(second.status).toBe("success");
    if (first.status !== "success" || second.status !== "success") return;
    expect(first.value.metadata?.version).toBe(1);

    const list = await provider.list();
    expect(list.status).toBe("success");
    if (list.status !== "success") return;
    expect(list.value).toHaveLength(2);
    expect(list.value[0].appId).toBe(second.value.appId);
  });

  it("assigns a UUID-shaped appId when built via crypto.randomUUID(), not derived from the name", async () => {
    const draft = buildOrganizationDraft({ appId: crypto.randomUUID(), name: "Any Org Name", type: "partner" });
    expect(draft.appId).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/);
  });

  it("rejects a duplicate AppId as a conflict without overwriting the existing organization", async () => {
    const provider = new MemoryOrganizationProvider();
    const created = await provider.create(organization());
    if (created.status !== "success") throw new Error("setup failed");
    const duplicate = await provider.create({ ...organization(), appId: created.value.appId });
    expect(duplicate.status).toBe("conflict");
  });

  it("increments RecordVersion on update and rejects a stale expectedVersion as a conflict", async () => {
    const provider = new MemoryOrganizationProvider();
    const created = await provider.create(organization());
    if (created.status !== "success") throw new Error("setup failed");

    const updated = await provider.update({ ...created.value, type: "partner" }, 1);
    expect(updated.status).toBe("success");
    if (updated.status === "success") expect(updated.value.metadata?.version).toBe(2);

    const stale = await provider.update({ ...created.value, type: "iu" }, 1);
    expect(stale.status).toBe("conflict");
  });

  it("keeps every instance independent — nothing is shared or durable across instances", async () => {
    const first = new MemoryOrganizationProvider();
    await first.create(organization());
    const second = new MemoryOrganizationProvider();
    expect(await second.list()).toEqual({ status: "success", value: [] });
  });

  it("has no delete method", () => {
    expect((new MemoryOrganizationProvider() as unknown as { delete?: unknown }).delete).toBeUndefined();
  });
});

describe("DelegatedSharePointOrganizationProvider", () => {
  it("returns a structured conflict when AppId already exists on create", async () => {
    const o = organization();
    const fetcher = vi.fn(async () => jsonResponse({ value: [graphItem(o)] }));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointOrganizationProvider(fakeController(async () => "token"), account, config);
    const result = await provider.create(o);
    expect(result.status).toBe("conflict");
  });

  it("creates successfully with RecordVersion 1 and normalizes the read-back organization", async () => {
    const o = organization();
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ value: [] }))
      .mockResolvedValueOnce(jsonResponse({ id: "10" }))
      .mockResolvedValueOnce(jsonResponse(graphItem(o, 1, "10")));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointOrganizationProvider(fakeController(async () => "token"), account, config);
    const result = await provider.create(o);
    expect(result.status).toBe("success");
    if (result.status === "success") {
      expect(result.value.metadata?.providerId).toBe("10");
      expect(result.value.metadata?.version).toBe(1);
    }
    // Confirms the create POST body's RecordVersion was 1, per Patch 8E's non-negotiable rule.
    const createCall = fetcher.mock.calls[1];
    const body = JSON.parse(String((createCall[1] as RequestInit).body));
    expect(body.fields.RecordVersion).toBe(1);
  });

  it("increments RecordVersion on update", async () => {
    const o = { ...organization(), metadata: { providerId: "10", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" as const } };
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse(graphItem(o, 1, "10"))) // resolveOrganizationItem's getOrganizationItem
      .mockResolvedValueOnce(new Response(null, { status: 200 })) // PATCH
      .mockResolvedValueOnce(jsonResponse(graphItem(o, 2, "10"))); // re-read after update
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointOrganizationProvider(fakeController(async () => "token"), account, config);
    const result = await provider.update(o, 1);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.value.metadata?.version).toBe(2);
    const patchCall = fetcher.mock.calls[1];
    expect((patchCall[1] as RequestInit).method).toBe("PATCH");
    expect((patchCall[1] as RequestInit).headers).toMatchObject({ "If-Match": '"1"' });
  });

  it("returns conflict without writing when RecordVersion no longer matches expectedVersion", async () => {
    const o = { ...organization(), metadata: { providerId: "10", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" as const } };
    const fetcher = vi.fn(async () => jsonResponse(graphItem(o, 2, "10")));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointOrganizationProvider(fakeController(async () => "token"), account, config);
    const result = await provider.update(o, 1);
    expect(result.status).toBe("conflict");
    expect(fetcher).toHaveBeenCalledTimes(1); // never attempted the write
  });

  it("surfaces a 412 (stale ETag) as a conflict carrying the current organization, never auto-merging or overwriting", async () => {
    const o = { ...organization(), metadata: { providerId: "10", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" as const } };
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse(graphItem(o, 1, "10"))) // resolve
      .mockResolvedValueOnce(new Response(null, { status: 412 })) // stale PATCH
      .mockResolvedValueOnce(jsonResponse(graphItem(o, 3, "10"))); // re-read to surface current
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointOrganizationProvider(fakeController(async () => "token"), account, config);
    const result = await provider.update(o, 1);
    expect(result.status).toBe("conflict");
    if (result.status === "conflict") expect(result.current.metadata?.version).toBe(3);
  });

  it("maps an expired Graph token to a network_error result instead of throwing", async () => {
    const provider = new DelegatedSharePointOrganizationProvider(
      fakeController(async () => {
        throw new Error("boom");
      }),
      account,
      config,
    );
    const result = await provider.list();
    expect(result.status).toBe("network_error");
  });

  it("rejects an invalid organization before ever acquiring a Graph token", async () => {
    const acquireGraphToken = vi.fn();
    const provider = new DelegatedSharePointOrganizationProvider(fakeController(acquireGraphToken), account, config);
    const result = await provider.create(organization({ name: "" }));
    expect(result.status).toBe("validation_error");
    expect(acquireGraphToken).not.toHaveBeenCalled();
  });

  it("has no delete method", () => {
    const provider = new DelegatedSharePointOrganizationProvider(fakeController(async () => "token"), account, config);
    expect((provider as unknown as { delete?: unknown }).delete).toBeUndefined();
  });
});

describe("selectOrganizationProvider", () => {
  it("falls back to the memory provider outside a browser context, using the canonical env var only", async () => {
    const { kind, provider } = await selectOrganizationProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryOrganizationProvider);
  });
});
