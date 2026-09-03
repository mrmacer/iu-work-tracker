import { describe, expect, it, vi } from "vitest";
import type { Organization } from "../lib/models";
import { buildOrganizationDraft, ORGANIZATION_TYPES } from "../lib/organization-provider";
import {
  createOrganizationItem,
  findOrganizationByAppId,
  fromSharePointItem,
  listOrganizationItems,
  resolveOrganizationItem,
  SharePointOrganizationsError,
  toSharePointFields,
  updateOrganizationItem,
  validateOrganizationSharePointLimits,
  type SharePointOrganizationConfig,
} from "../lib/sharepoint-organizations";

const config: SharePointOrganizationConfig = {
  siteId: "siu29.sharepoint.com,site,web",
  organizationsListId: "durable-organizations-list-id",
};

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

function testOrganization(overrides: Partial<Organization> = {}): Organization {
  return {
    ...buildOrganizationDraft({
      appId: "org-under-test",
      name: "North Valley SD",
      type: "district",
    }),
    ...overrides,
  };
}

function graphItem(o: Organization, version: number, overrides: Partial<{ id: string; eTag: string }> = {}) {
  return {
    id: overrides.id ?? "1",
    eTag: overrides.eTag ?? '"1"',
    createdDateTime: "2026-08-29T12:00:00Z",
    lastModifiedDateTime: "2026-08-29T12:00:00Z",
    fields: toSharePointFields(o, version),
  };
}

describe("Organization field mapping", () => {
  // OrganizationType is a live SharePoint Choice column, whose three allowed values are
  // exactly ORGANIZATION_TYPES' lowercase strings: district/partner/iu. A Graph Choice field's
  // value is written/read as a plain string identical to a text field — this proves every
  // allowed value round-trips exactly, with no casing coercion (never "District", etc.).
  it("round-trips every allowed OrganizationType Choice value as an exact lowercase string", () => {
    for (const type of ORGANIZATION_TYPES) {
      const o = testOrganization({ type });
      const fields = toSharePointFields(o, 1);
      expect(fields.OrganizationType).toBe(type);
      const mapped = fromSharePointItem(graphItem(o, 1));
      expect(mapped.type).toBe(type);
    }
  });

  it("round-trips an organization's Title/AppId/OrganizationType/RecordVersion", () => {
    const o = testOrganization();
    const fields = toSharePointFields(o, 1);
    expect(fields.Title).toBe("North Valley SD");
    expect(fields.AppId).toBe("org-under-test");
    expect(fields.OrganizationType).toBe("district");
    expect(fields.RecordVersion).toBe(1);

    const mapped = fromSharePointItem(graphItem(o, 1));
    expect(mapped.appId).toBe(o.appId);
    expect(mapped.name).toBe(o.name);
    expect(mapped.type).toBe("district");
    expect(mapped.metadata).toEqual({ providerId: "1", version: 1, createdAt: "2026-08-29T12:00:00Z", modifiedAt: "2026-08-29T12:00:00Z", syncState: "saved" });
  });

  it("rejects an unrecognized OrganizationType instead of silently coercing it", () => {
    const item = graphItem(testOrganization(), 1);
    item.fields.OrganizationType = "nonprofit";
    expect(() => fromSharePointItem(item)).toThrow(SharePointOrganizationsError);
  });

  it("rejects a missing/invalid RecordVersion", () => {
    const item = graphItem(testOrganization(), 1);
    item.fields.RecordVersion = "not-a-number";
    expect(() => fromSharePointItem(item)).toThrow(/RecordVersion/);
  });

  it("rejects an item missing SharePoint Created/Modified timestamps", () => {
    const item = graphItem(testOrganization(), 1);
    delete (item as { createdDateTime?: string }).createdDateTime;
    expect(() => fromSharePointItem(item)).toThrow(/Created\/Modified/);
  });

  it("flags an oversized text field via the SharePoint-compatible limit check, without truncating", () => {
    const oversized = testOrganization({ name: "x".repeat(256) });
    const issues = validateOrganizationSharePointLimits(oversized);
    expect(issues).toContainEqual(expect.objectContaining({ path: "name", code: "sharepoint_text_limit" }));
  });

  it("accepts an organization within every limit", () => {
    expect(validateOrganizationSharePointLimits(testOrganization())).toEqual([]);
  });

  it("never uses a Lookup-shaped field — fields are plain Title/AppId/OrganizationType/RecordVersion only", () => {
    const fields = toSharePointFields(testOrganization(), 1);
    expect(Object.keys(fields).sort()).toEqual(["AppId", "OrganizationType", "RecordVersion", "Title"]);
  });
});

describe("Organization Graph operations", () => {
  it("follows every @odata.nextLink when listing", async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({ value: [graphItem(testOrganization(), 1)], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next-page" }),
      )
      .mockResolvedValueOnce(jsonResponse({ value: [graphItem(testOrganization({ appId: "second" }), 1, { id: "2" })] })) as unknown as typeof fetch;

    const organizations = await listOrganizationItems(config, "token", fetcher);
    expect(organizations).toHaveLength(2);
    expect(fetcher).toHaveBeenCalledTimes(2);
  });

  it("creates an item, sends RecordVersion 1, and reads it back for SharePoint-owned id/timestamps", async () => {
    const o = testOrganization();
    const fetcherMock = vi.fn().mockResolvedValueOnce(jsonResponse({ id: "42" })).mockResolvedValueOnce(jsonResponse(graphItem(o, 1, { id: "42" })));
    const created = await createOrganizationItem(config, "token", o, fetcherMock as unknown as typeof fetch);
    expect(created.metadata?.providerId).toBe("42");
    expect(created.metadata?.version).toBe(1);
    const [, createInit] = fetcherMock.mock.calls[0] as [string, RequestInit];
    expect(createInit.method).toBe("POST");
    expect(JSON.parse(String(createInit.body)).fields.RecordVersion).toBe(1);
  });

  it("looks up an existing item by indexed AppId", async () => {
    const o = testOrganization();
    const fetcher = vi.fn(async (url: string) => {
      expect(url).toContain(encodeURIComponent(`fields/AppId eq '${o.appId}'`));
      return jsonResponse({ value: [graphItem(o, 1, { id: "7" })] });
    }) as unknown as typeof fetch;
    const found = await findOrganizationByAppId(config, "token", o.appId, fetcher);
    expect(found?.itemId).toBe("7");
  });

  it("resolves an update target by providerId and falls back to AppId lookup when it's gone", async () => {
    const o = testOrganization({ metadata: { providerId: "stale-id", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" } });
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ error: "not found" }, 404))
      .mockResolvedValueOnce(jsonResponse({ value: [graphItem(o, 1, { id: "fresh-id" })] })) as unknown as typeof fetch;
    const resolved = await resolveOrganizationItem(config, "token", o, fetcher);
    expect(resolved?.itemId).toBe("fresh-id");
  });

  it("sends If-Match on update and surfaces a 412 as a conflict carrying the current organization — never auto-merges or overwrites", async () => {
    const o = testOrganization();
    const fetcher = vi
      .fn()
      .mockImplementationOnce(async (_url: string, init: RequestInit) => {
        expect((init.headers as Record<string, string>)["If-Match"]).toBe('"stale"');
        return jsonResponse({ error: "precondition failed" }, 412);
      })
      .mockResolvedValueOnce(jsonResponse(graphItem(o, 2, { id: "5", eTag: '"2"' }))) as unknown as typeof fetch;

    await expect(updateOrganizationItem(config, "token", "5", '"stale"', o, 2, fetcher)).rejects.toMatchObject({
      kind: "conflict",
      current: expect.objectContaining({ appId: o.appId }),
    });
  });

  it("updates and reads back the incremented version on success", async () => {
    const o = testOrganization();
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(new Response(null, { status: 204 }))
      .mockResolvedValueOnce(jsonResponse(graphItem(o, 2, { id: "5" }))) as unknown as typeof fetch;
    const saved = await updateOrganizationItem(config, "token", "5", '"1"', o, 2, fetcher);
    expect(saved.metadata?.version).toBe(2);
  });

  it("maps a Graph 401 to an auth error", async () => {
    const fetcher = vi.fn(async () => jsonResponse({ error: "expired" }, 401)) as unknown as typeof fetch;
    await expect(listOrganizationItems(config, "expired-token", fetcher)).rejects.toMatchObject({ kind: "auth" });
  });
});
