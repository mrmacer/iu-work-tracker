import { describe, expect, it, vi } from "vitest";
import type { Contact } from "../lib/models";
import { buildContactDraft, CONTACT_STATUSES } from "../lib/contact-provider";
import {
  createContactItem,
  findContactByAppId,
  fromSharePointItem,
  listContactItems,
  resolveContactItem,
  SharePointContactsError,
  toSharePointFields,
  updateContactItem,
  validateContactSharePointLimits,
  type SharePointContactConfig,
} from "../lib/sharepoint-contacts";

const config: SharePointContactConfig = {
  siteId: "siu29.sharepoint.com,site,web",
  contactsListId: "contacts-list-id",
};

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

function testContact(overrides: Partial<Contact> = {}): Contact {
  return {
    ...buildContactDraft({
      appId: "contact-under-test",
      displayName: "Annie Milewski",
      role: "Superintendent",
      organizationId: "org-north-valley",
      email: "annie@example.org",
      status: "active",
      notes: "Leads district STEELS implementation.",
    }),
    ...overrides,
  };
}

function graphItem(c: Contact, version: number, overrides: Partial<{ id: string; eTag: string }> = {}) {
  return {
    id: overrides.id ?? "1",
    eTag: overrides.eTag ?? '"1"',
    createdDateTime: "2026-08-29T12:00:00Z",
    lastModifiedDateTime: "2026-08-29T12:00:00Z",
    fields: toSharePointFields(c, version),
  };
}

describe("Contact field mapping", () => {
  it("round-trips a contact, including every optional field present", () => {
    const c = testContact();
    const fields = toSharePointFields(c, 1);
    expect(fields.Title).toBe("Annie Milewski");
    expect(fields.AppId).toBe("contact-under-test");
    expect(fields.Role).toBe("Superintendent");
    expect(fields.OrganizationAppId).toBe("org-north-valley");
    expect(fields.Email).toBe("annie@example.org");
    expect(fields.Status).toBe("active");
    expect(fields.Notes).toBe("Leads district STEELS implementation.");
    expect(fields.RecordVersion).toBe(1);

    const mapped = fromSharePointItem(graphItem(c, 1));
    expect(mapped.appId).toBe(c.appId);
    expect(mapped.displayName).toBe(c.displayName);
    expect(mapped.role).toBe(c.role);
    expect(mapped.organizationId).toBe(c.organizationId);
    expect(mapped.email).toBe(c.email);
    expect(mapped.status).toBe("active");
    expect(mapped.notes).toBe(c.notes);
    expect(mapped.metadata).toEqual({ providerId: "1", version: 1, createdAt: "2026-08-29T12:00:00Z", modifiedAt: "2026-08-29T12:00:00Z", syncState: "saved" });
  });

  it("round-trips every allowed Status value as an exact lowercase string", () => {
    for (const status of CONTACT_STATUSES) {
      const c = testContact({ status });
      const fields = toSharePointFields(c, 1);
      expect(fields.Status).toBe(status);
      const mapped = fromSharePointItem(graphItem(c, 1));
      expect(mapped.status).toBe(status);
    }
  });

  it("maps a missing role/email/notes and a null organizationId without fabricating values", () => {
    const c = testContact({ organizationId: null });
    delete c.role;
    delete c.email;
    delete c.notes;
    const fields = toSharePointFields(c, 1);
    expect(fields.Role).toBe("");
    expect(fields.Email).toBe("");
    expect(fields.Notes).toBe("");
    expect(fields.OrganizationAppId).toBeNull();

    const mapped = fromSharePointItem(graphItem(c, 1));
    expect(mapped.role).toBeUndefined();
    expect(mapped.email).toBeUndefined();
    expect(mapped.notes).toBeUndefined();
    expect(mapped.organizationId).toBeNull();
  });

  it("rejects an unrecognized Status instead of silently coercing it", () => {
    const item = graphItem(testContact(), 1);
    item.fields.Status = "important";
    expect(() => fromSharePointItem(item)).toThrow(SharePointContactsError);
  });

  it("rejects a missing/invalid RecordVersion", () => {
    const item = graphItem(testContact(), 1);
    item.fields.RecordVersion = "not-a-number";
    expect(() => fromSharePointItem(item)).toThrow(/RecordVersion/);
  });

  it("rejects an item missing AppId", () => {
    const item = graphItem(testContact(), 1);
    item.fields.AppId = "";
    expect(() => fromSharePointItem(item)).toThrow(/AppId/);
  });

  it("flags an oversized text field via the SharePoint-compatible limit check, without truncating", () => {
    const oversized = testContact({ notes: "x".repeat(2001) });
    const issues = validateContactSharePointLimits(oversized);
    expect(issues).toContainEqual(expect.objectContaining({ path: "notes", code: "sharepoint_text_limit" }));
  });

  it("accepts a contact within every limit", () => {
    expect(validateContactSharePointLimits(testContact())).toEqual([]);
  });
});

describe("Contact Graph operations", () => {
  it("follows every @odata.nextLink when listing", async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({ value: [graphItem(testContact(), 1)], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next-page" }),
      )
      .mockResolvedValueOnce(jsonResponse({ value: [graphItem(testContact({ appId: "second" }), 1, { id: "2" })] })) as unknown as typeof fetch;

    const contacts = await listContactItems(config, "token", fetcher);
    expect(contacts).toHaveLength(2);
    expect(fetcher).toHaveBeenCalledTimes(2);
  });

  it("creates an item, sends RecordVersion 1, and reads it back for SharePoint-owned id/timestamps", async () => {
    const c = testContact();
    const fetcherMock = vi.fn().mockResolvedValueOnce(jsonResponse({ id: "42" })).mockResolvedValueOnce(jsonResponse(graphItem(c, 1, { id: "42" })));
    const created = await createContactItem(config, "token", c, fetcherMock as unknown as typeof fetch);
    expect(created.metadata?.providerId).toBe("42");
    expect(created.metadata?.version).toBe(1);
    const [, createInit] = fetcherMock.mock.calls[0] as [string, RequestInit];
    expect(createInit.method).toBe("POST");
    expect(JSON.parse(String(createInit.body)).fields.RecordVersion).toBe(1);
  });

  it("looks up an existing item by indexed AppId", async () => {
    const c = testContact();
    const fetcher = vi.fn(async (url: string) => {
      expect(url).toContain(encodeURIComponent(`fields/AppId eq '${c.appId}'`));
      return jsonResponse({ value: [graphItem(c, 1, { id: "7" })] });
    }) as unknown as typeof fetch;
    const found = await findContactByAppId(config, "token", c.appId, fetcher);
    expect(found?.itemId).toBe("7");
  });

  it("resolves an update target by providerId and falls back to AppId lookup when it's gone", async () => {
    const c = testContact({ metadata: { providerId: "stale-id", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" } });
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ error: "not found" }, 404))
      .mockResolvedValueOnce(jsonResponse({ value: [graphItem(c, 1, { id: "fresh-id" })] })) as unknown as typeof fetch;
    const resolved = await resolveContactItem(config, "token", c, fetcher);
    expect(resolved?.itemId).toBe("fresh-id");
  });

  it("sends If-Match on update and surfaces a 412 as a conflict carrying the current contact", async () => {
    const c = testContact();
    const fetcher = vi
      .fn()
      .mockImplementationOnce(async (_url: string, init: RequestInit) => {
        expect((init.headers as Record<string, string>)["If-Match"]).toBe('"stale"');
        return jsonResponse({ error: "precondition failed" }, 412);
      })
      .mockResolvedValueOnce(jsonResponse(graphItem(c, 2, { id: "5", eTag: '"2"' }))) as unknown as typeof fetch;

    await expect(updateContactItem(config, "token", "5", '"stale"', c, 2, fetcher)).rejects.toMatchObject({
      kind: "conflict",
      current: expect.objectContaining({ appId: c.appId }),
    });
  });

  it("updates and reads back the incremented version on success", async () => {
    const c = testContact();
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(new Response(null, { status: 204 }))
      .mockResolvedValueOnce(jsonResponse(graphItem(c, 2, { id: "5" }))) as unknown as typeof fetch;
    const saved = await updateContactItem(config, "token", "5", '"1"', c, 2, fetcher);
    expect(saved.metadata?.version).toBe(2);
  });

  it("maps a Graph 401 to an auth error", async () => {
    const fetcher = vi.fn(async () => jsonResponse({ error: "expired" }, 401)) as unknown as typeof fetch;
    await expect(listContactItems(config, "expired-token", fetcher)).rejects.toMatchObject({ kind: "auth" });
  });
});
