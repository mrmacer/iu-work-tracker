import { afterEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import type { Contact } from "../lib/models";
import {
  buildContactDraft,
  CONTACT_STATUSES,
  DelegatedSharePointContactProvider,
  MemoryContactProvider,
  normalizeContactEmail,
  normalizeContactName,
  selectContactProvider,
  validateContactShape,
} from "../lib/contact-provider";
import { toSharePointFields } from "../lib/sharepoint-contacts";
import type { MicrosoftAuthController } from "../lib/microsoft-auth";

const account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
const config = { siteId: "site-id", contactsListId: "list-id" };

function fakeController(acquireGraphToken: () => Promise<string>): MicrosoftAuthController {
  return { acquireGraphToken } as unknown as MicrosoftAuthController;
}

function contact(overrides: Partial<Parameters<typeof buildContactDraft>[0]> = {}): Contact {
  return buildContactDraft({
    appId: "contact-under-test",
    displayName: "Annie Milewski",
    role: "Superintendent",
    organizationId: "org-north-valley",
    email: "annie@example.org",
    status: "active",
    notes: "Leads district STEELS implementation.",
    ...overrides,
  });
}

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

function graphItem(c: Contact, version = 1, id = "1") {
  return {
    id,
    eTag: '"1"',
    createdDateTime: "2026-08-29T12:00:00Z",
    lastModifiedDateTime: "2026-08-29T12:00:00Z",
    fields: toSharePointFields(c, version),
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("Contact model / status", () => {
  it("builds a fresh, unsaved draft with metadata.version 0 — routes create() first", () => {
    expect(contact().metadata?.version).toBe(0);
  });

  it("requires a non-empty displayName", () => {
    const issues = validateContactShape(contact({ displayName: "  " }));
    expect(issues).toContainEqual(expect.objectContaining({ path: "displayName", code: "required" }));
  });

  it("accepts exactly the five approved statuses", () => {
    for (const status of CONTACT_STATUSES) {
      expect(validateContactShape(contact({ status }))).toEqual([]);
    }
  });

  it("rejects a status outside the approved set", () => {
    const issues = validateContactShape(contact({ status: "important" as Contact["status"] }));
    expect(issues).toContainEqual(expect.objectContaining({ path: "status", code: "invalid_status" }));
  });

  it("accepts an omitted role", () => {
    const draft = contact();
    delete draft.role;
    expect(validateContactShape(draft)).toEqual([]);
  });

  it("accepts an omitted email", () => {
    const draft = contact();
    delete draft.email;
    expect(validateContactShape(draft)).toEqual([]);
  });

  it("rejects a malformed email instead of silently coercing it", () => {
    const issues = validateContactShape(contact({ email: "not-an-email" }));
    expect(issues).toContainEqual(expect.objectContaining({ path: "email", code: "invalid_email" }));
  });

  it("accepts an omitted notes field", () => {
    const draft = contact();
    delete draft.notes;
    expect(validateContactShape(draft)).toEqual([]);
  });

  it("accepts a null organizationId", () => {
    expect(validateContactShape(contact({ organizationId: null }))).toEqual([]);
  });
});

describe("Normalization", () => {
  it("normalizes names by trim + lowercase, no fuzzy matching", () => {
    expect(normalizeContactName("  Annie Milewski  ")).toBe("annie milewski");
    expect(normalizeContactName("ANNIE MILEWSKI")).toBe("annie milewski");
  });

  it("normalizes emails by trim + lowercase", () => {
    expect(normalizeContactEmail("  Annie@Example.ORG ")).toBe("annie@example.org");
  });
});

describe("MemoryContactProvider", () => {
  it("assigns SharePoint-shaped provider metadata on create and lists newest first", async () => {
    const provider = new MemoryContactProvider();
    const first = await provider.create(contact());
    const second = await provider.create(contact({ appId: "second-contact", email: "kim@example.org", displayName: "Kim Rivera" }));
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

  it("rejects a duplicate AppId as a conflict without overwriting the existing contact", async () => {
    const provider = new MemoryContactProvider();
    const created = await provider.create(contact());
    if (created.status !== "success") throw new Error("setup failed");
    const duplicate = await provider.create({ ...contact(), appId: created.value.appId });
    expect(duplicate.status).toBe("conflict");
  });

  it("increments RecordVersion on update and rejects a stale expectedVersion as a conflict", async () => {
    const provider = new MemoryContactProvider();
    const created = await provider.create(contact());
    if (created.status !== "success") throw new Error("setup failed");

    const updated = await provider.update({ ...created.value, status: "dormant" }, 1);
    expect(updated.status).toBe("success");
    if (updated.status === "success") expect(updated.value.metadata?.version).toBe(2);

    const stale = await provider.update({ ...created.value, status: "archived" }, 1);
    expect(stale.status).toBe("conflict");
  });

  it("keeps every instance independent — nothing is shared or durable across instances", async () => {
    const first = new MemoryContactProvider();
    await first.create(contact());
    const second = new MemoryContactProvider();
    expect(await second.list()).toEqual({ status: "success", value: [] });
  });

  it("has no delete method", () => {
    expect((new MemoryContactProvider() as unknown as { delete?: unknown }).delete).toBeUndefined();
  });
});

describe("DelegatedSharePointContactProvider", () => {
  it("returns a structured conflict when AppId already exists on create", async () => {
    const c = contact();
    const fetcher = vi.fn(async () => jsonResponse({ value: [graphItem(c)] }));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointContactProvider(fakeController(async () => "token"), account, config);
    const result = await provider.create(c);
    expect(result.status).toBe("conflict");
  });

  it("creates successfully and normalizes the read-back contact", async () => {
    const c = contact();
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ value: [] }))
      .mockResolvedValueOnce(jsonResponse({ id: "10" }))
      .mockResolvedValueOnce(jsonResponse(graphItem(c, 1, "10")));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointContactProvider(fakeController(async () => "token"), account, config);
    const result = await provider.create(c);
    expect(result.status).toBe("success");
    if (result.status === "success") expect(result.value.metadata?.providerId).toBe("10");
  });

  it("returns conflict without writing when RecordVersion no longer matches expectedVersion", async () => {
    const c = { ...contact(), metadata: { providerId: "10", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" as const } };
    const fetcher = vi.fn(async () => jsonResponse(graphItem(c, 2, "10")));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointContactProvider(fakeController(async () => "token"), account, config);
    const result = await provider.update(c, 1);
    expect(result.status).toBe("conflict");
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it("maps an expired Graph token to a network_error result instead of throwing", async () => {
    const provider = new DelegatedSharePointContactProvider(
      fakeController(async () => {
        throw new Error("boom");
      }),
      account,
      config,
    );
    const result = await provider.list();
    expect(result.status).toBe("network_error");
  });

  it("rejects an invalid contact before ever acquiring a Graph token", async () => {
    const acquireGraphToken = vi.fn();
    const provider = new DelegatedSharePointContactProvider(fakeController(acquireGraphToken), account, config);
    const result = await provider.create(contact({ displayName: "" }));
    expect(result.status).toBe("validation_error");
    expect(acquireGraphToken).not.toHaveBeenCalled();
  });

  it("has no delete method", () => {
    const provider = new DelegatedSharePointContactProvider(fakeController(async () => "token"), account, config);
    expect((provider as unknown as { delete?: unknown }).delete).toBeUndefined();
  });
});

describe("selectContactProvider", () => {
  it("falls back to the memory provider outside a browser context — never a fake/hardcoded SharePoint list id", async () => {
    const { kind, provider } = await selectContactProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryContactProvider);
  });
});
