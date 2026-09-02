import { afterEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import type { Project } from "../lib/models";
import {
  buildProjectDraft,
  DelegatedSharePointProjectProvider,
  MemoryProjectProvider,
  PROJECT_STATUSES,
  selectProjectProvider,
  validateProjectShape,
} from "../lib/project-provider";
import { toSharePointFields } from "../lib/sharepoint-projects";
import type { MicrosoftAuthController } from "../lib/microsoft-auth";

const account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
const config = { siteId: "site-id", projectsListId: "list-id" };

function fakeController(acquireGraphToken: () => Promise<string>): MicrosoftAuthController {
  return { acquireGraphToken } as unknown as MicrosoftAuthController;
}

function project(overrides: Partial<Parameters<typeof buildProjectDraft>[0]> = {}): Project {
  return buildProjectDraft({
    appId: "project-under-test",
    name: "STEM Ecosystem Expansion",
    description: "Cross-sector partnership development.",
    status: "planning",
    color: "blue",
    startDate: "2026-09-01",
    targetDate: "2027-06-30",
    stemOrbit: true,
    ...overrides,
  });
}

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

function graphItem(p: Project, version = 1, id = "1") {
  return {
    id,
    eTag: '"1"',
    createdDateTime: "2026-08-29T12:00:00Z",
    lastModifiedDateTime: "2026-08-29T12:00:00Z",
    fields: toSharePointFields(p, version),
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("Project model", () => {
  it("builds a fresh, unsaved draft with metadata.version 0 — routes create() first", () => {
    const p = project();
    expect(p.metadata?.version).toBe(0);
  });

  it("requires a non-empty name", () => {
    const issues = validateProjectShape(project({ name: "  " }));
    expect(issues).toContainEqual(expect.objectContaining({ path: "name", code: "required" }));
  });

  it("accepts exactly the four allowed statuses, including the new paused status", () => {
    for (const status of PROJECT_STATUSES) {
      expect(validateProjectShape(project({ status }))).toEqual([]);
    }
  });

  it("rejects a status outside the allowed set", () => {
    const issues = validateProjectShape(project({ status: "archived" as Project["status"] }));
    expect(issues).toContainEqual(expect.objectContaining({ path: "status", code: "invalid_status" }));
  });

  it("accepts null or omitted optional dates and STEM/ORBIT connection", () => {
    expect(validateProjectShape(project({ startDate: null, targetDate: null }))).toEqual([]);
    const withoutStemOrbit = project();
    delete withoutStemOrbit.stemOrbit;
    expect(validateProjectShape(withoutStemOrbit)).toEqual([]);
  });

  it("rejects a malformed date instead of silently coercing it", () => {
    const issues = validateProjectShape(project({ startDate: "not-a-date" }));
    expect(issues).toContainEqual(expect.objectContaining({ path: "startDate", code: "invalid_date" }));
  });
});

describe("MemoryProjectProvider", () => {
  it("assigns SharePoint-shaped provider metadata on create and lists newest first", async () => {
    const provider = new MemoryProjectProvider();
    const first = await provider.create(project());
    const second = await provider.create(project({ appId: "second-project" }));
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

  it("rejects a duplicate AppId as a conflict without overwriting the existing project", async () => {
    const provider = new MemoryProjectProvider();
    const created = await provider.create(project());
    if (created.status !== "success") throw new Error("setup failed");
    const duplicate = await provider.create({ ...project(), appId: created.value.appId });
    expect(duplicate.status).toBe("conflict");
  });

  it("increments RecordVersion on update and rejects a stale expectedVersion as a conflict", async () => {
    const provider = new MemoryProjectProvider();
    const created = await provider.create(project());
    if (created.status !== "success") throw new Error("setup failed");

    const updated = await provider.update({ ...created.value, status: "active" }, 1);
    expect(updated.status).toBe("success");
    if (updated.status === "success") expect(updated.value.metadata?.version).toBe(2);

    const stale = await provider.update({ ...created.value, status: "paused" }, 1);
    expect(stale.status).toBe("conflict");
  });

  it("keeps every instance independent — nothing is shared or durable across instances", async () => {
    const first = new MemoryProjectProvider();
    await first.create(project());
    const second = new MemoryProjectProvider();
    expect(await second.list()).toEqual({ status: "success", value: [] });
  });

  it("has no delete method", () => {
    expect((new MemoryProjectProvider() as unknown as { delete?: unknown }).delete).toBeUndefined();
  });
});

describe("DelegatedSharePointProjectProvider", () => {
  it("returns a structured conflict when AppId already exists on create", async () => {
    const p = project();
    const fetcher = vi.fn(async () => jsonResponse({ value: [graphItem(p)] }));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointProjectProvider(fakeController(async () => "token"), account, config);
    const result = await provider.create(p);
    expect(result.status).toBe("conflict");
  });

  it("creates successfully and normalizes the read-back project", async () => {
    const p = project();
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ value: [] }))
      .mockResolvedValueOnce(jsonResponse({ id: "10" }))
      .mockResolvedValueOnce(jsonResponse(graphItem(p, 1, "10")));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointProjectProvider(fakeController(async () => "token"), account, config);
    const result = await provider.create(p);
    expect(result.status).toBe("success");
    if (result.status === "success") expect(result.value.metadata?.providerId).toBe("10");
  });

  it("returns conflict without writing when RecordVersion no longer matches expectedVersion", async () => {
    const p = { ...project(), metadata: { providerId: "10", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" as const } };
    const fetcher = vi.fn(async () => jsonResponse(graphItem(p, 2, "10")));
    vi.stubGlobal("fetch", fetcher);
    const provider = new DelegatedSharePointProjectProvider(fakeController(async () => "token"), account, config);
    const result = await provider.update(p, 1);
    expect(result.status).toBe("conflict");
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it("maps an expired Graph token to a network_error result instead of throwing", async () => {
    const provider = new DelegatedSharePointProjectProvider(
      fakeController(async () => {
        throw new Error("boom");
      }),
      account,
      config,
    );
    const result = await provider.list();
    expect(result.status).toBe("network_error");
  });

  it("rejects an invalid project before ever acquiring a Graph token", async () => {
    const acquireGraphToken = vi.fn();
    const provider = new DelegatedSharePointProjectProvider(fakeController(acquireGraphToken), account, config);
    const result = await provider.create(project({ name: "" }));
    expect(result.status).toBe("validation_error");
    expect(acquireGraphToken).not.toHaveBeenCalled();
  });

  it("has no delete method", () => {
    const provider = new DelegatedSharePointProjectProvider(fakeController(async () => "token"), account, config);
    expect((provider as unknown as { delete?: unknown }).delete).toBeUndefined();
  });
});

describe("selectProjectProvider", () => {
  it("falls back to the memory provider outside a browser context — never a fake/hardcoded SharePoint list id", async () => {
    const { kind, provider } = await selectProjectProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryProjectProvider);
  });
});
