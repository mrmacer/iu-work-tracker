import { describe, expect, it, vi } from "vitest";
import type { Project } from "../lib/models";
import { buildProjectDraft, PROJECT_STATUSES } from "../lib/project-provider";
import {
  createProjectItem,
  findProjectByAppId,
  fromSharePointItem,
  listProjectItems,
  resolveProjectItem,
  SharePointProjectsError,
  toSharePointFields,
  updateProjectItem,
  validateProjectSharePointLimits,
  type SharePointProjectConfig,
} from "../lib/sharepoint-projects";

const config: SharePointProjectConfig = {
  siteId: "siu29.sharepoint.com,site,web",
  projectsListId: "durable-projects-list-id",
};

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

function testProject(overrides: Partial<Project> = {}): Project {
  return {
    ...buildProjectDraft({
      appId: "project-under-test",
      name: "STEM Ecosystem Expansion",
      description: "Cross-sector partnership development.",
      status: "planning",
      color: "blue",
      startDate: "2026-09-01",
      targetDate: "2027-06-30",
      stemOrbit: true,
    }),
    ...overrides,
  };
}

function graphItem(p: Project, version: number, overrides: Partial<{ id: string; eTag: string }> = {}) {
  return {
    id: overrides.id ?? "1",
    eTag: overrides.eTag ?? '"1"',
    createdDateTime: "2026-08-29T12:00:00Z",
    lastModifiedDateTime: "2026-08-29T12:00:00Z",
    fields: toSharePointFields(p, version),
  };
}

describe("Project field mapping", () => {
  // ProjectStatus is a live SharePoint Choice column (Patch 7B — kept as Choice, never
  // converted to text), whose four allowed values are exactly PROJECT_STATUSES' lowercase
  // strings: active/planning/paused/complete. A Graph Choice field's value is written/read as
  // a plain string identical to a text field — this proves every allowed value round-trips
  // exactly, with no casing coercion (never "Planning", "Active", etc. on the wire).
  it("round-trips every allowed ProjectStatus Choice value as an exact lowercase string", () => {
    for (const status of PROJECT_STATUSES) {
      const p = testProject({ status });
      const fields = toSharePointFields(p, 1);
      expect(fields.ProjectStatus).toBe(status);
      const mapped = fromSharePointItem(graphItem(p, 1));
      expect(mapped.status).toBe(status);
    }
  });

  it("round-trips a project, including the new paused-status and date/STEM-ORBIT fields", () => {
    const p = testProject({ status: "paused" });
    const fields = toSharePointFields(p, 1);
    expect(fields.ProjectStatus).toBe("paused");
    expect(fields.StartDate).toBe("2026-09-01T00:00:00.000Z");
    expect(fields.TargetDate).toBe("2027-06-30T00:00:00.000Z");
    expect(fields.StemOrbit).toBe(true);

    const mapped = fromSharePointItem(graphItem(p, 1));
    expect(mapped.appId).toBe(p.appId);
    expect(mapped.name).toBe(p.name);
    expect(mapped.status).toBe("paused");
    expect(mapped.startDate).toBe("2026-09-01");
    expect(mapped.targetDate).toBe("2027-06-30");
    expect(mapped.stemOrbit).toBe(true);
    expect(mapped.metadata).toEqual({ providerId: "1", version: 1, createdAt: "2026-08-29T12:00:00Z", modifiedAt: "2026-08-29T12:00:00Z", syncState: "saved" });
  });

  it("maps null optional dates and a false STEM/ORBIT flag without fabricating values", () => {
    const p = testProject({ startDate: null, targetDate: null, stemOrbit: false });
    const fields = toSharePointFields(p, 1);
    expect(fields.StartDate).toBeNull();
    expect(fields.TargetDate).toBeNull();
    expect(fields.StemOrbit).toBe(false);
    const mapped = fromSharePointItem(graphItem(p, 1));
    expect(mapped.startDate).toBeNull();
    expect(mapped.targetDate).toBeNull();
  });

  it("rejects an unrecognized ProjectStatus instead of silently coercing it", () => {
    const item = graphItem(testProject(), 1);
    item.fields.ProjectStatus = "archived";
    expect(() => fromSharePointItem(item)).toThrow(SharePointProjectsError);
  });

  it("rejects a missing/invalid RecordVersion", () => {
    const item = graphItem(testProject(), 1);
    item.fields.RecordVersion = "not-a-number";
    expect(() => fromSharePointItem(item)).toThrow(/RecordVersion/);
  });

  it("flags an oversized text field via the SharePoint-compatible limit check, without truncating", () => {
    const oversized = testProject({ description: "x".repeat(1001) });
    const issues = validateProjectSharePointLimits(oversized);
    expect(issues).toContainEqual(expect.objectContaining({ path: "description", code: "sharepoint_text_limit" }));
  });

  it("accepts a project within every limit", () => {
    expect(validateProjectSharePointLimits(testProject())).toEqual([]);
  });
});

describe("Project Graph operations", () => {
  it("follows every @odata.nextLink when listing", async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({ value: [graphItem(testProject(), 1)], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next-page" }),
      )
      .mockResolvedValueOnce(jsonResponse({ value: [graphItem(testProject({ appId: "second" }), 1, { id: "2" })] })) as unknown as typeof fetch;

    const projects = await listProjectItems(config, "token", fetcher);
    expect(projects).toHaveLength(2);
    expect(fetcher).toHaveBeenCalledTimes(2);
  });

  it("creates an item, sends RecordVersion 1, and reads it back for SharePoint-owned id/timestamps", async () => {
    const p = testProject();
    const fetcherMock = vi.fn().mockResolvedValueOnce(jsonResponse({ id: "42" })).mockResolvedValueOnce(jsonResponse(graphItem(p, 1, { id: "42" })));
    const created = await createProjectItem(config, "token", p, fetcherMock as unknown as typeof fetch);
    expect(created.metadata?.providerId).toBe("42");
    expect(created.metadata?.version).toBe(1);
    const [, createInit] = fetcherMock.mock.calls[0] as [string, RequestInit];
    expect(createInit.method).toBe("POST");
    expect(JSON.parse(String(createInit.body)).fields.RecordVersion).toBe(1);
  });

  it("looks up an existing item by indexed AppId", async () => {
    const p = testProject();
    const fetcher = vi.fn(async (url: string) => {
      expect(url).toContain(encodeURIComponent(`fields/AppId eq '${p.appId}'`));
      return jsonResponse({ value: [graphItem(p, 1, { id: "7" })] });
    }) as unknown as typeof fetch;
    const found = await findProjectByAppId(config, "token", p.appId, fetcher);
    expect(found?.itemId).toBe("7");
  });

  it("resolves an update target by providerId and falls back to AppId lookup when it's gone", async () => {
    const p = testProject({ metadata: { providerId: "stale-id", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" } });
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ error: "not found" }, 404))
      .mockResolvedValueOnce(jsonResponse({ value: [graphItem(p, 1, { id: "fresh-id" })] })) as unknown as typeof fetch;
    const resolved = await resolveProjectItem(config, "token", p, fetcher);
    expect(resolved?.itemId).toBe("fresh-id");
  });

  it("sends If-Match on update and surfaces a 412 as a conflict carrying the current project", async () => {
    const p = testProject();
    const fetcher = vi
      .fn()
      .mockImplementationOnce(async (_url: string, init: RequestInit) => {
        expect((init.headers as Record<string, string>)["If-Match"]).toBe('"stale"');
        return jsonResponse({ error: "precondition failed" }, 412);
      })
      .mockResolvedValueOnce(jsonResponse(graphItem(p, 2, { id: "5", eTag: '"2"' }))) as unknown as typeof fetch;

    await expect(updateProjectItem(config, "token", "5", '"stale"', p, 2, fetcher)).rejects.toMatchObject({
      kind: "conflict",
      current: expect.objectContaining({ appId: p.appId }),
    });
  });

  it("updates and reads back the incremented version on success", async () => {
    const p = testProject();
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(new Response(null, { status: 204 }))
      .mockResolvedValueOnce(jsonResponse(graphItem(p, 2, { id: "5" }))) as unknown as typeof fetch;
    const saved = await updateProjectItem(config, "token", "5", '"1"', p, 2, fetcher);
    expect(saved.metadata?.version).toBe(2);
  });

  it("maps a Graph 401 to an auth error", async () => {
    const fetcher = vi.fn(async () => jsonResponse({ error: "expired" }, 401)) as unknown as typeof fetch;
    await expect(listProjectItems(config, "expired-token", fetcher)).rejects.toMatchObject({ kind: "auth" });
  });
});
