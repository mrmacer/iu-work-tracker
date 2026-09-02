import { describe, expect, it } from "vitest";
import {
  buildContactRelationshipSummary,
  selectConnectedProjects,
  selectConnectedWorkRecords,
} from "../lib/contact-relationships";
import { WORK_RECORD_SCHEMA_VERSION, type Project, type WorkRecord } from "../lib/models";

// Patch 8C — pure derivation, no I/O, no AI. Every test here proves the relationship stays
// exactly Contact.appId → WorkRecord.contactIds → WorkRecord.projectIds → Project, and that
// nothing is ever inferred from a name/email/organization/free-text mention.

function project(overrides: Partial<Project> = {}): Project {
  return { appId: "project-steels", name: "STEELS Implementation", description: "", status: "active", color: "blue", ...overrides };
}

function workRecord(overrides: Partial<WorkRecord> = {}): WorkRecord {
  return {
    appId: "wr-1",
    title: "Some work",
    activityDate: "2026-09-01",
    activityType: "Professional learning",
    description: "",
    detailedNotes: "",
    durationMinutes: 60,
    status: "complete",
    engagementScope: "none",
    projectIds: [],
    organizationIds: [],
    contactIds: [],
    categoryIds: [],
    reach: { educatorsLeaders: 0, studentsFamilies: 0, workforceCommunity: 0, other: 0 },
    evidenceSummary: "",
    evidenceReferenceIds: [],
    output: "",
    outcome: "",
    nextStep: "",
    followUpNeeded: false,
    followUpDate: null,
    orbit: { reportable: false, primaryDeliverable: null, supportingDeliverables: [], stemPocMinutes: 0, tacMinutes: 0, evidence: "" },
    schemaVersion: WORK_RECORD_SCHEMA_VERSION,
    metadata: { version: 1, createdAt: "2026-08-01T00:00:00.000Z", modifiedAt: "2026-08-01T00:00:00.000Z", syncState: "saved" },
    isSample: false,
    ...overrides,
  };
}

describe("selectConnectedWorkRecords", () => {
  it("includes only Work Records whose contactIds contains the exact Contact appId", () => {
    const target = workRecord({ appId: "wr-connected", contactIds: ["contact-annie"] });
    const other = workRecord({ appId: "wr-unrelated", contactIds: ["contact-kim"] });
    const result = selectConnectedWorkRecords([target, other], "contact-annie");
    expect(result.map((r) => r.appId)).toEqual(["wr-connected"]);
  });

  it("never infers a connection from name, email, organization, or free text — only the explicit ID array", () => {
    // A record that merely mentions "Annie" in its title/description is NOT connected —
    // selectConnectedWorkRecords never reads those fields at all, only contactIds.
    const mentionsOnly = workRecord({ appId: "wr-mention", title: "Called Annie about the grant", description: "Talked to Annie Milewski", contactIds: [] });
    const result = selectConnectedWorkRecords([mentionsOnly], "contact-annie");
    expect(result).toEqual([]);
  });

  it("sorts newest activityDate first, matching History's own business-date convention", () => {
    const oldest = workRecord({ appId: "wr-oldest", activityDate: "2026-08-01", contactIds: ["contact-annie"] });
    const newest = workRecord({ appId: "wr-newest", activityDate: "2026-09-01", contactIds: ["contact-annie"] });
    const middle = workRecord({ appId: "wr-middle", activityDate: "2026-08-15", contactIds: ["contact-annie"] });
    const result = selectConnectedWorkRecords([oldest, newest, middle], "contact-annie");
    expect(result.map((r) => r.appId)).toEqual(["wr-newest", "wr-middle", "wr-oldest"]);
  });

  it("returns an empty array when there are no connected records", () => {
    expect(selectConnectedWorkRecords([workRecord({ contactIds: [] })], "contact-annie")).toEqual([]);
  });
});

describe("selectConnectedProjects", () => {
  it("unions projectIds across every connected Work Record", () => {
    const records = [
      workRecord({ appId: "wr-1", projectIds: ["project-steels"] }),
      workRecord({ appId: "wr-2", projectIds: ["project-ai"] }),
    ];
    const projects = [project({ appId: "project-steels" }), project({ appId: "project-ai", name: "AI in Education" })];
    const result = selectConnectedProjects(records, projects);
    expect(result.map((p) => p.appId).sort()).toEqual(["project-ai", "project-steels"]);
  });

  it("deduplicates a project referenced by multiple connected Work Records", () => {
    const records = [
      workRecord({ appId: "wr-1", projectIds: ["project-steels"] }),
      workRecord({ appId: "wr-2", projectIds: ["project-steels"] }),
    ];
    const result = selectConnectedProjects(records, [project()]);
    expect(result).toHaveLength(1);
  });

  it("silently skips an unresolvable historical project ID instead of crashing or fabricating a Project", () => {
    const records = [workRecord({ projectIds: ["project-does-not-exist-anymore"] })];
    expect(() => selectConnectedProjects(records, [project()])).not.toThrow();
    expect(selectConnectedProjects(records, [project()])).toEqual([]);
  });

  it("returns an empty array when no connected Work Record references any project", () => {
    expect(selectConnectedProjects([workRecord({ projectIds: [] })], [project()])).toEqual([]);
  });
});

describe("buildContactRelationshipSummary", () => {
  it("sums durationMinutes only across connected Work Records", () => {
    const records = [
      workRecord({ appId: "wr-1", contactIds: ["contact-annie"], durationMinutes: 45 }),
      workRecord({ appId: "wr-2", contactIds: ["contact-annie"], durationMinutes: 30 }),
      workRecord({ appId: "wr-3", contactIds: ["contact-kim"], durationMinutes: 999 }), // unrelated — must not count
    ];
    const summary = buildContactRelationshipSummary(records, [], "contact-annie");
    expect(summary.totalMinutes).toBe(75);
    expect(summary.workRecordCount).toBe(2);
  });

  it("uses the newest connected Work Record's activityDate as lastInteractionDate — never SharePoint Created/Modified, never today", () => {
    const records = [
      workRecord({ appId: "wr-1", contactIds: ["contact-annie"], activityDate: "2026-08-20", metadata: { version: 1, createdAt: "2099-01-01T00:00:00.000Z", modifiedAt: "2099-01-01T00:00:00.000Z", syncState: "saved" } }),
      workRecord({ appId: "wr-2", contactIds: ["contact-annie"], activityDate: "2026-09-01" }),
    ];
    const summary = buildContactRelationshipSummary(records, [], "contact-annie");
    expect(summary.lastInteractionDate).toBe("2026-09-01");
  });

  it("an unrelated newer Work Record never affects lastInteractionDate", () => {
    const records = [
      workRecord({ appId: "wr-connected", contactIds: ["contact-annie"], activityDate: "2026-08-01" }),
      workRecord({ appId: "wr-unrelated-newer", contactIds: ["contact-kim"], activityDate: "2026-09-15" }),
    ];
    const summary = buildContactRelationshipSummary(records, [], "contact-annie");
    expect(summary.lastInteractionDate).toBe("2026-08-01");
  });

  it("returns a clean, empty summary when there are no connected records — never crashes, never fabricates", () => {
    const summary = buildContactRelationshipSummary([workRecord({ contactIds: [] })], [project()], "contact-annie");
    expect(summary).toEqual({
      connectedWorkRecords: [],
      recentWorkRecords: [],
      connectedProjects: [],
      totalMinutes: 0,
      workRecordCount: 0,
      lastInteractionDate: null,
    });
  });

  it("limits recentWorkRecords to the 5 most recent while workRecordCount reflects the true total", () => {
    const records = Array.from({ length: 8 }, (_, i) =>
      workRecord({ appId: `wr-${i}`, contactIds: ["contact-annie"], activityDate: `2026-08-${String(i + 1).padStart(2, "0")}` }),
    );
    const summary = buildContactRelationshipSummary(records, [], "contact-annie");
    expect(summary.recentWorkRecords).toHaveLength(5);
    expect(summary.workRecordCount).toBe(8);
    expect(summary.recentWorkRecords[0].appId).toBe("wr-7"); // newest (Aug 8) first
  });
});
