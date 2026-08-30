import { describe, expect, it } from "vitest";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import { REFERENCE_DATA } from "../lib/reference-data";
import { buildWorkRecordDraftFromAnalysis } from "../lib/inbox-intelligence-work-record";
import {
  buildInboxIntelligenceRecord,
  computeInboxIntelligenceSummary,
  EmailAnalysisSchema,
  type EmailAnalysis,
} from "../lib/inbox-intelligence-models";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import { validateWorkRecord } from "../lib/validation";

function baseRecord(): WorkRecord {
  return {
    appId: "draft-under-test",
    title: "",
    activityDate: "2026-08-29",
    activityType: "",
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
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
    isSample: false,
  };
}

function analysis(overrides: Partial<EmailAnalysis> = {}): EmailAnalysis {
  return EmailAnalysisSchema.parse({
    summary: "Test summary",
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
    suggestedWorkRecord: { title: "Suggested title", description: "Suggested description" },
    ...overrides,
  });
}

describe("buildWorkRecordDraftFromAnalysis", () => {
  it("is a pure mapping that never touches a DataProvider — a fully-matched draft is valid and still unsaved", () => {
    const activityType = REFERENCE_DATA.settings.activityTypes[0];
    const draft = buildWorkRecordDraftFromAnalysis(analysis({ suggestedWorkType: activityType }), REFERENCE_DATA, baseRecord());
    expect(draft.metadata.version).toBe(0); // save() would still route this through createWorkRecord, never update
    expect(draft.title).toBe("Suggested title");
    expect(draft.description).toBe("Suggested description");
    expect(validateWorkRecord(draft, REFERENCE_DATA).valid).toBe(true);
  });

  it("matches an activityType only on an exact case-insensitive name, else leaves it for the user", () => {
    const exact = REFERENCE_DATA.settings.activityTypes[0];
    const matched = buildWorkRecordDraftFromAnalysis(analysis({ suggestedWorkType: exact.toUpperCase() }), REFERENCE_DATA, baseRecord());
    expect(matched.activityType).toBe(exact);

    const unmatched = buildWorkRecordDraftFromAnalysis(
      analysis({ suggestedWorkType: "Something not in the vocabulary" }),
      REFERENCE_DATA,
      baseRecord(),
    );
    expect(unmatched.activityType).toBe("");
  });

  it("never invents an organization/project — only exact-name matches are connected", () => {
    const realOrg = REFERENCE_DATA.organizations.find((org) => org.type !== "district")!;
    const realProject = REFERENCE_DATA.projects[0];
    const draft = buildWorkRecordDraftFromAnalysis(
      analysis({ organizations: [realOrg.name, "Not A Real Org"], projects: [realProject.name, "Not A Real Project"] }),
      REFERENCE_DATA,
      baseRecord(),
    );
    expect(draft.organizationIds).toEqual([realOrg.appId]);
    expect(draft.projectIds).toEqual([realProject.appId]);
  });

  it("sets engagementScope to specific only when a canonical district is matched, keeping the draft valid", () => {
    const district = REFERENCE_DATA.organizations.find((org) => org.type === "district")!;
    const activityType = REFERENCE_DATA.settings.activityTypes[0];
    const draft = buildWorkRecordDraftFromAnalysis(
      analysis({ districts: [district.name], suggestedWorkType: activityType }),
      REFERENCE_DATA,
      baseRecord(),
    );
    expect(draft.engagementScope).toBe("specific");
    expect(draft.organizationIds).toContain(district.appId);
    expect(validateWorkRecord(draft, REFERENCE_DATA).valid).toBe(true);

    const noDistrictMatch = buildWorkRecordDraftFromAnalysis(analysis({ districts: ["Unknown District"] }), REFERENCE_DATA, baseRecord());
    expect(noDistrictMatch.engagementScope).toBe("none");
    expect(noDistrictMatch.organizationIds).toEqual([]);
  });

  it("prefills follow-up from the extracted action item without fabricating a due date", () => {
    const draft = buildWorkRecordDraftFromAnalysis(
      analysis({ actionItems: [{ action: "Reply to the district", dueDate: "2026-09-05", owner: "me" }], followUp: "" }),
      REFERENCE_DATA,
      baseRecord(),
    );
    expect(draft.followUpNeeded).toBe(true);
    expect(draft.followUpDate).toBe("2026-09-05");
    expect(draft.nextStep).toBe("Reply to the district");
  });
});

describe("buildInboxIntelligenceRecord", () => {
  it("builds a fresh, unsaved durable record with no raw-email-capable field", () => {
    const record = buildInboxIntelligenceRecord(analysis(), "a short excerpt", REFERENCE_DATA, "2026-08-29T12:00:00.000Z");
    expect(record.metadata.version).toBe(0); // provider.create() routes this, never update()
    expect(record.status).toBe("open");
    expect(record.resolvedAt).toBeNull();
    expect(record.linkedWorkRecordAppId).toBeNull();
    expect(record.sourceType).toBe("pasted-email");
    expect(record.sourceExcerpt).toBe("a short excerpt");
    // Exhaustive key check: nothing beyond the documented shape exists (in particular, no rawEmail-shaped field).
    expect(Object.keys(record).sort()).toEqual(
      [
        "appId",
        "schemaVersion",
        "sourceType",
        "analyzedAt",
        "sourceExcerpt",
        "analysis",
        "matchedOrganizationIds",
        "matchedDistrictIds",
        "matchedProjectIds",
        "status",
        "resolvedAt",
        "linkedWorkRecordAppId",
        "metadata",
      ].sort(),
    );
  });

  it("resolves matched*Ids using the same exact-match rule as the Work Record mapping", () => {
    const district = REFERENCE_DATA.organizations.find((org) => org.type === "district")!;
    const project = REFERENCE_DATA.projects[0];
    const record = buildInboxIntelligenceRecord(
      analysis({ districts: [district.name, "Unknown District"], projects: [project.name] }),
      "",
      REFERENCE_DATA,
      "2026-08-29T12:00:00.000Z",
    );
    expect(record.matchedDistrictIds).toEqual([district.appId]);
    expect(record.matchedProjectIds).toEqual([project.appId]);
  });
});

describe("computeInboxIntelligenceSummary", () => {
  it("counts records by status", () => {
    const record = buildInboxIntelligenceRecord(analysis(), "", REFERENCE_DATA, "2026-08-29T12:00:00.000Z");
    const summary = computeInboxIntelligenceSummary([
      { ...record, status: "open" },
      { ...record, status: "open" },
      { ...record, status: "waiting" },
      { ...record, status: "resolved" },
    ]);
    expect(summary).toEqual({ openCount: 2, waitingCount: 1, resolvedCount: 1 });
  });
});

describe("SessionInboxIntelligenceProvider", () => {
  function record() {
    return buildInboxIntelligenceRecord(analysis(), "excerpt", REFERENCE_DATA, "2026-08-29T12:00:00.000Z");
  }

  it("assigns SharePoint-shaped provider metadata on create and lists newest first", async () => {
    const provider = new SessionInboxIntelligenceProvider();
    const first = await provider.create(record());
    const second = await provider.create(record());
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
    const provider = new SessionInboxIntelligenceProvider();
    const created = await provider.create(record());
    if (created.status !== "success") throw new Error("setup failed");
    const duplicate = await provider.create({ ...record(), appId: created.value.appId });
    expect(duplicate.status).toBe("conflict");
  });

  it("increments RecordVersion on update and rejects a stale expectedVersion as a conflict", async () => {
    const provider = new SessionInboxIntelligenceProvider();
    const created = await provider.create(record());
    if (created.status !== "success") throw new Error("setup failed");

    const updated = await provider.update({ ...created.value, status: "waiting" }, 1);
    expect(updated.status).toBe("success");
    if (updated.status === "success") expect(updated.value.metadata.version).toBe(2);

    const stale = await provider.update({ ...created.value, status: "resolved" }, 1);
    expect(stale.status).toBe("conflict");
  });

  it("keeps every instance independent — nothing is shared or durable across instances", async () => {
    const first = new SessionInboxIntelligenceProvider();
    await first.create(record());
    const second = new SessionInboxIntelligenceProvider();
    const list = await second.list();
    expect(list).toEqual({ status: "success", value: [] });
  });
});
