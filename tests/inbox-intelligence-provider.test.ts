import { describe, expect, it } from "vitest";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import { REFERENCE_DATA } from "../lib/reference-data";
import { buildWorkRecordDraftFromAnalysis } from "../lib/inbox-intelligence-work-record";
import { EmailAnalysisSchema, type EmailAnalysis, type InboxIntelligenceRecord } from "../lib/inbox-intelligence-models";
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

describe("SessionInboxIntelligenceProvider", () => {
  function record(overrides: Partial<InboxIntelligenceRecord> = {}): InboxIntelligenceRecord {
    return {
      appId: crypto.randomUUID(),
      createdAt: new Date().toISOString(),
      sourceExcerpt: "excerpt",
      analysis: analysis(),
      linkedWorkRecordAppId: null,
      ...overrides,
    };
  }

  it("lists saved records newest first", () => {
    const provider = new SessionInboxIntelligenceProvider();
    provider.save(record({ createdAt: "2026-08-01T00:00:00Z" }));
    provider.save(record({ createdAt: "2026-08-02T00:00:00Z" }));
    const list = provider.list();
    expect(list).toHaveLength(2);
    expect(list[0].createdAt).toBe("2026-08-02T00:00:00Z");
  });

  it("computes needsAttention and open-action counts from saved records", () => {
    const provider = new SessionInboxIntelligenceProvider();
    provider.save(record({ analysis: analysis({ needsAttention: true, actionItems: [{ action: "a", dueDate: null, owner: "me" }] }) }));
    provider.save(record({ analysis: analysis({ needsAttention: false, actionItems: [] }) }));
    expect(provider.summary()).toEqual({ needsAttentionCount: 1, openActionCount: 1 });
  });

  it("keeps every instance independent — nothing is shared or durable across instances", () => {
    const first = new SessionInboxIntelligenceProvider();
    first.save(record());
    const second = new SessionInboxIntelligenceProvider();
    expect(second.list()).toEqual([]);
  });
});
