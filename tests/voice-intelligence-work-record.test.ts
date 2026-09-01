import { describe, expect, it } from "vitest";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import type { VoiceCandidate } from "../lib/voice-intelligence-models";
import {
  buildWorkRecordDraftFromVoiceCandidate,
  parseDeterministicDurationMinutes,
} from "../lib/voice-intelligence-work-record";

function baseRecord(overrides: Partial<WorkRecord> = {}): WorkRecord {
  return {
    appId: "draft",
    title: "",
    activityDate: "2026-08-31",
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
    ...overrides,
  };
}

function candidate(overrides: Partial<VoiceCandidate> = {}): VoiceCandidate {
  return {
    type: "COMPLETED_WORK",
    title: "Met with North Schuylkill about science resources",
    detail: "Went well, covered the new lab kits.",
    sourceExcerpt: "I met with North Schuylkill this morning about science resources",
    durationText: null,
    ...overrides,
  };
}

describe("parseDeterministicDurationMinutes — explicit forms only", () => {
  const supported: [string, number][] = [
    ["45 minutes", 45],
    ["30 minutes", 30],
    ["1 hour", 60],
    ["about an hour", 60],
    ["2 hours", 120],
    ["about 2 hours", 120],
    ["1.5 hours", 90],
    ["90 minutes", 90],
  ];
  for (const [input, expected] of supported) {
    it(`converts "${input}" to ${expected} minutes`, () => {
      expect(parseDeterministicDurationMinutes(input)).toBe(expected);
    });
  }

  const unsupported = [
    "maybe two hours",
    "most of the morning",
    "a while",
    "forever",
    "a long meeting",
    "a couple hours",
    "really long",
    "took forever",
    "spent a while",
    "a minute",
  ];
  for (const input of unsupported) {
    it(`leaves "${input}" unconverted (returns null)`, () => {
      expect(parseDeterministicDurationMinutes(input)).toBeNull();
    });
  }
});

describe("buildWorkRecordDraftFromVoiceCandidate", () => {
  it("maps candidate title to Work Record title", () => {
    const draft = buildWorkRecordDraftFromVoiceCandidate(candidate({ title: "Met with the STEELS committee" }), baseRecord());
    expect(draft.title).toBe("Met with the STEELS committee");
  });

  it("maps candidate detail to Work Record description", () => {
    const draft = buildWorkRecordDraftFromVoiceCandidate(candidate({ detail: "Finalized the agenda." }), baseRecord());
    expect(draft.description).toBe("Finalized the agenda.");
  });

  it("never inserts sourceExcerpt into description or detailedNotes", () => {
    const draft = buildWorkRecordDraftFromVoiceCandidate(
      candidate({ detail: "Finalized the agenda.", sourceExcerpt: "AI PROVENANCE TEXT SHOULD NOT APPEAR" }),
      baseRecord(),
    );
    expect(draft.description).not.toContain("AI PROVENANCE TEXT SHOULD NOT APPEAR");
    expect(draft.detailedNotes).not.toContain("AI PROVENANCE TEXT SHOULD NOT APPEAR");
  });

  it("converts a safe explicit durationText into durationMinutes", () => {
    const draft = buildWorkRecordDraftFromVoiceCandidate(candidate({ durationText: "about an hour" }), baseRecord({ durationMinutes: 60 }));
    expect(draft.durationMinutes).toBe(60);
    const draft2 = buildWorkRecordDraftFromVoiceCandidate(candidate({ durationText: "45 minutes" }), baseRecord({ durationMinutes: 60 }));
    expect(draft2.durationMinutes).toBe(45);
  });

  it("leaves durationMinutes at the base default when durationText is unsupported", () => {
    const draft = buildWorkRecordDraftFromVoiceCandidate(candidate({ durationText: "maybe two hours" }), baseRecord({ durationMinutes: 60 }));
    expect(draft.durationMinutes).toBe(60);
  });

  it("leaves durationMinutes at the base default when durationText is null", () => {
    const draft = buildWorkRecordDraftFromVoiceCandidate(candidate({ durationText: null }), baseRecord({ durationMinutes: 60 }));
    expect(draft.durationMinutes).toBe(60);
  });

  it("does not infer a category / activityType from candidate text", () => {
    const base = baseRecord({ activityType: "" });
    const draft = buildWorkRecordDraftFromVoiceCandidate(candidate(), base);
    expect(draft.activityType).toBe("");
  });

  it("does not infer a project", () => {
    const draft = buildWorkRecordDraftFromVoiceCandidate(candidate(), baseRecord({ projectIds: [] }));
    expect(draft.projectIds).toEqual([]);
  });

  it("does not infer a district or organization", () => {
    const draft = buildWorkRecordDraftFromVoiceCandidate(candidate(), baseRecord({ organizationIds: [] }));
    expect(draft.organizationIds).toEqual([]);
  });

  it("does not infer reach/reporting or ORBIT", () => {
    const draft = buildWorkRecordDraftFromVoiceCandidate(candidate(), baseRecord());
    expect(draft.orbit.reportable).toBe(false);
    expect(draft.orbit.primaryDeliverable).toBeNull();
    expect(draft.reach).toEqual({ educatorsLeaders: 0, studentsFamilies: 0, workforceCommunity: 0, other: 0 });
  });

  it("uses the CURRENT (already-edited) candidate values, not any notion of an original", () => {
    // buildWorkRecordDraftFromVoiceCandidate has no access to "original" AI output at all —
    // it only ever receives whatever candidate object the caller passes in, proving the
    // human-edited-state requirement structurally rather than by simulated edit history.
    const edited = candidate({ title: "Human-edited title", detail: "Human-edited detail" });
    const draft = buildWorkRecordDraftFromVoiceCandidate(edited, baseRecord());
    expect(draft.title).toBe("Human-edited title");
    expect(draft.description).toBe("Human-edited detail");
  });
});
