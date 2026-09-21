// Patch 7.3 — Voice Intelligence SELECT + ROUTE: the pure routing helpers
// (lib/voice-routing.ts) and the additive session fields that carry routing state
// (lib/voice-intelligence-session.ts). No I/O, no AI, no providers.
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { REFERENCE_DATA } from "../lib/reference-data";
import type { Organization, Project } from "../lib/models";
import { VOICE_CANDIDATE_TYPES } from "../lib/voice-intelligence-models";
import {
  VOICE_SESSION_STORAGE_KEY,
  VOICE_SESSION_SCHEMA_VERSION,
  loadVoiceSession,
  saveVoiceSession,
  type VoiceReviewCandidate,
  type VoiceWorkingSessionV1,
} from "../lib/voice-intelligence-session";
import {
  VOICE_DESTINATIONS,
  formatCandidateForCopy,
  groupSelectedCandidates,
  hasVoiceDestination,
  matchOrganizationCandidates,
  matchProjectCandidates,
  processedChipLabel,
} from "../lib/voice-routing";
import { createFakeStorage } from "./voice-session-test-utils";

function candidate(overrides: Partial<VoiceReviewCandidate> = {}): VoiceReviewCandidate {
  return {
    id: "c1",
    type: "COMPLETED_WORK",
    title: "A title",
    detail: "Some detail.",
    sourceExcerpt: "an excerpt",
    durationText: null,
    selected: true,
    ...overrides,
  };
}

describe("destinations — only four durable destinations exist", () => {
  it("routes exactly Completed work, Person, Organization, District, and Project", () => {
    const ready = VOICE_CANDIDATE_TYPES.filter(hasVoiceDestination);
    expect(ready).toEqual(["COMPLETED_WORK", "PERSON", "ORGANIZATION", "DISTRICT", "PROJECT"]);
    expect(VOICE_DESTINATIONS.COMPLETED_WORK).toEqual({ kind: "work-record", label: "Work Records" });
    expect(VOICE_DESTINATIONS.PERSON.kind).toBe("contact");
    expect(VOICE_DESTINATIONS.ORGANIZATION.kind).toBe("organization");
    expect(VOICE_DESTINATIONS.DISTRICT.kind).toBe("organization"); // District is Organization.type === "district"
    expect(VOICE_DESTINATIONS.PROJECT.kind).toBe("project");
  });

  it("honestly has NO destination for Action, Idea, Decision, Question, and Knowledge", () => {
    for (const type of ["ACTION", "IDEA", "DECISION", "QUESTION", "KNOWLEDGE"] as const) {
      expect(hasVoiceDestination(type)).toBe(false);
      expect(VOICE_DESTINATIONS[type].kind).toBe("none");
    }
  });
});

describe("groupSelectedCandidates", () => {
  it("groups only SELECTED candidates by type, ready destinations apart from unsupported ones", () => {
    const groups = groupSelectedCandidates([
      candidate({ id: "w1", type: "COMPLETED_WORK" }),
      candidate({ id: "w2", type: "COMPLETED_WORK" }),
      candidate({ id: "p1", type: "PERSON" }),
      candidate({ id: "a1", type: "ACTION" }),
      candidate({ id: "k1", type: "KNOWLEDGE" }),
      candidate({ id: "w3", type: "COMPLETED_WORK", selected: false }),
    ]);
    expect(groups.ready.map((g) => [g.type, g.candidates.map((c) => c.id)])).toEqual([
      ["COMPLETED_WORK", ["w1", "w2"]],
      ["PERSON", ["p1"]],
    ]);
    expect(groups.unsupported.map((g) => g.type)).toEqual(["ACTION", "KNOWLEDGE"]);
  });

  it("never drops an unsupported type and never includes an unselected candidate", () => {
    const groups = groupSelectedCandidates([
      candidate({ id: "i1", type: "IDEA" }),
      candidate({ id: "d1", type: "DECISION" }),
      candidate({ id: "q1", type: "QUESTION" }),
      candidate({ id: "x", type: "PROJECT", selected: false }),
    ]);
    expect(groups.unsupported.map((g) => g.type)).toEqual(["IDEA", "DECISION", "QUESTION"]);
    expect(groups.ready).toEqual([]);
  });

  it("returns nothing when nothing is selected", () => {
    expect(groupSelectedCandidates([candidate({ selected: false })])).toEqual({ ready: [], unsupported: [] });
  });
});

describe("matchOrganizationCandidates — exact normalized name only", () => {
  const orgs: Organization[] = [
    { appId: "d1", name: "North Valley SD", type: "district" },
    { appId: "p1", name: "FutureWorks Partnership", type: "partner" },
    { appId: "i1", name: "Intermediate Unit", type: "iu" },
    { appId: "p2", name: "north  valley sd", type: "partner" }, // same normalized name, different type
  ];

  it("matches case- and whitespace-insensitively", () => {
    expect(matchOrganizationCandidates("  FUTUREWORKS   partnership ", orgs, "ORGANIZATION").map((o) => o.appId)).toEqual(["p1"]);
  });

  it("an ORGANIZATION candidate matches only non-district Organizations; a DISTRICT candidate only districts", () => {
    expect(matchOrganizationCandidates("North Valley SD", orgs, "DISTRICT").map((o) => o.appId)).toEqual(["d1"]);
    expect(matchOrganizationCandidates("North Valley SD", orgs, "ORGANIZATION").map((o) => o.appId)).toEqual(["p2"]);
  });

  it("returns EVERY exact match so a same-name ambiguity reaches the human", () => {
    const dupes: Organization[] = [
      { appId: "a", name: "Riverbend", type: "partner" },
      { appId: "b", name: "riverbend", type: "iu" },
    ];
    expect(matchOrganizationCandidates("Riverbend", dupes, "ORGANIZATION").map((o) => o.appId)).toEqual(["a", "b"]);
  });

  it("never fuzzy-matches: partial, prefix, typo, and empty names find nothing", () => {
    for (const name of ["North Valley", "Nort Valley SD", "FutureWorks", "", "   "]) {
      expect(matchOrganizationCandidates(name, orgs, "DISTRICT")).toEqual([]);
      expect(matchOrganizationCandidates(name, orgs, "ORGANIZATION")).toEqual([]);
    }
  });
});

describe("matchProjectCandidates — exact normalized name only", () => {
  it("strong match: exact normalized project name", () => {
    expect(matchProjectCandidates("steels   implementation", REFERENCE_DATA.projects).map((p) => p.appId)).toEqual(["project-steels"]);
  });

  it("ambiguous: every exact match is returned for human selection", () => {
    const projects: Project[] = [
      { appId: "a", name: "Makerspace", description: "", status: "active", color: "blue" },
      { appId: "b", name: "makerspace", description: "", status: "planning", color: "coral" },
    ];
    expect(matchProjectCandidates("Makerspace", projects).map((p) => p.appId)).toEqual(["a", "b"]);
  });

  it("no fuzzy matching: a partial or reworded name finds nothing", () => {
    expect(matchProjectCandidates("STEELS", REFERENCE_DATA.projects)).toEqual([]);
    expect(matchProjectCandidates("Steels Implementation Project", REFERENCE_DATA.projects)).toEqual([]);
    expect(matchProjectCandidates("", REFERENCE_DATA.projects)).toEqual([]);
  });
});

describe("processedChipLabel — only from a recorded successful outcome", () => {
  it("is null for an untouched candidate of every type", () => {
    for (const type of VOICE_CANDIDATE_TYPES) expect(processedChipLabel(candidate({ type }))).toBeNull();
  });

  it("labels matched vs created Organization / Project / Contact", () => {
    expect(processedChipLabel(candidate({ type: "ORGANIZATION", routing: { type: "matched", entityAppId: "o" } }))).toBe("Matched Organization ✓");
    expect(processedChipLabel(candidate({ type: "DISTRICT", routing: { type: "created", entityAppId: "o" } }))).toBe("Created Organization ✓");
    expect(processedChipLabel(candidate({ type: "PROJECT", routing: { type: "matched", entityAppId: "p" } }))).toBe("Matched Project ✓");
    expect(processedChipLabel(candidate({ type: "PROJECT", routing: { type: "created", entityAppId: "p" } }))).toBe("Created Project ✓");
    expect(processedChipLabel(candidate({ type: "PERSON", contactDecision: { type: "matched", contactAppId: "c" }, contactCreated: true }))).toBe("Created Contact ✓");
  });

  it("an ignored decision or a merely matched Contact earns no processed chip", () => {
    expect(processedChipLabel(candidate({ type: "PROJECT", routing: { type: "ignored" } }))).toBeNull();
    expect(processedChipLabel(candidate({ type: "PERSON", contactDecision: { type: "matched", contactAppId: "c" } }))).toBeNull();
  });
});

describe("formatCandidateForCopy", () => {
  it("formats type, title, and detail as plain text — and omits an empty detail", () => {
    expect(formatCandidateForCopy(candidate({ type: "KNOWLEDGE", title: "Kits need labels", detail: "  Label every bin.  " }))).toBe(
      "Knowledge: Kits need labels\nLabel every bin.",
    );
    expect(formatCandidateForCopy(candidate({ type: "ACTION", title: "Call Kim", detail: " " }))).toBe("Action: Call Kim");
  });
});

describe("session schema — routing state is additive and restores", () => {
  function session(candidates: VoiceReviewCandidate[], extra: Partial<VoiceWorkingSessionV1> = {}): VoiceWorkingSessionV1 {
    return {
      schemaVersion: VOICE_SESSION_SCHEMA_VERSION,
      transcript: "t",
      phase: "review",
      analyzed: true,
      candidates,
      usage: null,
      savedAt: "",
      ...extra,
    };
  }

  it("round-trips routing decisions, contactCreated, logged state, selection, and routingOpen", () => {
    const storage = createFakeStorage();
    saveVoiceSession(
      storage,
      session(
        [
          candidate({ id: "o", type: "ORGANIZATION", routing: { type: "created", entityAppId: "org-1" } }),
          candidate({ id: "p", type: "PROJECT", routing: { type: "matched", entityAppId: "project-steels" }, selected: false }),
          candidate({ id: "n", type: "PERSON", contactDecision: { type: "matched", contactAppId: "c1" }, contactCreated: true }),
          candidate({ id: "i", type: "PROJECT", routing: { type: "ignored" } }),
          candidate({ id: "w", loggedWorkRecordAppId: "wr-1" }),
        ],
        { routingOpen: true },
      ),
    );
    const loaded = loadVoiceSession(storage)!;
    expect(loaded.routingOpen).toBe(true);
    expect(loaded.candidates.map((c) => c.routing ?? null)).toEqual([
      { type: "created", entityAppId: "org-1" },
      { type: "matched", entityAppId: "project-steels" },
      null,
      { type: "ignored" },
      null,
    ]);
    expect(loaded.candidates[1].selected).toBe(false);
    expect(loaded.candidates[2].contactCreated).toBe(true);
    expect(loaded.candidates[4].loggedWorkRecordAppId).toBe("wr-1");
  });

  it("a Patch 7.2 session (no routing fields at all) still restores — the schema change is purely additive", () => {
    const storage = createFakeStorage();
    storage.data.set(
      VOICE_SESSION_STORAGE_KEY,
      JSON.stringify({
        schemaVersion: 1,
        transcript: "old",
        phase: "review",
        analyzed: true,
        candidates: [{ id: "x", type: "IDEA", title: "t", detail: "", sourceExcerpt: "s", durationText: null, selected: true }],
        usage: null,
        savedAt: "2026-09-20T00:00:00.000Z",
      }),
    );
    const loaded = loadVoiceSession(storage);
    expect(loaded).not.toBeNull();
    expect(loaded!.candidates[0].routing).toBeUndefined();
    expect(loaded!.routingOpen).toBeUndefined();
  });

  it("rejects a malformed routing decision instead of restoring it", () => {
    const storage = createFakeStorage();
    storage.data.set(
      VOICE_SESSION_STORAGE_KEY,
      JSON.stringify({
        schemaVersion: 1,
        transcript: "t",
        phase: "review",
        analyzed: true,
        candidates: [{ id: "x", type: "PROJECT", title: "t", detail: "", sourceExcerpt: "s", durationText: null, selected: true, routing: { type: "matched" } }],
        usage: null,
        savedAt: "",
      }),
    );
    expect(loadVoiceSession(storage)).toBeNull();
  });
});

describe("module boundary", () => {
  it("lib/voice-routing.ts is pure — no fetch, storage, provider, SharePoint, or AI reference", () => {
    const code = readFileSync("lib/voice-routing.ts", "utf-8").replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
    expect(code).not.toMatch(/fetch\(|sessionStorage|localStorage|sharepoint|anthropic|createOrganization|createProject|saveProject|saveOrganization/i);
  });
});
