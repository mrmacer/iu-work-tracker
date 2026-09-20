// Patch 7.2 — the Voice Intelligence working-session helper (lib/voice-intelligence-session.ts),
// tested directly against a small in-memory fake Storage (this environment's global jsdom
// Storage is unreliable under Node 25 — see tests/voice-session-test-utils.ts).
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import {
  VOICE_SESSION_SCHEMA_VERSION,
  VOICE_SESSION_STORAGE_KEY,
  clearVoiceSession,
  emptyVoiceSession,
  isVoiceSessionEmpty,
  loadVoiceSession,
  saveVoiceSession,
  withCandidateLogged,
  type VoiceReviewCandidate,
  type VoiceSessionStorage,
  type VoiceWorkingSessionV1,
} from "../lib/voice-intelligence-session";
import { createFakeStorage } from "./voice-session-test-utils";

function candidate(overrides: Partial<VoiceReviewCandidate> = {}): VoiceReviewCandidate {
  return {
    id: "cand-1",
    type: "COMPLETED_WORK",
    title: "Met with North Schuylkill about science resources",
    detail: "Went well, about an hour long.",
    sourceExcerpt: "I met with North Schuylkill this morning",
    durationText: "about an hour",
    selected: true,
    ...overrides,
  };
}

function session(overrides: Partial<VoiceWorkingSessionV1> = {}): VoiceWorkingSessionV1 {
  return {
    schemaVersion: VOICE_SESSION_SCHEMA_VERSION,
    transcript: "I met with North Schuylkill this morning about science resources.",
    phase: "review",
    analyzed: true,
    candidates: [candidate(), candidate({ id: "cand-2", type: "ACTION", title: "Send Kim the materials", detail: "", durationText: null })],
    usage: { model: "claude-opus-5", inputTokens: 900, outputTokens: 420 },
    savedAt: "",
    ...overrides,
  };
}

function roundTrip(input: VoiceWorkingSessionV1): VoiceWorkingSessionV1 {
  const storage = createFakeStorage();
  expect(saveVoiceSession(storage, input)).toBe(true);
  const loaded = loadVoiceSession(storage);
  expect(loaded).not.toBeNull();
  return loaded!;
}

describe("storage key and schema version", () => {
  it("uses one clearly namespaced, versioned key and schema version 1", () => {
    expect(VOICE_SESSION_STORAGE_KEY).toBe("iu-work-tracker.voice-intelligence-session.v1");
    expect(VOICE_SESSION_SCHEMA_VERSION).toBe(1);
  });
});

describe("save + load round trip", () => {
  it("round-trips the whole session and stamps savedAt", () => {
    const loaded = roundTrip(session());
    expect(loaded.schemaVersion).toBe(1);
    expect(loaded.candidates).toHaveLength(2);
    expect(Number.isNaN(Date.parse(loaded.savedAt))).toBe(false);
  });

  it("preserves the transcript", () => {
    expect(roundTrip(session({ transcript: "A long\nmulti-line ramble — with “quotes” and emoji 🎙" })).transcript).toBe(
      "A long\nmulti-line ramble — with “quotes” and emoji 🎙",
    );
  });

  it("preserves the candidate collection in order", () => {
    expect(roundTrip(session()).candidates.map((c) => c.id)).toEqual(["cand-1", "cand-2"]);
  });

  it("preserves an edited title, edited detail, and edited type", () => {
    const edited = session({
      candidates: [candidate({ title: "Human-edited title", detail: "Human-edited detail", type: "DECISION" })],
    });
    const [loaded] = roundTrip(edited).candidates;
    expect(loaded.title).toBe("Human-edited title");
    expect(loaded.detail).toBe("Human-edited detail");
    expect(loaded.type).toBe("DECISION");
  });

  it("restores an in-progress edit that the AI-output schema would reject (cleared title/detail)", () => {
    const [loaded] = roundTrip(session({ candidates: [candidate({ title: "", detail: "" })] })).candidates;
    expect(loaded.title).toBe("");
    expect(loaded.detail).toBe("");
  });

  it("preserves selected and deselected state", () => {
    const loaded = roundTrip(session({ candidates: [candidate({ selected: false }), candidate({ id: "cand-2", selected: true })] }));
    expect(loaded.candidates.map((c) => c.selected)).toEqual([false, true]);
  });

  it("preserves sourceExcerpt and durationText (including a removed/null duration)", () => {
    const loaded = roundTrip(session({ candidates: [candidate(), candidate({ id: "cand-2", durationText: null })] }));
    expect(loaded.candidates[0].sourceExcerpt).toBe("I met with North Schuylkill this morning");
    expect(loaded.candidates[0].durationText).toBe("about an hour");
    expect(loaded.candidates[1].durationText).toBeNull();
  });

  it("preserves the stable candidate id exactly", () => {
    const loaded = roundTrip(session({ candidates: [candidate({ id: "3f2c8d0e-stable-id" })] }));
    expect(loaded.candidates[0].id).toBe("3f2c8d0e-stable-id");
  });

  it("preserves analyzed state, phase, and usage", () => {
    const loaded = roundTrip(session({ analyzed: true, phase: "review" }));
    expect(loaded.analyzed).toBe(true);
    expect(loaded.phase).toBe("review");
    expect(loaded.usage).toEqual({ model: "claude-opus-5", inputTokens: 900, outputTokens: 420 });
  });

  it("preserves an un-analyzed transcript-only session without inventing candidates", () => {
    const loaded = roundTrip(session({ analyzed: false, phase: "paste", candidates: [], usage: null }));
    expect(loaded.analyzed).toBe(false);
    expect(loaded.candidates).toEqual([]);
    expect(loaded.transcript).toContain("North Schuylkill");
  });

  it("preserves a PERSON candidate's Contact-match decision and a candidate's Logged marker", () => {
    const loaded = roundTrip(
      session({
        candidates: [
          candidate({ id: "p1", type: "PERSON", contactDecision: { type: "matched", contactAppId: "contact-1" } }),
          candidate({ id: "p2", type: "PERSON", contactDecision: { type: "ignored" } }),
          candidate({ id: "w1", loggedWorkRecordAppId: "work-record-7" }),
        ],
      }),
    );
    expect(loaded.candidates[0].contactDecision).toEqual({ type: "matched", contactAppId: "contact-1" });
    expect(loaded.candidates[1].contactDecision).toEqual({ type: "ignored" });
    expect(loaded.candidates[2].loggedWorkRecordAppId).toBe("work-record-7");
  });

  it("stores only the Voice session under its one key — nothing else", () => {
    const storage = createFakeStorage();
    saveVoiceSession(storage, session());
    expect([...storage.data.keys()]).toEqual([VOICE_SESSION_STORAGE_KEY]);
    const raw = storage.data.get(VOICE_SESSION_STORAGE_KEY)!;
    expect(raw).not.toMatch(/access_?token|api_?key|authorization|bearer|secret|password/i); // inputTokens/outputTokens are model usage counts, not credentials
  });
});

describe("invalid stored data fails safely", () => {
  function storageWith(raw: string): ReturnType<typeof createFakeStorage> {
    const storage = createFakeStorage();
    storage.data.set(VOICE_SESSION_STORAGE_KEY, raw);
    return storage;
  }

  it("returns null and removes the entry for malformed JSON — without throwing", () => {
    const storage = storageWith("{not json");
    expect(loadVoiceSession(storage)).toBeNull();
    expect(storage.data.has(VOICE_SESSION_STORAGE_KEY)).toBe(false);
  });

  it("returns null and removes the entry for a wrong schemaVersion", () => {
    const storage = storageWith(JSON.stringify({ ...session(), schemaVersion: 2 }));
    expect(loadVoiceSession(storage)).toBeNull();
    expect(storage.data.has(VOICE_SESSION_STORAGE_KEY)).toBe(false);
  });

  it("returns null for missing required structure (no candidates array / no transcript)", () => {
    const noCandidates = { ...session() } as Record<string, unknown>;
    delete noCandidates.candidates;
    expect(loadVoiceSession(storageWith(JSON.stringify(noCandidates)))).toBeNull();
    const noTranscript = { ...session() } as Record<string, unknown>;
    delete noTranscript.transcript;
    expect(loadVoiceSession(storageWith(JSON.stringify(noTranscript)))).toBeNull();
  });

  it("returns null for a candidate with an unknown type or a missing id", () => {
    expect(loadVoiceSession(storageWith(JSON.stringify({ ...session(), candidates: [{ ...candidate(), type: "NOT_A_TYPE" }] })))).toBeNull();
    const noId = { ...candidate() } as Record<string, unknown>;
    delete noId.id;
    expect(loadVoiceSession(storageWith(JSON.stringify({ ...session(), candidates: [noId] })))).toBeNull();
  });

  it("returns null for non-object JSON (null, array, string)", () => {
    for (const raw of ["null", "[]", '"a string"', "42"]) expect(loadVoiceSession(storageWith(raw))).toBeNull();
  });

  it("returns null when nothing is stored, or when storage is unavailable", () => {
    expect(loadVoiceSession(createFakeStorage())).toBeNull();
    expect(loadVoiceSession(null)).toBeNull();
  });
});

describe("clear", () => {
  it("removes the session so a later load finds nothing", () => {
    const storage = createFakeStorage();
    saveVoiceSession(storage, session());
    expect(loadVoiceSession(storage)).not.toBeNull();
    expect(clearVoiceSession(storage)).toBe(true);
    expect(storage.data.has(VOICE_SESSION_STORAGE_KEY)).toBe(false);
    expect(loadVoiceSession(storage)).toBeNull();
  });

  it("saving an empty session removes the key instead of storing an empty shell", () => {
    const storage = createFakeStorage();
    saveVoiceSession(storage, session());
    expect(saveVoiceSession(storage, emptyVoiceSession())).toBe(true);
    expect(storage.data.has(VOICE_SESSION_STORAGE_KEY)).toBe(false);
    expect(isVoiceSessionEmpty(emptyVoiceSession())).toBe(true);
    expect(isVoiceSessionEmpty(session({ transcript: "   ", analyzed: false, candidates: [] }))).toBe(true);
    expect(isVoiceSessionEmpty(session({ analyzed: true, candidates: [] }))).toBe(false); // analyzed → "no useful candidates" is still a session
  });
});

describe("storage failure safety", () => {
  const throwing = (message: string): VoiceSessionStorage => ({
    getItem: () => {
      throw new Error(message);
    },
    setItem: () => {
      throw new Error(message);
    },
    removeItem: () => {
      throw new Error(message);
    },
  });

  it("a storage READ failure returns null instead of throwing", () => {
    expect(() => loadVoiceSession(throwing("SecurityError"))).not.toThrow();
    expect(loadVoiceSession(throwing("SecurityError"))).toBeNull();
  });

  it("a storage WRITE failure (quota/blocked) returns false instead of throwing", () => {
    expect(() => saveVoiceSession(throwing("QuotaExceededError"), session())).not.toThrow();
    expect(saveVoiceSession(throwing("QuotaExceededError"), session())).toBe(false);
  });

  it("a storage REMOVE failure returns false instead of throwing", () => {
    expect(() => clearVoiceSession(throwing("SecurityError"))).not.toThrow();
    expect(clearVoiceSession(throwing("SecurityError"))).toBe(false);
  });

  it("missing storage (null) fails safely on every operation", () => {
    expect(saveVoiceSession(null, session())).toBe(false);
    expect(clearVoiceSession(null)).toBe(false);
    expect(loadVoiceSession(null)).toBeNull();
  });
});

describe("withCandidateLogged", () => {
  it("marks only the named candidate and leaves every other candidate untouched", () => {
    const before = session();
    const after = withCandidateLogged(before, "cand-1", "work-record-1");
    expect(after.candidates[0].loggedWorkRecordAppId).toBe("work-record-1");
    expect(after.candidates[1]).toEqual(before.candidates[1]);
    expect(after.transcript).toBe(before.transcript);
  });

  it("leaves the session unchanged for an unknown candidate id", () => {
    const before = session();
    expect(withCandidateLogged(before, "nope", "work-record-1").candidates).toEqual(before.candidates);
  });
});

describe("module boundary — temporary sessionStorage only", () => {
  it("references no localStorage, SharePoint, Graph, Anthropic, IndexedDB, or cookie API", () => {
    const source = readFileSync("lib/voice-intelligence-session.ts", "utf-8");
    // Strip comments so the explanatory header (which names what is NOT used) doesn't count.
    const code = source.replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
    expect(code).not.toMatch(/localStorage|indexedDB|document\.cookie|sharepoint|graph|anthropic|fetch\(/i);
    expect(code).toMatch(/sessionStorage/);
  });
});
