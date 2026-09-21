import { z } from "zod";
import type { ContactMatchDecision } from "./contact-matching";
import { VoiceCandidateTypeSchema, type VoiceCandidate } from "./voice-intelligence-models";

// Patch 7.2 — Voice Intelligence WORKING SESSION persistence.
//
// ANALYZE ONCE. KEEP THE WORKING SESSION. A pasted transcript can segment into many candidates
// that the human processes one at a time; navigating around Work Tracker (including the
// Log-as-work handoff, which ends on another screen after a successful save) must not throw the
// analysis away and force a second paid Anthropic call.
//
// This is deliberately TEMPORARY, per-tab state in `sessionStorage` — never localStorage,
// SharePoint, the DataProvider, IndexedDB, cookies, or the URL. A transcript can be sensitive
// drive-home reflection; it survives navigation, unmount/remount, the Work Record handoff, and a
// same-tab refresh, and it disappears when the tab closes. It is NOT durable institutional data
// and nothing here talks to Microsoft Graph or Anthropic. See docs/AI_HANDOFF.md "Voice
// Intelligence working session (Patch 7.2)".

export const VOICE_SESSION_STORAGE_KEY = "iu-work-tracker.voice-intelligence-session.v1";
export const VOICE_SESSION_SCHEMA_VERSION = 1 as const;

/** The only Storage surface this module needs — lets tests pass a small in-memory fake. */
export type VoiceSessionStorage = Pick<Storage, "getItem" | "setItem" | "removeItem">;

/**
 * Local browser review state only — an AI candidate plus the fields the review UI needs that
 * the model never produces: a STABLE id (must survive save/restore so edits and Logged state keep
 * their identity), whether the user has it selected, a PERSON candidate's transient Contact-match
 * review decision (Patch 8D), and — Patch 7.2 — the appId of a Work Record the human successfully
 * saved from this candidate. `loggedWorkRecordAppId` is Voice-session UI state ONLY: it is never
 * written to a Work Record or SharePoint.
 */
/**
 * Patch 7.3 — the human's routing decision for an ORGANIZATION / DISTRICT / PROJECT candidate.
 * `matched` links an EXISTING durable Organization/Project (by appId); `created` records the
 * durable entity the human explicitly created from this candidate through the existing form;
 * `ignored` means "not for this item". Session UI state only — never written to the entity.
 */
export type VoiceRoutingDecision =
  // `entityName` is display-only: the entity's name at decision time, used when the entity is not
  // (or no longer) in the current reference list — e.g. an in-memory Preview record after a refresh.
  | { type: "matched"; entityAppId: string; entityName?: string }
  | { type: "created"; entityAppId: string; entityName?: string }
  | { type: "ignored" };

export type VoiceReviewCandidate = VoiceCandidate & {
  id: string;
  /**
   * Patch 7.3 — CHECKBOX SEMANTICS: `selected` means "include this candidate in the current
   * processing batch" (the Process selected tray). Unchecked candidates stay in the working
   * session untouched, and are NOT processed right now — unchecking never deletes anything.
   * Only "Remove" deletes a candidate.
   */
  selected: boolean;
  contactDecision?: ContactMatchDecision;
  /** PERSON only: the Contact behind a "matched" contactDecision was created here via Add Person. */
  contactCreated?: boolean;
  /** ORGANIZATION / DISTRICT / PROJECT only — see VoiceRoutingDecision. */
  routing?: VoiceRoutingDecision;
  loggedWorkRecordAppId?: string | null;
};

export type VoiceSessionUsage = { model: string; inputTokens: number; outputTokens: number };

export type VoiceWorkingSessionV1 = {
  schemaVersion: typeof VOICE_SESSION_SCHEMA_VERSION;
  transcript: string;
  /** Which Voice screen the human was on, so a restore reproduces exactly what they saw. */
  phase: "paste" | "review";
  /** True once an explicit Analyze succeeded — `candidates` are then the reviewed analysis. */
  analyzed: boolean;
  candidates: VoiceReviewCandidate[];
  usage: VoiceSessionUsage | null;
  /** Patch 7.3 — whether the "Process selected" panel is open (UI state, restored with the session). */
  routingOpen?: boolean;
  savedAt: string;
};

// Deliberately LENIENT relative to VoiceCandidateSchema (the AI-output schema): a human may have
// cleared a title or detail mid-edit, and that in-progress edit must restore as-is rather than
// being rejected for violating the model's min-length rules.
const StoredCandidateSchema = z.object({
  type: VoiceCandidateTypeSchema,
  title: z.string(),
  detail: z.string(),
  sourceExcerpt: z.string(),
  durationText: z.string().nullable(),
  id: z.string().min(1),
  selected: z.boolean(),
  contactDecision: z
    .union([
      z.object({ type: z.literal("matched"), contactAppId: z.string().min(1) }),
      z.object({ type: z.literal("ignored") }),
    ])
    .optional(),
  contactCreated: z.boolean().optional(),
  routing: z
    .union([
      z.object({ type: z.literal("matched"), entityAppId: z.string().min(1), entityName: z.string().optional() }),
      z.object({ type: z.literal("created"), entityAppId: z.string().min(1), entityName: z.string().optional() }),
      z.object({ type: z.literal("ignored") }),
    ])
    .optional(),
  loggedWorkRecordAppId: z.string().nullable().optional(),
});

const StoredSessionSchema = z.object({
  schemaVersion: z.literal(VOICE_SESSION_SCHEMA_VERSION),
  transcript: z.string(),
  phase: z.enum(["paste", "review"]),
  analyzed: z.boolean(),
  candidates: z.array(StoredCandidateSchema).max(100),
  usage: z
    .object({ model: z.string(), inputTokens: z.number(), outputTokens: z.number() })
    .nullable(),
  routingOpen: z.boolean().optional(),
  savedAt: z.string(),
});

/** The browser's sessionStorage, or null when it is unavailable/blocked (e.g. SSR, privacy modes). */
export function getBrowserSessionStorage(): VoiceSessionStorage | null {
  try {
    if (typeof window === "undefined") return null;
    return window.sessionStorage ?? null;
  } catch {
    return null;
  }
}

export function emptyVoiceSession(): VoiceWorkingSessionV1 {
  return {
    schemaVersion: VOICE_SESSION_SCHEMA_VERSION,
    transcript: "",
    phase: "paste",
    analyzed: false,
    candidates: [],
    usage: null,
    routingOpen: false,
    savedAt: "",
  };
}

/** Nothing worth keeping: no transcript text and no analysis. Such a session is removed, not stored. */
export function isVoiceSessionEmpty(session: VoiceWorkingSessionV1): boolean {
  return !session.analyzed && session.candidates.length === 0 && session.transcript.trim().length === 0;
}

export function clearVoiceSession(storage: VoiceSessionStorage | null): boolean {
  if (!storage) return false;
  try {
    storage.removeItem(VOICE_SESSION_STORAGE_KEY);
    return true;
  } catch {
    // Storage is best-effort: the screen keeps working from in-memory state.
    return false;
  }
}

/**
 * Restores the working session, or returns null. Never throws and never surfaces a parse error:
 * malformed JSON, a wrong schemaVersion, or a structurally invalid session is discarded (and the
 * bad entry removed) so the human simply sees the normal empty Voice screen. Never calls
 * Anthropic — restoring is a pure read.
 */
export function loadVoiceSession(storage: VoiceSessionStorage | null): VoiceWorkingSessionV1 | null {
  if (!storage) return null;
  let raw: string | null;
  try {
    raw = storage.getItem(VOICE_SESSION_STORAGE_KEY);
  } catch {
    return null;
  }
  if (raw === null) return null;
  try {
    const parsed = StoredSessionSchema.safeParse(JSON.parse(raw));
    if (parsed.success) return parsed.data;
  } catch {
    // fall through to discard
  }
  clearVoiceSession(storage);
  return null;
}

/**
 * Writes the working session (or removes the key when the session is empty). Returns whether
 * storage now reflects `session` — false on any storage failure (quota, blocked, unavailable),
 * in which case the caller simply continues with in-memory state.
 */
export function saveVoiceSession(storage: VoiceSessionStorage | null, session: VoiceWorkingSessionV1): boolean {
  if (!storage) return false;
  if (isVoiceSessionEmpty(session)) return clearVoiceSession(storage);
  try {
    storage.setItem(VOICE_SESSION_STORAGE_KEY, JSON.stringify({ ...session, savedAt: new Date().toISOString() }));
    return true;
  } catch {
    return false;
  }
}

/**
 * Pure: marks ONE candidate as Logged. Called only after the existing Work Record save pathway
 * reports success (never on the "Log as work" click) — see app/VoiceIntelligence.tsx. Every other
 * candidate is returned untouched, and an unknown id leaves the session unchanged.
 */
export function withCandidateLogged(
  session: VoiceWorkingSessionV1,
  candidateId: string,
  workRecordAppId: string,
): VoiceWorkingSessionV1 {
  return {
    ...session,
    candidates: session.candidates.map((candidate) =>
      candidate.id === candidateId ? { ...candidate, loggedWorkRecordAppId: workRecordAppId } : candidate,
    ),
  };
}
