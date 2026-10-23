import { z } from "zod";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "./models";
import type { MicrosoftSessionStatus } from "./microsoft-session";

// Patch 9A (Option B) — Work Record DRAFT CHECKPOINT for re-authentication.
//
// The problem: a Work Record save can fail because the Microsoft sign-in expired. The wizard
// keeps the draft, but while it is open the whole app shell is inert, so the only "Sign in"
// control is unreachable — and signing in is a full-page redirect that would throw the draft
// away. This module is the safety net: just before the user-initiated redirect the draft is
// copied into a narrowly scoped, versioned, expiring `sessionStorage` entry; after sign-in it
// is restored into the wizard FOR THE HUMAN TO REVIEW AND SAVE. It is never auto-submitted.
//
// Boundaries (all enforced by tests): sessionStorage only (tab-scoped — closing the tab ends it,
// like the Voice working session); one namespaced key; a 12-hour expiry; no tokens, credentials,
// or unrelated app state — only the draft, the wizard step, and which account it belongs to;
// bound to that account (never restored into a different one); cleared only after that draft is
// successfully saved or the human explicitly discards it. No SharePoint, no localStorage.

export const WORK_RECORD_CHECKPOINT_KEY = "iu-work-tracker.work-record-draft-checkpoint.v1";
export const WORK_RECORD_CHECKPOINT_SCHEMA_VERSION = 1 as const;
/** A checkpoint older than this is discarded instead of restored. */
export const WORK_RECORD_CHECKPOINT_TTL_MS = 12 * 60 * 60 * 1000;

export type CheckpointStorage = Pick<Storage, "getItem" | "setItem" | "removeItem">;

export type WorkRecordCheckpoint = {
  schemaVersion: typeof WORK_RECORD_CHECKPOINT_SCHEMA_VERSION;
  /** ISO timestamp the checkpoint was written (drives expiry). */
  savedAt: string;
  /** MSAL homeAccountId of the account that was signed in when the draft was checkpointed, or null if unknown. */
  accountId: string | null;
  /** That account's username, display only (e.g. "restore as someone@example.edu"). */
  accountLabel: string | null;
  /** The wizard step (1–5) the human was on. */
  step: number;
  draft: WorkRecord;
};

export function getCheckpointStorage(): CheckpointStorage | null {
  try {
    if (typeof window === "undefined") return null;
    return window.sessionStorage ?? null;
  } catch {
    return null;
  }
}

const reachSchema = z.object({
  educatorsLeaders: z.number(),
  studentsFamilies: z.number(),
  workforceCommunity: z.number(),
  other: z.number(),
});

// STRUCTURAL validation against the current Work Record model (types, enums, the record
// schemaVersion). Deliberately NOT a reference-ID check: durable Projects/Contacts/Organizations
// load asynchronously, so an ID may be legitimately unknown at restore time — the explicit Save
// still runs the full validation (validateWorkRecord) exactly as for any other entry.
const draftSchema = z.object({
  appId: z.string().min(1),
  title: z.string(),
  activityDate: z.string(),
  activityType: z.string(),
  description: z.string(),
  detailedNotes: z.string(),
  durationMinutes: z.number(),
  status: z.enum(["complete", "draft"]),
  engagementScope: z.enum(["none", "specific", "regional", "allDistricts"]),
  projectIds: z.array(z.string()),
  organizationIds: z.array(z.string()),
  contactIds: z.array(z.string()),
  categoryIds: z.array(z.string()),
  reach: reachSchema,
  evidenceSummary: z.string(),
  evidenceReferenceIds: z.array(z.string()),
  output: z.string(),
  outcome: z.string(),
  nextStep: z.string(),
  followUpNeeded: z.boolean(),
  followUpDate: z.string().nullable(),
  orbit: z.object({
    reportable: z.boolean(),
    primaryDeliverable: z.string().nullable(),
    supportingDeliverables: z.array(z.string()),
    stemPocMinutes: z.number(),
    tacMinutes: z.number(),
    evidence: z.string(),
  }),
  schemaVersion: z.literal(WORK_RECORD_SCHEMA_VERSION),
  metadata: z.object({
    providerId: z.string().optional(),
    version: z.number(),
    createdAt: z.string(),
    modifiedAt: z.string(),
    syncState: z.enum(["saved", "saving", "sync-error"]).optional(),
  }),
  isSample: z.boolean(),
});

const checkpointSchema = z.object({
  schemaVersion: z.literal(WORK_RECORD_CHECKPOINT_SCHEMA_VERSION),
  savedAt: z.string(),
  accountId: z.string().min(1).nullable(),
  accountLabel: z.string().nullable(),
  step: z.number().int().min(1).max(5),
  draft: draftSchema,
});

export function removeWorkRecordCheckpoint(storage: CheckpointStorage | null): boolean {
  if (!storage) return false;
  try {
    storage.removeItem(WORK_RECORD_CHECKPOINT_KEY);
    return true;
  } catch {
    return false;
  }
}

/**
 * Writes the checkpoint and READS IT BACK. Returns false if it could not be stored (storage
 * unavailable, blocked, or full) — the caller must then NOT start the sign-in redirect, because
 * redirecting without a safe copy would lose the entry.
 */
export function saveWorkRecordCheckpoint(
  storage: CheckpointStorage | null,
  input: { accountId: string | null; accountLabel: string | null; step: number; draft: WorkRecord },
  now: () => number = Date.now,
): boolean {
  if (!storage) return false;
  const checkpoint: WorkRecordCheckpoint = {
    schemaVersion: WORK_RECORD_CHECKPOINT_SCHEMA_VERSION,
    savedAt: new Date(now()).toISOString(),
    accountId: input.accountId,
    accountLabel: input.accountLabel,
    step: Math.min(5, Math.max(1, Math.round(input.step))),
    draft: input.draft,
  };
  try {
    const serialized = JSON.stringify(checkpoint);
    storage.setItem(WORK_RECORD_CHECKPOINT_KEY, serialized);
    return storage.getItem(WORK_RECORD_CHECKPOINT_KEY) === serialized;
  } catch {
    return false;
  }
}

export type CheckpointLoadResult =
  | { status: "none" }
  | { status: "found"; checkpoint: WorkRecordCheckpoint }
  /** Present but unusable — already removed. `reason` is for an honest user message. */
  | { status: "discarded"; reason: "expired" | "damaged" | "outdated" };

/**
 * Reads the checkpoint. Never throws and never surfaces a parse error. An expired, damaged
 * (malformed / structurally invalid), or outdated (different checkpoint or Work Record schema
 * version) checkpoint is removed and reported so the app can say so instead of restoring junk.
 */
export function loadWorkRecordCheckpoint(storage: CheckpointStorage | null, now: () => number = Date.now): CheckpointLoadResult {
  if (!storage) return { status: "none" };
  let raw: string | null;
  try {
    raw = storage.getItem(WORK_RECORD_CHECKPOINT_KEY);
  } catch {
    return { status: "none" };
  }
  if (raw === null) return { status: "none" };

  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    removeWorkRecordCheckpoint(storage);
    return { status: "discarded", reason: "damaged" };
  }

  const result = checkpointSchema.safeParse(parsed);
  if (!result.success) {
    removeWorkRecordCheckpoint(storage);
    return { status: "discarded", reason: isOutdated(parsed) ? "outdated" : "damaged" };
  }

  const savedAt = Date.parse(result.data.savedAt);
  if (!Number.isFinite(savedAt) || now() - savedAt > WORK_RECORD_CHECKPOINT_TTL_MS || savedAt - now() > WORK_RECORD_CHECKPOINT_TTL_MS) {
    removeWorkRecordCheckpoint(storage);
    return { status: "discarded", reason: Number.isFinite(savedAt) ? "expired" : "damaged" };
  }
  // Compile-time guarantee that the schema and the WorkRecord type cannot drift apart.
  const checkpoint: WorkRecordCheckpoint = result.data;
  return { status: "found", checkpoint };
}

function isOutdated(value: unknown): boolean {
  if (!value || typeof value !== "object") return false;
  const record = value as { schemaVersion?: unknown; draft?: { schemaVersion?: unknown } };
  return (
    (typeof record.schemaVersion === "number" && record.schemaVersion !== WORK_RECORD_CHECKPOINT_SCHEMA_VERSION) ||
    (typeof record.draft?.schemaVersion === "number" && record.draft.schemaVersion !== WORK_RECORD_SCHEMA_VERSION)
  );
}

/**
 * Clears the checkpoint ONLY if it belongs to this draft (same appId). A save or an explicit
 * discard of some unrelated entry therefore never destroys a different draft's safe copy.
 */
export function clearWorkRecordCheckpointForDraft(storage: CheckpointStorage | null, draftAppId: string): boolean {
  if (!storage) return false;
  try {
    const raw = storage.getItem(WORK_RECORD_CHECKPOINT_KEY);
    if (raw === null) return false;
    let owner: unknown;
    try {
      owner = (JSON.parse(raw) as { draft?: { appId?: unknown } })?.draft?.appId;
    } catch {
      owner = undefined; // unreadable → it is junk, no draft can claim it: remove below
    }
    if (owner === undefined || owner === draftAppId) {
      storage.removeItem(WORK_RECORD_CHECKPOINT_KEY);
      return true;
    }
    return false;
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------------------------------------
// Recovery decision — pure. Never restores into a different or unconfirmed account on its own.
// ---------------------------------------------------------------------------------------------

export type RecoveryDecision =
  /** Same account, signed in, nothing in the way: restore into the wizard (for review — never submit). */
  | { kind: "restore" }
  /** Signed in as a DIFFERENT account: never restore; keep the checkpoint for its owner. */
  | { kind: "different-account"; ownerLabel: string | null; currentLabel: string | null }
  /** The owner cannot sign in right now (signed out / expired / failed / choosing): keep it, explain. */
  | { kind: "sign-in-needed"; ownerLabel: string | null; reason: "signed-out" | "sign-in-required" | "auth-failed" | "account-choice-required" }
  /** The checkpoint has no account identity (or Microsoft isn't configured): the human must choose. */
  | { kind: "unconfirmed" };

export function evaluateCheckpointRecovery(checkpoint: WorkRecordCheckpoint, session: MicrosoftSessionStatus | null): RecoveryDecision {
  if (checkpoint.accountId === null || session === null) return { kind: "unconfirmed" };
  const ownerLabel = checkpoint.accountLabel;
  switch (session.status) {
    case "connected":
      return session.account.homeAccountId === checkpoint.accountId
        ? { kind: "restore" }
        : { kind: "different-account", ownerLabel, currentLabel: session.account.username ?? null };
    case "sign-in-required":
      // A cached-but-expired account that is NOT the owner is a different person's session.
      return session.account.homeAccountId === checkpoint.accountId
        ? { kind: "sign-in-needed", ownerLabel, reason: "sign-in-required" }
        : { kind: "different-account", ownerLabel, currentLabel: session.account.username ?? null };
    case "auth-failed":
      return { kind: "sign-in-needed", ownerLabel, reason: "auth-failed" };
    case "account-choice-required":
      return { kind: "sign-in-needed", ownerLabel, reason: "account-choice-required" };
    case "signed-out":
    default:
      return { kind: "sign-in-needed", ownerLabel, reason: "signed-out" };
  }
}

export type RecoveryResolution =
  | { status: "none" }
  | { status: "discarded"; reason: "expired" | "damaged" | "outdated" }
  | { status: "found"; checkpoint: WorkRecordCheckpoint; decision: RecoveryDecision };

/**
 * Loads the checkpoint (if any) and decides what to do with it. `inspectSession` is only
 * consulted when a usable checkpoint exists, so a normal page load costs nothing extra.
 */
export async function resolveCheckpointRecovery(
  storage: CheckpointStorage | null,
  inspectSession: () => Promise<MicrosoftSessionStatus | null>,
  now: () => number = Date.now,
): Promise<RecoveryResolution> {
  const loaded = loadWorkRecordCheckpoint(storage, now);
  if (loaded.status === "none") return { status: "none" };
  if (loaded.status === "discarded") return loaded;
  let session: MicrosoftSessionStatus | null = null;
  try {
    session = await inspectSession();
  } catch {
    session = null;
  }
  return { status: "found", checkpoint: loaded.checkpoint, decision: evaluateCheckpointRecovery(loaded.checkpoint, session) };
}
