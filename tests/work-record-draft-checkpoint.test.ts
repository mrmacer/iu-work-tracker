// Patch 9A (Option B) — the Work Record draft checkpoint used to survive a Microsoft re-sign-in.
// Pure module tests against the in-memory fake Storage (this environment's jsdom storage is
// unreliable under Node 25 — see tests/voice-session-test-utils.ts).
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import type { AccountInfo } from "@azure/msal-browser";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import type { MicrosoftSessionStatus } from "../lib/microsoft-session";
import {
  WORK_RECORD_CHECKPOINT_KEY,
  WORK_RECORD_CHECKPOINT_SCHEMA_VERSION,
  WORK_RECORD_CHECKPOINT_TTL_MS,
  clearWorkRecordCheckpointForDraft,
  evaluateCheckpointRecovery,
  loadWorkRecordCheckpoint,
  removeWorkRecordCheckpoint,
  resolveCheckpointRecovery,
  saveWorkRecordCheckpoint,
  type CheckpointStorage,
  type WorkRecordCheckpoint,
} from "../lib/work-record-draft-checkpoint";
import { createFakeStorage } from "./voice-session-test-utils";

/** A draft with EVERY field set to a non-default value, so a dropped field cannot hide. */
function fullDraft(overrides: Partial<WorkRecord> = {}): WorkRecord {
  return {
    appId: "draft-1",
    title: "District STEELS planning meeting",
    activityDate: "2026-09-18",
    activityType: "District meeting",
    description: "Planned the fall rollout.",
    detailedNotes: "Long notes —\nwith a newline, “quotes”, and emoji 🎓.",
    durationMinutes: 95,
    status: "complete",
    engagementScope: "specific",
    projectIds: ["project-steels", "project-ai"],
    organizationIds: ["org-north-valley", "org-futureworks"],
    contactIds: ["contact-1"],
    categoryIds: ["cat-1", "cat-2"],
    reach: { educatorsLeaders: 12, studentsFamilies: 30, workforceCommunity: 4, other: 1 },
    evidenceSummary: "Agenda and sign-in sheet.",
    evidenceReferenceIds: ["ev-1"],
    output: "Rollout plan",
    outcome: "Agreed on timeline",
    nextStep: "Send the draft calendar",
    followUpNeeded: true,
    followUpDate: "2026-09-25",
    orbit: { reportable: true, primaryDeliverable: "deliverable-1", supportingDeliverables: ["deliverable-2"], stemPocMinutes: 45, tacMinutes: 20, evidence: "Orbit evidence" },
    schemaVersion: WORK_RECORD_SCHEMA_VERSION,
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
    isSample: false,
    ...overrides,
  };
}

const T0 = Date.parse("2026-09-21T10:00:00.000Z");
const owner = { homeAccountId: "owner-id", username: "owner@example.edu", name: "Owner" } as AccountInfo;
const stranger = { homeAccountId: "stranger-id", username: "stranger@example.edu", name: "Stranger" } as AccountInfo;

function write(storage: CheckpointStorage, overrides: Partial<{ accountId: string | null; accountLabel: string | null; step: number; draft: WorkRecord; now: number }> = {}) {
  return saveWorkRecordCheckpoint(
    storage,
    { accountId: "owner-id", accountLabel: "owner@example.edu", step: 4, draft: fullDraft(), ...overrides },
    () => overrides.now ?? T0,
  );
}

function found(storage: CheckpointStorage, now = T0): WorkRecordCheckpoint {
  const result = loadWorkRecordCheckpoint(storage, () => now);
  if (result.status !== "found") throw new Error(`expected a checkpoint, got ${result.status}`);
  return result.checkpoint;
}

describe("key, version, and expiry policy", () => {
  it("uses one narrowly scoped, versioned key, schema version 1, and a documented 12-hour expiry", () => {
    expect(WORK_RECORD_CHECKPOINT_KEY).toBe("iu-work-tracker.work-record-draft-checkpoint.v1");
    expect(WORK_RECORD_CHECKPOINT_SCHEMA_VERSION).toBe(1);
    expect(WORK_RECORD_CHECKPOINT_TTL_MS).toBe(12 * 60 * 60 * 1000);
  });
});

describe("round trip — every draft field survives", () => {
  it("restores the entire Work Record draft exactly, plus the step and the owning account", () => {
    const storage = createFakeStorage();
    expect(write(storage)).toBe(true);
    const checkpoint = found(storage);
    expect(checkpoint.draft).toEqual(fullDraft());
    expect(checkpoint.step).toBe(4);
    expect(checkpoint.accountId).toBe("owner-id");
    expect(checkpoint.accountLabel).toBe("owner@example.edu");
    expect(checkpoint.savedAt).toBe(new Date(T0).toISOString());
  });

  it.each([
    ["title / description / notes", ["title", "description", "detailedNotes"] as const],
    ["date / duration / type", ["activityDate", "durationMinutes", "activityType"] as const],
    ["project / organization / contact / category selections", ["projectIds", "organizationIds", "contactIds", "categoryIds"] as const],
    ["scope / reach / evidence", ["engagementScope", "reach", "evidenceSummary", "evidenceReferenceIds"] as const],
    ["output / outcome / next step / follow-up", ["output", "outcome", "nextStep", "followUpNeeded", "followUpDate"] as const],
    ["ORBIT block", ["orbit"] as const],
  ])("preserves %s", (_label, fields) => {
    const storage = createFakeStorage();
    write(storage);
    const restored = found(storage).draft;
    for (const field of fields) expect(restored[field]).toEqual(fullDraft()[field]);
  });

  it("preserves the wizard step for every step 1–5", () => {
    for (const step of [1, 2, 3, 4, 5]) {
      const storage = createFakeStorage();
      write(storage, { step });
      expect(found(storage).step).toBe(step);
    }
  });

  it("preserves a draft with unusual-but-valid content (empty selections, null follow-up, an existing record's version)", () => {
    const storage = createFakeStorage();
    const draft = fullDraft({ projectIds: [], organizationIds: [], contactIds: [], followUpNeeded: false, followUpDate: null, metadata: { providerId: "7", version: 3, createdAt: "a", modifiedAt: "b", syncState: "sync-error" } });
    write(storage, { draft });
    expect(found(storage).draft).toEqual(draft);
  });

  it("stores ONLY the draft, step, and account — nothing token- or app-state-like", () => {
    const storage = createFakeStorage();
    write(storage);
    expect([...storage.data.keys()]).toEqual([WORK_RECORD_CHECKPOINT_KEY]);
    const stored = JSON.parse(storage.data.get(WORK_RECORD_CHECKPOINT_KEY)!);
    expect(Object.keys(stored).sort()).toEqual(["accountId", "accountLabel", "draft", "savedAt", "schemaVersion", "step"]);
    expect(storage.data.get(WORK_RECORD_CHECKPOINT_KEY)).not.toMatch(/access_?token|refresh_?token|id_?token|password|secret|bearer|authorization/i);
  });

  it("works with an unknown account (accountId null)", () => {
    const storage = createFakeStorage();
    write(storage, { accountId: null, accountLabel: null });
    expect(found(storage).accountId).toBeNull();
  });
});

describe("corrupt, outdated, and expired checkpoints are handled gracefully", () => {
  const seed = (raw: string) => {
    const storage = createFakeStorage();
    storage.data.set(WORK_RECORD_CHECKPOINT_KEY, raw);
    return storage;
  };

  it("nothing stored → none", () => {
    expect(loadWorkRecordCheckpoint(createFakeStorage(), () => T0)).toEqual({ status: "none" });
    expect(loadWorkRecordCheckpoint(null, () => T0)).toEqual({ status: "none" });
  });

  it("malformed JSON → discarded as damaged, removed, and never throws", () => {
    const storage = seed("{ not json");
    expect(loadWorkRecordCheckpoint(storage, () => T0)).toEqual({ status: "discarded", reason: "damaged" });
    expect(storage.data.has(WORK_RECORD_CHECKPOINT_KEY)).toBe(false);
  });

  it.each([["null"], ["[]"], ['"text"'], ["42"], ["{}"]])("structurally invalid JSON %s → damaged and removed", (raw) => {
    const storage = seed(raw);
    expect(loadWorkRecordCheckpoint(storage, () => T0)).toEqual({ status: "discarded", reason: "damaged" });
    expect(storage.data.has(WORK_RECORD_CHECKPOINT_KEY)).toBe(false);
  });

  it("a draft missing required fields or with a wrong type → damaged", () => {
    const storage = createFakeStorage();
    write(storage);
    const stored = JSON.parse(storage.data.get(WORK_RECORD_CHECKPOINT_KEY)!);
    delete stored.draft.reach;
    storage.data.set(WORK_RECORD_CHECKPOINT_KEY, JSON.stringify(stored));
    expect(loadWorkRecordCheckpoint(storage, () => T0)).toEqual({ status: "discarded", reason: "damaged" });

    write(storage);
    const wrongType = JSON.parse(storage.data.get(WORK_RECORD_CHECKPOINT_KEY)!);
    wrongType.draft.durationMinutes = "ninety";
    storage.data.set(WORK_RECORD_CHECKPOINT_KEY, JSON.stringify(wrongType));
    expect(loadWorkRecordCheckpoint(storage, () => T0).status).toBe("discarded");
  });

  it("an out-of-range step is damaged", () => {
    const storage = createFakeStorage();
    write(storage);
    const stored = JSON.parse(storage.data.get(WORK_RECORD_CHECKPOINT_KEY)!);
    stored.step = 9;
    storage.data.set(WORK_RECORD_CHECKPOINT_KEY, JSON.stringify(stored));
    expect(loadWorkRecordCheckpoint(storage, () => T0)).toEqual({ status: "discarded", reason: "damaged" });
  });

  it("a different CHECKPOINT schema version is outdated, not restored", () => {
    const storage = createFakeStorage();
    write(storage);
    const stored = JSON.parse(storage.data.get(WORK_RECORD_CHECKPOINT_KEY)!);
    stored.schemaVersion = 2;
    storage.data.set(WORK_RECORD_CHECKPOINT_KEY, JSON.stringify(stored));
    expect(loadWorkRecordCheckpoint(storage, () => T0)).toEqual({ status: "discarded", reason: "outdated" });
    expect(storage.data.has(WORK_RECORD_CHECKPOINT_KEY)).toBe(false);
  });

  it("a draft from an OLDER Work Record model version is outdated, not restored", () => {
    const storage = createFakeStorage();
    write(storage);
    const stored = JSON.parse(storage.data.get(WORK_RECORD_CHECKPOINT_KEY)!);
    stored.draft.schemaVersion = WORK_RECORD_SCHEMA_VERSION - 1;
    storage.data.set(WORK_RECORD_CHECKPOINT_KEY, JSON.stringify(stored));
    expect(loadWorkRecordCheckpoint(storage, () => T0)).toEqual({ status: "discarded", reason: "outdated" });
  });

  it("expires after exactly the TTL: still valid at the limit, discarded one millisecond later", () => {
    const storage = createFakeStorage();
    write(storage);
    expect(loadWorkRecordCheckpoint(storage, () => T0 + WORK_RECORD_CHECKPOINT_TTL_MS).status).toBe("found");
    expect(loadWorkRecordCheckpoint(storage, () => T0 + WORK_RECORD_CHECKPOINT_TTL_MS + 1)).toEqual({ status: "discarded", reason: "expired" });
    expect(storage.data.has(WORK_RECORD_CHECKPOINT_KEY)).toBe(false);
  });

  it("a timestamp far in the FUTURE (clock tampering) is not trusted forever", () => {
    const storage = createFakeStorage();
    write(storage, { now: T0 + 30 * 24 * 60 * 60 * 1000 });
    expect(loadWorkRecordCheckpoint(storage, () => T0).status).toBe("discarded");
  });

  it("an unparseable timestamp is damaged", () => {
    const storage = createFakeStorage();
    write(storage);
    const stored = JSON.parse(storage.data.get(WORK_RECORD_CHECKPOINT_KEY)!);
    stored.savedAt = "not a date";
    storage.data.set(WORK_RECORD_CHECKPOINT_KEY, JSON.stringify(stored));
    expect(loadWorkRecordCheckpoint(storage, () => T0)).toEqual({ status: "discarded", reason: "damaged" });
  });
});

describe("storage failure", () => {
  const throwing = (): CheckpointStorage => ({
    getItem: () => { throw new Error("SecurityError"); },
    setItem: () => { throw new Error("QuotaExceededError"); },
    removeItem: () => { throw new Error("SecurityError"); },
  });

  it("a write failure returns false (so the caller must NOT redirect) instead of throwing", () => {
    expect(() => write(throwing())).not.toThrow();
    expect(write(throwing())).toBe(false);
  });

  it("missing storage returns false for every operation", () => {
    expect(saveWorkRecordCheckpoint(null, { accountId: null, accountLabel: null, step: 1, draft: fullDraft() })).toBe(false);
    expect(removeWorkRecordCheckpoint(null)).toBe(false);
    expect(clearWorkRecordCheckpointForDraft(null, "x")).toBe(false);
  });

  it("a write that does not read back identically (silently dropped/truncated) is reported as a failure", () => {
    const lossy: CheckpointStorage = { getItem: () => null, setItem: () => undefined, removeItem: () => undefined };
    expect(write(lossy)).toBe(false);
  });

  it("read and remove failures never throw", () => {
    expect(loadWorkRecordCheckpoint(throwing(), () => T0)).toEqual({ status: "none" });
    expect(removeWorkRecordCheckpoint(throwing())).toBe(false);
    expect(clearWorkRecordCheckpointForDraft(throwing(), "x")).toBe(false);
  });
});

describe("clearing — only for the same draft, or an explicit removal", () => {
  it("clearing for the SAME draft appId removes the checkpoint", () => {
    const storage = createFakeStorage();
    write(storage);
    expect(clearWorkRecordCheckpointForDraft(storage, "draft-1")).toBe(true);
    expect(storage.data.has(WORK_RECORD_CHECKPOINT_KEY)).toBe(false);
  });

  it("clearing for a DIFFERENT draft (a save or discard of some other entry) leaves it untouched", () => {
    const storage = createFakeStorage();
    write(storage);
    expect(clearWorkRecordCheckpointForDraft(storage, "some-other-draft")).toBe(false);
    expect(storage.data.has(WORK_RECORD_CHECKPOINT_KEY)).toBe(true);
    expect(found(storage).draft.appId).toBe("draft-1");
  });

  it("nothing stored → nothing to clear", () => {
    expect(clearWorkRecordCheckpointForDraft(createFakeStorage(), "draft-1")).toBe(false);
  });

  it("explicit removal removes it", () => {
    const storage = createFakeStorage();
    write(storage);
    expect(removeWorkRecordCheckpoint(storage)).toBe(true);
    expect(storage.data.has(WORK_RECORD_CHECKPOINT_KEY)).toBe(false);
  });
});

describe("evaluateCheckpointRecovery — never into a different or unconfirmed account", () => {
  const storage = createFakeStorage();
  write(storage);
  const checkpoint = found(storage);

  it("same signed-in account → restore", () => {
    expect(evaluateCheckpointRecovery(checkpoint, { status: "connected", account: owner })).toEqual({ kind: "restore" });
  });

  it("a DIFFERENT signed-in account → different-account (never restore)", () => {
    expect(evaluateCheckpointRecovery(checkpoint, { status: "connected", account: stranger })).toEqual({
      kind: "different-account",
      ownerLabel: "owner@example.edu",
      currentLabel: "stranger@example.edu",
    });
  });

  it("a different account that merely has an expired cached session is also never the owner", () => {
    expect(evaluateCheckpointRecovery(checkpoint, { status: "sign-in-required", account: stranger }).kind).toBe("different-account");
  });

  it("the owner's own expired session → sign-in-needed (no restore until authenticated)", () => {
    expect(evaluateCheckpointRecovery(checkpoint, { status: "sign-in-required", account: owner })).toEqual({
      kind: "sign-in-needed",
      ownerLabel: "owner@example.edu",
      reason: "sign-in-required",
    });
  });

  it.each<[MicrosoftSessionStatus, string]>([
    [{ status: "signed-out" }, "signed-out"],
    [{ status: "auth-failed", account: null, message: "x" }, "auth-failed"],
    [{ status: "account-choice-required", accounts: [owner, stranger] }, "account-choice-required"],
  ])("%j → sign-in-needed (%s), never restored", (session, reason) => {
    expect(evaluateCheckpointRecovery(checkpoint, session)).toMatchObject({ kind: "sign-in-needed", reason });
  });

  it("a checkpoint with no account identity requires an explicit human decision, even when signed in", () => {
    const anonymous = { ...checkpoint, accountId: null, accountLabel: null };
    expect(evaluateCheckpointRecovery(anonymous, { status: "connected", account: owner })).toEqual({ kind: "unconfirmed" });
  });

  it("when Microsoft sign-in is not available at all (no session info), the identity cannot be confirmed → explicit decision", () => {
    expect(evaluateCheckpointRecovery(checkpoint, null)).toEqual({ kind: "unconfirmed" });
  });
});

describe("resolveCheckpointRecovery", () => {
  it("does not even inspect the Microsoft session when there is no checkpoint", async () => {
    let inspected = false;
    const result = await resolveCheckpointRecovery(createFakeStorage(), async () => { inspected = true; return null; }, () => T0);
    expect(result).toEqual({ status: "none" });
    expect(inspected).toBe(false);
  });

  it("reports an unusable checkpoint without inspecting the session", async () => {
    const storage = createFakeStorage();
    storage.data.set(WORK_RECORD_CHECKPOINT_KEY, "garbage");
    let inspected = false;
    expect(await resolveCheckpointRecovery(storage, async () => { inspected = true; return null; }, () => T0)).toEqual({ status: "discarded", reason: "damaged" });
    expect(inspected).toBe(false);
  });

  it("decides restore for the owner, and treats an inspection that throws as 'cannot confirm'", async () => {
    const storage = createFakeStorage();
    write(storage);
    const ok = await resolveCheckpointRecovery(storage, async () => ({ status: "connected", account: owner }), () => T0);
    expect(ok).toMatchObject({ status: "found", decision: { kind: "restore" } });
    const failed = await resolveCheckpointRecovery(storage, async () => { throw new Error("boom"); }, () => T0);
    expect(failed).toMatchObject({ status: "found", decision: { kind: "unconfirmed" } });
  });

  it("leaves the checkpoint in place — resolving never consumes it", async () => {
    const storage = createFakeStorage();
    write(storage);
    await resolveCheckpointRecovery(storage, async () => ({ status: "connected", account: owner }), () => T0);
    expect(storage.data.has(WORK_RECORD_CHECKPOINT_KEY)).toBe(true);
  });
});

describe("module boundary", () => {
  it("is sessionStorage-scoped only: no localStorage, cookies, SharePoint, Graph, fetch, or token handling", () => {
    const code = readFileSync("lib/work-record-draft-checkpoint.ts", "utf-8").replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
    expect(code).not.toMatch(/localStorage|indexedDB|document\.cookie|sharepoint|graph\.microsoft|fetch\(|acquireGraphToken|accessToken/i);
    expect(code).toMatch(/sessionStorage/);
  });
});
