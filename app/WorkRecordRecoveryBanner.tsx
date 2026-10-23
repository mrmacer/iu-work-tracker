"use client";

// Patch 9A (Option B) — shown only when a Work Record draft was set aside before a Microsoft
// re-sign-in (lib/work-record-draft-checkpoint.ts) and could NOT simply be restored on its own:
// the owner isn't signed in yet, a different account is signed in, the owner can't be confirmed,
// the wizard is already in use, or the saved copy was unusable. A pure view — every action is
// passed in, and nothing here writes anywhere. A draft is never restored into a different
// account, and never submitted: restoring only reopens the form for the human to review.

import type { WorkRecordCheckpoint } from "../lib/work-record-draft-checkpoint";

export type RecoveryPrompt =
  | { kind: "ready"; checkpoint: WorkRecordCheckpoint }
  | { kind: "unconfirmed"; checkpoint: WorkRecordCheckpoint }
  | { kind: "different-account"; ownerLabel: string | null; currentLabel: string | null }
  | {
      kind: "sign-in-needed";
      ownerLabel: string | null;
      reason: "signed-out" | "sign-in-required" | "auth-failed" | "account-choice-required";
    }
  | { kind: "unusable"; reason: "expired" | "damaged" | "outdated" };

const UNUSABLE_REASON: Record<"expired" | "damaged" | "outdated", string> = {
  expired: "it was kept for too long",
  damaged: "it was damaged",
  outdated: "it was saved by an older version of the app",
};

export default function WorkRecordRecoveryBanner({
  prompt,
  onRestore,
  onDiscard,
  onDismiss,
}: {
  prompt: RecoveryPrompt;
  onRestore: (checkpoint: WorkRecordCheckpoint) => void;
  onDiscard: () => void;
  onDismiss: () => void;
}) {
  const owner = "ownerLabel" in prompt && prompt.ownerLabel ? prompt.ownerLabel : "the account you were signed in with";

  return (
    <div className="storage-banner attention" role="alert" aria-label="Unfinished work entry">
      <span>Unfinished work entry</span>{" "}
      {prompt.kind === "unusable" && (
        <>
          A work entry that was set aside before signing in could not be restored ({UNUSABLE_REASON[prompt.reason]}) and was removed.
          <div className="recovery-actions">
            <button type="button" className="ghost-button" onClick={onDismiss}>
              Dismiss
            </button>
          </div>
        </>
      )}
      {prompt.kind === "sign-in-needed" && (
        <>
          An unfinished work entry is kept safe in this tab. Sign in with Microsoft as {owner} (top right) to restore it. It will not be
          submitted — you review it and save it yourself.
          <div className="recovery-actions">
            <button type="button" className="ghost-button" onClick={onDiscard}>
              Discard saved entry
            </button>
          </div>
        </>
      )}
      {prompt.kind === "different-account" && (
        <>
          An unfinished work entry belongs to {owner}, but you are signed in as {prompt.currentLabel ?? "a different account"}. It is not
          restored into a different account. Sign in as {owner} to restore it, or discard it.
          <div className="recovery-actions">
            <button type="button" className="ghost-button" onClick={onDiscard}>
              Discard saved entry
            </button>
          </div>
        </>
      )}
      {prompt.kind === "unconfirmed" && (
        <>
          An unfinished work entry is kept safe in this tab, but the account it belongs to cannot be confirmed. Restore it only if it is
          yours. It will not be submitted.
          <div className="recovery-actions">
            <button type="button" className="secondary-button" onClick={() => onRestore(prompt.checkpoint)}>
              Restore entry
            </button>
            <button type="button" className="ghost-button" onClick={onDiscard}>
              Discard saved entry
            </button>
          </div>
        </>
      )}
      {prompt.kind === "ready" && (
        <>
          An unfinished work entry from before you signed in is ready to restore. Restoring reopens it in the entry form for you to review —
          it is not submitted.
          <div className="recovery-actions">
            <button type="button" className="secondary-button" onClick={() => onRestore(prompt.checkpoint)}>
              Restore entry
            </button>
            <button type="button" className="ghost-button" onClick={onDiscard}>
              Discard saved entry
            </button>
          </div>
        </>
      )}
    </div>
  );
}
