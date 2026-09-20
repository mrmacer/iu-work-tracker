"use client";

import { useRef, useState } from "react";
import type { AnalyzeTranscriptResult } from "../lib/anthropic-voice-analysis";
import type { Contact, ReferenceData, WorkRecord } from "../lib/models";
import { MAX_TRANSCRIPT_LENGTH } from "../lib/voice-intelligence-config";
import {
  VOICE_CANDIDATE_TYPES,
  VOICE_CANDIDATE_TYPE_LABELS,
  type VoiceCandidate,
  type VoiceCandidateType,
} from "../lib/voice-intelligence-models";
import { buildWorkRecordDraftFromVoiceCandidate } from "../lib/voice-intelligence-work-record";
import {
  emptyVoiceSession,
  getBrowserSessionStorage,
  loadVoiceSession,
  saveVoiceSession,
  withCandidateLogged,
  type VoiceReviewCandidate,
  type VoiceSessionStorage,
  type VoiceWorkingSessionV1,
} from "../lib/voice-intelligence-session";
import ContactFormModal, { emptyContactDraft } from "./ContactFormModal";
import ContactMatchPanel from "./ContactMatchPanel";
import { matchContactCandidates } from "../lib/contact-matching";
import type { ContactResult } from "../lib/contact-provider";

/**
 * Local browser review state only — see VoiceReviewCandidate in lib/voice-intelligence-session.ts.
 * Patch 7.2: this state now lives in a TEMPORARY per-tab working session (sessionStorage) so it
 * survives navigation/remount, but it is still never written to SharePoint, Work Records, Inbox
 * Intelligence, Organizations, Projects, a Knowledge Base, or localStorage — Voice Intelligence
 * still has no durable persistence (see docs/AI_HANDOFF.md "Voice Intelligence working session
 * (Patch 7.2)"). A Contact created via "Add Person" IS durably saved (it reuses the real Contact
 * creation path, see app/ContactFormModal.tsx), but the fact that THIS transcript's candidate
 * matches it lives only in the tab-scoped working session.
 */
type ReviewCandidate = VoiceReviewCandidate;

function toReviewCandidates(candidates: VoiceCandidate[]): ReviewCandidate[] {
  return candidates.map((candidate) => ({ ...candidate, id: crypto.randomUUID(), selected: true }));
}

export default function VoiceIntelligence({
  openLog,
  createDraftRecord,
  references,
  saveContact,
  updateContact,
  storage,
}: {
  openLog: (record?: WorkRecord, onSaved?: (saved: WorkRecord) => void) => void;
  createDraftRecord: () => WorkRecord;
  references: ReferenceData;
  saveContact: (contact: Contact) => Promise<ContactResult<Contact>>;
  updateContact: (contact: Contact, expectedVersion: number) => Promise<ContactResult<Contact>>;
  /** Test seam only: production passes nothing and the browser's sessionStorage is used. `null` disables persistence. */
  storage?: VoiceSessionStorage | null;
}) {
  // Patch 7.2 — the ONE temporary, per-tab Voice working session. Restored on mount with a pure
  // sessionStorage read (never an Anthropic call); every meaningful change goes through commit().
  const [sessionStore] = useState<VoiceSessionStorage | null>(() =>
    storage !== undefined ? storage : getBrowserSessionStorage(),
  );
  const [restoredSession] = useState<VoiceWorkingSessionV1 | null>(() => loadVoiceSession(sessionStore));
  const [session, setSession] = useState<VoiceWorkingSessionV1>(() => restoredSession ?? emptyVoiceSession());
  // Always the latest committed session, so callbacks that outlive a render (the Work Record
  // save callback, the async analyze result) never write a stale snapshot.
  const latestSession = useRef(session);
  const [persisted, setPersisted] = useState(true); // false once a storage write has failed
  const [restored, setRestored] = useState(restoredSession !== null);
  const [analyzing, setAnalyzing] = useState(false);
  const [error, setError] = useState("");

  const { transcript, phase, candidates, usage } = session;

  const commit = (update: (current: VoiceWorkingSessionV1) => VoiceWorkingSessionV1) => {
    const next = update(latestSession.current);
    latestSession.current = next;
    setSession(next);
    // Storage failure never blocks the screen: state above is already updated in memory.
    setPersisted(saveVoiceSession(sessionStore, next));
    setRestored(false);
  };

  // No AI request happens until this is explicitly called by the "Analyze transcript" click —
  // loading this screen, and loading Home, both cost zero Anthropic requests.
  const analyze = async () => {
    if (analyzing) return; // guards against double-click/rapid-Enter re-entrancy
    if (!transcript.trim()) {
      setError("Paste a transcript before analyzing.");
      return;
    }
    if (transcript.length > MAX_TRANSCRIPT_LENGTH) {
      setError(
        `This transcript is too long (${transcript.length.toLocaleString()} of ${MAX_TRANSCRIPT_LENGTH.toLocaleString()} characters allowed). Split it into smaller pieces and analyze each separately.`,
      );
      return;
    }
    setAnalyzing(true);
    setError("");
    try {
      const response = await fetch("/api/voice-intelligence", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ rawTranscript: transcript }),
      });
      const result = (await response.json().catch(() => null)) as AnalyzeTranscriptResult | null;
      if (!result || result.status !== "success") {
        setError(result?.message ?? "The transcript could not be analyzed. Try again.");
        return; // transcript is deliberately left in place — never cleared on failure
      }
      commit((current) => ({
        ...current,
        candidates: toReviewCandidates(result.analysis.candidates),
        usage: result.usage,
        phase: "review",
        analyzed: true,
      }));
    } catch {
      setError("The AI service could not be reached. Check your connection and try again.");
    } finally {
      setAnalyzing(false);
    }
  };

  const patchCandidate = (
    id: string,
    patch: Partial<Pick<ReviewCandidate, "type" | "title" | "detail" | "durationText" | "selected" | "contactDecision">>,
  ) =>
    commit((current) => ({
      ...current,
      candidates: current.candidates.map((candidate) => (candidate.id === id ? { ...candidate, ...patch } : candidate)),
    }));

  const removeCandidate = (id: string) =>
    commit((current) => ({ ...current, candidates: current.candidates.filter((candidate) => candidate.id !== id) }));

  // Patch 8D — deterministic Contact matching for PERSON candidates only. Zero AI calls, zero
  // network calls: see lib/contact-matching.ts. contactDecision is transient exactly like
  // every other Voice review field — see the ReviewCandidate doc comment above.
  const [addPersonCandidateId, setAddPersonCandidateId] = useState<string | null>(null);
  const addPersonCandidate = candidates.find((candidate) => candidate.id === addPersonCandidateId) ?? null;

  // Opens the existing Log Work form prefilled from the CURRENT edited candidate state —
  // never the original model output. Performs zero persistence: no Work Record is created,
  // no provider is called, and this candidate's own review state is left exactly as-is. The
  // human still reviews and explicitly saves through the existing, unmodified save path.
  //
  // Patch 7.2: the working session is persisted FIRST (so it is recoverable no matter where the
  // human goes next), and the existing wizard is opened with an onSaved callback. That callback
  // fires only when the existing Work Record save pathway reports SUCCESS — never on this click,
  // a cancel, a validation failure, or a provider error — and marks just this candidate Logged
  // (Voice-session UI state only; nothing is added to the Work Record).
  const logAsWork = (candidate: ReviewCandidate) => {
    commit((current) => current);
    const draft = buildWorkRecordDraftFromVoiceCandidate(candidate, createDraftRecord());
    openLog(draft, (savedWorkRecord) => commit((current) => withCandidateLogged(current, candidate.id, savedWorkRecord.appId)));
  };

  const backToTranscript = () => commit((current) => ({ ...current, phase: "paste" }));

  // The deliberate destructive action for the temporary working session: clears the transcript,
  // every candidate and review decision, and removes the sessionStorage entry.
  const startOver = () => {
    commit(() => emptyVoiceSession());
    setError("");
    setAddPersonCandidateId(null);
  };

  const sessionActive = persisted && (session.analyzed || transcript.trim().length > 0);

  const selectedCount = candidates.filter((candidate) => candidate.selected).length;

  return (
    <div className="screen-inner">
      <div className="page-heading">
        <div>
          <p className="eyebrow">AI-assisted intake</p>
          <h1>Voice Intelligence</h1>
          <p>
            Paste a transcript from a voice note, meeting reflection, drive-home ramble, or daily debrief. The AI
            will break it into separate candidates for you to review.
          </p>
        </div>
      </div>

      {sessionActive && (
        <p className="muted-copy voice-session-status" role="status">
          <strong>{restored ? "Working session restored." : "Working session saved in this tab."}</strong>{" "}
          Closing this tab clears the temporary Voice Intelligence session.
        </p>
      )}

      {phase === "paste" && (
        <section className="panel">
          <div className="form-stack">
            <label>
              <span>Paste a transcript</span>
              <textarea
                rows={16}
                value={transcript}
                onChange={(event) => commit((current) => ({ ...current, transcript: event.target.value }))}
                placeholder="Paste the transcript — no need to clean it up first. Ramble is fine; the AI will sort it out."
              />
            </label>
            <p className="muted-copy">
              {transcript.length.toLocaleString()} / {MAX_TRANSCRIPT_LENGTH.toLocaleString()} characters
            </p>
            {error && (
              <div className="form-error" role="alert">
                {error}
              </div>
            )}
          </div>
          <footer className="log-footer">
            <button className="ghost-button" onClick={startOver} disabled={analyzing}>
              Clear
            </button>
            <button className="primary-action" onClick={() => void analyze()} disabled={analyzing || !transcript.trim()}>
              {analyzing ? "Analyzing…" : "Analyze transcript"}
            </button>
          </footer>
        </section>
      )}

      {phase === "review" && (
        <section className="panel">
          <p className="eyebrow">AI-suggested — review before doing anything with these</p>
          <p className="muted-copy">Selected candidates are ready for review. Nothing has been saved yet.</p>

          {candidates.length === 0 ? (
            <div className="empty">
              <span>○</span>
              <strong>No useful candidates found</strong>
              <p>The transcript may have been too short or too vague to segment. You can edit it and try again.</p>
              <button onClick={backToTranscript}>Edit transcript</button>
            </div>
          ) : (
            <>
              <div className="candidate-summary">
                <span>{candidates.length} candidate{candidates.length === 1 ? "" : "s"}</span>
                <span>{selectedCount} selected</span>
                <span>{candidates.length - selectedCount} ignored</span>
              </div>
              <div className="candidate-list">
                {candidates.map((candidate) => (
                  <CandidateCard
                    key={candidate.id}
                    candidate={candidate}
                    references={references}
                    onPatch={(patch) => patchCandidate(candidate.id, patch)}
                    onRemove={() => removeCandidate(candidate.id)}
                    onLogAsWork={() => logAsWork(candidate)}
                    onAddPerson={() => setAddPersonCandidateId(candidate.id)}
                  />
                ))}
              </div>
            </>
          )}

          {usage && (
            <p className="muted-copy">
              Model: {usage.model} · {usage.inputTokens.toLocaleString()} in / {usage.outputTokens.toLocaleString()} out tokens
            </p>
          )}

          <footer className="log-footer">
            <button className="primary-action" onClick={startOver}>
              Analyze another transcript
            </button>
          </footer>
        </section>
      )}

      {addPersonCandidate && (
        <ContactFormModal
          contact={{ ...emptyContactDraft(), displayName: addPersonCandidate.title }}
          contacts={references.contacts}
          organizations={references.organizations}
          onCancel={() => setAddPersonCandidateId(null)}
          onSaved={(savedContact) => {
            patchCandidate(addPersonCandidate.id, { contactDecision: { type: "matched", contactAppId: savedContact.appId } });
            setAddPersonCandidateId(null);
          }}
          saveContact={saveContact}
          updateContact={updateContact}
        />
      )}
    </div>
  );
}

function CandidateCard({
  candidate,
  references,
  onPatch,
  onRemove,
  onLogAsWork,
  onAddPerson,
}: {
  candidate: ReviewCandidate;
  references: ReferenceData;
  onPatch: (patch: Partial<Pick<ReviewCandidate, "type" | "title" | "detail" | "durationText" | "selected" | "contactDecision">>) => void;
  onRemove: () => void;
  onLogAsWork: () => void;
  onAddPerson: () => void;
}) {
  return (
    <div className={`candidate-card${candidate.selected ? "" : " deselected"}`}>
      <div className="candidate-card-head">
        <input
          type="checkbox"
          aria-label={candidate.selected ? "Deselect candidate" : "Select candidate"}
          checked={candidate.selected}
          onChange={(event) => onPatch({ selected: event.target.checked })}
        />
        <select
          aria-label="Candidate type"
          value={candidate.type}
          onChange={(event) => onPatch({ type: event.target.value as VoiceCandidateType })}
        >
          {VOICE_CANDIDATE_TYPES.map((type) => (
            <option key={type} value={type}>
              {VOICE_CANDIDATE_TYPE_LABELS[type]}
            </option>
          ))}
        </select>
        {candidate.durationText && (
          <span className="candidate-duration">
            {candidate.durationText}
            <button type="button" aria-label="Remove duration" onClick={() => onPatch({ durationText: null })}>
              ×
            </button>
          </span>
        )}
        {candidate.loggedWorkRecordAppId && <span className="candidate-chip">Logged ✓</span>}
        {candidate.type === "COMPLETED_WORK" &&
          (candidate.loggedWorkRecordAppId ? (
            // Already logged: de-emphasized, not disabled — logging again stays possible on purpose.
            <button type="button" className="ghost-button" onClick={onLogAsWork}>
              Log again
            </button>
          ) : (
            <button type="button" className="candidate-log-button" onClick={onLogAsWork}>
              Log as work
            </button>
          ))}
        <button type="button" className="ghost-button" onClick={onRemove}>
          Remove
        </button>
      </div>
      <input
        className="candidate-title"
        aria-label="Candidate title"
        value={candidate.title}
        onChange={(event) => onPatch({ title: event.target.value })}
      />
      <textarea
        className="candidate-detail"
        aria-label="Candidate detail"
        rows={2}
        value={candidate.detail}
        onChange={(event) => onPatch({ detail: event.target.value })}
      />
      <p className="candidate-source">&ldquo;{candidate.sourceExcerpt}&rdquo;</p>
      {candidate.type === "PERSON" && (
        <ContactMatchPanel
          personName={candidate.title}
          candidates={matchContactCandidates(candidate.title, references.contacts)}
          decision={candidate.contactDecision}
          contacts={references.contacts}
          organizations={references.organizations}
          onMatch={(contactAppId) => onPatch({ contactDecision: { type: "matched", contactAppId } })}
          onIgnore={() => onPatch({ contactDecision: { type: "ignored" } })}
          onReset={() => onPatch({ contactDecision: undefined })}
          onAddPerson={onAddPerson}
        />
      )}
    </div>
  );
}
