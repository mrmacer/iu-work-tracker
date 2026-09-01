"use client";

import { useState } from "react";
import type { AnalyzeTranscriptResult, AnalyzeTranscriptUsage } from "../lib/anthropic-voice-analysis";
import type { WorkRecord } from "../lib/models";
import { MAX_TRANSCRIPT_LENGTH } from "../lib/voice-intelligence-config";
import {
  VOICE_CANDIDATE_TYPES,
  VOICE_CANDIDATE_TYPE_LABELS,
  type VoiceCandidate,
  type VoiceCandidateType,
} from "../lib/voice-intelligence-models";
import { buildWorkRecordDraftFromVoiceCandidate } from "../lib/voice-intelligence-work-record";

type Phase = "paste" | "review";

/**
 * Local browser review state only — an AI candidate plus the two fields the review UI needs
 * that the model never produces: a stable React key and whether the user has it selected.
 * Nothing here is ever written to SharePoint, Work Records, Inbox Intelligence, People,
 * Organizations, Projects, a Knowledge Base, or any browser storage — see
 * docs/AI_HANDOFF.md "Voice Intelligence V1".
 */
type ReviewCandidate = VoiceCandidate & { id: string; selected: boolean };

function toReviewCandidates(candidates: VoiceCandidate[]): ReviewCandidate[] {
  return candidates.map((candidate) => ({ ...candidate, id: crypto.randomUUID(), selected: true }));
}

export default function VoiceIntelligence({
  openLog,
  createDraftRecord,
}: {
  openLog: (record?: WorkRecord, onSaved?: (saved: WorkRecord) => void) => void;
  createDraftRecord: () => WorkRecord;
}) {
  const [phase, setPhase] = useState<Phase>("paste");
  const [transcript, setTranscript] = useState("");
  const [analyzing, setAnalyzing] = useState(false);
  const [error, setError] = useState("");
  const [candidates, setCandidates] = useState<ReviewCandidate[]>([]);
  const [usage, setUsage] = useState<AnalyzeTranscriptUsage | null>(null);

  const clearTranscript = () => {
    setTranscript("");
    setError("");
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
      setCandidates(toReviewCandidates(result.analysis.candidates));
      setUsage(result.usage);
      setPhase("review");
    } catch {
      setError("The AI service could not be reached. Check your connection and try again.");
    } finally {
      setAnalyzing(false);
    }
  };

  const patchCandidate = (id: string, patch: Partial<Pick<ReviewCandidate, "type" | "title" | "detail" | "durationText" | "selected">>) =>
    setCandidates((current) => current.map((candidate) => (candidate.id === id ? { ...candidate, ...patch } : candidate)));

  const removeCandidate = (id: string) =>
    setCandidates((current) => current.filter((candidate) => candidate.id !== id));

  // Opens the existing Log Work form prefilled from the CURRENT edited candidate state —
  // never the original model output. Performs zero persistence: no Work Record is created,
  // no provider is called, and this candidate's own review state is left exactly as-is. The
  // human still reviews and explicitly saves through the existing, unmodified save path.
  const logAsWork = (candidate: ReviewCandidate) => {
    const draft = buildWorkRecordDraftFromVoiceCandidate(candidate, createDraftRecord());
    openLog(draft);
  };

  const backToTranscript = () => {
    setPhase("paste");
  };

  const startOver = () => {
    setPhase("paste");
    setCandidates([]);
    setUsage(null);
    setError("");
    clearTranscript();
  };

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

      {phase === "paste" && (
        <section className="panel">
          <div className="form-stack">
            <label>
              <span>Paste a transcript</span>
              <textarea
                rows={16}
                value={transcript}
                onChange={(event) => setTranscript(event.target.value)}
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
            <button className="ghost-button" onClick={clearTranscript} disabled={analyzing}>
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
                    onPatch={(patch) => patchCandidate(candidate.id, patch)}
                    onRemove={() => removeCandidate(candidate.id)}
                    onLogAsWork={() => logAsWork(candidate)}
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
    </div>
  );
}

function CandidateCard({
  candidate,
  onPatch,
  onRemove,
  onLogAsWork,
}: {
  candidate: ReviewCandidate;
  onPatch: (patch: Partial<Pick<ReviewCandidate, "type" | "title" | "detail" | "durationText" | "selected">>) => void;
  onRemove: () => void;
  onLogAsWork: () => void;
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
        {candidate.type === "COMPLETED_WORK" && (
          <button type="button" className="candidate-log-button" onClick={onLogAsWork}>
            Log as work
          </button>
        )}
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
    </div>
  );
}
