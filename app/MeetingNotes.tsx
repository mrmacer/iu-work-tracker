"use client";

import { useState } from "react";
import type { AnalyzeMeetingResult, AnalyzeMeetingUsage } from "../lib/anthropic-meeting-analysis";
import { MAX_MEETING_CONTENT_LENGTH } from "../lib/meeting-intelligence-config";
import {
  MEETING_CANDIDATE_TYPES,
  MEETING_CANDIDATE_TYPE_LABELS,
  type MeetingCandidate,
  type MeetingCandidateType,
} from "../lib/meeting-intelligence-models";
import { buildWorkRecordDraftFromMeetingCandidate } from "../lib/meeting-intelligence-work-record";
import { buildDraftMinutes, type MeetingDraft } from "../lib/meeting-minutes";
import type { WorkRecord } from "../lib/models";

const MEETING_TYPES = [
  "District Meeting",
  "Internal IU Meeting",
  "Project Meeting",
  "STEM Collaborative",
  "Professional Development",
  "Partner Meeting",
  "Planning Meeting",
  "Other",
];

const todayIso = () =>
  new Intl.DateTimeFormat("en-CA", { year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());

function emptyMeetingDraft(): MeetingDraft {
  return { title: "", date: todayIso(), meetingType: "", attendeesText: "", agendaText: "", notesText: "" };
}

/**
 * Local browser review state only — an AI candidate plus the two fields the review UI needs
 * that the model never produces: a stable React key and whether the user has it selected.
 * Nothing here — meeting details, agenda, notes, or candidates — is ever written to
 * SharePoint, Work Records, Inbox Intelligence, a Knowledge Base, or any browser storage. See
 * docs/AI_HANDOFF.md "Meeting Notes V1".
 */
type ReviewMeetingCandidate = MeetingCandidate & { id: string; selected: boolean };

function toReviewCandidates(candidates: MeetingCandidate[]): ReviewMeetingCandidate[] {
  return candidates.map((candidate) => ({ ...candidate, id: crypto.randomUUID(), selected: true }));
}

export default function MeetingNotes({
  openLog,
  createDraftRecord,
}: {
  openLog: (record?: WorkRecord, onSaved?: (saved: WorkRecord) => void) => void;
  createDraftRecord: () => WorkRecord;
}) {
  const [draft, setDraft] = useState<MeetingDraft>(emptyMeetingDraft);
  const [analyzing, setAnalyzing] = useState(false);
  const [error, setError] = useState("");
  const [candidates, setCandidates] = useState<ReviewMeetingCandidate[]>([]);
  const [usage, setUsage] = useState<AnalyzeMeetingUsage | null>(null);
  const [copied, setCopied] = useState(false);

  const patchDraft = (patch: Partial<MeetingDraft>) => setDraft((current) => ({ ...current, ...patch }));

  const hasAnalyzableContent = Boolean(draft.agendaText.trim() || draft.notesText.trim());
  const combinedLength = draft.agendaText.length + draft.notesText.length;

  // No AI request happens until this is explicitly called by the "Analyze Meeting" click —
  // loading this screen, editing fields, and loading Home all cost zero Anthropic requests.
  const analyze = async () => {
    if (analyzing) return; // guards against double-click/rapid-Enter re-entrancy
    if (!hasAnalyzableContent) {
      setError("Add agenda items or notes before analyzing.");
      return;
    }
    if (combinedLength > MAX_MEETING_CONTENT_LENGTH) {
      setError(
        `The agenda and notes are too long combined (${combinedLength.toLocaleString()} of ${MAX_MEETING_CONTENT_LENGTH.toLocaleString()} characters allowed). Trim it and try again.`,
      );
      return;
    }
    setAnalyzing(true);
    setError("");
    try {
      const response = await fetch("/api/meeting-intelligence", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          title: draft.title,
          date: draft.date,
          meetingType: draft.meetingType,
          attendeesText: draft.attendeesText,
          agendaText: draft.agendaText,
          notesText: draft.notesText,
        }),
      });
      const result = (await response.json().catch(() => null)) as AnalyzeMeetingResult | null;
      if (!result || result.status !== "success") {
        setError(result?.message ?? "The meeting could not be analyzed. Try again.");
        return; // meeting content is deliberately left in place — never cleared on failure
      }
      setCandidates(toReviewCandidates(result.analysis.candidates));
      setUsage(result.usage);
    } catch {
      setError("The AI service could not be reached. Check your connection and try again.");
    } finally {
      setAnalyzing(false);
    }
  };

  const patchCandidate = (
    id: string,
    patch: Partial<Pick<ReviewMeetingCandidate, "type" | "title" | "detail" | "ownerText" | "dueText" | "durationText" | "selected">>,
  ) => setCandidates((current) => current.map((candidate) => (candidate.id === id ? { ...candidate, ...patch } : candidate)));

  const removeCandidate = (id: string) =>
    setCandidates((current) => current.filter((candidate) => candidate.id !== id));

  // Opens the existing Log Work form prefilled from the CURRENT edited candidate state —
  // never the original model output. Performs zero persistence: no Work Record is created,
  // no provider is called, and this candidate's own review state is left exactly as-is.
  const logAsWork = (candidate: ReviewMeetingCandidate) => {
    const draftRecord = buildWorkRecordDraftFromMeetingCandidate(candidate, createDraftRecord());
    openLog(draftRecord);
  };

  const minutesText = buildDraftMinutes(draft, candidates);

  const copyMinutes = async () => {
    try {
      await navigator.clipboard.writeText(minutesText);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 2000);
    } catch {
      setError("Minutes could not be copied automatically. Select and copy the text manually.");
    }
  };

  const selectedCount = candidates.filter((candidate) => candidate.selected).length;

  return (
    <div className="screen-inner">
      <div className="page-heading">
        <div>
          <p className="eyebrow">AI-assisted intake</p>
          <h1>Meeting Notes</h1>
          <p>Before the meeting, jot an agenda. During and after, take notes. Then analyze for a reviewable summary.</p>
        </div>
      </div>
      <p className="muted-copy">Meeting Notes V1 is a review workspace. Meeting content is not saved yet.</p>

      <section className="panel">
        <div className="form-two">
          <label>
            <span>Meeting title</span>
            <input value={draft.title} onChange={(event) => patchDraft({ title: event.target.value })} placeholder="e.g. STEELS quarterly planning" />
          </label>
          <label>
            <span>Date</span>
            <input type="date" value={draft.date} onChange={(event) => patchDraft({ date: event.target.value })} />
          </label>
          <label>
            <span>Meeting type</span>
            <select value={draft.meetingType} onChange={(event) => patchDraft({ meetingType: event.target.value })}>
              <option value="">Choose a meeting type…</option>
              {MEETING_TYPES.map((type) => (
                <option key={type} value={type}>
                  {type}
                </option>
              ))}
            </select>
          </label>
          <label>
            <span>
              Attendees <small>plain text</small>
            </span>
            <input value={draft.attendeesText} onChange={(event) => patchDraft({ attendeesText: event.target.value })} placeholder="e.g. Greg, Annie, Kim" />
          </label>
        </div>
      </section>

      <div className="meeting-columns">
        <section className="panel">
          <p className="eyebrow">Agenda</p>
          <textarea
            rows={12}
            value={draft.agendaText}
            onChange={(event) => patchDraft({ agendaText: event.target.value })}
            placeholder="Paste or type agenda items — no need to structure them."
          />
        </section>
        <section className="panel">
          <p className="eyebrow">General notes</p>
          <textarea
            rows={12}
            value={draft.notesText}
            onChange={(event) => patchDraft({ notesText: event.target.value })}
            placeholder="Take notes during the meeting, or paste them afterward."
          />
        </section>
      </div>

      <section className="panel">
        <p className="muted-copy">
          {combinedLength.toLocaleString()} / {MAX_MEETING_CONTENT_LENGTH.toLocaleString()} combined characters
        </p>
        {error && (
          <div className="form-error" role="alert">
            {error}
          </div>
        )}
        <footer className="log-footer">
          <span />
          <button className="primary-action" onClick={() => void analyze()} disabled={analyzing || !hasAnalyzableContent}>
            {analyzing ? "Analyzing…" : candidates.length > 0 ? "Re-analyze meeting" : "Analyze Meeting"}
          </button>
        </footer>
      </section>

      {candidates.length > 0 && (
        <section className="panel">
          <p className="eyebrow">Meeting Intelligence — review before doing anything with these</p>
          <p className="muted-copy">Selected candidates are ready for review. Nothing has been saved yet.</p>
          <div className="candidate-summary">
            <span>{candidates.length} candidate{candidates.length === 1 ? "" : "s"}</span>
            <span>{selectedCount} selected</span>
            <span>{candidates.length - selectedCount} ignored</span>
          </div>
          <div className="candidate-list">
            {candidates.map((candidate) => (
              <MeetingCandidateCard
                key={candidate.id}
                candidate={candidate}
                onPatch={(patch) => patchCandidate(candidate.id, patch)}
                onRemove={() => removeCandidate(candidate.id)}
                onLogAsWork={() => logAsWork(candidate)}
              />
            ))}
          </div>
          {usage && (
            <p className="muted-copy">
              Model: {usage.model} · {usage.inputTokens.toLocaleString()} in / {usage.outputTokens.toLocaleString()} out tokens
            </p>
          )}
        </section>
      )}

      {candidates.length > 0 && (
        <section className="panel">
          <p className="eyebrow">Draft Minutes</p>
          <pre className="draft-minutes">{minutesText}</pre>
          <footer className="log-footer">
            <span />
            <button className="primary-action" onClick={() => void copyMinutes()}>
              {copied ? "Copied!" : "Copy Minutes"}
            </button>
          </footer>
        </section>
      )}
    </div>
  );
}

function MeetingCandidateCard({
  candidate,
  onPatch,
  onRemove,
  onLogAsWork,
}: {
  candidate: ReviewMeetingCandidate;
  onPatch: (
    patch: Partial<Pick<ReviewMeetingCandidate, "type" | "title" | "detail" | "ownerText" | "dueText" | "durationText" | "selected">>,
  ) => void;
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
          onChange={(event) => onPatch({ type: event.target.value as MeetingCandidateType })}
        >
          {MEETING_CANDIDATE_TYPES.map((type) => (
            <option key={type} value={type}>
              {MEETING_CANDIDATE_TYPE_LABELS[type]}
            </option>
          ))}
        </select>
        {candidate.ownerText && (
          <span className="candidate-chip">
            {candidate.ownerText}
            <button type="button" aria-label="Remove owner" onClick={() => onPatch({ ownerText: null })}>
              ×
            </button>
          </span>
        )}
        {candidate.dueText && (
          <span className="candidate-chip">
            {candidate.dueText}
            <button type="button" aria-label="Remove due date" onClick={() => onPatch({ dueText: null })}>
              ×
            </button>
          </span>
        )}
        {candidate.durationText && (
          <span className="candidate-chip">
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
      {candidate.sourceExcerpt && <p className="candidate-source">&ldquo;{candidate.sourceExcerpt}&rdquo;</p>}
    </div>
  );
}
