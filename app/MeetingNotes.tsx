"use client";

import { useState } from "react";
import type { ActiveProviderKind } from "../lib/data-provider";
import type { AnalyzeMeetingResult, AnalyzeMeetingUsage } from "../lib/anthropic-meeting-analysis";
import { MAX_MEETING_CONTENT_LENGTH } from "../lib/meeting-intelligence-config";
import {
  MEETING_CANDIDATE_TYPES,
  MEETING_CANDIDATE_TYPE_LABELS,
  MEETING_RECORD_SCHEMA_VERSION,
  type MeetingCandidate,
  type MeetingCandidateType,
  type MeetingRecord,
  type ReviewedMeetingCandidate,
} from "../lib/meeting-intelligence-models";
import { buildWorkRecordDraftFromMeetingCandidate } from "../lib/meeting-intelligence-work-record";
import { buildDraftMinutes, type MeetingDraft } from "../lib/meeting-minutes";
import type { MeetingRecordResult } from "../lib/meeting-record-provider";
import type { ProviderMetadata, WorkRecord } from "../lib/models";

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

function emptyMetadata(): ProviderMetadata {
  return { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" };
}

/**
 * Local browser review state only — an AI candidate plus the two fields the review UI needs
 * that the model never produces: a stable React key and whether the user has it selected.
 * See docs/AI_HANDOFF.md "Meeting Notes durability (Patch 6B)" — the client-only `id` is
 * never part of the durable record; it is regenerated every time a saved meeting is reopened.
 */
type ReviewCandidate = MeetingCandidate & { id: string; selected: boolean };

function toReviewCandidates(candidates: MeetingCandidate[]): ReviewCandidate[] {
  return candidates.map((candidate) => ({ ...candidate, id: crypto.randomUUID(), selected: true }));
}

/** Strips the transient UI id — this is the exact shape that becomes MeetingRecord.reviewedCandidates.
 * Field-by-field (not an `id`-omitting spread) so nothing beyond the documented reviewed-candidate
 * shape can leak in silently. */
function stripReviewIds(candidates: ReviewCandidate[]): ReviewedMeetingCandidate[] {
  return candidates.map(({ type, title, detail, sourceExcerpt, ownerText, dueText, durationText, selected }) => ({
    type,
    title,
    detail,
    sourceExcerpt,
    ownerText,
    dueText,
    durationText,
    selected,
  }));
}

/** Dirty-state fingerprint: candidate content only, deliberately excluding the UI-only id so
 * regenerating ids on reopen never falsely shows unsaved changes. */
function snapshotOf(draft: MeetingDraft, candidates: ReviewCandidate[], analysisModel: string | null, analyzedAt: string | null): string {
  return JSON.stringify({ draft, candidates: stripReviewIds(candidates), analysisModel, analyzedAt });
}

type Identity = { appId: string; metadata: ProviderMetadata };

function freshIdentity(): Identity {
  return { appId: crypto.randomUUID(), metadata: emptyMetadata() };
}

export default function MeetingNotes({
  openLog,
  createDraftRecord,
  records,
  saveRecord,
  updateRecord,
  loadFailed,
  storageMode,
}: {
  openLog: (record?: WorkRecord, onSaved?: (saved: WorkRecord) => void) => void;
  createDraftRecord: () => WorkRecord;
  records: MeetingRecord[];
  saveRecord: (record: MeetingRecord) => Promise<MeetingRecordResult<MeetingRecord>>;
  updateRecord: (record: MeetingRecord, expectedVersion: number) => Promise<MeetingRecordResult<MeetingRecord>>;
  loadFailed: boolean;
  storageMode: ActiveProviderKind;
}) {
  const [draft, setDraft] = useState<MeetingDraft>(emptyMeetingDraft);
  const [identity, setIdentity] = useState<Identity>(freshIdentity);
  const [analyzing, setAnalyzing] = useState(false);
  const [error, setError] = useState("");
  const [candidates, setCandidates] = useState<ReviewCandidate[]>([]);
  const [usage, setUsage] = useState<AnalyzeMeetingUsage | null>(null);
  const [analysisModel, setAnalysisModel] = useState<string | null>(null);
  const [analyzedAt, setAnalyzedAt] = useState<string | null>(null);
  const [baseline, setBaseline] = useState(() => snapshotOf(emptyMeetingDraft(), [], null, null));
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState("");
  const [copied, setCopied] = useState(false);

  const isDirty = snapshotOf(draft, candidates, analysisModel, analyzedAt) !== baseline;

  const patchDraft = (patch: Partial<MeetingDraft>) => setDraft((current) => ({ ...current, ...patch }));

  const hasAnalyzableContent = Boolean(draft.agendaText.trim() || draft.notesText.trim());
  const combinedLength = draft.agendaText.length + draft.notesText.length;

  // No AI request happens until this is explicitly called by the "Analyze Meeting"/
  // "Re-analyze meeting" click — loading this screen, editing fields, saving, reopening a
  // saved meeting, and loading Home all cost zero Anthropic requests.
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
    if (
      candidates.length > 0 &&
      !window.confirm(
        "Re-analyzing will replace the current Meeting Intelligence candidates with a new AI analysis. Your agenda and notes will remain.",
      )
    ) {
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
      setAnalysisModel(result.usage.model);
      setAnalyzedAt(new Date().toISOString());
    } catch {
      setError("The AI service could not be reached. Check your connection and try again.");
    } finally {
      setAnalyzing(false);
    }
  };

  const patchCandidate = (
    id: string,
    patch: Partial<Pick<ReviewCandidate, "type" | "title" | "detail" | "ownerText" | "dueText" | "durationText" | "selected">>,
  ) => setCandidates((current) => current.map((candidate) => (candidate.id === id ? { ...candidate, ...patch } : candidate)));

  const removeCandidate = (id: string) =>
    setCandidates((current) => current.filter((candidate) => candidate.id !== id));

  // Opens the existing Log Work form prefilled from the CURRENT edited candidate state —
  // never the original model output. Performs zero persistence: no Work Record is created,
  // no MeetingRecord is created or updated, and this candidate's own review state is left
  // exactly as-is. Saving a meeting and logging its completed work are separate, explicit,
  // uncoupled user actions — see docs/AI_HANDOFF.md "Meeting Notes durability (Patch 6B)".
  const logAsWork = (candidate: ReviewCandidate) => {
    const draftRecord = buildWorkRecordDraftFromMeetingCandidate(candidate, createDraftRecord());
    openLog(draftRecord);
  };

  const resetEditorTo = (
    nextDraft: MeetingDraft,
    nextCandidates: ReviewCandidate[],
    nextIdentity: Identity,
    nextAnalysisModel: string | null,
    nextAnalyzedAt: string | null,
  ) => {
    setDraft(nextDraft);
    setCandidates(nextCandidates);
    setIdentity(nextIdentity);
    setAnalysisModel(nextAnalysisModel);
    setAnalyzedAt(nextAnalyzedAt);
    setUsage(null);
    setError("");
    setSaveError("");
    setBaseline(snapshotOf(nextDraft, nextCandidates, nextAnalysisModel, nextAnalyzedAt));
  };

  const newMeeting = () => {
    if (isDirty && !window.confirm("Discard the changes in this meeting?")) return;
    resetEditorTo(emptyMeetingDraft(), [], freshIdentity(), null, null);
  };

  const openMeeting = (record: MeetingRecord) => {
    if (record.appId === identity.appId && !isDirty) return; // already open, nothing to do
    if (isDirty && !window.confirm("Discard the changes in this meeting?")) return;
    const nextDraft: MeetingDraft = {
      title: record.title,
      date: record.meetingDate,
      meetingType: record.meetingType,
      attendeesText: record.attendeesText,
      agendaText: record.agendaText,
      notesText: record.notesText,
    };
    const nextCandidates = record.reviewedCandidates.map((candidate) => ({ ...candidate, id: crypto.randomUUID() }));
    resetEditorTo(nextDraft, nextCandidates, { appId: record.appId, metadata: record.metadata }, record.analysisModel, record.analyzedAt);
  };

  const saveMeeting = async () => {
    if (saving) return;
    setSaving(true);
    setSaveError("");
    const minutesText = buildDraftMinutes(draft, candidates);
    const record: MeetingRecord = {
      appId: identity.appId,
      schemaVersion: MEETING_RECORD_SCHEMA_VERSION,
      title: draft.title.trim(),
      meetingDate: draft.date,
      meetingType: draft.meetingType,
      attendeesText: draft.attendeesText,
      agendaText: draft.agendaText,
      notesText: draft.notesText,
      reviewedCandidates: stripReviewIds(candidates),
      minutesText,
      analysisModel,
      analyzedAt,
      metadata: identity.metadata,
    };
    const result = identity.metadata.version > 0 ? await updateRecord(record, identity.metadata.version) : await saveRecord(record);
    setSaving(false);
    if (result.status !== "success") {
      setSaveError(
        result.status === "validation_error" ? (result.errors[0]?.message ?? "Check the meeting and try again.") : result.message,
      );
      return;
    }
    setIdentity({ appId: result.value.appId, metadata: result.value.metadata });
    setBaseline(snapshotOf(draft, candidates, analysisModel, analyzedAt));
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
  const isSavedMeeting = identity.metadata.version > 0;

  return (
    <div className="screen-inner">
      <div className="page-heading">
        <div>
          <p className="eyebrow">AI-assisted intake</p>
          <h1>Meeting Notes</h1>
          <p>Before the meeting, jot an agenda. During and after, take notes. Then analyze for a reviewable summary.</p>
        </div>
      </div>
      <p className="muted-copy">
        {storageMode === "sharepoint"
          ? "SharePoint DEV connected — meeting records are saved to your IU Work Tracker workspace when you choose Save Meeting."
          : "Preview mode — meeting records are not durable in this mode."}
        {isDirty && <span className="sample-label" style={{ marginLeft: 8 }}>Unsaved changes</span>}
      </p>

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
        <footer className="log-footer">
          <button className="ghost-button" onClick={newMeeting}>
            New Meeting
          </button>
          <div>
            {saveError && <span style={{ marginRight: 10 }}>{saveError}</span>}
            <button className="primary-action" onClick={() => void saveMeeting()} disabled={saving || !isDirty}>
              {saving ? "Saving…" : "Save Meeting"}
            </button>
          </div>
        </footer>
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

      {analyzedAt && (
        <section className="panel">
          <p className="eyebrow">Meeting Intelligence — review before doing anything with these</p>
          <p className="muted-copy">Selected candidates are ready for review.{!isSavedMeeting && " Nothing has been saved yet."}</p>
          {candidates.length === 0 ? (
            <div className="empty">
              <span>○</span>
              <strong>No useful candidates found</strong>
              <p>The agenda or notes may have been too short or too vague to segment. You can edit them and re-analyze.</p>
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
                  <MeetingCandidateCard
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
        </section>
      )}

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

      <section className="panel list-panel">
        <div className="panel-heading">
          <h2>Saved Meetings</h2>
          <span className="sample-label">{records.length}</span>
        </div>
        {loadFailed && <p className="muted-copy">Saved meetings could not be loaded. Try reloading the page.</p>}
        {records.length ? (
          records.map((record) => <SavedMeetingRow key={record.appId} record={record} onOpen={() => openMeeting(record)} />)
        ) : (
          !loadFailed && <p className="muted-copy">No meetings saved yet.</p>
        )}
      </section>
    </div>
  );
}

function SavedMeetingRow({ record, onOpen }: { record: MeetingRecord; onOpen: () => void }) {
  const summary = record.reviewedCandidates.find((candidate) => candidate.type === "SUMMARY");
  return (
    <button
      className="record-row"
      style={{ width: "100%", cursor: "pointer", flexWrap: "wrap", minHeight: "auto", padding: "10px 4px" }}
      onClick={onOpen}
    >
      <span className="record-dot orbit" />
      <span>
        <strong>{record.title || "Untitled meeting"}</strong>
        <small>
          {record.meetingDate}
          {record.meetingType ? ` · ${record.meetingType}` : ""}
          {summary ? ` · ${summary.title}` : ""}
        </small>
      </span>
    </button>
  );
}

function MeetingCandidateCard({
  candidate,
  onPatch,
  onRemove,
  onLogAsWork,
}: {
  candidate: ReviewCandidate;
  onPatch: (
    patch: Partial<Pick<ReviewCandidate, "type" | "title" | "detail" | "ownerText" | "dueText" | "durationText" | "selected">>,
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
