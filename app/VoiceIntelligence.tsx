"use client";

import { useRef, useState } from "react";
import type { AnalyzeTranscriptResult } from "../lib/anthropic-voice-analysis";
import type { Contact, Organization, Project, ReferenceData, WorkRecord } from "../lib/models";
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
import OrganizationFormModal, { emptyOrganizationDraft } from "./OrganizationFormModal";
import ProjectFormModal, { emptyProjectDraft } from "./ProjectFormModal";
import VoiceRoutingPanel from "./VoiceRoutingPanel";
import { matchContactCandidates } from "../lib/contact-matching";
import type { ContactResult } from "../lib/contact-provider";
import type { OrganizationResult } from "../lib/organization-provider";
import type { ProjectResult } from "../lib/project-provider";
import { processedChipLabel } from "../lib/voice-routing";

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
  saveProject,
  updateProject,
  saveOrganization,
  updateOrganization,
  storage,
}: {
  openLog: (record?: WorkRecord, onSaved?: (saved: WorkRecord) => void) => void;
  createDraftRecord: () => WorkRecord;
  references: ReferenceData;
  saveContact: (contact: Contact) => Promise<ContactResult<Contact>>;
  updateContact: (contact: Contact, expectedVersion: number) => Promise<ContactResult<Contact>>;
  // Patch 7.3 — the EXISTING durable Project/Organization providers (via IUWorkTracker), used
  // only when the human explicitly saves the reused Create Project / Create Organization form.
  saveProject: (project: Project) => Promise<ProjectResult<Project>>;
  updateProject: (project: Project, expectedVersion: number) => Promise<ProjectResult<Project>>;
  saveOrganization: (organization: Organization) => Promise<OrganizationResult<Organization>>;
  updateOrganization: (organization: Organization, expectedVersion: number) => Promise<OrganizationResult<Organization>>;
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
    patch: Partial<Pick<ReviewCandidate, "type" | "title" | "detail" | "durationText" | "selected" | "contactDecision" | "contactCreated" | "routing">>,
  ) =>
    commit((current) => {
      const nextCandidates = current.candidates.map((candidate) => (candidate.id === id ? { ...candidate, ...patch } : candidate));
      // Nothing left selected → nothing to process, so the panel closes (selection itself is kept as-is).
      const keepOpen = current.routingOpen && nextCandidates.some((candidate) => candidate.selected);
      return { ...current, candidates: nextCandidates, routingOpen: keepOpen };
    });

  // Patch 7.3 — selection controls. Selecting means "include in the current processing batch";
  // deselecting NEVER deletes a candidate (only Remove does).
  const setAllSelected = (selected: boolean) =>
    commit((current) => ({
      ...current,
      candidates: current.candidates.map((candidate) => ({ ...candidate, selected })),
      routingOpen: selected ? current.routingOpen : false,
    }));
  const openRouting = () => commit((current) => ({ ...current, routingOpen: true }));
  const closeRouting = () => {
    setSkippedIds(new Set());
    commit((current) => ({ ...current, routingOpen: false }));
  };

  const removeCandidate = (id: string) =>
    commit((current) => ({ ...current, candidates: current.candidates.filter((candidate) => candidate.id !== id) }));

  // Patch 8D — deterministic Contact matching for PERSON candidates only. Zero AI calls, zero
  // network calls: see lib/contact-matching.ts. contactDecision is transient exactly like
  // every other Voice review field — see the ReviewCandidate doc comment above.
  const [addPersonCandidateId, setAddPersonCandidateId] = useState<string | null>(null);
  const addPersonCandidate = candidates.find((candidate) => candidate.id === addPersonCandidateId) ?? null;
  // Patch 7.3 — transient routing-panel state: completed-work items skipped in THIS visit (skipping
  // is not durable and loses nothing), and the candidate whose Create Organization/Project form is open.
  const [skippedIds, setSkippedIds] = useState<Set<string>>(new Set());
  const [createEntityId, setCreateEntityId] = useState<string | null>(null);
  const createEntityCandidate = candidates.find((candidate) => candidate.id === createEntityId) ?? null;

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
  const routingVisible = Boolean(session.routingOpen) && selectedCount > 0;

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
                <span>{candidates.length - selectedCount} not selected</span>
              </div>
              {selectedCount > 0 && (
                <div className="voice-tray" role="group" aria-label="Selected candidates">
                  <button type="button" className="primary-action" onClick={openRouting} disabled={routingVisible}>
                    Process selected
                  </button>
                  {selectedCount < candidates.length && (
                    <button type="button" className="ghost-button" onClick={() => setAllSelected(true)}>
                      Select all
                    </button>
                  )}
                  <button type="button" className="ghost-button" onClick={() => setAllSelected(false)}>
                    Clear selection
                  </button>
                  <span className="muted-copy">Checked candidates are included when you process. Unchecked ones stay here untouched.</span>
                </div>
              )}
              {routingVisible && (
                <VoiceRoutingPanel
                  candidates={candidates}
                  references={references}
                  skippedIds={skippedIds}
                  onSkip={(id) => setSkippedIds((current) => new Set(current).add(id))}
                  onUnskip={() => setSkippedIds(new Set())}
                  onLogAsWork={(candidate) => logAsWork(candidate as ReviewCandidate)}
                  onPatch={patchCandidate}
                  onAddPerson={setAddPersonCandidateId}
                  onCreateEntity={setCreateEntityId}
                  onClose={closeRouting}
                />
              )}
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
            patchCandidate(addPersonCandidate.id, { contactDecision: { type: "matched", contactAppId: savedContact.appId }, contactCreated: true });
            setAddPersonCandidateId(null);
          }}
          saveContact={saveContact}
          updateContact={updateContact}
        />
      )}

      {createEntityCandidate && (createEntityCandidate.type === "ORGANIZATION" || createEntityCandidate.type === "DISTRICT") && (
        // The existing Organization form and provider, prefilled only with the candidate's name
        // (District → type "district"). Only the human's explicit Save writes anything.
        <OrganizationFormModal
          organization={{
            ...emptyOrganizationDraft(),
            name: createEntityCandidate.title,
            type: createEntityCandidate.type === "DISTRICT" ? "district" : "partner",
          }}
          organizations={references.organizations}
          onCancel={() => setCreateEntityId(null)}
          onSaved={(saved) => {
            patchCandidate(createEntityCandidate.id, { routing: { type: "created", entityAppId: saved.appId, entityName: saved.name } });
            setCreateEntityId(null);
          }}
          saveOrganization={saveOrganization}
          updateOrganization={updateOrganization}
        />
      )}

      {createEntityCandidate && createEntityCandidate.type === "PROJECT" && (
        // The existing Project form and provider, prefilled with the candidate's name/detail.
        <ProjectFormModal
          project={{ ...emptyProjectDraft(references.projects.length), name: createEntityCandidate.title, description: createEntityCandidate.detail }}
          onCancel={() => setCreateEntityId(null)}
          onSaved={(saved) => {
            patchCandidate(createEntityCandidate.id, { routing: { type: "created", entityAppId: saved.appId, entityName: saved.name } });
            setCreateEntityId(null);
          }}
          saveProject={saveProject}
          updateProject={updateProject}
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
  onPatch: (patch: Partial<Pick<ReviewCandidate, "type" | "title" | "detail" | "durationText" | "selected" | "contactDecision" | "contactCreated">>) => void;
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
          title="Include in the current processing batch (unchecking never deletes — use Remove for that)"
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
        {processedChipLabel(candidate) && <span className="candidate-chip">{processedChipLabel(candidate)}</span>}
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
        <button type="button" className="ghost-button" onClick={onRemove} title="Delete this candidate from the working session">
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
          // contactCreated is reset with every decision made here, so a "Created Contact ✓" from an
          // earlier Add Person can never outlive the decision it described (the routing panel does the same).
          onMatch={(contactAppId) => onPatch({ contactDecision: { type: "matched", contactAppId }, contactCreated: false })}
          onIgnore={() => onPatch({ contactDecision: { type: "ignored" }, contactCreated: false })}
          onReset={() => onPatch({ contactDecision: undefined, contactCreated: false })}
          onAddPerson={onAddPerson}
        />
      )}
    </div>
  );
}
