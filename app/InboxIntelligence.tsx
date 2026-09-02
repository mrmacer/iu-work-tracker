"use client";

import { useState } from "react";
import { MAX_EMAIL_LENGTH } from "../lib/anthropic-config";
import type { AnalyzeEmailResult, AnalyzeEmailUsage } from "../lib/anthropic-email-analysis";
import ContactMatchPanel, { resolveMatchedContacts } from "./ContactMatchPanel";
import ContactFormModal, { emptyContactDraft } from "./ContactFormModal";
import { matchContactCandidates, type ContactMatchDecision } from "../lib/contact-matching";
import type { ContactResult } from "../lib/contact-provider";
import { buildWorkRecordDraftFromAnalysis } from "../lib/inbox-intelligence-work-record";
import {
  buildInboxIntelligenceRecord,
  computeInboxIntelligenceSummary,
  resolveEmailAnalysisEntities,
  type EmailAnalysis,
  type InboxIntelligenceRecord,
  type InboxIntelligenceStatus,
} from "../lib/inbox-intelligence-models";
import type { InboxIntelligenceResult } from "../lib/inbox-intelligence-provider";
import type { Contact, ReferenceData, WorkRecord } from "../lib/models";

type Phase = "paste" | "review" | "saved";

function excerpt(rawEmail: string): string {
  const trimmed = rawEmail.trim().replace(/\s+/g, " ");
  return trimmed.length > 200 ? `${trimmed.slice(0, 200)}…` : trimmed;
}

function csvToList(value: string): string[] {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function resultMessage(result: InboxIntelligenceResult<unknown>): string {
  if (result.status === "success") return "";
  if (result.status === "validation_error") return result.errors[0]?.message ?? "Check the record and try again.";
  return result.message;
}

function lastModified(record: InboxIntelligenceRecord): string {
  return new Intl.DateTimeFormat("en-US", { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }).format(
    new Date(record.metadata.modifiedAt),
  );
}

export default function InboxIntelligence({
  references,
  openLog,
  createDraftRecord,
  records,
  saveRecord,
  updateRecord,
  saveContact,
  updateContact,
}: {
  references: ReferenceData;
  openLog: (record?: WorkRecord, onSaved?: (saved: WorkRecord) => void) => void;
  createDraftRecord: () => WorkRecord;
  records: InboxIntelligenceRecord[];
  saveRecord: (record: InboxIntelligenceRecord) => Promise<InboxIntelligenceResult<InboxIntelligenceRecord>>;
  updateRecord: (record: InboxIntelligenceRecord, expectedVersion: number) => Promise<InboxIntelligenceResult<InboxIntelligenceRecord>>;
  saveContact: (contact: Contact) => Promise<ContactResult<Contact>>;
  updateContact: (contact: Contact, expectedVersion: number) => Promise<ContactResult<Contact>>;
}) {
  const [phase, setPhase] = useState<Phase>("paste");
  const [rawEmail, setRawEmail] = useState("");
  const [analyzing, setAnalyzing] = useState(false);
  const [error, setError] = useState("");
  const [analysis, setAnalysis] = useState<EmailAnalysis | null>(null);
  const [usage, setUsage] = useState<AnalyzeEmailUsage | null>(null);
  const [sourceExcerpt, setSourceExcerpt] = useState("");
  const [analyzedAt, setAnalyzedAt] = useState("");
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState("");
  const [justSaved, setJustSaved] = useState<InboxIntelligenceRecord | null>(null);
  const [rowError, setRowError] = useState("");
  const [busyAppId, setBusyAppId] = useState<string | null>(null);
  // Patch 8D — People review state for the CURRENT (unsaved) analysis only. Keyed by the exact
  // detected name string. Never persisted directly: at Save, only the accepted "matched"
  // decisions become matchedContactIds — an "ignored" decision or an unreviewed person leaves
  // no trace, exactly like Organization/Project matching leaves no "was this reviewed" record.
  const [personDecisions, setPersonDecisions] = useState<Record<string, ContactMatchDecision>>({});
  const [addPersonName, setAddPersonName] = useState<string | null>(null);

  const summary = computeInboxIntelligenceSummary(records);

  const clearAll = () => {
    setRawEmail("");
    setError("");
  };

  const analyze = async () => {
    if (analyzing) return; // guards against double-click/rapid-Enter re-entrancy
    if (!rawEmail.trim()) {
      setError("Paste an email before analyzing.");
      return;
    }
    if (rawEmail.length > MAX_EMAIL_LENGTH) {
      setError(`This email is too long (${rawEmail.length} of ${MAX_EMAIL_LENGTH} characters allowed). Trim it and try again.`);
      return;
    }
    setAnalyzing(true);
    setError("");
    try {
      const response = await fetch("/api/inbox-intelligence", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ rawEmail }),
      });
      const result = (await response.json().catch(() => null)) as AnalyzeEmailResult | null;
      if (!result || result.status !== "success") {
        setError(result?.message ?? "The email could not be analyzed. Try again.");
        return;
      }
      setAnalysis(result.analysis);
      setUsage(result.usage);
      setSourceExcerpt(excerpt(rawEmail));
      setAnalyzedAt(new Date().toISOString());
      setPersonDecisions({});
      setPhase("review");
    } catch {
      setError("The AI service could not be reached. Check your connection and try again.");
    } finally {
      setAnalyzing(false);
    }
  };

  const patchAnalysis = (patch: Partial<EmailAnalysis>) =>
    setAnalysis((current) => (current ? { ...current, ...patch } : current));

  const startOver = () => {
    setPhase("paste");
    setAnalysis(null);
    setUsage(null);
    setSourceExcerpt("");
    setJustSaved(null);
    setSaveError("");
    setPersonDecisions({});
    setAddPersonName(null);
    clearAll();
  };

  // Patch 8D — deterministic, zero-AI, zero-network. See lib/contact-matching.ts.
  const decideMatch = (personName: string, contactAppId: string) =>
    setPersonDecisions((current) => ({ ...current, [personName]: { type: "matched", contactAppId } }));
  const decideIgnore = (personName: string) =>
    setPersonDecisions((current) => ({ ...current, [personName]: { type: "ignored" } }));
  const resetDecision = (personName: string) =>
    setPersonDecisions((current) => {
      const next = { ...current };
      delete next[personName];
      return next;
    });
  const onPersonAdded = (personName: string, savedContact: Contact) => {
    decideMatch(personName, savedContact.appId);
    setAddPersonName(null);
  };

  const saveToInbox = async () => {
    if (!analysis || saving) return; // one Save click, no re-entrancy — generates zero Anthropic calls
    setSaving(true);
    setSaveError("");
    // AI MAY PROPOSE. THE HUMAN DECIDES. Only explicit "Match Existing" / "Add Person" review
    // decisions become matchedContactIds — never auto-resolved the way Organization/District/
    // Project are inside buildInboxIntelligenceRecord() itself. See lib/contact-matching.ts.
    const matchedContactIds = [
      ...new Set(
        Object.values(personDecisions)
          .filter((decision): decision is { type: "matched"; contactAppId: string } => decision.type === "matched")
          .map((decision) => decision.contactAppId),
      ),
    ];
    const record = buildInboxIntelligenceRecord(analysis, sourceExcerpt, references, analyzedAt, matchedContactIds);
    const result = await saveRecord(record);
    setSaving(false);
    if (result.status !== "success") {
      setSaveError(resultMessage(result));
      return;
    }
    setJustSaved(result.value);
    setPhase("saved");
  };

  const createWorkRecordFrom = (record: InboxIntelligenceRecord) => {
    const draft = buildWorkRecordDraftFromAnalysis(record.analysis, references, createDraftRecord());
    openLog(draft, (savedWorkRecord) => {
      void linkWorkRecord(record, savedWorkRecord.appId);
    });
  };

  const linkWorkRecord = async (record: InboxIntelligenceRecord, workRecordAppId: string) => {
    const result = await updateRecord({ ...record, linkedWorkRecordAppId: workRecordAppId }, record.metadata.version);
    if (result.status !== "success") {
      setRowError(`The Work Record was created, but linking it back to this Inbox item failed: ${resultMessage(result)}`);
    }
  };

  const updateStatus = async (record: InboxIntelligenceRecord, status: InboxIntelligenceStatus) => {
    if (busyAppId) return;
    setBusyAppId(record.appId);
    setRowError("");
    const result = await updateRecord(
      { ...record, status, resolvedAt: status === "resolved" ? new Date().toISOString() : null },
      record.metadata.version,
    );
    setBusyAppId(null);
    if (result.status !== "success") setRowError(resultMessage(result));
  };

  return (
    <div className="screen-inner">
      <div className="page-heading">
        <div>
          <p className="eyebrow">AI-assisted intake</p>
          <h1>Inbox Intelligence</h1>
          <p>Paste a work email. Review what AI finds. Track it until it&rsquo;s resolved.</p>
        </div>
      </div>

      <div className="metric-strip">
        <Metric value={String(summary.openCount)} label="needs attention" />
        <Metric value={String(summary.waitingCount)} label="waiting" />
        <Metric value={String(summary.resolvedCount)} label="resolved" />
      </div>

      {phase === "paste" && (
        <section className="panel">
          <div className="form-stack">
            <label>
              <span>Paste an email</span>
              <textarea
                rows={14}
                value={rawEmail}
                onChange={(event) => setRawEmail(event.target.value)}
                placeholder="Paste the whole email — subject, sender, recipients, timestamps, body, signature, and any quoted thread. No need to clean it up first."
              />
            </label>
            <p className="muted-copy">
              {rawEmail.length.toLocaleString()} / {MAX_EMAIL_LENGTH.toLocaleString()} characters
            </p>
            {error && (
              <div className="form-error" role="alert">
                {error}
              </div>
            )}
          </div>
          <footer className="log-footer">
            <button className="ghost-button" onClick={clearAll} disabled={analyzing}>
              Clear
            </button>
            <button className="primary-action" onClick={() => void analyze()} disabled={analyzing || !rawEmail.trim()}>
              {analyzing ? "Analyzing…" : "Analyze email"}
            </button>
          </footer>
        </section>
      )}

      {phase === "review" && analysis && (
        <section className="panel">
          <p className="eyebrow">AI-suggested — review before saving</p>
          <div className="form-stack">
            <label>
              <span>Summary</span>
              <textarea rows={2} value={analysis.summary} onChange={(event) => patchAnalysis({ summary: event.target.value })} />
            </label>
            <div className="form-two">
              <label>
                <span>Priority</span>
                <select
                  value={analysis.priority}
                  onChange={(event) => patchAnalysis({ priority: event.target.value as EmailAnalysis["priority"] })}
                >
                  <option value="high">High</option>
                  <option value="medium">Medium</option>
                  <option value="low">Low</option>
                </select>
              </label>
              <div className="toggle-line">
                <input
                  aria-label="Needs attention"
                  type="checkbox"
                  checked={analysis.needsAttention}
                  onChange={(event) => patchAnalysis({ needsAttention: event.target.checked })}
                />
                <span>
                  <strong>Needs attention</strong>
                  <small>AI flagged this email as worth prioritizing.</small>
                </span>
              </div>
            </div>

            {analysis.actionItems.length > 0 && (
              <fieldset>
                <legend>Action items</legend>
                {analysis.actionItems.map((item, index) => (
                  <div className="form-two" key={index}>
                    <label>
                      <span>Action</span>
                      <input
                        value={item.action}
                        onChange={(event) => {
                          const actionItems = [...analysis.actionItems];
                          actionItems[index] = { ...item, action: event.target.value };
                          patchAnalysis({ actionItems });
                        }}
                      />
                    </label>
                    <label>
                      <span>Owner</span>
                      <select
                        value={item.owner}
                        onChange={(event) => {
                          const actionItems = [...analysis.actionItems];
                          actionItems[index] = { ...item, owner: event.target.value as typeof item.owner };
                          patchAnalysis({ actionItems });
                        }}
                      >
                        <option value="me">Me</option>
                        <option value="sender">Sender</option>
                        <option value="other">Other</option>
                        <option value="unknown">Unknown</option>
                      </select>
                    </label>
                    <label>
                      <span>Due date</span>
                      <input
                        type="date"
                        value={item.dueDate ?? ""}
                        onChange={(event) => {
                          const actionItems = [...analysis.actionItems];
                          actionItems[index] = { ...item, dueDate: event.target.value || null };
                          patchAnalysis({ actionItems });
                        }}
                      />
                    </label>
                    <button
                      type="button"
                      className="ghost-button"
                      onClick={() => patchAnalysis({ actionItems: analysis.actionItems.filter((_, i) => i !== index) })}
                    >
                      Remove
                    </button>
                  </div>
                ))}
              </fieldset>
            )}

            <label>
              <span>Follow-up</span>
              <input value={analysis.followUp} onChange={(event) => patchAnalysis({ followUp: event.target.value })} />
            </label>
            <label>
              <span>Tags <small>comma-separated</small></span>
              <input value={analysis.tags.join(", ")} onChange={(event) => patchAnalysis({ tags: csvToList(event.target.value) })} />
            </label>
            <div className="form-two">
              <label>
                <span>People <small>comma-separated</small></span>
                <input value={analysis.people.join(", ")} onChange={(event) => patchAnalysis({ people: csvToList(event.target.value) })} />
              </label>
              <label>
                <span>Organizations <small>comma-separated</small></span>
                <input
                  value={analysis.organizations.join(", ")}
                  onChange={(event) => patchAnalysis({ organizations: csvToList(event.target.value) })}
                />
              </label>
            </div>
            <div className="form-two">
              <label>
                <span>Districts <small>comma-separated</small></span>
                <input value={analysis.districts.join(", ")} onChange={(event) => patchAnalysis({ districts: csvToList(event.target.value) })} />
              </label>
              <label>
                <span>Projects <small>comma-separated</small></span>
                <input value={analysis.projects.join(", ")} onChange={(event) => patchAnalysis({ projects: csvToList(event.target.value) })} />
              </label>
            </div>

            {analysis.people.length > 0 && (
              <fieldset>
                <legend>People — possible Contact matches</legend>
                <p className="muted-copy">
                  AI detected these names. Deterministic matching only suggests who they might be — you decide. Nothing is linked until you choose.
                </p>
                {(() => {
                  const entityMatches = resolveEmailAnalysisEntities(analysis, references);
                  const organizationContext = [...entityMatches.organizationIds, ...entityMatches.districtIds];
                  return analysis.people.map((personName) => (
                    <ContactMatchPanel
                      key={personName}
                      personName={personName}
                      candidates={matchContactCandidates(personName, references.contacts, { organizationIds: organizationContext })}
                      decision={personDecisions[personName]}
                      contacts={references.contacts}
                      organizations={references.organizations}
                      onMatch={(contactAppId) => decideMatch(personName, contactAppId)}
                      onIgnore={() => decideIgnore(personName)}
                      onReset={() => resetDecision(personName)}
                      onAddPerson={() => setAddPersonName(personName)}
                    />
                  ));
                })()}
              </fieldset>
            )}

            <fieldset>
              <legend>Suggested work record</legend>
              <label>
                <span>Title</span>
                <input
                  value={analysis.suggestedWorkRecord.title}
                  onChange={(event) =>
                    patchAnalysis({ suggestedWorkRecord: { ...analysis.suggestedWorkRecord, title: event.target.value } })
                  }
                />
              </label>
              <label>
                <span>Description</span>
                <textarea
                  rows={2}
                  value={analysis.suggestedWorkRecord.description}
                  onChange={(event) =>
                    patchAnalysis({ suggestedWorkRecord: { ...analysis.suggestedWorkRecord, description: event.target.value } })
                  }
                />
              </label>
            </fieldset>

            {usage && (
              <p className="muted-copy">
                Model: {usage.model} · {usage.inputTokens.toLocaleString()} in / {usage.outputTokens.toLocaleString()} out tokens
              </p>
            )}
            {saveError && (
              <div className="form-error" role="alert">
                {saveError}
              </div>
            )}
          </div>
          <footer className="log-footer">
            <button className="ghost-button" onClick={startOver} disabled={saving}>
              Discard
            </button>
            <button className="primary-action" onClick={() => void saveToInbox()} disabled={saving}>
              {saving ? "Saving…" : "Save to Inbox"}
            </button>
          </footer>
        </section>
      )}

      {addPersonName && (
        <ContactFormModal
          contact={{ ...emptyContactDraft(), displayName: addPersonName }}
          contacts={references.contacts}
          organizations={references.organizations}
          onCancel={() => setAddPersonName(null)}
          onSaved={(savedContact) => onPersonAdded(addPersonName, savedContact)}
          saveContact={saveContact}
          updateContact={updateContact}
        />
      )}

      {phase === "saved" && justSaved && (
        <section className="panel">
          <div className="non-orbit-note">
            <span>✓</span>
            <div>
              <strong>Saved to Inbox.</strong>
              <p>This intelligence record now appears below. Create a Work Record from it any time.</p>
            </div>
          </div>
          <footer className="log-footer">
            <button className="ghost-button" onClick={startOver}>
              Analyze another email
            </button>
            <button className="primary-action" onClick={() => createWorkRecordFrom(justSaved)}>
              Create Work Record
            </button>
          </footer>
        </section>
      )}

      {rowError && (
        <div className="form-error" role="alert">
          {rowError}
        </div>
      )}

      <InboxSection
        title="Needs attention"
        records={records.filter((record) => record.status === "open")}
        busyAppId={busyAppId}
        onCreateWorkRecord={createWorkRecordFrom}
        onUpdateStatus={updateStatus}
        references={references}
        empty="Nothing needs attention right now."
      />
      <InboxSection
        title="Waiting"
        records={records.filter((record) => record.status === "waiting")}
        busyAppId={busyAppId}
        onCreateWorkRecord={createWorkRecordFrom}
        onUpdateStatus={updateStatus}
        references={references}
        empty="Nothing is waiting on someone else."
      />
      <InboxSection
        title="Recent / resolved"
        records={records.filter((record) => record.status === "resolved").slice(0, 5)}
        busyAppId={busyAppId}
        onCreateWorkRecord={createWorkRecordFrom}
        onUpdateStatus={updateStatus}
        references={references}
        empty="Nothing resolved yet."
      />
    </div>
  );
}

function InboxSection({
  title,
  records,
  busyAppId,
  onCreateWorkRecord,
  onUpdateStatus,
  references,
  empty,
}: {
  title: string;
  records: InboxIntelligenceRecord[];
  busyAppId: string | null;
  onCreateWorkRecord: (record: InboxIntelligenceRecord) => void;
  onUpdateStatus: (record: InboxIntelligenceRecord, status: InboxIntelligenceStatus) => void;
  references: ReferenceData;
  empty: string;
}) {
  return (
    <section className="panel list-panel">
      <div className="panel-heading">
        <h2>{title}</h2>
        <span className="sample-label">{records.length}</span>
      </div>
      {records.length ? (
        records.map((record) => (
          <InboxRow
            key={record.appId}
            record={record}
            busy={busyAppId === record.appId}
            onCreateWorkRecord={onCreateWorkRecord}
            onUpdateStatus={onUpdateStatus}
            references={references}
          />
        ))
      ) : (
        <p className="muted-copy">{empty}</p>
      )}
    </section>
  );
}

function InboxRow({
  record,
  busy,
  onCreateWorkRecord,
  onUpdateStatus,
  references,
}: {
  record: InboxIntelligenceRecord;
  busy: boolean;
  onCreateWorkRecord: (record: InboxIntelligenceRecord) => void;
  onUpdateStatus: (record: InboxIntelligenceRecord, status: InboxIntelligenceStatus) => void;
  references: ReferenceData;
}) {
  const relatedProject = record.matchedProjectIds[0]
    ? references.projects.find((project) => project.appId === record.matchedProjectIds[0])?.name
    : null;
  const relatedOrg = record.matchedDistrictIds[0]
    ? references.organizations.find((org) => org.appId === record.matchedDistrictIds[0])?.name
    : record.matchedOrganizationIds[0]
      ? references.organizations.find((org) => org.appId === record.matchedOrganizationIds[0])?.name
      : null;
  // Patch 8D — read-only display of the human-reviewed Contact matches from Save time. There is
  // no post-save re-review UI, matching the existing Organization/District/Project precedent
  // above (also resolved once and shown read-only here, never re-editable from this list).
  const matchedContacts = resolveMatchedContacts(record.matchedContactIds, references.contacts);

  return (
    <div className="record-row" style={{ cursor: "default", flexWrap: "wrap", minHeight: "auto", padding: "10px 4px" }}>
      <span className={`record-dot ${record.analysis.needsAttention ? "iu" : "orbit"}`} />
      <span>
        <strong>{record.analysis.suggestedWorkRecord.title}</strong>
        <small>
          {record.analysis.priority} priority · {record.analysis.actionItems.length} action item
          {record.analysis.actionItems.length === 1 ? "" : "s"}
          {record.analysis.followUp ? ` · ${record.analysis.followUp}` : ""}
          {relatedProject ? ` · ${relatedProject}` : relatedOrg ? ` · ${relatedOrg}` : ""}
          {matchedContacts.length ? ` · ${matchedContacts.map((c) => c.displayName).join(", ")}` : ""} · updated {lastModified(record)}
        </small>
      </span>
      <span style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
        {record.status === "open" && (
          <>
            <button className="ghost-button" disabled={busy} onClick={() => onUpdateStatus(record, "waiting")}>
              Mark waiting
            </button>
            <button className="ghost-button" disabled={busy} onClick={() => onUpdateStatus(record, "resolved")}>
              Resolve
            </button>
          </>
        )}
        {record.status === "waiting" && (
          <>
            <button className="ghost-button" disabled={busy} onClick={() => onUpdateStatus(record, "resolved")}>
              Resolve
            </button>
            <button className="ghost-button" disabled={busy} onClick={() => onUpdateStatus(record, "open")}>
              Reopen
            </button>
          </>
        )}
        {record.status === "resolved" && (
          <button className="ghost-button" disabled={busy} onClick={() => onUpdateStatus(record, "open")}>
            Reopen
          </button>
        )}
        {record.linkedWorkRecordAppId ? (
          <span className="muted-copy">Linked to a Work Record</span>
        ) : (
          <button className="ghost-button" disabled={busy} onClick={() => onCreateWorkRecord(record)}>
            Create Work Record
          </button>
        )}
      </span>
    </div>
  );
}

function Metric({ value, label }: { value: string; label: string }) {
  return (
    <div className="metric">
      <strong>{value}</strong>
      <span>{label}</span>
    </div>
  );
}
