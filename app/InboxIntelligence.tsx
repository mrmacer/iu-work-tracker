"use client";

import { useRef, useState } from "react";
import { MAX_EMAIL_LENGTH } from "../lib/anthropic-config";
import type { AnalyzeEmailResult, AnalyzeEmailUsage } from "../lib/anthropic-email-analysis";
import { buildWorkRecordDraftFromAnalysis } from "../lib/inbox-intelligence-work-record";
import type { EmailAnalysis, InboxIntelligenceRecord } from "../lib/inbox-intelligence-models";
import { SessionInboxIntelligenceProvider } from "../lib/inbox-intelligence-provider";
import type { ReferenceData, WorkRecord } from "../lib/models";

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

export default function InboxIntelligence({
  references,
  openLog,
  createDraftRecord,
}: {
  references: ReferenceData;
  openLog: (record?: WorkRecord) => void;
  createDraftRecord: () => WorkRecord;
}) {
  const provider = useRef(new SessionInboxIntelligenceProvider());
  const [phase, setPhase] = useState<Phase>("paste");
  const [rawEmail, setRawEmail] = useState("");
  const [analyzing, setAnalyzing] = useState(false);
  const [error, setError] = useState("");
  const [analysis, setAnalysis] = useState<EmailAnalysis | null>(null);
  const [usage, setUsage] = useState<AnalyzeEmailUsage | null>(null);
  const [sourceExcerpt, setSourceExcerpt] = useState("");
  // Mirrors the session-only provider's contents so render never reads ref.current directly.
  const [recent, setRecent] = useState<InboxIntelligenceRecord[]>([]);
  const needsAttentionCount = recent.filter((record) => record.analysis.needsAttention).length;
  const openActionCount = recent.reduce((sum, record) => sum + record.analysis.actionItems.length, 0);

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
    clearAll();
  };

  const save = () => {
    if (!analysis) return;
    provider.current.save({
      appId: crypto.randomUUID(),
      createdAt: new Date().toISOString(),
      sourceExcerpt,
      analysis,
      linkedWorkRecordAppId: null,
    });
    setRecent(provider.current.list());
    setPhase("saved");
  };

  const createWorkRecord = () => {
    if (!analysis) return;
    const draft = buildWorkRecordDraftFromAnalysis(analysis, references, createDraftRecord());
    openLog(draft);
  };

  return (
    <div className="screen-inner">
      <div className="page-heading">
        <div>
          <p className="eyebrow">AI-assisted intake</p>
          <h1>Inbox Intelligence</h1>
          <p>Paste a work email. Review what AI finds. Decide what becomes a work record.</p>
        </div>
      </div>

      <div className="metric-strip">
        <Metric value={String(needsAttentionCount)} label="needs attention" />
        <Metric value={String(openActionCount)} label="open actions" />
        <Metric value={String(recent.length)} label="analyzed this session" />
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
          </div>
          <footer className="log-footer">
            <button className="ghost-button" onClick={startOver}>
              Discard
            </button>
            <button className="primary-action" onClick={save}>
              Save
            </button>
          </footer>
        </section>
      )}

      {phase === "saved" && analysis && (
        <section className="panel">
          <div className="non-orbit-note">
            <span>✓</span>
            <div>
              <strong>Saved for this session.</strong>
              <p>This intelligence record is not yet a Work Record. Create one below if this email should become one.</p>
            </div>
          </div>
          <footer className="log-footer">
            <button className="ghost-button" onClick={startOver}>
              Analyze another email
            </button>
            <button className="primary-action" onClick={createWorkRecord}>
              Create Work Record
            </button>
          </footer>
        </section>
      )}

      <section className="panel list-panel">
        <div className="panel-heading">
          <h2>Recently analyzed</h2>
          <span className="sample-label">This session only</span>
        </div>
        {recent.length ? (
          recent.slice(0, 5).map((record) => (
            <div className="record-row" key={record.appId} style={{ cursor: "default" }}>
              <span className={`record-dot ${record.analysis.needsAttention ? "iu" : "orbit"}`} />
              <span>
                <strong>{record.analysis.suggestedWorkRecord.title}</strong>
                <small>{record.sourceExcerpt || record.analysis.summary}</small>
              </span>
              <b>{record.analysis.priority}</b>
            </div>
          ))
        ) : (
          <p className="muted-copy">Nothing analyzed yet this session.</p>
        )}
      </section>
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
