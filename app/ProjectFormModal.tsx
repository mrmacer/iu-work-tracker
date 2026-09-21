"use client";

import { useEffect, useRef, useState } from "react";
import type { Project } from "../lib/models";
import {
  buildProjectDraft,
  PROJECT_STATUSES,
  validateProjectShape,
  type ProjectResult,
  type ProjectStatus,
} from "../lib/project-provider";

// Patch 7 — Durable Projects' single Create/Edit Project form. Extracted to its own file in
// Patch 7.3 (exactly as Patch 8D extracted ContactFormModal) so both the Projects screen
// (app/IUWorkTracker.tsx) and Voice Intelligence's "Create Project" routing
// (app/VoiceIntelligence.tsx) share this ONE form and the ONE ProjectProvider — never a second
// Project form or provider. The only change from the original inline version: onSaved now
// receives the saved Project (existing callers ignore the argument).

// Cycled deterministically for new durable projects (see globals.css .project-mark.<color>) —
// the create form deliberately has no color picker (not part of the Patch 7 spec), so every
// new project keeps the exact same visual card design as the five seeded ones.
export const PROJECT_COLORS = ["blue", "coral", "lime", "purple", "yellow"] as const;

export function emptyProjectDraft(existingCount: number): Project {
  return buildProjectDraft({
    appId: crypto.randomUUID(),
    name: "",
    description: "",
    status: "planning",
    color: PROJECT_COLORS[existingCount % PROJECT_COLORS.length],
    startDate: null,
    targetDate: null,
    stemOrbit: false,
  });
}

/**
 * A single compact overlay handles both Create and Edit — see docs/AI_HANDOFF.md "Durable
 * Projects (Patch 7)". Create vs. update is decided the same way every other durable resource
 * in this codebase decides it: `project.metadata.version > 0` means the project already has a
 * durable identity, so submitting routes through updateProject(); otherwise saveProject().
 * Updating a project never touches its connected Work Records — this component only ever
 * calls the Project provider.
 */
export default function ProjectFormModal({
  project,
  onCancel,
  onSaved,
  saveProject,
  updateProject,
}: {
  project: Project;
  onCancel: () => void;
  onSaved: (saved: Project) => void;
  saveProject: (project: Project) => Promise<ProjectResult<Project>>;
  updateProject: (project: Project, expectedVersion: number) => Promise<ProjectResult<Project>>;
}) {
  const [draft, setDraft] = useState<Project>(project);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const nameRef = useRef<HTMLInputElement>(null);
  const isEditing = (draft.metadata?.version ?? 0) > 0;

  useEffect(() => {
    nameRef.current?.focus();
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onCancel();
      }
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [onCancel]);

  const patch = (change: Partial<Project>) => setDraft((current) => ({ ...current, ...change }));

  const submit = async () => {
    if (saving) return;
    const shapeIssues = validateProjectShape(draft);
    if (shapeIssues.length) {
      setError(shapeIssues[0].message);
      return;
    }
    setSaving(true);
    setError("");
    const result = isEditing ? await updateProject(draft, draft.metadata!.version) : await saveProject(draft);
    setSaving(false);
    if (result.status !== "success") {
      setError(result.status === "validation_error" ? (result.errors[0]?.message ?? "Check the project and try again.") : result.message);
      return;
    }
    onSaved(result.value);
  };

  return (
    <div className="modal-backdrop" role="presentation">
      <section className="project-modal" role="dialog" aria-modal="true" aria-labelledby="project-modal-title">
        <header className="log-header">
          <h2 id="project-modal-title">{isEditing ? "Edit Project" : "Create Project"}</h2>
          <button onClick={onCancel} aria-label="Close">
            ×
          </button>
        </header>
        <div className="log-content">
          <div className="form-stack">
            <label>
              <span>
                Project name <b>*</b>
              </span>
              <input ref={nameRef} value={draft.name} onChange={(event) => patch({ name: event.target.value })} placeholder="e.g. STEM Ecosystem" />
            </label>
            <label>
              <span>Description</span>
              <textarea value={draft.description} onChange={(event) => patch({ description: event.target.value })} placeholder="A sentence is enough." rows={2} />
            </label>
            <label>
              <span>Status</span>
              <select value={draft.status} onChange={(event) => patch({ status: event.target.value as ProjectStatus })}>
                {PROJECT_STATUSES.map((status) => (
                  <option key={status} value={status}>
                    {status[0].toUpperCase() + status.slice(1)}
                  </option>
                ))}
              </select>
            </label>
            <div className="form-two">
              <label>
                <span>Start date</span>
                <input type="date" value={draft.startDate ?? ""} onChange={(event) => patch({ startDate: event.target.value || null })} />
              </label>
              <label>
                <span>Target date</span>
                <input type="date" value={draft.targetDate ?? ""} onChange={(event) => patch({ targetDate: event.target.value || null })} />
              </label>
            </div>
            <div className="toggle-line">
              <input
                aria-label="STEM / ORBIT connection"
                type="checkbox"
                checked={draft.stemOrbit ?? false}
                onChange={(event) => patch({ stemOrbit: event.target.checked })}
              />
              <span>
                <strong>STEM / ORBIT connection</strong>
                <small>Optional — does not classify any Work Record automatically.</small>
              </span>
            </div>
          </div>
        </div>
        {error && (
          <div className="form-error" role="alert">
            {error}
          </div>
        )}
        <footer className="log-footer">
          <button className="ghost-button" onClick={onCancel}>
            Cancel
          </button>
          <button className="primary-action" onClick={() => void submit()} disabled={saving || !draft.name.trim()}>
            {saving ? "Saving…" : isEditing ? "Save Changes" : "Create Project"}
          </button>
        </footer>
      </section>
    </div>
  );
}

