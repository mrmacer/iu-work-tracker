"use client";

import { useEffect, useRef, useState } from "react";
import type { Organization } from "../lib/models";
import {
  buildOrganizationDraft,
  normalizeOrganizationName,
  ORGANIZATION_TYPES,
  validateOrganizationShape,
  type OrganizationResult,
  type OrganizationTypeValue,
} from "../lib/organization-provider";

// Patch 8E — Durable Organizations' single Create/Edit Organization form. Extracted to its own
// file in Patch 7.3 (exactly as Patch 8D extracted ContactFormModal) so both the Organizations
// screen (app/IUWorkTracker.tsx) and Voice Intelligence's "Create Organization" routing
// (app/VoiceIntelligence.tsx) share this ONE form and the ONE OrganizationProvider — never a
// second Organization form or provider. The only change from the original inline version:
// onSaved now receives the saved Organization (existing callers ignore the argument).

export const ORGANIZATION_TYPE_LABELS: Record<OrganizationTypeValue, string> = {
  district: "District",
  partner: "Partner",
  iu: "IU",
};

export function emptyOrganizationDraft(): Organization {
  return buildOrganizationDraft({ appId: crypto.randomUUID(), name: "", type: "partner" });
}

/**
 * A single compact overlay handles both Create and Edit — mirrors ContactFormModal/
 * ProjectFormModal exactly. Duplicate detection is conservative and deterministic (trim +
 * collapse whitespace + lowercase, exact match only — never fuzzy): a name-only match is an
 * informational warning that never blocks Save, since Organization has no email field to serve
 * as a stronger signal.
 */
export default function OrganizationFormModal({
  organization,
  organizations,
  onCancel,
  onSaved,
  saveOrganization,
  updateOrganization,
}: {
  organization: Organization;
  organizations: Organization[];
  onCancel: () => void;
  onSaved: (saved: Organization) => void;
  saveOrganization: (organization: Organization) => Promise<OrganizationResult<Organization>>;
  updateOrganization: (organization: Organization, expectedVersion: number) => Promise<OrganizationResult<Organization>>;
}) {
  const [draft, setDraft] = useState<Organization>(organization);
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

  const patch = (change: Partial<Organization>) => setDraft((current) => ({ ...current, ...change }));

  const otherOrganizations = organizations.filter((item) => item.appId !== draft.appId);
  const nameDuplicate = draft.name.trim()
    ? otherOrganizations.find((item) => normalizeOrganizationName(item.name) === normalizeOrganizationName(draft.name))
    : undefined;

  const submit = async () => {
    if (saving) return;
    const shapeIssues = validateOrganizationShape(draft);
    if (shapeIssues.length) {
      setError(shapeIssues[0].message);
      return;
    }
    setSaving(true);
    setError("");
    const result = isEditing ? await updateOrganization(draft, draft.metadata!.version) : await saveOrganization(draft);
    setSaving(false);
    if (result.status !== "success") {
      setError(result.status === "validation_error" ? (result.errors[0]?.message ?? "Check the organization and try again.") : result.message);
      return;
    }
    onSaved(result.value);
  };

  return (
    <div className="modal-backdrop" role="presentation">
      <section className="project-modal" role="dialog" aria-modal="true" aria-labelledby="organization-modal-title">
        <header className="log-header">
          <h2 id="organization-modal-title">{isEditing ? "Edit Organization" : "Add Organization"}</h2>
          <button onClick={onCancel} aria-label="Close">
            ×
          </button>
        </header>
        <div className="log-content">
          <div className="form-stack">
            <label>
              <span>
                Name <b>*</b>
              </span>
              <input
                ref={nameRef}
                value={draft.name}
                onChange={(event) => patch({ name: event.target.value })}
                placeholder="e.g. North Valley SD"
              />
            </label>
            {nameDuplicate && (
              <p className="muted-copy" role="status">
                Another organization is already named &ldquo;{nameDuplicate.name}&rdquo;. This is just a heads up — Save is not blocked.
              </p>
            )}
            <label>
              <span>Type</span>
              <select value={draft.type} onChange={(event) => patch({ type: event.target.value as OrganizationTypeValue })}>
                {ORGANIZATION_TYPES.map((type) => (
                  <option key={type} value={type}>
                    {ORGANIZATION_TYPE_LABELS[type]}
                  </option>
                ))}
              </select>
            </label>
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
            {saving ? "Saving…" : isEditing ? "Save Changes" : "Add Organization"}
          </button>
        </footer>
      </section>
    </div>
  );
}
