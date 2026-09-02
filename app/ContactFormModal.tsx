"use client";

import { useEffect, useRef, useState } from "react";
import {
  buildContactDraft,
  CONTACT_STATUSES,
  normalizeContactEmail,
  normalizeContactName,
  validateContactShape,
  type ContactResult,
} from "../lib/contact-provider";
import type { Contact, ContactStatus, Organization } from "../lib/models";

// Patch 8B — Durable Contacts' single Create/Edit Contact form. Extracted to its own file in
// Patch 8D so both the Contacts screen (app/IUWorkTracker.tsx) and Intelligence review flows
// (app/InboxIntelligence.tsx, app/VoiceIntelligence.tsx — "Add Person", lib/contact-matching.ts)
// share this exact path rather than a second Contact form or a second ContactProvider. See
// docs/AI_HANDOFF.md "Intelligence Contact matching (Patch 8D)".

export const CONTACT_STATUS_LABELS: Record<ContactStatus, string> = {
  active: "Active",
  developing: "Developing",
  occasional: "Occasional",
  dormant: "Dormant",
  archived: "Archived",
};

export function emptyContactDraft(): Contact {
  return buildContactDraft({
    appId: crypto.randomUUID(),
    displayName: "",
    role: undefined,
    organizationId: null,
    email: undefined,
    status: "active",
    notes: undefined,
  });
}

/**
 * A single compact overlay handles both Create and Edit — see docs/AI_HANDOFF.md "Durable
 * Contacts (Patch 8B)". Create vs. update is decided the same way every other durable resource
 * decides it: `contact.metadata.version > 0` means the contact already has a durable identity.
 * Duplicate detection is conservative and deterministic (trim + lowercase, exact match only —
 * never fuzzy, never first-name-only): a name-only match is an informational warning that
 * never blocks Save; an email match requires an explicit second confirmation ("Create Anyway")
 * before it will save, since email is the strongest available matching signal.
 */
export default function ContactFormModal({
  contact,
  contacts,
  organizations,
  onCancel,
  onSaved,
  saveContact,
  updateContact,
}: {
  contact: Contact;
  contacts: Contact[];
  organizations: Organization[];
  onCancel: () => void;
  onSaved: (savedContact: Contact) => void;
  saveContact: (contact: Contact) => Promise<ContactResult<Contact>>;
  updateContact: (contact: Contact, expectedVersion: number) => Promise<ContactResult<Contact>>;
}) {
  const [draft, setDraft] = useState<Contact>(contact);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [emailDuplicateConfirmed, setEmailDuplicateConfirmed] = useState(false);
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

  const patch = (change: Partial<Contact>) => setDraft((current) => ({ ...current, ...change }));

  const otherContacts = contacts.filter((item) => item.appId !== draft.appId);
  const emailDuplicate =
    draft.email && draft.email.trim()
      ? otherContacts.find((item) => item.email && normalizeContactEmail(item.email) === normalizeContactEmail(draft.email!))
      : undefined;
  const nameDuplicate = draft.displayName.trim()
    ? otherContacts.find((item) => normalizeContactName(item.displayName) === normalizeContactName(draft.displayName))
    : undefined;

  const submit = async () => {
    if (saving) return;
    const shapeIssues = validateContactShape(draft);
    if (shapeIssues.length) {
      setError(shapeIssues[0].message);
      return;
    }
    if (emailDuplicate && !emailDuplicateConfirmed) {
      // Require an explicit second decision — never silently create a second Contact with the
      // same email, and never silently merge. See docs/AI_HANDOFF.md "Durable Contacts (Patch
      // 8B)" duplicate-detection rules.
      setEmailDuplicateConfirmed(true);
      return;
    }
    setSaving(true);
    setError("");
    const result = isEditing ? await updateContact(draft, draft.metadata!.version) : await saveContact(draft);
    setSaving(false);
    if (result.status !== "success") {
      setError(result.status === "validation_error" ? (result.errors[0]?.message ?? "Check the contact and try again.") : result.message);
      return;
    }
    onSaved(result.value);
  };

  return (
    <div className="modal-backdrop" role="presentation">
      <section className="project-modal" role="dialog" aria-modal="true" aria-labelledby="contact-modal-title">
        <header className="log-header">
          <h2 id="contact-modal-title">{isEditing ? "Edit Contact" : "Add Contact"}</h2>
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
                value={draft.displayName}
                onChange={(event) => {
                  patch({ displayName: event.target.value });
                  setEmailDuplicateConfirmed(false);
                }}
                placeholder="e.g. Annie Milewski"
              />
            </label>
            {nameDuplicate && (
              <p className="muted-copy" role="status">
                Another contact is already named &ldquo;{nameDuplicate.displayName}&rdquo;. Two real people can share a name — this is just a heads up.
              </p>
            )}
            <label>
              <span>Role</span>
              <input value={draft.role ?? ""} onChange={(event) => patch({ role: event.target.value || undefined })} placeholder="e.g. Superintendent" />
            </label>
            <label>
              <span>Organization</span>
              <select value={draft.organizationId ?? ""} onChange={(event) => patch({ organizationId: event.target.value || null })}>
                <option value="">No organization</option>
                {organizations.map((org) => (
                  <option key={org.appId} value={org.appId}>
                    {org.name}
                  </option>
                ))}
              </select>
            </label>
            <label>
              <span>Email</span>
              <input
                type="email"
                value={draft.email ?? ""}
                onChange={(event) => {
                  patch({ email: event.target.value || undefined });
                  setEmailDuplicateConfirmed(false);
                }}
                placeholder="name@example.org"
              />
            </label>
            {emailDuplicate && (
              <div className="form-error" role="alert">
                Another contact ({emailDuplicate.displayName}) already uses this email address. Click {isEditing ? "Save Changes" : "Add Contact"} again to
                create a separate contact anyway, or change the email.
              </div>
            )}
            <label>
              <span>Relationship status</span>
              <select value={draft.status} onChange={(event) => patch({ status: event.target.value as ContactStatus })}>
                {CONTACT_STATUSES.map((status) => (
                  <option key={status} value={status}>
                    {CONTACT_STATUS_LABELS[status]}
                  </option>
                ))}
              </select>
            </label>
            <label>
              <span>Notes</span>
              <textarea
                value={draft.notes ?? ""}
                onChange={(event) => patch({ notes: event.target.value || undefined })}
                placeholder="A sentence or two — not a biography."
                rows={2}
              />
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
          <button className="primary-action" onClick={() => void submit()} disabled={saving || !draft.displayName.trim()}>
            {saving ? "Saving…" : isEditing ? "Save Changes" : "Add Contact"}
          </button>
        </footer>
      </section>
    </div>
  );
}
