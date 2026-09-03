"use client";

// Patch 8F — Reviewed Contact Import. IMPORTS PROVIDE EVIDENCE. DURABLE CONTACTS REMAIN
// HUMAN-REVIEWED TRUTH. Nothing from an imported file may silently overwrite an existing
// Contact, and nothing may automatically become a durable Contact — every creation still runs
// through the existing ContactFormModal / ContactProvider.create() path, one explicit human
// Save at a time. See docs/AI_HANDOFF.md "Reviewed Contact Import (Patch 8F)".
//
// The entire review session (parsed rows, decisions, which rows were created-here) is transient
// component state — nothing here is persisted, and nothing is written until the human explicitly
// saves an individual Contact.

import { useMemo, useState } from "react";
import ContactFormModal, { emptyContactDraft } from "./ContactFormModal";
import ContactMatchPanel from "./ContactMatchPanel";
import {
  buildContactImportCandidates,
  parseContactImportRows,
  type ContactImportCandidate,
  type ContactImportRawRow,
} from "../lib/contact-import";
import type { ContactResult } from "../lib/contact-provider";
import type { ContactMatchDecision } from "../lib/contact-matching";
import type { Contact, Organization } from "../lib/models";

export default function ContactImport({
  contacts,
  organizations,
  saveContact,
  updateContact,
  onBack,
}: {
  contacts: Contact[];
  organizations: Organization[];
  saveContact: (contact: Contact) => Promise<ContactResult<Contact>>;
  updateContact: (contact: Contact, expectedVersion: number) => Promise<ContactResult<Contact>>;
  onBack: () => void;
}) {
  const [fileName, setFileName] = useState<string | null>(null);
  const [error, setError] = useState("");
  const [rawRows, setRawRows] = useState<ContactImportRawRow[]>([]);
  // Keyed by candidate id ("<fileName>:<sourceRowNumber>") — matched/ignored, reusing
  // lib/contact-matching.ts's own decision type exactly, the same as Inbox/Voice review.
  const [decisions, setDecisions] = useState<Record<string, ContactMatchDecision>>({});
  // Tracks which rows' "matched" decision came from Create New rather than Match Existing —
  // session-summary bookkeeping only, never persisted, never a separate matching algorithm.
  const [createdRowIds, setCreatedRowIds] = useState<Set<string>>(new Set());
  const [addPersonForId, setAddPersonForId] = useState<string | null>(null);

  // Recomputed on every render from the CURRENT `contacts` prop — which the app already updates
  // reactively after any saveContact() call (the same mechanism Patch 8D's Inbox/Voice "Add
  // Person" relies on). A Contact created earlier in this same review session is therefore
  // automatically visible to every other row's matching with no separate session-contact list.
  const candidates = useMemo(
    () => (fileName ? buildContactImportCandidates(rawRows, fileName, organizations, contacts) : []),
    [rawRows, fileName, organizations, contacts],
  );

  const readFileText = (file: File): Promise<string> =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result ?? ""));
      reader.onerror = () => reject(reader.error ?? new Error("The file could not be read."));
      reader.readAsText(file);
    });

  const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = ""; // allow re-selecting the same file name after Clear
    if (!file) return;
    setError("");
    let text: string;
    try {
      text = await readFileText(file);
    } catch {
      setError("The file could not be read.");
      return;
    }
    const result = parseContactImportRows(text);
    if (result.status !== "success") {
      setError(result.message);
      setFileName(null);
      setRawRows([]);
      return;
    }
    setFileName(file.name);
    setRawRows(result.rows);
    setDecisions({});
    setCreatedRowIds(new Set());
  };

  const clearImport = () => {
    setFileName(null);
    setRawRows([]);
    setError("");
    setDecisions({});
    setCreatedRowIds(new Set());
  };

  const decideMatch = (id: string, contactAppId: string) => setDecisions((current) => ({ ...current, [id]: { type: "matched", contactAppId } }));
  const decideIgnore = (id: string) => setDecisions((current) => ({ ...current, [id]: { type: "ignored" } }));
  const resetDecision = (id: string) => {
    setDecisions((current) => {
      const next = { ...current };
      delete next[id];
      return next;
    });
    setCreatedRowIds((current) => {
      if (!current.has(id)) return current;
      const next = new Set(current);
      next.delete(id);
      return next;
    });
  };
  const onPersonCreated = (id: string, savedContact: Contact) => {
    decideMatch(id, savedContact.appId);
    setCreatedRowIds((current) => new Set(current).add(id));
    setAddPersonForId(null);
  };

  const addPersonCandidate = addPersonForId ? candidates.find((c) => c.id === addPersonForId) : undefined;

  const total = candidates.length;
  const reviewed = candidates.filter((c) => decisions[c.id]).length;
  const unreviewed = total - reviewed;
  const createdCount = candidates.filter((c) => decisions[c.id]?.type === "matched" && createdRowIds.has(c.id)).length;
  const matchedExistingCount = candidates.filter((c) => decisions[c.id]?.type === "matched" && !createdRowIds.has(c.id)).length;
  const ignoredCount = candidates.filter((c) => decisions[c.id]?.type === "ignored").length;

  return (
    <div className="screen-inner">
      <button type="button" className="ghost-button" onClick={onBack} style={{ marginLeft: 0, marginBottom: 13 }}>
        ← Back to Contacts
      </button>
      <PageHeading
        eyebrow="Bring in existing people"
        title="Import Contacts"
        copy="Import provides evidence. You decide what becomes a Contact — nothing is created or changed without an explicit review."
      />

      <section className="panel">
        <p className="eyebrow">Choose a CSV file</p>
        <p className="muted-copy">
          Recognized columns: Name (or Full Name) — required; Email (or Email Address); Role (or Title, Job Title); Organization (or
          Organisation, District, Company). Any other column is ignored.
        </p>
        <div className="form-stack">
          <label>
            <span>CSV file</span>
            <input type="file" accept=".csv,text/csv" aria-label="Choose CSV file" onChange={(event) => void handleFileChange(event)} />
          </label>
        </div>
        {error && (
          <div className="form-error" role="alert">
            {error}
          </div>
        )}
        {fileName && (
          <p className="muted-copy" style={{ marginTop: 9 }}>
            {fileName} · {total} row{total === 1 ? "" : "s"}{" "}
            <button type="button" className="ghost-button" onClick={clearImport} style={{ marginLeft: 6 }}>
              Clear
            </button>
          </p>
        )}
      </section>

      {total > 0 && (
        <>
          <div className="metric-strip">
            <Metric value={String(total)} label="total rows" />
            <Metric value={String(unreviewed)} label="unreviewed" />
            <Metric value={String(matchedExistingCount)} label="matched existing" />
            <Metric value={String(createdCount)} label="created" />
            <Metric value={String(ignoredCount)} label="ignored" />
          </div>

          <section className="panel list-panel">
            {candidates.map((candidate) => (
              <ImportRow
                key={candidate.id}
                candidate={candidate}
                decision={decisions[candidate.id]}
                contacts={contacts}
                organizations={organizations}
                onMatch={(contactAppId) => decideMatch(candidate.id, contactAppId)}
                onIgnore={() => decideIgnore(candidate.id)}
                onReset={() => resetDecision(candidate.id)}
                onAddPerson={() => setAddPersonForId(candidate.id)}
              />
            ))}
          </section>
        </>
      )}

      {addPersonCandidate && (
        <ContactFormModal
          contact={{
            ...emptyContactDraft(),
            displayName: addPersonCandidate.displayName,
            email: addPersonCandidate.email,
            role: addPersonCandidate.role,
            organizationId: addPersonCandidate.resolvedOrganizationId,
          }}
          contacts={contacts}
          organizations={organizations}
          onCancel={() => setAddPersonForId(null)}
          onSaved={(savedContact) => onPersonCreated(addPersonCandidate.id, savedContact)}
          saveContact={saveContact}
          updateContact={updateContact}
        />
      )}
    </div>
  );
}

function ImportRow({
  candidate,
  decision,
  contacts,
  organizations,
  onMatch,
  onIgnore,
  onReset,
  onAddPerson,
}: {
  candidate: ContactImportCandidate;
  decision: ContactMatchDecision | undefined;
  contacts: Contact[];
  organizations: Organization[];
  onMatch: (contactAppId: string) => void;
  onIgnore: () => void;
  onReset: () => void;
  onAddPerson: () => void;
}) {
  const resolvedOrganizationName = candidate.resolvedOrganizationId
    ? organizations.find((org) => org.appId === candidate.resolvedOrganizationId)?.name
    : undefined;

  return (
    <div className="import-row">
      <div className="import-row-evidence">
        <span className="muted-copy">Row {candidate.sourceRowNumber}</span>
        <strong>{candidate.displayName || "(no name in this row)"}</strong>
        {candidate.email && <span className="muted-copy">{candidate.email}</span>}
        {candidate.role && <span className="muted-copy">{candidate.role}</span>}
        {candidate.organizationText && (
          <span className="muted-copy">
            Organization: {candidate.organizationText}
            {resolvedOrganizationName ? ` → ${resolvedOrganizationName}` : " (not resolved — will be left unassigned)"}
          </span>
        )}
      </div>
      <ContactMatchPanel
        personName={candidate.displayName || `Row ${candidate.sourceRowNumber}`}
        candidates={candidate.matches}
        decision={decision}
        contacts={contacts}
        organizations={organizations}
        onMatch={onMatch}
        onIgnore={onIgnore}
        onReset={onReset}
        onAddPerson={onAddPerson}
      />
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

function PageHeading({ eyebrow, title, copy }: { eyebrow: string; title: string; copy: string }) {
  return (
    <div className="page-heading">
      <div>
        <p className="eyebrow">{eyebrow}</p>
        <h1>{title}</h1>
        <p>{copy}</p>
      </div>
    </div>
  );
}
