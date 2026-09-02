"use client";

// Patch 8D — shared presentational review UI for one Intelligence-detected person, used by both
// Inbox Intelligence (app/InboxIntelligence.tsx) and Voice Intelligence (app/VoiceIntelligence.tsx)
// so the two never grow inconsistent matching UIs. Purely a view over
// lib/contact-matching.ts's deterministic output — this component itself makes no matching
// decisions, no AI calls, and no network calls. See docs/AI_HANDOFF.md "Intelligence Contact
// matching (Patch 8D)".
//
// AI MAY PROPOSE. THE HUMAN DECIDES. Nothing here commits a match — every candidate, even a
// "strong" exact-email match, is a suggestion until the human clicks "Match Existing".

import type { ContactMatchCandidate, ContactMatchDecision } from "../lib/contact-matching";
import type { Contact, Organization } from "../lib/models";

const STRENGTH_LABELS: Record<ContactMatchCandidate["strength"], string> = {
  strong: "Exact email match",
  useful: "Name + organization match",
  review: "Name match — please confirm",
  none: "",
};

function organizationName(organizations: Organization[], organizationId: string | null | undefined): string {
  return organizationId ? (organizations.find((org) => org.appId === organizationId)?.name ?? "") : "";
}

export default function ContactMatchPanel({
  personName,
  candidates,
  decision,
  contacts,
  organizations,
  onMatch,
  onIgnore,
  onReset,
  onAddPerson,
}: {
  personName: string;
  candidates: ContactMatchCandidate[];
  decision: ContactMatchDecision | undefined;
  /** The full current Contact reference set (seed + durable) — used ONLY to resolve a
   * matched contact's display name/organization for the "matched" state below. Deliberately
   * NOT `candidates`: a matched decision must keep displaying correctly even if a later edit
   * to the detected name changes what `candidates` recomputes to. */
  contacts: Contact[];
  organizations: Organization[];
  onMatch: (contactAppId: string) => void;
  onIgnore: () => void;
  onReset: () => void;
  onAddPerson: () => void;
}) {
  if (decision?.type === "matched") {
    const matchedContact = contacts.find((c) => c.appId === decision.contactAppId);
    const org = matchedContact ? organizationName(organizations, matchedContact.organizationId) : "";
    return (
      <div className="contact-match-panel matched">
        <span>
          <strong>{personName}</strong>
          <small>Matched Contact: {matchedContact?.displayName ?? "Contact"}{org ? ` · ${org}` : ""}</small>
        </span>
        <button type="button" className="ghost-button" onClick={onReset}>
          Change
        </button>
      </div>
    );
  }

  if (decision?.type === "ignored") {
    return (
      <div className="contact-match-panel ignored">
        <span>
          <strong>{personName}</strong>
          <small>Ignored for this item</small>
        </span>
        <button type="button" className="ghost-button" onClick={onReset}>
          Reconsider
        </button>
      </div>
    );
  }

  return (
    <div className="contact-match-panel">
      <span>
        <strong>{personName}</strong>
        <small>
          {candidates.length === 0
            ? "No reliable match found."
            : candidates.length === 1
              ? "Possible match:"
              : `${candidates.length} possible matches:`}
        </small>
      </span>
      {candidates.length > 0 && (
        <ul className="contact-match-candidates">
          {candidates.map((candidate) => (
            <li key={candidate.contact.appId}>
              <span>
                {candidate.contact.displayName}
                {organizationName(organizations, candidate.contact.organizationId) ? ` · ${organizationName(organizations, candidate.contact.organizationId)}` : ""}
                <small className="muted-copy"> — {STRENGTH_LABELS[candidate.strength]}</small>
              </span>
              <button type="button" className="ghost-button" onClick={() => onMatch(candidate.contact.appId)}>
                Match Existing
              </button>
            </li>
          ))}
        </ul>
      )}
      <div className="contact-match-actions">
        <button type="button" className="ghost-button" onClick={onAddPerson}>
          Add Person
        </button>
        <button type="button" className="ghost-button" onClick={onIgnore}>
          Ignore
        </button>
      </div>
    </div>
  );
}

/** Resolves a saved contact's basic display info for a persisted (already-decided) matched-person
 * row — e.g. InboxIntelligenceRecord.matchedContactIds after Save. Pure lookup, no matching logic. */
export function resolveMatchedContacts(contactIds: string[], contacts: Contact[]): Contact[] {
  const byAppId = new Map(contacts.map((c) => [c.appId, c]));
  return contactIds.map((id) => byAppId.get(id)).filter((c): c is Contact => Boolean(c));
}
