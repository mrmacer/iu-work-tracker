import type { Contact } from "./models";
import { normalizeContactEmail, normalizeContactName } from "./contact-provider";

// Patch 8D — deterministic Contact identity resolution for Intelligence-detected people. AI
// proposes evidence (a name, sometimes an email); this module NEVER calls AI, NEVER performs
// network I/O, and NEVER picks a Contact on the caller's behalf — it only classifies how
// strong the evidence is so a human can decide. See docs/AI_HANDOFF.md "Intelligence Contact
// matching (Patch 8D)".
//
// The matching ladder, strongest to weakest:
//   strong  — exact normalized email match.
//   useful  — exact normalized full name match AND the Contact's organization is among the
//             organization/district IDs already resolved as relevant to the same Intelligence
//             item (e.g. an email's matchedOrganizationIds/matchedDistrictIds).
//   review  — exact normalized full name match only, with no organization evidence.
//   none    — anything weaker: an empty string, or a single-token name ("Annie") — a first
//             name alone is never sufficient identity evidence and must never resolve to a
//             candidate, no matter how few Contacts share it.
//
// Archived Contacts are never excluded from matching — Contact.status describes the durable
// relationship, not whether the person remains a real, matchable identity (see Contact Detail,
// Patch 8C, which shows connected work for archived Contacts the same way).

export type ContactMatchStrength = "strong" | "useful" | "review" | "none";

export type ContactMatchCandidate = {
  contact: Contact;
  strength: ContactMatchStrength;
  reasons: string[];
};

/** A person-candidate's human review decision. Transient by default — durable only where the
 * calling Intelligence record already persists comparable reviewed entity links (Inbox). */
export type ContactMatchDecision = { type: "matched"; contactAppId: string } | { type: "ignored" };

function collapseWhitespace(value: string): string {
  return value.trim().replace(/\s+/g, " ");
}

/** Reuses the exact normalization Contact duplicate-detection already uses
 * (lib/contact-provider.ts), plus internal-whitespace collapsing — AI-extracted names are
 * freer-form than a typed form field ("Annie   Milewski", stray newlines, etc.). */
function normalizedName(value: string): string {
  return normalizeContactName(collapseWhitespace(value));
}

/** A single token ("Annie") is never sufficient identity evidence — only a multi-token full
 * name counts as reviewable evidence at all. */
function isFullName(value: string): boolean {
  return normalizedName(value).split(" ").filter(Boolean).length >= 2;
}

export type ContactMatchOptions = {
  /** An email associated with the detected person, when the source evidence included one. */
  email?: string | null;
  /** Stable Organization/District appIds already resolved as relevant to this same
   * Intelligence item — used only to elevate a name-only match to "useful" when the
   * candidate Contact's own organizationId is among them. Never a free-text organization-name
   * comparison — see docs/DATA_MODEL.md "Reference and configuration entities". */
  organizationIds?: string[];
};

/**
 * Deterministic, side-effect-free, zero AI/zero network calls. Returns every plausible Contact
 * for `detectedName` — never a single "best" pick — so ambiguity (two Contacts sharing a name)
 * is always surfaced for human review rather than resolved by array order (see Step F, Patch
 * 8D instructions). An empty result means "no reliable match found" — the caller renders that
 * as an empty state, never a fabricated candidate.
 */
export function matchContactCandidates(
  detectedName: string,
  contacts: Contact[],
  options: ContactMatchOptions = {},
): ContactMatchCandidate[] {
  if (!detectedName.trim()) return [];

  // STRONG: exact normalized email match is the strongest signal and is evaluated
  // independently of name evidence — a matched email always wins.
  if (options.email && options.email.trim()) {
    const targetEmail = normalizeContactEmail(options.email);
    const emailMatches = contacts.filter((contact) => contact.email && normalizeContactEmail(contact.email) === targetEmail);
    if (emailMatches.length) {
      return emailMatches.map((contact) => ({ contact, strength: "strong", reasons: ["Exact email match"] }));
    }
  }

  if (!isFullName(detectedName)) return [];

  const target = normalizedName(detectedName);
  const nameMatches = contacts.filter((contact) => normalizedName(contact.displayName) === target);
  if (!nameMatches.length) return [];

  const organizationIds = new Set(options.organizationIds ?? []);
  return nameMatches.map((contact) => {
    const organizationEvidence = Boolean(contact.organizationId && organizationIds.has(contact.organizationId));
    return {
      contact,
      strength: organizationEvidence ? "useful" : "review",
      reasons: organizationEvidence ? ["Name and organization match"] : ["Name match only — organization not confirmed"],
    };
  });
}
