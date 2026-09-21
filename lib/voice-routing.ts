import type { Organization, Project } from "./models";
import { normalizeOrganizationName } from "./organization-provider";
import {
  VOICE_CANDIDATE_TYPES,
  VOICE_CANDIDATE_TYPE_LABELS,
  type VoiceCandidateType,
} from "./voice-intelligence-models";
import type { VoiceReviewCandidate } from "./voice-intelligence-session";

// Patch 7.3 — Voice Intelligence SELECT + ROUTE. AI EXTRACTS. HUMAN SELECTS. THE SYSTEM ROUTES.
//
// Pure, deterministic, zero-I/O helpers: which destination a candidate type has, how a selected
// batch groups, and the smallest exact-name matching for Organization / Project candidates.
// No AI, no network, no provider — every durable write still happens only through the existing
// providers' own forms after an explicit human Save (see app/VoiceIntelligence.tsx).

export type VoiceDestinationKind = "work-record" | "contact" | "organization" | "project" | "none";

export type VoiceDestination = { kind: VoiceDestinationKind; label: string };

/**
 * The single source of truth for where each candidate type can go TODAY. Only four durable
 * destinations exist in this codebase (Work Records, Contacts, Organizations, Projects) —
 * Action Center is merely a derived view over Inbox Intelligence records, and there is no
 * Knowledge / Idea / Decision / Question domain at all. Those types are honestly "none" rather
 * than fake-persisted.
 */
export const VOICE_DESTINATIONS: Record<VoiceCandidateType, VoiceDestination> = {
  COMPLETED_WORK: { kind: "work-record", label: "Work Records" },
  PERSON: { kind: "contact", label: "Contacts" },
  ORGANIZATION: { kind: "organization", label: "Organizations" },
  // District is Organization.type === "district" — there is no separate District model.
  DISTRICT: { kind: "organization", label: "Organizations (districts)" },
  PROJECT: { kind: "project", label: "Projects" },
  ACTION: { kind: "none", label: "" },
  IDEA: { kind: "none", label: "" },
  DECISION: { kind: "none", label: "" },
  QUESTION: { kind: "none", label: "" },
  KNOWLEDGE: { kind: "none", label: "" },
};

export function hasVoiceDestination(type: VoiceCandidateType): boolean {
  return VOICE_DESTINATIONS[type].kind !== "none";
}

export type VoiceCandidateGroup = {
  type: VoiceCandidateType;
  label: string;
  candidates: VoiceReviewCandidate[];
};

/**
 * Splits the SELECTED candidates into groups by type — ready-destination groups and
 * no-destination groups — in stable candidate-type order. Unselected candidates are never
 * included (and never touched); unsupported types are never dropped.
 */
export function groupSelectedCandidates(candidates: VoiceReviewCandidate[]): {
  ready: VoiceCandidateGroup[];
  unsupported: VoiceCandidateGroup[];
} {
  const groups = VOICE_CANDIDATE_TYPES.map((type) => ({
    type,
    label: VOICE_CANDIDATE_TYPE_LABELS[type],
    candidates: candidates.filter((candidate) => candidate.selected && candidate.type === type),
  })).filter((group) => group.candidates.length > 0);
  return {
    ready: groups.filter((group) => hasVoiceDestination(group.type)),
    unsupported: groups.filter((group) => !hasVoiceDestination(group.type)),
  };
}

function matchesByExactName<T extends { name: string }>(name: string, items: T[]): T[] {
  const normalized = normalizeOrganizationName(name);
  if (!normalized) return [];
  return items.filter((item) => normalizeOrganizationName(item.name) === normalized);
}

/**
 * Exact normalized-name (trim + collapse whitespace + lowercase) matches only — never fuzzy,
 * never AI. Mirrors resolveEmailAnalysisEntities: an ORGANIZATION candidate matches only
 * non-district Organizations, a DISTRICT candidate only district-type ones. Returns EVERY exact
 * match so a same-name ambiguity always reaches the human (never picked by array order).
 */
export function matchOrganizationCandidates(name: string, organizations: Organization[], kind: "ORGANIZATION" | "DISTRICT"): Organization[] {
  return matchesByExactName(
    name,
    organizations.filter((organization) => (kind === "DISTRICT" ? organization.type === "district" : organization.type !== "district")),
  );
}

/** Exact normalized-name Project matches only — never fuzzy, never AI; every exact match is returned. */
export function matchProjectCandidates(name: string, projects: Project[]): Project[] {
  return matchesByExactName(name, projects);
}

/**
 * The small "processed" chip for a candidate whose durable action already SUCCEEDED, or null.
 * ("Logged ✓" for completed work is rendered separately.) Only ever derived from session state
 * that was set after a successful durable action — never from a button click.
 */
export function processedChipLabel(candidate: VoiceReviewCandidate): string | null {
  if (candidate.type === "PERSON") {
    return candidate.contactCreated && candidate.contactDecision?.type === "matched" ? "Created Contact ✓" : null;
  }
  const noun = candidate.type === "PROJECT" ? "Project" : candidate.type === "ORGANIZATION" || candidate.type === "DISTRICT" ? "Organization" : null;
  if (!noun || !candidate.routing) return null;
  if (candidate.routing.type === "matched") return `Matched ${noun} ✓`;
  if (candidate.routing.type === "created") return `Created ${noun} ✓`;
  return null;
}

/** Plain text a human can paste elsewhere for a candidate with no durable destination yet. */
export function formatCandidateForCopy(candidate: VoiceReviewCandidate): string {
  const heading = `${VOICE_CANDIDATE_TYPE_LABELS[candidate.type]}: ${candidate.title}`;
  return candidate.detail.trim() ? `${heading}\n${candidate.detail.trim()}` : heading;
}
