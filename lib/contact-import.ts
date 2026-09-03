import { CsvParseError, parseCsv } from "./csv-parser";
import { matchContactCandidates, type ContactMatchCandidate } from "./contact-matching";
import { normalizeOrganizationName } from "./organization-provider";
import type { Contact, Organization } from "./models";

// Patch 8F — Reviewed Contact Import. IMPORTS PROVIDE EVIDENCE. DURABLE CONTACTS REMAIN
// HUMAN-REVIEWED TRUTH. This module only parses a CSV file into transient candidate rows and
// deterministically classifies matching evidence — it never persists anything, never calls AI,
// and never decides anything on the human's behalf. See docs/AI_HANDOFF.md "Reviewed Contact
// Import (Patch 8F)".
//
// Deliberately reuses, rather than reimplements:
//   - lib/contact-matching.ts's matchContactCandidates() for all duplicate-detection —  no
//     second matching algorithm exists anywhere in this file.
//   - lib/organization-provider.ts's normalizeOrganizationName() for Organization resolution
//     normalization — the exact same rule Patch 8E's own duplicate-name warning uses.

// ---------------------------------------------------------------------------------------------
// Header recognition — CSV/XLSX headers vary by source; only these four Contact concepts are
// ever recognized. Every other column is silently ignored (never absorbed into Notes or any
// other field) — see the header-alias table below for the exact recognized aliases.
// ---------------------------------------------------------------------------------------------

export type ImportHeaderConcept = "name" | "email" | "role" | "organization";

/** Recognized header aliases, already lowercase/whitespace-normalized for direct comparison
 * against normalizeHeader()'s output. Unknown headers map to no concept and are ignored. */
export const IMPORT_HEADER_ALIASES: Record<ImportHeaderConcept, string[]> = {
  name: ["name", "full name"],
  email: ["email", "email address"],
  role: ["role", "title", "job title"],
  organization: ["organization", "organisation", "district", "company"],
};

/** trim + collapse internal whitespace + lowercase — the same normalization discipline every
 * other deterministic matcher in this codebase uses. */
function normalizeHeader(value: string): string {
  return value.trim().replace(/\s+/g, " ").toLowerCase();
}

function headerConceptFor(header: string): ImportHeaderConcept | null {
  const normalized = normalizeHeader(header);
  for (const concept of Object.keys(IMPORT_HEADER_ALIASES) as ImportHeaderConcept[]) {
    if (IMPORT_HEADER_ALIASES[concept].includes(normalized)) return concept;
  }
  return null;
}

// ---------------------------------------------------------------------------------------------
// Raw parsing — CSV text to plain evidence rows, before any Organization resolution or Contact
// matching happens. No I/O, no persistence.
// ---------------------------------------------------------------------------------------------

export type ContactImportRawRow = {
  /** The spreadsheet-visible row number (header is row 1, so the first data row is 2) — shown
   * to the human for cross-referencing their source file, never persisted anywhere. */
  sourceRowNumber: number;
  displayName: string;
  email?: string;
  role?: string;
  organizationText?: string;
};

export type ContactImportParseResult =
  | { status: "success"; rows: ContactImportRawRow[] }
  | { status: "error"; message: string };

/**
 * Parses CSV text into raw evidence rows. A missing Name-equivalent column stops parsing with a
 * clear, user-facing error rather than guessing. Unknown columns are recognized and then
 * discarded — never mapped into Notes or any other field.
 */
export function parseContactImportRows(csvText: string): ContactImportParseResult {
  let parsedRows: string[][];
  try {
    parsedRows = parseCsv(csvText);
  } catch (error) {
    return { status: "error", message: error instanceof CsvParseError ? error.message : "The file could not be read as CSV." };
  }

  if (parsedRows.length === 0) return { status: "error", message: "The file is empty." };

  const [headerRow, ...dataRows] = parsedRows;
  const columnConcepts = headerRow.map(headerConceptFor);
  if (!columnConcepts.includes("name")) {
    return {
      status: "error",
      message: 'The file has no Name column. Add a "Name" or "Full Name" column and try again.',
    };
  }

  const cellFor = (cells: string[], concept: ImportHeaderConcept): string | undefined => {
    const columnIndex = columnConcepts.indexOf(concept);
    if (columnIndex === -1) return undefined;
    const value = cells[columnIndex]?.trim();
    return value ? value : undefined;
  };

  const rows: ContactImportRawRow[] = dataRows.map((cells, index) => ({
    sourceRowNumber: index + 2, // header occupies row 1
    displayName: cellFor(cells, "name") ?? "",
    email: cellFor(cells, "email"),
    role: cellFor(cells, "role"),
    organizationText: cellFor(cells, "organization"),
  }));

  return { status: "success", rows };
}

// ---------------------------------------------------------------------------------------------
// Organization resolution — exact normalized name only. Never fuzzy, never partial, never
// creates an Organization. Resolves against whatever merged (seed + durable, Patch 8E) set the
// caller currently has loaded.
// ---------------------------------------------------------------------------------------------

export function resolveOrganizationByExactName(organizationText: string | undefined, organizations: Organization[]): string | null {
  if (!organizationText || !organizationText.trim()) return null;
  const target = normalizeOrganizationName(organizationText);
  return organizations.find((organization) => normalizeOrganizationName(organization.name) === target)?.appId ?? null;
}

// ---------------------------------------------------------------------------------------------
// Transient import candidate — one row's full derived evidence, including its current Contact
// match candidates. Never added to the durable Contact model, never given ProviderMetadata,
// never written to SharePoint, and never assembled into any import-history record. A candidate
// is rebuilt fresh (via buildContactImportCandidates()) whenever the effective Contact set
// changes, so a Contact created earlier in the same review session is immediately visible to
// later rows' matching — nothing here caches a stale match list.
// ---------------------------------------------------------------------------------------------

export type ContactImportCandidate = {
  id: string;
  sourceFileName: string;
  sourceRowNumber: number;
  displayName: string;
  email?: string;
  role?: string;
  organizationText?: string;
  resolvedOrganizationId: string | null;
  matches: ContactMatchCandidate[];
};

/**
 * Pure: builds one candidate per raw row, resolving its Organization text and computing its
 * current Contact match candidates against `contacts` (the caller's current effective Contact
 * set — seed + durable, and including anything created earlier in this review session).
 * Deliberately excludes the human's review decision — that is tracked separately by the caller,
 * keyed by `id`, so recomputing candidates (e.g. after a same-session Create New) never loses or
 * has to reconcile against a previous decision.
 */
export function buildContactImportCandidates(
  rows: ContactImportRawRow[],
  sourceFileName: string,
  organizations: Organization[],
  contacts: Contact[],
): ContactImportCandidate[] {
  return rows.map((row) => {
    const resolvedOrganizationId = resolveOrganizationByExactName(row.organizationText, organizations);
    const matches = matchContactCandidates(row.displayName, contacts, {
      email: row.email,
      organizationIds: resolvedOrganizationId ? [resolvedOrganizationId] : [],
    });
    return {
      id: `${sourceFileName}:${row.sourceRowNumber}`,
      sourceFileName,
      sourceRowNumber: row.sourceRowNumber,
      displayName: row.displayName,
      email: row.email,
      role: row.role,
      organizationText: row.organizationText,
      resolvedOrganizationId,
      matches,
    };
  });
}
