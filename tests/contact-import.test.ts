import { describe, expect, it } from "vitest";
import {
  buildContactImportCandidates,
  parseContactImportRows,
  resolveOrganizationByExactName,
  type ContactImportRawRow,
} from "../lib/contact-import";
import type { Contact, Organization } from "../lib/models";

function organization(overrides: Partial<Organization> = {}): Organization {
  return { appId: "org-north-valley", name: "North Valley SD", type: "district", ...overrides };
}

function contact(overrides: Partial<Contact> = {}): Contact {
  return { appId: "contact-annie", displayName: "Annie Milewski", organizationId: "org-north-valley", status: "active", ...overrides };
}

describe("parseContactImportRows — header aliases", () => {
  it.each(["Name", "Full Name", "name", "  full   name  "])("recognizes %s as the Name column", (header) => {
    const result = parseContactImportRows(`${header}\nJordan Example`);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.rows[0].displayName).toBe("Jordan Example");
  });

  it.each(["Email", "Email Address"])("recognizes %s as the Email column", (header) => {
    const result = parseContactImportRows(`Name,${header}\nJordan,jordan@example.test`);
    if (result.status !== "success") throw new Error("expected success");
    expect(result.rows[0].email).toBe("jordan@example.test");
  });

  it.each(["Role", "Title", "Job Title"])("recognizes %s as the Role column", (header) => {
    const result = parseContactImportRows(`Name,${header}\nJordan,Coordinator`);
    if (result.status !== "success") throw new Error("expected success");
    expect(result.rows[0].role).toBe("Coordinator");
  });

  it.each(["Organization", "Organisation", "District", "Company"])("recognizes %s as the Organization column", (header) => {
    const result = parseContactImportRows(`Name,${header}\nJordan,North Valley SD`);
    if (result.status !== "success") throw new Error("expected success");
    expect(result.rows[0].organizationText).toBe("North Valley SD");
  });

  it("stops with a clear error when no Name-equivalent column exists", () => {
    const result = parseContactImportRows("Email,Role\njordan@example.test,Coordinator");
    expect(result.status).toBe("error");
    if (result.status !== "error") return;
    expect(result.message).toMatch(/Name column/);
  });

  it("ignores unknown columns entirely — never mapped into any field", () => {
    const result = parseContactImportRows("Name,Phone,Notes\nJordan,555-0100,Met at conference");
    if (result.status !== "success") throw new Error("expected success");
    expect(result.rows[0]).toEqual({ sourceRowNumber: 2, displayName: "Jordan", email: undefined, role: undefined, organizationText: undefined });
  });

  it("treats an empty file as an error", () => {
    expect(parseContactImportRows("").status).toBe("error");
  });

  it("surfaces a CSV parse error (e.g. an unterminated quote) as a clear message", () => {
    const result = parseContactImportRows('Name\n"Jordan');
    expect(result.status).toBe("error");
    if (result.status !== "error") return;
    expect(result.message).toMatch(/never closed/);
  });
});

describe("parseContactImportRows — row shape", () => {
  it("numbers rows as spreadsheet-visible row numbers (header is row 1)", () => {
    const result = parseContactImportRows("Name\nJordan\nTaylor\nMorgan");
    if (result.status !== "success") throw new Error("expected success");
    expect(result.rows.map((r) => r.sourceRowNumber)).toEqual([2, 3, 4]);
  });

  it("treats an empty optional cell as undefined, not an empty string", () => {
    const result = parseContactImportRows("Name,Email,Role\nMorgan,,Partner Lead");
    if (result.status !== "success") throw new Error("expected success");
    expect(result.rows[0].email).toBeUndefined();
    expect(result.rows[0].role).toBe("Partner Lead");
  });

  it("preserves every source row independently, including exact duplicates", () => {
    const csv = "Name,Email\nJordan Example,jordan@example.test\nJordan Example,jordan@example.test";
    const result = parseContactImportRows(csv);
    if (result.status !== "success") throw new Error("expected success");
    expect(result.rows).toHaveLength(2);
    expect(result.rows[0].sourceRowNumber).toBe(2);
    expect(result.rows[1].sourceRowNumber).toBe(3);
  });
});

describe("resolveOrganizationByExactName", () => {
  const organizations = [organization(), organization({ appId: "org-futureworks", name: "FutureWorks Partnership", type: "partner" })];

  it("resolves an exact (normalized) name match", () => {
    expect(resolveOrganizationByExactName("North Valley SD", organizations)).toBe("org-north-valley");
    expect(resolveOrganizationByExactName("  north   valley sd ", organizations)).toBe("org-north-valley");
  });

  it("returns null for an unresolved Organization — never fabricates or partially matches", () => {
    expect(resolveOrganizationByExactName("Unknown Organization", organizations)).toBeNull();
    expect(resolveOrganizationByExactName("North Valley", organizations)).toBeNull(); // partial match, deliberately not resolved
  });

  it("returns null for an absent or empty organization text", () => {
    expect(resolveOrganizationByExactName(undefined, organizations)).toBeNull();
    expect(resolveOrganizationByExactName("   ", organizations)).toBeNull();
  });
});

describe("buildContactImportCandidates — Organization resolution", () => {
  it("attaches the resolved Organization appId to the candidate", () => {
    const rows: ContactImportRawRow[] = [{ sourceRowNumber: 2, displayName: "Taylor Example", organizationText: "North Valley SD" }];
    const candidates = buildContactImportCandidates(rows, "roster.csv", [organization()], []);
    expect(candidates[0].resolvedOrganizationId).toBe("org-north-valley");
  });

  it("leaves resolvedOrganizationId null when the source text doesn't match any Organization", () => {
    const rows: ContactImportRawRow[] = [{ sourceRowNumber: 2, displayName: "Morgan Example", organizationText: "Unknown Organization" }];
    const candidates = buildContactImportCandidates(rows, "roster.csv", [organization()], []);
    expect(candidates[0].resolvedOrganizationId).toBeNull();
  });
});

describe("buildContactImportCandidates — Contact matching (reuses lib/contact-matching.ts)", () => {
  it("an exact email match is a strong candidate", () => {
    const rows: ContactImportRawRow[] = [{ sourceRowNumber: 2, displayName: "Annie M.", email: "annie@northvalley.org" }];
    const c = contact({ email: "annie@northvalley.org" });
    const candidates = buildContactImportCandidates(rows, "roster.csv", [], [c]);
    expect(candidates[0].matches).toEqual([{ contact: c, strength: "strong", reasons: ["Exact email match"] }]);
  });

  it("full name + resolved Organization is a useful candidate", () => {
    const rows: ContactImportRawRow[] = [{ sourceRowNumber: 2, displayName: "Annie Milewski", organizationText: "North Valley SD" }];
    const candidates = buildContactImportCandidates(rows, "roster.csv", [organization()], [contact()]);
    expect(candidates[0].matches[0].strength).toBe("useful");
  });

  it("full name only (no organization evidence) is a review candidate", () => {
    const rows: ContactImportRawRow[] = [{ sourceRowNumber: 2, displayName: "Annie Milewski" }];
    const candidates = buildContactImportCandidates(rows, "roster.csv", [], [contact()]);
    expect(candidates[0].matches[0].strength).toBe("review");
  });

  it("first-name-only evidence produces no candidates at all", () => {
    const rows: ContactImportRawRow[] = [{ sourceRowNumber: 2, displayName: "Annie" }];
    const candidates = buildContactImportCandidates(rows, "roster.csv", [], [contact()]);
    expect(candidates[0].matches).toEqual([]);
  });

  it("retains every plausible candidate when multiple Contacts share the same name", () => {
    const first = contact({ appId: "contact-annie-1" });
    const second = contact({ appId: "contact-annie-2", organizationId: "org-other" });
    const rows: ContactImportRawRow[] = [{ sourceRowNumber: 2, displayName: "Annie Milewski" }];
    const candidates = buildContactImportCandidates(rows, "roster.csv", [], [first, second]);
    expect(candidates[0].matches.map((m) => m.contact.appId).sort()).toEqual(["contact-annie-1", "contact-annie-2"]);
  });
});

describe("Transient provenance", () => {
  it("sourceFileName and sourceRowNumber live only on the transient candidate — never assembled into any durable-looking record", () => {
    const rows: ContactImportRawRow[] = [{ sourceRowNumber: 5, displayName: "Jordan Example" }];
    const candidates = buildContactImportCandidates(rows, "district-roster.csv", [], []);
    expect(candidates[0].sourceFileName).toBe("district-roster.csv");
    expect(candidates[0].sourceRowNumber).toBe(5);
    expect(candidates[0].id).toBe("district-roster.csv:5");
    // The candidate type carries no ProviderMetadata-shaped field.
    expect(candidates[0]).not.toHaveProperty("metadata");
  });

  it("recomputes matches fresh against whatever Contact set is passed in — a Contact created earlier in the same review session becomes visible to a later call", () => {
    const rows: ContactImportRawRow[] = [{ sourceRowNumber: 3, displayName: "Jordan Example" }];
    const beforeCreate = buildContactImportCandidates(rows, "roster.csv", [], []);
    expect(beforeCreate[0].matches).toEqual([]);

    const createdThisSession = contact({ appId: "contact-jordan-new", displayName: "Jordan Example" });
    const afterCreate = buildContactImportCandidates(rows, "roster.csv", [], [createdThisSession]);
    expect(afterCreate[0].matches.map((m) => m.contact.appId)).toEqual(["contact-jordan-new"]);
  });
});
