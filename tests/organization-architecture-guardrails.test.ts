// Patch 8E — static architecture guardrails. District remains Organization.type === "district":
// no separate District model, provider, list, or codec; no Organization.contactIds/projectIds;
// no expanded OrganizationType taxonomy beyond district/partner/iu; no second Organizations
// list environment variable; no new Microsoft permissions.
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { ORGANIZATION_TYPES } from "../lib/organization-provider";

function source(path: string): string {
  return readFileSync(new URL(`../${path}`, import.meta.url), "utf-8");
}

describe("No District model, provider, list, or codec", () => {
  it("no district-provider.ts / sharepoint-districts.ts files exist", () => {
    for (const path of ["lib/district-provider.ts", "lib/sharepoint-districts.ts"]) {
      expect(() => source(path)).toThrow();
    }
  });

  it("lib/models.ts defines no separate District type", () => {
    expect(source("lib/models.ts")).not.toMatch(/type District\b/);
  });

  it("districts derive only from Organization.type === \"district\" — never a separate district list", () => {
    const validation = source("lib/validation.ts");
    expect(validation).toMatch(/\.type === "district"/);
    const models = source("lib/inbox-intelligence-models.ts");
    expect(models).toMatch(/\.type === "district"/);
  });
});

describe("No Organization.contactIds or Organization.projectIds", () => {
  it("the Organization type block in lib/models.ts contains neither field", () => {
    const models = source("lib/models.ts");
    const start = models.indexOf("export type Organization = {");
    const end = models.indexOf("\n};", start);
    const organizationBlock = models.slice(start, end);
    expect(organizationBlock).not.toMatch(/contactIds/);
    expect(organizationBlock).not.toMatch(/projectIds/);
  });
});

describe("Contact.organizationId and WorkRecord.organizationIds shapes are unchanged", () => {
  it("Contact.organizationId remains singular (string | null), not an array", () => {
    const models = source("lib/models.ts");
    expect(models).toMatch(/organizationId:\s*string \| null;/);
    expect(models).not.toMatch(/organizationIds:\s*string\[\];\s*\n\s*role\?/); // never on Contact
  });

  it("WorkRecord.organizationIds remains string[]", () => {
    const models = source("lib/models.ts");
    const start = models.indexOf("export type WorkRecord = {");
    const end = models.indexOf("\n};", start);
    const workRecordBlock = models.slice(start, end);
    expect(workRecordBlock).toMatch(/organizationIds:\s*string\[\];/);
  });
});

describe("matchedOrganizationIds / matchedDistrictIds preserved, unmodified by Patch 8E", () => {
  it("both fields still exist on InboxIntelligenceRecord", () => {
    const models = source("lib/inbox-intelligence-models.ts");
    expect(models).toMatch(/matchedOrganizationIds:\s*string\[\];/);
    expect(models).toMatch(/matchedDistrictIds:\s*string\[\];/);
  });
});

describe("OrganizationType taxonomy is not expanded in this patch", () => {
  it("ORGANIZATION_TYPES is exactly district, partner, iu — no manufacturer/nonprofit/business/community/other", () => {
    expect(ORGANIZATION_TYPES).toEqual(["district", "partner", "iu"]);
  });
});

describe("No second Organizations list environment variable", () => {
  it("only NEXT_PUBLIC_SHAREPOINT_IU_ORGANIZATIONS_LIST_ID is referenced by the provider", () => {
    const provider = source("lib/organization-provider.ts");
    const envVarMatches = provider.match(/NEXT_PUBLIC_SHAREPOINT_[A-Z_]*ORGANIZATION[A-Z_]*/g) ?? [];
    expect(new Set(envVarMatches)).toEqual(new Set(["NEXT_PUBLIC_SHAREPOINT_IU_ORGANIZATIONS_LIST_ID"]));
  });
});

describe("No AI calls, no auto-creation from Intelligence, no fuzzy matching in the Organization domain", () => {
  it("lib/organization-provider.ts and lib/sharepoint-organizations.ts make no AI/network calls of their own beyond Graph fetch", () => {
    for (const path of ["lib/organization-provider.ts", "lib/sharepoint-organizations.ts"]) {
      const text = source(path);
      expect(text).not.toMatch(/anthropic/i);
      expect(text).not.toMatch(/levenshtein|fuzzball|fuse\.js|string-similarity/i);
    }
  });

  it("Inbox/Voice Intelligence never call an Organization create/save function — matching stays read-only against references.organizations", () => {
    const inbox = source("app/InboxIntelligence.tsx");
    const voice = source("app/VoiceIntelligence.tsx");
    expect(inbox).not.toMatch(/saveOrganization|updateOrganization/);
    expect(voice).not.toMatch(/saveOrganization|updateOrganization/);
  });
});

describe("No Microsoft permission changes", () => {
  it("lib/microsoft-auth-config.ts declares no new scopes beyond the existing set", () => {
    const config = source("lib/microsoft-auth-config.ts");
    expect(config).not.toMatch(/Sites\.Manage\.All|Sites\.FullControl\.All|People\.Read|Contacts\.Read/);
  });
});
