// Patch 8F — static architecture guardrails. Imports provide evidence; durable Contacts remain
// human-reviewed truth. No Contact schema change, no durable provenance, no import-history
// model, no new SharePoint list/columns, no Organization auto-create, no fuzzy matching, no AI
// calls, no new Microsoft permissions, no bulk-import actions.
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";

function source(path: string): string {
  return readFileSync(new URL(`../${path}`, import.meta.url), "utf-8");
}

describe("Contact model unchanged", () => {
  it("the Contact type block in lib/models.ts has none of the disallowed fields", () => {
    const models = source("lib/models.ts");
    const start = models.indexOf("export type Contact = {");
    const end = models.indexOf("\n};", start);
    const contactBlock = models.slice(start, end);
    for (const disallowed of ["firstName", "lastName", "phone", "organizationIds", "projectIds", "waitingOn", "importance", "tags", "address", "linkedIn", "socialProfiles"]) {
      expect(contactBlock.toLowerCase()).not.toMatch(new RegExp(disallowed.toLowerCase()));
    }
  });

  it("Contact.organizationId remains singular", () => {
    const models = source("lib/models.ts");
    expect(models).toMatch(/organizationId:\s*string \| null;/);
  });

  it("no durable provenance field (source/sourceLabel/importedFrom/importSource) exists on Contact", () => {
    const models = source("lib/models.ts");
    const start = models.indexOf("export type Contact = {");
    const end = models.indexOf("\n};", start);
    const contactBlock = models.slice(start, end);
    expect(contactBlock).not.toMatch(/source|importedFrom|importSource/i);
  });
});

describe("No import-history model or persistence", () => {
  it("lib/contact-import.ts declares no ProviderMetadata field and performs no Graph/fetch I/O", () => {
    const text = source("lib/contact-import.ts");
    expect(text).not.toMatch(/metadata\??:\s*ProviderMetadata/); // no field typed as durable provider metadata
    expect(text).not.toMatch(/from ["'].*sharepoint/i); // no import of any sharepoint-* codec module
    expect(text).not.toMatch(/fetch\(/);
  });

  it("no lib/sharepoint-contact-import*.ts or lib/import-history*.ts file exists", () => {
    for (const path of ["lib/sharepoint-contact-import.ts", "lib/import-history.ts", "lib/import-history-provider.ts"]) {
      expect(() => source(path)).toThrow();
    }
  });

  it("IU_Contacts SharePoint codec is untouched by this patch — no import-specific field mapping added to it", () => {
    const codec = source("lib/sharepoint-contacts.ts");
    expect(codec).not.toMatch(/ImportedFrom|SourceFile|SourceRow|ContactImport/i);
  });
});

describe("No Organization auto-create; exact-name resolution only", () => {
  it("resolveOrganizationByExactName never calls create/save, and no fuzzy-matching library is used", () => {
    const text = source("lib/contact-import.ts");
    expect(text).not.toMatch(/saveOrganization|createOrganization|OrganizationProvider/);
    expect(text).not.toMatch(/levenshtein|fuzzball|fuse\.js|string-similarity/i);
  });
});

describe("No fuzzy matching / second duplicate-detection algorithm", () => {
  it("lib/contact-import.ts reuses matchContactCandidates from lib/contact-matching.ts rather than reimplementing it", () => {
    const text = source("lib/contact-import.ts");
    expect(text).toMatch(/import\s*{\s*matchContactCandidates/);
  });

  it("no fuzzy-matching library appears anywhere in the import module or the CSV parser", () => {
    for (const path of ["lib/contact-import.ts", "lib/csv-parser.ts"]) {
      expect(source(path)).not.toMatch(/levenshtein|fuzzball|fuse\.js|string-similarity/i);
    }
  });
});

describe("No AI calls anywhere in the import path", () => {
  it("lib/csv-parser.ts, lib/contact-import.ts, and app/ContactImport.tsx never reference Anthropic/OpenAI", () => {
    for (const path of ["lib/csv-parser.ts", "lib/contact-import.ts", "app/ContactImport.tsx"]) {
      expect(source(path)).not.toMatch(/anthropic|openai/i);
    }
  });
});

describe("No new Microsoft permissions", () => {
  it("lib/microsoft-auth-config.ts declares no new scopes", () => {
    const config = source("lib/microsoft-auth-config.ts");
    expect(config).not.toMatch(/Sites\.Manage\.All|Sites\.FullControl\.All|People\.Read|Contacts\.Read/);
  });
});

describe("No bulk-import actions", () => {
  it("app/ContactImport.tsx contains no Save All / Create All / Import All / Accept All / Auto Match All control", () => {
    const text = source("app/ContactImport.tsx");
    expect(text).not.toMatch(/save all|create all|import all|accept all|auto match all/i);
  });
});

describe("Match Existing never updates a Contact; Create New uses the existing provider path", () => {
  it("decideMatch only writes to local decision state — no updateContact call anywhere near it", () => {
    const text = source("app/ContactImport.tsx");
    const decideMatchBlock = text.slice(text.indexOf("const decideMatch"), text.indexOf("const decideIgnore"));
    expect(decideMatchBlock).not.toMatch(/updateContact/);
  });

  it("Create New renders the existing ContactFormModal, not a second form", () => {
    const text = source("app/ContactImport.tsx");
    expect(text).toMatch(/import ContactFormModal/);
    expect(text).toMatch(/<ContactFormModal/);
  });
});

describe("No new SharePoint list/column configuration", () => {
  it("no NEXT_PUBLIC_SHAREPOINT_*IMPORT* environment variable is referenced anywhere", () => {
    for (const path of ["lib/contact-import.ts", "lib/csv-parser.ts", "app/ContactImport.tsx"]) {
      expect(source(path)).not.toMatch(/NEXT_PUBLIC_SHAREPOINT/);
    }
  });
});

describe("CSV-only, no XLSX", () => {
  it("no xlsx/spreadsheet parsing code is imported or invoked anywhere in the import path", () => {
    for (const path of ["lib/contact-import.ts", "lib/csv-parser.ts", "app/ContactImport.tsx"]) {
      expect(source(path)).not.toMatch(/from ["'].*xlsx|readXlsx|parseXlsx|SheetJS/i);
    }
  });

  it("the file input only accepts CSV", () => {
    const text = source("app/ContactImport.tsx");
    expect(text).toMatch(/accept="\.csv,text\/csv"/);
  });

  it("package.json declares no spreadsheet-parsing dependency", () => {
    const pkg = JSON.parse(source("package.json")) as { dependencies?: Record<string, string>; devDependencies?: Record<string, string> };
    const allDeps = { ...pkg.dependencies, ...pkg.devDependencies };
    for (const name of Object.keys(allDeps)) {
      expect(name.toLowerCase()).not.toMatch(/xlsx|sheetjs|papaparse|csv-parse/);
    }
  });
});
