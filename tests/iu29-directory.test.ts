// Patch 9B — the real IU29 Organization directory, the non-selectable legacy compatibility
// collection, and the rule that keeps historical records readable and editable without ever
// letting a NEW record select a retired placeholder. Pure data/validation tests plus both
// providers (in-memory, and the SharePoint codec path with a mocked Graph). No real network.
import { afterEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { DelegatedSharePointDataProvider, MemoryDataProvider } from "../lib/data-provider";
import type { MicrosoftAuthController } from "../lib/microsoft-auth";
import { WORK_RECORD_SCHEMA_VERSION, type Organization, type WorkRecord } from "../lib/models";
import {
  legacyOrganizationsReferencedBy,
  mergeOrganizationDirectory,
  organizationsForLookup,
} from "../lib/organization-directory";
import { normalizeOrganizationName } from "../lib/organization-provider";
import { CONTACTS, LEGACY_ORGANIZATIONS, ORGANIZATIONS, REFERENCE_DATA } from "../lib/reference-data";
import { SAMPLE_RECORDS } from "../lib/sample-data";
import { toSharePointFields } from "../lib/sharepoint-work-records";
import { validateWorkRecord } from "../lib/validation";

const OFFICIAL_DISTRICTS = [
  "Blue Mountain School District",
  "Mahanoy Area School District",
  "Minersville Area School District",
  "North Schuylkill School District",
  "Pine Grove Area School District",
  "Pottsville Area School District",
  "Saint Clair Area School District",
  "Schuylkill Haven Area School District",
  "Shenandoah Valley School District",
  "Tamaqua Area School District",
  "Tri-Valley School District",
  "Williams Valley School District",
];

const durable = (name: string, type: Organization["type"] = "district", appId = `durable-${name}`): Organization => ({
  appId,
  name,
  type,
  metadata: { providerId: "1", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" },
});

function record(overrides: Partial<WorkRecord> = {}): WorkRecord {
  return {
    ...structuredClone(SAMPLE_RECORDS.find((item) => !item.orbit.reportable && item.engagementScope === "none")!),
    appId: "record-under-test",
    title: "A saved entry",
    organizationIds: [],
    projectIds: [],
    contactIds: [],
    categoryIds: [],
    evidenceReferenceIds: [],
    engagementScope: "none",
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
    schemaVersion: WORK_RECORD_SCHEMA_VERSION,
    isSample: false,
    ...overrides,
  };
}

describe("the seeded IU29 directory", () => {
  it("contains exactly the twelve official districts, as type district, with the approved full names", () => {
    const districts = ORGANIZATIONS.filter((organization) => organization.type === "district");
    expect(districts.map((organization) => organization.name)).toEqual(OFFICIAL_DISTRICTS);
  });

  it("spells Saint Clair as the district's own site does — never 'St. Clair'", () => {
    expect(ORGANIZATIONS.some((organization) => organization.name === "Saint Clair Area School District")).toBe(true);
    expect(ORGANIZATIONS.some((organization) => /^St\.? Clair/i.test(organization.name))).toBe(false);
  });

  it("keeps org-iu as the IU, with its corrected official name", () => {
    expect(ORGANIZATIONS.find((organization) => organization.appId === "org-iu")).toEqual({
      appId: "org-iu",
      name: "Schuylkill Intermediate Unit 29",
      type: "iu",
    });
  });

  it("adds PAIMS as a PARTNER with the stable id org-paims — never a district", () => {
    expect(ORGANIZATIONS.find((organization) => organization.appId === "org-paims")).toEqual({ appId: "org-paims", name: "PAIMS", type: "partner" });
  });

  it("is fourteen seeds: twelve districts + the IU + PAIMS, with no separate District model", () => {
    expect(ORGANIZATIONS).toHaveLength(14);
    expect(ORGANIZATIONS.filter((organization) => organization.type === "iu")).toHaveLength(1);
    expect(ORGANIZATIONS.filter((organization) => organization.type === "partner")).toHaveLength(1);
    expect(Object.keys(REFERENCE_DATA).some((key) => /district/i.test(key))).toBe(false);
  });

  it("has unique appIds and unique normalized names across active and legacy seeds", () => {
    const all = [...ORGANIZATIONS, ...LEGACY_ORGANIZATIONS];
    expect(new Set(all.map((organization) => organization.appId)).size).toBe(all.length);
    expect(new Set(all.map((organization) => normalizeOrganizationName(organization.name))).size).toBe(all.length);
  });

  it("contains none of the fictional placeholders as active entries", () => {
    for (const fictional of ["North Valley SD", "Riverbend Area SD", "FutureWorks Partnership", "Intermediate Unit"]) {
      expect(ORGANIZATIONS.some((organization) => organization.name === fictional)).toBe(false);
    }
  });

  it("preserves the three retired placeholder appIds, with their original names, in the legacy collection only", () => {
    expect(LEGACY_ORGANIZATIONS).toEqual([
      { appId: "org-north-valley", name: "North Valley SD", type: "district" },
      { appId: "org-riverbend", name: "Riverbend Area SD", type: "district" },
      { appId: "org-futureworks", name: "FutureWorks Partnership", type: "partner" },
    ]);
    expect(REFERENCE_DATA.organizations.some((organization) => LEGACY_ORGANIZATIONS.some((legacy) => legacy.appId === organization.appId))).toBe(false);
    expect(REFERENCE_DATA.legacyOrganizations).toBe(LEGACY_ORGANIZATIONS);
  });

  it("keeps seed Contacts and sample Work Records on the clearly fictional placeholders — never an arbitrary real district or PAIMS", () => {
    const active = new Set(ORGANIZATIONS.map((organization) => organization.appId));
    const legacy = new Set(LEGACY_ORGANIZATIONS.map((organization) => organization.appId));
    const referenced = [
      ...CONTACTS.flatMap((contact) => (contact.organizationId ? [contact.organizationId] : [])),
      ...SAMPLE_RECORDS.flatMap((sample) => sample.organizationIds),
    ];
    // Only the IU itself (a real, non-district organization) or a retired fictional placeholder.
    for (const id of referenced) expect(id === "org-iu" || legacy.has(id), id).toBe(true);
    for (const id of ["org-north-valley", "org-riverbend", "org-futureworks"]) expect(active.has(id)).toBe(false);
    expect(CONTACTS.find((contact) => contact.appId === "contact-north-valley-lead")?.organizationId).toBe("org-north-valley");
    expect(CONTACTS.find((contact) => contact.appId === "contact-futureworks")?.organizationId).toBe("org-futureworks");
  });

  it("sample Work Records that carry a placeholder are valid only as EXISTING records (retained), and invalid as new ones", () => {
    for (const sample of SAMPLE_RECORDS) {
      const legacyIds = sample.organizationIds.filter((id) => LEGACY_ORGANIZATIONS.some((organization) => organization.appId === id));
      expect(validateWorkRecord(sample, REFERENCE_DATA, { retainedLegacyOrganizationIds: legacyIds }).valid, sample.appId).toBe(true);
      if (legacyIds.length) expect(validateWorkRecord(sample, REFERENCE_DATA).valid, sample.appId).toBe(false);
    }
  });
});

describe("mergeOrganizationDirectory", () => {
  it("with no durable Organizations: selectable = the active seeds, legacy = the retired placeholders", () => {
    const directory = mergeOrganizationDirectory(ORGANIZATIONS, LEGACY_ORGANIZATIONS, []);
    expect(directory.selectable).toEqual(ORGANIZATIONS);
    expect(directory.legacy).toEqual(LEGACY_ORGANIZATIONS);
  });

  it("a durable Organization is added to the selectable set without touching any seed", () => {
    const gillingham = durable("Gillingham Charter School", "partner");
    const directory = mergeOrganizationDirectory(ORGANIZATIONS, LEGACY_ORGANIZATIONS, [gillingham]);
    expect(directory.selectable).toHaveLength(15);
    expect(directory.selectable.at(-1)).toBe(gillingham);
    expect(directory.legacy).toEqual(LEGACY_ORGANIZATIONS);
  });

  it("a durable Organization with the same NORMALIZED name as a seed replaces it — one picker entry, no duplicate", () => {
    const sameName = durable("  blue   MOUNTAIN school district ", "district", "durable-uuid-1");
    const directory = mergeOrganizationDirectory(ORGANIZATIONS, LEGACY_ORGANIZATIONS, [sameName]);
    const matches = directory.selectable.filter((organization) => normalizeOrganizationName(organization.name) === "blue mountain school district");
    expect(matches).toEqual([sameName]);
    expect(directory.selectable).toHaveLength(14);
  });

  it("the replaced seed's appId stays RESOLVABLE (legacy) — nothing that links to it breaks — and the durable appId is never altered", () => {
    const sameName = durable("Schuylkill Intermediate Unit 29", "iu", "durable-uuid-iu");
    const directory = mergeOrganizationDirectory(ORGANIZATIONS, LEGACY_ORGANIZATIONS, [sameName]);
    expect(directory.selectable.find((organization) => organization.appId === "durable-uuid-iu")).toBe(sameName);
    expect(directory.selectable.some((organization) => organization.appId === "org-iu")).toBe(false);
    expect(directory.legacy.map((organization) => organization.appId)).toContain("org-iu");
  });

  it("de-duplicates by appId when a provider has ALREADY merged durable records into the list it passes in", () => {
    const gillingham = durable("Gillingham Charter School", "partner");
    const directory = mergeOrganizationDirectory([...ORGANIZATIONS, gillingham], LEGACY_ORGANIZATIONS, [gillingham]);
    expect(directory.selectable.filter((organization) => organization.appId === gillingham.appId)).toHaveLength(1);
  });

  it("does not mutate its inputs", () => {
    const before = JSON.stringify([ORGANIZATIONS, LEGACY_ORGANIZATIONS]);
    mergeOrganizationDirectory(ORGANIZATIONS, LEGACY_ORGANIZATIONS, [durable("Blue Mountain School District")]);
    expect(JSON.stringify([ORGANIZATIONS, LEGACY_ORGANIZATIONS])).toBe(before);
  });

  it("organizationsForLookup resolves selectable AND legacy ids; legacyOrganizationsReferencedBy exposes only the ones a record already has", () => {
    const references = { organizations: ORGANIZATIONS, legacyOrganizations: LEGACY_ORGANIZATIONS };
    const lookup = organizationsForLookup(references);
    expect(lookup).toHaveLength(17);
    expect(lookup.find((organization) => organization.appId === "org-north-valley")?.name).toBe("North Valley SD");
    expect(legacyOrganizationsReferencedBy(["org-riverbend", "org-blue-mountain"], references).map((organization) => organization.appId)).toEqual(["org-riverbend"]);
    expect(legacyOrganizationsReferencedBy([], references)).toEqual([]);
    expect(organizationsForLookup({ organizations: ORGANIZATIONS })).toEqual(ORGANIZATIONS); // legacy is optional
  });
});

describe("Work Record validation and retired placeholders", () => {
  it("a NEW record can select a real district", () => {
    expect(validateWorkRecord(record({ engagementScope: "specific", organizationIds: ["org-pottsville-area"] }), REFERENCE_DATA).valid).toBe(true);
  });

  it.each(["org-north-valley", "org-riverbend", "org-futureworks"])("a NEW record cannot select retired %s", (id) => {
    const result = validateWorkRecord(record({ engagementScope: id === "org-futureworks" ? "none" : "specific", organizationIds: [id] }), REFERENCE_DATA);
    expect(result.valid).toBe(false);
    if (!result.valid) expect(result.issues.map((issue) => issue.code)).toContain("retired_organization");
  });

  it("an EXISTING record keeps the retired placeholders it already has (explicitly retained), including a retired district for specific scope", () => {
    const stored = record({ engagementScope: "specific", organizationIds: ["org-north-valley", "org-futureworks"] });
    expect(validateWorkRecord(stored, REFERENCE_DATA, { retainedLegacyOrganizationIds: ["org-north-valley", "org-futureworks"] }).valid).toBe(true);
  });

  it("retaining one legacy id does not make another one valid", () => {
    const edited = record({ engagementScope: "specific", organizationIds: ["org-north-valley", "org-riverbend"] });
    const result = validateWorkRecord(edited, REFERENCE_DATA, { retainedLegacyOrganizationIds: ["org-north-valley"] });
    expect(result.valid).toBe(false);
    if (!result.valid) expect(result.issues.filter((issue) => issue.code === "retired_organization")).toHaveLength(1);
  });

  it("an unknown organization id is still rejected as not canonical", () => {
    const result = validateWorkRecord(record({ organizationIds: ["org-made-up"] }), REFERENCE_DATA);
    expect(result.valid).toBe(false);
    if (!result.valid) expect(result.issues.map((issue) => issue.code)).toContain("invalid_id");
  });

  it("the scope/district rules still hold for real districts (district ids only with specific scope)", () => {
    expect(validateWorkRecord(record({ engagementScope: "regional", organizationIds: ["org-mahanoy-area"] }), REFERENCE_DATA).valid).toBe(false);
    expect(validateWorkRecord(record({ engagementScope: "specific" }), REFERENCE_DATA).valid).toBe(false);
  });

  it("PAIMS and the IU are non-district Organizations: valid with no scope, rejected as 'specific' districts", () => {
    expect(validateWorkRecord(record({ organizationIds: ["org-paims", "org-iu"] }), REFERENCE_DATA).valid).toBe(true);
    expect(validateWorkRecord(record({ engagementScope: "specific", organizationIds: ["org-paims"] }), REFERENCE_DATA).valid).toBe(false);
  });
});

describe("MemoryDataProvider — existing records stay editable, new ones cannot use retired placeholders", () => {
  const stored = (): WorkRecord =>
    record({
      appId: "historic",
      engagementScope: "specific",
      organizationIds: ["org-north-valley", "org-futureworks"],
      metadata: { providerId: "memory:1", version: 1, createdAt: "a", modifiedAt: "a", syncState: "saved" },
    });

  it("creating a NEW record with a retired placeholder is rejected", async () => {
    const provider = new MemoryDataProvider([]);
    const result = await provider.createWorkRecord(record({ engagementScope: "specific", organizationIds: ["org-north-valley"] }));
    expect(result.status).toBe("validation_error");
  });

  it("an existing record that has retired placeholders can be edited and saved (title change)", async () => {
    const provider = new MemoryDataProvider([stored()]);
    const result = await provider.updateWorkRecord({ ...stored(), title: "Edited title" }, 1);
    expect(result.status).toBe("success");
    if (result.status === "success") {
      expect(result.value.title).toBe("Edited title");
      expect(result.value.organizationIds).toEqual(["org-north-valley", "org-futureworks"]); // never silently remapped
    }
  });

  it("an existing record can DROP a retired placeholder and replace it with a real district", async () => {
    const provider = new MemoryDataProvider([stored()]);
    const result = await provider.updateWorkRecord({ ...stored(), organizationIds: ["org-blue-mountain", "org-futureworks"] }, 1);
    expect(result.status).toBe("success");
  });

  it("an existing record can NOT gain a retired placeholder it did not already have", async () => {
    const withoutLegacy = record({ appId: "historic", engagementScope: "specific", organizationIds: ["org-blue-mountain"], metadata: { providerId: "memory:1", version: 1, createdAt: "a", modifiedAt: "a", syncState: "saved" } });
    const provider = new MemoryDataProvider([withoutLegacy]);
    const result = await provider.updateWorkRecord({ ...withoutLegacy, organizationIds: ["org-blue-mountain", "org-riverbend"] }, 1);
    expect(result.status).toBe("validation_error");
    if (result.status === "validation_error") expect(result.errors.map((issue) => issue.code)).toContain("retired_organization");
  });

  it("when a durable Organization replaces a seed, an existing record linked to the seed stays editable", async () => {
    const linked = record({ appId: "historic", organizationIds: ["org-iu"], metadata: { providerId: "memory:1", version: 1, createdAt: "a", modifiedAt: "a", syncState: "saved" } });
    const provider = new MemoryDataProvider([linked]);
    provider.setDurableOrganizations([durable("Schuylkill Intermediate Unit 29", "iu", "durable-uuid-iu")]);
    expect((await provider.updateWorkRecord({ ...linked, title: "Still editable" }, 1)).status).toBe("success");
    // …while a brand-new record must use the durable one, not the replaced seed id.
    expect((await provider.createWorkRecord(record({ appId: "fresh", organizationIds: ["org-iu"] }))).status).toBe("validation_error");
    expect((await provider.createWorkRecord(record({ appId: "fresh2", organizationIds: ["durable-uuid-iu"] }))).status).toBe("success");
  });

  it("exposes only selectable Organizations through getOrganizations (so no consumer offers a retired one)", async () => {
    const provider = new MemoryDataProvider([]);
    provider.setDurableOrganizations([durable("Gillingham Charter School", "partner")]);
    const organizations = await provider.getOrganizations();
    expect(organizations).toHaveLength(15);
    expect(organizations.some((organization) => LEGACY_ORGANIZATIONS.some((legacy) => legacy.appId === organization.appId))).toBe(false);
  });
});

describe("DelegatedSharePointDataProvider — the stored record decides which retired ids are retained", () => {
  const account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
  const config = { siteId: "site-id", workRecordsListId: "list-id" };
  const controller = { acquireGraphToken: async () => "token" } as unknown as MicrosoftAuthController;
  const json = (value: unknown) => new Response(JSON.stringify(value), { status: 200, headers: { "Content-Type": "application/json" } });

  function graphItem(source: WorkRecord, version: number) {
    return { id: "10", eTag: '"1"', createdDateTime: "2026-08-28T12:00:00Z", lastModifiedDateTime: "2026-08-28T12:00:00Z", fields: toSharePointFields(source, version) };
  }

  const base = (organizationIds: string[]): WorkRecord =>
    record({ appId: "historic", engagementScope: "specific", organizationIds, metadata: { providerId: "10", version: 1, createdAt: "x", modifiedAt: "x", syncState: "saved" } });

  afterEach(() => vi.unstubAllGlobals());

  /** GET item → the stored version; PATCH → ok; the read-back GET → version 2. */
  function stubGraph(storedRecord: WorkRecord) {
    let reads = 0;
    const fetcher = vi.fn(async (_url: string, init?: RequestInit) => {
      if (init?.method === "PATCH") return json({});
      reads += 1;
      return json(graphItem(storedRecord, reads === 1 ? 1 : 2));
    });
    vi.stubGlobal("fetch", fetcher);
    return fetcher;
  }

  it("a stored record WITH retired placeholders can be updated", async () => {
    const storedRecord = base(["org-north-valley"]);
    const fetcher = stubGraph(storedRecord);
    const provider = new DelegatedSharePointDataProvider(controller, account, config);
    const result = await provider.updateWorkRecord({ ...storedRecord, title: "Edited" }, 1);
    expect(result.status).toBe("success");
    expect(fetcher.mock.calls.some(([, init]) => init?.method === "PATCH")).toBe(true);
  });

  it("an update that ADDS a retired placeholder the stored record lacks is rejected before any write", async () => {
    const storedRecord = base(["org-blue-mountain"]);
    const fetcher = stubGraph(storedRecord);
    const provider = new DelegatedSharePointDataProvider(controller, account, config);
    const result = await provider.updateWorkRecord({ ...storedRecord, organizationIds: ["org-blue-mountain", "org-riverbend"] }, 1);
    expect(result.status).toBe("validation_error");
    expect(fetcher.mock.calls.some(([, init]) => init?.method === "PATCH")).toBe(false);
  });

  it("an ordinary update with no retired placeholders never needs the extra check (unchanged behavior)", async () => {
    const storedRecord = base(["org-blue-mountain"]);
    stubGraph(storedRecord);
    const provider = new DelegatedSharePointDataProvider(controller, account, config);
    expect((await provider.updateWorkRecord({ ...storedRecord, title: "Plain edit" }, 1)).status).toBe("success");
  });

  it("creating with a retired placeholder is rejected before ever acquiring a token", async () => {
    const acquire = vi.fn(async () => "token");
    const provider = new DelegatedSharePointDataProvider({ acquireGraphToken: acquire } as unknown as MicrosoftAuthController, account, config);
    const result = await provider.createWorkRecord(record({ engagementScope: "specific", organizationIds: ["org-north-valley"] }));
    expect(result.status).toBe("validation_error");
    expect(acquire).not.toHaveBeenCalled();
  });
});

describe("no schema, permission, or SharePoint-seeding change", () => {
  it("the seeds are code-level only: no SharePoint codec or provider references the seed directory", async () => {
    const { readFileSync } = await import("node:fs");
    for (const path of ["lib/sharepoint-organizations.ts", "lib/organization-provider.ts"]) {
      expect(readFileSync(path, "utf-8"), path).not.toMatch(/from ["'][^"']*reference-data|\bLEGACY_ORGANIZATIONS\b|org-paims|org-blue-mountain/);
    }
  });
});
