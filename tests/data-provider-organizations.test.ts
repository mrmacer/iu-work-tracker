// Patch 8E — proves the reference-merge boundary (lib/data-provider.ts "setDurableOrganizations")
// makes a durable Organization available to every existing consumer of references.organizations
// automatically, without touching those consumers individually — the same architecture already
// proven by setDurableProjects/setDurableContacts.
import { describe, expect, it } from "vitest";
import { buildOrganizationDraft } from "../lib/organization-provider";
import { MemoryDataProvider } from "../lib/data-provider";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import { REFERENCE_DATA } from "../lib/reference-data";
import { SAMPLE_RECORDS } from "../lib/sample-data";

function newRecord(overrides: Partial<WorkRecord> = {}): WorkRecord {
  const source = structuredClone(SAMPLE_RECORDS.find((record) => !record.orbit.reportable)!);
  return {
    ...source,
    appId: `test-${crypto.randomUUID()}`,
    title: "Development test work record",
    activityDate: "2026-08-26",
    activityType: "Resource development",
    engagementScope: "none",
    organizationIds: [],
    projectIds: [],
    contactIds: [],
    categoryIds: [],
    evidenceReferenceIds: [],
    isSample: false,
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
    schemaVersion: WORK_RECORD_SCHEMA_VERSION,
    ...overrides,
  };
}

describe("static + durable Organization merge", () => {
  it("getOrganizations() returns exactly the seed set before any durable organization is added", async () => {
    const provider = new MemoryDataProvider([]);
    const organizations = await provider.getOrganizations();
    expect(organizations.map((o) => o.appId)).toEqual(REFERENCE_DATA.organizations.map((o) => o.appId));
  });

  it("getOrganizations() returns seed organizations plus the durable one after setDurableOrganizations()", async () => {
    const provider = new MemoryDataProvider([]);
    const durable = buildOrganizationDraft({ appId: crypto.randomUUID(), name: "New Regional Partner", type: "partner" });
    provider.setDurableOrganizations([durable]);
    const organizations = await provider.getOrganizations();
    expect(organizations).toHaveLength(REFERENCE_DATA.organizations.length + 1);
    expect(organizations.map((o) => o.appId)).toContain(durable.appId);
    // Every original seed organization is still present, unmutated.
    for (const seed of REFERENCE_DATA.organizations) {
      expect(organizations.find((o) => o.appId === seed.appId)).toEqual(seed);
    }
  });

  it("setDurableOrganizations() replaces the durable set on each call rather than accumulating duplicates across calls", async () => {
    const provider = new MemoryDataProvider([]);
    const first = buildOrganizationDraft({ appId: crypto.randomUUID(), name: "First Durable Org", type: "partner" });
    provider.setDurableOrganizations([first]);
    provider.setDurableOrganizations([first]); // re-sync with the same durable set, as the app does on every render
    const organizations = await provider.getOrganizations();
    expect(organizations.filter((o) => o.appId === first.appId)).toHaveLength(1);
  });
});

describe("WorkRecord.organizationIds validates a durable Organization", () => {
  it("rejects a WorkRecord referencing an organizationId that is not yet known as durable", async () => {
    const provider = new MemoryDataProvider([]);
    const durableAppId = crypto.randomUUID();
    const result = await provider.createWorkRecord(newRecord({ organizationIds: [durableAppId] }));
    expect(result.status).toBe("validation_error");
  });

  it("accepts a WorkRecord referencing a durable Organization once setDurableOrganizations() has run", async () => {
    const provider = new MemoryDataProvider([]);
    const durable = buildOrganizationDraft({ appId: crypto.randomUUID(), name: "New Regional Partner", type: "partner" });
    provider.setDurableOrganizations([durable]);
    const result = await provider.createWorkRecord(newRecord({ organizationIds: [durable.appId] }));
    expect(result.status).toBe("success");
  });

  it("a durable district participates in engagementScope validation exactly like a seed district", async () => {
    const provider = new MemoryDataProvider([]);
    const durableDistrict = buildOrganizationDraft({ appId: crypto.randomUUID(), name: "New District", type: "district" });
    provider.setDurableOrganizations([durableDistrict]);

    // "specific" scope requires at least one canonical district — a durable one counts.
    const withDurableDistrict = await provider.createWorkRecord(
      newRecord({ engagementScope: "specific", organizationIds: [durableDistrict.appId] }),
    );
    expect(withDurableDistrict.status).toBe("success");

    // A durable district selected outside "specific" scope is rejected, same as a seed district.
    const wrongScope = await provider.createWorkRecord(
      newRecord({ engagementScope: "none", organizationIds: [durableDistrict.appId] }),
    );
    expect(wrongScope.status).toBe("validation_error");
  });
});
