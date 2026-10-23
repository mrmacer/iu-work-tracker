import type { Organization, ReferenceData } from "./models";
import { normalizeOrganizationName } from "./organization-provider";

// Patch 9B — how the SELECTABLE Organization directory and the NON-SELECTABLE compatibility
// collection are derived from the code-level seeds plus the durable (SharePoint) Organizations.
//
//   selectable = active seeds + durable Organizations   (what every picker, matcher, importer, and
//                                                       the directory screen see — unchanged API)
//   legacy     = retired placeholder seeds + any active seed SHADOWED by a durable Organization
//                with the same normalized name         (resolvable and editable-in-place only)
//
// Pure: no I/O, no AI, no SharePoint. Durable appIds and their relationships are never touched.

export type OrganizationDirectory = {
  selectable: Organization[];
  legacy: Organization[];
};

/**
 * `organizations` may already contain durable entries (a provider merges them into its own
 * reference set); `durable` is the app's own list of durable Organizations. They are de-duplicated
 * by appId (durable wins), then any seed whose normalized name equals a durable Organization's
 * moves to `legacy`: the durable record is the one the user created, so it is the one offered —
 * and the seed's appId stays resolvable (e.g. `org-iu` if a durable "Schuylkill Intermediate Unit 29"
 * exists), so nothing that links to it breaks.
 */
export function mergeOrganizationDirectory(
  organizations: Organization[],
  legacySeeds: Organization[],
  durable: Organization[],
): OrganizationDirectory {
  const byId = new Map<string, Organization>();
  for (const organization of [...organizations, ...durable]) byId.set(organization.appId, organization);
  const all = [...byId.values()];

  const durableOrganizations = all.filter((organization) => organization.metadata !== undefined);
  const durableNames = new Set(durableOrganizations.map((organization) => normalizeOrganizationName(organization.name)));
  const seeds = all.filter((organization) => organization.metadata === undefined);

  const shadowed = seeds.filter((seed) => durableNames.has(normalizeOrganizationName(seed.name)));
  const shadowedIds = new Set(shadowed.map((seed) => seed.appId));
  const legacyIds = new Set(legacySeeds.map((seed) => seed.appId));

  return {
    selectable: all.filter((organization) => !shadowedIds.has(organization.appId) && !legacyIds.has(organization.appId)),
    // A legacy id that also exists as a durable record is not "legacy" — the durable record serves it.
    legacy: [...legacySeeds.filter((seed) => !byId.get(seed.appId)?.metadata), ...shadowed],
  };
}

/** Every Organization an existing record might reference: selectable + legacy. Display/lookup only. */
export function organizationsForLookup(references: Pick<ReferenceData, "organizations" | "legacyOrganizations">): Organization[] {
  const selectable = references.organizations;
  const known = new Set(selectable.map((organization) => organization.appId));
  return [...selectable, ...(references.legacyOrganizations ?? []).filter((organization) => !known.has(organization.appId))];
}

/**
 * The legacy Organizations a given record already references — the ONLY legacy ones a picker may
 * still show (so the human can see and deselect them). A legacy Organization the record does not
 * already reference is never shown.
 */
export function legacyOrganizationsReferencedBy(
  ids: readonly string[],
  references: Pick<ReferenceData, "legacyOrganizations">,
): Organization[] {
  const referenced = new Set(ids);
  return (references.legacyOrganizations ?? []).filter((organization) => referenced.has(organization.appId));
}
