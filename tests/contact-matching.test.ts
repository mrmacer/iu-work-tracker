import { describe, expect, it } from "vitest";
import { matchContactCandidates } from "../lib/contact-matching";
import type { Contact } from "../lib/models";

// Patch 8D — deterministic Contact identity resolution. Every test here proves: zero AI, zero
// network, never a silent commit, never a first-name-only match, and archived Contacts stay
// matchable (status is a relationship concern, not an identity-eligibility filter).

function contact(overrides: Partial<Contact> = {}): Contact {
  return {
    appId: "contact-annie",
    displayName: "Annie Milewski",
    organizationId: "org-north-valley",
    status: "active",
    ...overrides,
  };
}

describe("matchContactCandidates — email (strong)", () => {
  it("returns a strong match on an exact normalized email, independent of name shape", () => {
    const c = contact({ email: "Annie@NorthValley.org" });
    const result = matchContactCandidates("Annie", [c], { email: "annie@northvalley.org" });
    expect(result).toEqual([{ contact: c, strength: "strong", reasons: ["Exact email match"] }]);
  });

  it("trims and lowercases both sides before comparing", () => {
    const c = contact({ email: "  Annie@NorthValley.org  " });
    const result = matchContactCandidates("Annie Milewski", [c], { email: " ANNIE@northvalley.ORG " });
    expect(result[0].strength).toBe("strong");
  });

  it("an email evidence miss falls through to name-based matching instead of returning empty", () => {
    const c = contact({ email: "annie@northvalley.org" });
    const result = matchContactCandidates("Annie Milewski", [c], { email: "someone-else@example.org", organizationIds: [] });
    expect(result).toEqual([{ contact: c, strength: "review", reasons: ["Name match only — organization not confirmed"] }]);
  });
});

describe("matchContactCandidates — full name + organization (useful)", () => {
  it("elevates a full-name match to useful when the Contact's organization is in the provided context", () => {
    const c = contact({ organizationId: "org-north-valley" });
    const result = matchContactCandidates("Annie Milewski", [c], { organizationIds: ["org-north-valley", "org-other"] });
    expect(result).toEqual([{ contact: c, strength: "useful", reasons: ["Name and organization match"] }]);
  });
});

describe("matchContactCandidates — full name only (review)", () => {
  it("a full-name match with no organization evidence is review-tier, never auto-elevated", () => {
    const c = contact();
    const result = matchContactCandidates("Annie Milewski", [c]);
    expect(result).toEqual([{ contact: c, strength: "review", reasons: ["Name match only — organization not confirmed"] }]);
  });

  it("does not elevate to useful when the Contact's organization is absent from the context", () => {
    const c = contact({ organizationId: "org-north-valley" });
    const result = matchContactCandidates("Annie Milewski", [c], { organizationIds: ["org-some-other-district"] });
    expect(result[0].strength).toBe("review");
  });

  it("a Contact with no organizationId at all is never elevated to useful", () => {
    const c = contact({ organizationId: null });
    const result = matchContactCandidates("Annie Milewski", [c], { organizationIds: ["org-north-valley"] });
    expect(result[0].strength).toBe("review");
  });
});

describe("matchContactCandidates — too weak (none)", () => {
  it("a first-name-only detection never matches, no matter how few Contacts share it", () => {
    const c = contact();
    expect(matchContactCandidates("Annie", [c])).toEqual([]);
  });

  it("an empty or whitespace-only detected name never matches", () => {
    const c = contact();
    expect(matchContactCandidates("", [c])).toEqual([]);
    expect(matchContactCandidates("   ", [c])).toEqual([]);
  });

  it("a full name with zero Contact matches returns an empty array, not a fabricated candidate", () => {
    const c = contact({ displayName: "Kim Rivera" });
    expect(matchContactCandidates("Annie Milewski", [c])).toEqual([]);
  });

  it("an empty Contacts list always returns no candidates", () => {
    expect(matchContactCandidates("Annie Milewski", [])).toEqual([]);
  });
});

describe("matchContactCandidates — normalization", () => {
  it("is case-insensitive on the detected name", () => {
    const c = contact({ displayName: "Annie Milewski" });
    expect(matchContactCandidates("ANNIE milewski", [c])).toHaveLength(1);
  });

  it("collapses irregular internal whitespace in the detected name before comparing", () => {
    const c = contact({ displayName: "Annie Milewski" });
    expect(matchContactCandidates("Annie   Milewski", [c])).toHaveLength(1);
    expect(matchContactCandidates("  Annie Milewski  ", [c])).toHaveLength(1);
  });
});

describe("matchContactCandidates — multiple matches", () => {
  it("returns every same-name Contact as a separate candidate rather than picking one by array order", () => {
    const first = contact({ appId: "contact-annie-1", organizationId: "org-north-valley" });
    const second = contact({ appId: "contact-annie-2", organizationId: "org-schuylkill-haven" });
    const result = matchContactCandidates("Annie Milewski", [first, second]);
    expect(result.map((r) => r.contact.appId).sort()).toEqual(["contact-annie-1", "contact-annie-2"]);
  });

  it("each candidate in a multi-match result is independently strength-classified", () => {
    const useful = contact({ appId: "contact-annie-1", organizationId: "org-north-valley" });
    const reviewOnly = contact({ appId: "contact-annie-2", organizationId: "org-schuylkill-haven" });
    const result = matchContactCandidates("Annie Milewski", [useful, reviewOnly], { organizationIds: ["org-north-valley"] });
    const byId = Object.fromEntries(result.map((r) => [r.contact.appId, r.strength]));
    expect(byId["contact-annie-1"]).toBe("useful");
    expect(byId["contact-annie-2"]).toBe("review");
  });
});

describe("matchContactCandidates — archived Contacts remain matchable", () => {
  it("an archived Contact is still returned as a candidate — status is never an identity-eligibility filter", () => {
    const archived = contact({ status: "archived" });
    const result = matchContactCandidates("Annie Milewski", [archived]);
    expect(result).toEqual([{ contact: archived, strength: "review", reasons: ["Name match only — organization not confirmed"] }]);
  });

  it("an archived Contact still gets a strong email match", () => {
    const archived = contact({ status: "archived", email: "annie@northvalley.org" });
    const result = matchContactCandidates("Annie", [archived], { email: "annie@northvalley.org" });
    expect(result[0].strength).toBe("strong");
  });
});
