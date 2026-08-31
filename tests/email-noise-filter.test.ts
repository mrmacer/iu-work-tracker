import { describe, expect, it } from "vitest";
import { stripDeterministicEmailNoise } from "../lib/email-noise-filter";
import { EXTERNAL_SENDER_WARNING, IMAGE_PLACEHOLDERS, UNSUBSCRIBE_FOOTER, buildEmail } from "./fixtures/email-noise";

describe("stripDeterministicEmailNoise", () => {
  it("removes the external-sender warning banner", () => {
    const email = buildEmail("Hi Sam, can you send the agenda by Friday?", EXTERNAL_SENDER_WARNING);
    const cleaned = stripDeterministicEmailNoise(email);
    expect(cleaned).not.toMatch(/caution/i);
    expect(cleaned).not.toMatch(/do not click links/i);
    expect(cleaned).toContain("send the agenda by Friday");
  });

  it("removes standalone image placeholder lines", () => {
    const email = buildEmail("Please see the attached flyer.", IMAGE_PLACEHOLDERS);
    const cleaned = stripDeterministicEmailNoise(email);
    expect(cleaned).not.toMatch(/\[image\]/i);
    expect(cleaned).not.toMatch(/\[cid:/i);
    expect(cleaned).not.toMatch(/\[logo\]/i);
    expect(cleaned).toContain("attached flyer");
  });

  it("removes standalone unsubscribe/footer-link lines", () => {
    const email = buildEmail("The newsletter is out this week.", UNSUBSCRIBE_FOOTER);
    const cleaned = stripDeterministicEmailNoise(email);
    expect(cleaned).not.toMatch(/unsubscribe/i);
    expect(cleaned).not.toMatch(/manage preferences/i);
    expect(cleaned).not.toMatch(/view in browser/i);
    expect(cleaned).toContain("newsletter is out");
  });

  it("never removes a line that merely contains, rather than exactly matches, a noise pattern", () => {
    const email = "Please unsubscribe the retired staff list from the old distribution group by Friday.";
    expect(stripDeterministicEmailNoise(email)).toBe(email);
  });

  it("leaves signatures, quoted threads, and calendar blocks untouched — conservative by design", () => {
    const email = buildEmail(
      "Greg, can you send the revised agenda by Friday?",
      "Jordan Smith\nDirector of Curriculum\nPhone: 570-555-1234\nFax: 570-555-5678",
      "-----Original Message-----\nFrom: Pat Alvarez\nSent: Monday\nPlease send the grant draft by August 20.",
      "When: Tuesday, September 15, 2026 2:00 PM\nOrganizer: Pat Alvarez",
    );
    expect(stripDeterministicEmailNoise(email)).toBe(email);
  });

  it("preserves blank-line paragraph structure while collapsing excess blank lines left by removed blocks", () => {
    const email = "Real content.\n\n[image]\n[logo]\n\nMore real content.";
    const cleaned = stripDeterministicEmailNoise(email);
    expect(cleaned).toBe("Real content.\n\nMore real content.");
  });
});
