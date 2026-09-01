import { describe, expect, it } from "vitest";
import { buildDraftMinutes, type DraftMinutesCandidate, type MeetingDraft } from "../lib/meeting-minutes";

function draft(overrides: Partial<MeetingDraft> = {}): MeetingDraft {
  return {
    title: "STEELS quarterly planning",
    date: "2026-09-01",
    meetingType: "District Meeting",
    attendeesText: "Greg, Annie, Kim",
    agendaText: "1. Grant status\n2. Fall meeting date",
    notesText: "",
    ...overrides,
  };
}

function candidate(overrides: Partial<DraftMinutesCandidate> = {}): DraftMinutesCandidate {
  return {
    type: "ACTION",
    title: "Call the district about the venue",
    detail: "Confirm venue details.",
    ownerText: "Annie",
    dueText: "Friday",
    selected: true,
    ...overrides,
  };
}

describe("buildDraftMinutes — no AI call, pure composition", () => {
  it("includes meeting details", () => {
    const text = buildDraftMinutes(draft(), []);
    expect(text).toContain("STEELS quarterly planning");
    expect(text).toContain("2026-09-01");
    expect(text).toContain("District Meeting");
    expect(text).toContain("Attendees: Greg, Annie, Kim");
  });

  it("includes the agenda", () => {
    const text = buildDraftMinutes(draft(), []);
    expect(text).toContain("Agenda");
    expect(text).toContain("1. Grant status");
  });

  it("omits the agenda section when there is no agenda text", () => {
    const text = buildDraftMinutes(draft({ agendaText: "" }), []);
    expect(text).not.toContain("Agenda");
  });

  it("includes the SUMMARY candidate's detail", () => {
    const text = buildDraftMinutes(draft(), [candidate({ type: "SUMMARY", title: "Meeting summary", detail: "Grant on track, fall meeting moved." })]);
    expect(text).toContain("Summary");
    expect(text).toContain("Grant on track, fall meeting moved.");
  });

  it("falls back to the SUMMARY candidate's title when detail is empty", () => {
    const text = buildDraftMinutes(draft(), [candidate({ type: "SUMMARY", title: "Short summary title", detail: "" })]);
    expect(text).toContain("Short summary title");
  });

  it("includes selected DECISION candidates", () => {
    const text = buildDraftMinutes(draft(), [candidate({ type: "DECISION", title: "Move fall meeting to October" })]);
    expect(text).toContain("Decisions");
    expect(text).toContain("- Move fall meeting to October");
  });

  it("includes selected ACTION candidates with owner and due rendered only when present", () => {
    const text = buildDraftMinutes(draft(), [
      candidate({ type: "ACTION", title: "Call the district about the venue", ownerText: "Annie", dueText: "Friday" }),
      candidate({ type: "ACTION", title: "Send the updated agenda", ownerText: null, dueText: null }),
    ]);
    expect(text).toContain("Action Items");
    expect(text).toContain("- Call the district about the venue (Owner: Annie · Due: Friday)");
    expect(text).toContain("- Send the updated agenda");
    expect(text).not.toContain("Send the updated agenda (");
  });

  it("excludes ignored (deselected) candidates", () => {
    const text = buildDraftMinutes(draft(), [candidate({ type: "DECISION", title: "Deselected decision", selected: false })]);
    expect(text).not.toContain("Deselected decision");
    expect(text).not.toContain("Decisions");
  });

  it("excludes removed candidates (simply absent from the array)", () => {
    const text = buildDraftMinutes(draft(), [candidate({ type: "DECISION", title: "Still present decision" })]);
    expect(text).toContain("Still present decision");
    expect(text).not.toContain("Removed decision that was never passed in");
  });

  it("excludes IDEA/QUESTION/KNOWLEDGE/FOLLOW_UP_AGENDA from the minutes text", () => {
    const text = buildDraftMinutes(draft(), [
      candidate({ type: "IDEA", title: "Partner map idea", ownerText: null, dueText: null }),
      candidate({ type: "QUESTION", title: "DEP funding question", ownerText: null, dueText: null }),
      candidate({ type: "KNOWLEDGE", title: "Impact numbers work best", ownerText: null, dueText: null }),
      candidate({ type: "FOLLOW_UP_AGENDA", title: "STEM site map for next month", ownerText: null, dueText: null }),
    ]);
    expect(text).not.toContain("Partner map idea");
    expect(text).not.toContain("DEP funding question");
    expect(text).not.toContain("Impact numbers work best");
    expect(text).not.toContain("STEM site map for next month");
  });

  it("reflects user-edited candidate text, not any notion of original AI output", () => {
    const text = buildDraftMinutes(draft(), [candidate({ type: "ACTION", title: "Human-edited action title", ownerText: "Human-edited owner" })]);
    expect(text).toContain("Human-edited action title");
    expect(text).toContain("Human-edited owner");
  });

  it("requires no AI call — buildDraftMinutes is a plain synchronous function", () => {
    expect(buildDraftMinutes).not.toHaveProperty("then");
    expect(typeof buildDraftMinutes(draft(), [])).toBe("string");
  });
});
