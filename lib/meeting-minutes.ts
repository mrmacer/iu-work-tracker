import type { MeetingCandidateType } from "./meeting-intelligence-models";

/**
 * Minimal client-side meeting model — deliberately no SharePoint IDs, no persistence
 * metadata, no entity/relationship fields. See docs/AI_HANDOFF.md "Meeting Notes V1": this
 * lives only in React state in app/MeetingNotes.tsx, never SharePoint, never browser storage.
 */
export type MeetingDraft = {
  title: string;
  date: string;
  meetingType: string;
  attendeesText: string;
  agendaText: string;
  notesText: string;
};

/** The minimal candidate shape buildDraftMinutes needs — deliberately decoupled from the
 * review UI's own state type (id, sourceExcerpt, etc. are irrelevant to composing minutes). */
export type DraftMinutesCandidate = {
  type: MeetingCandidateType;
  title: string;
  detail: string;
  ownerText: string | null;
  dueText: string | null;
  selected: boolean;
};

/**
 * Deterministic, client-side-only composition of plain-text draft minutes from the CURRENT
 * reviewed state — meeting details, agenda, and whichever candidates are currently selected.
 * No AI call: this is pure string composition, matching docs/AI_HANDOFF.md "Meeting Notes
 * V1" — "Do NOT make a second AI call to generate minutes." Unselected (ignored) candidates
 * and removed candidates (already absent from the array) are both excluded. IDEA/QUESTION/
 * KNOWLEDGE/FOLLOW_UP_AGENDA candidates are intentionally left out of the minutes text itself
 * — they remain reviewable intelligence, not minutes content.
 */
export function buildDraftMinutes(draft: MeetingDraft, candidates: DraftMinutesCandidate[]): string {
  const selected = candidates.filter((candidate) => candidate.selected);
  const summary = selected.find((candidate) => candidate.type === "SUMMARY");
  const decisions = selected.filter((candidate) => candidate.type === "DECISION");
  const actions = selected.filter((candidate) => candidate.type === "ACTION");

  const lines: string[] = [];
  lines.push(draft.title.trim() || "Untitled meeting");
  const metaParts = [draft.date.trim(), draft.meetingType.trim()].filter(Boolean);
  if (metaParts.length) lines.push(metaParts.join(" · "));
  if (draft.attendeesText.trim()) lines.push(`Attendees: ${draft.attendeesText.trim()}`);

  if (draft.agendaText.trim()) {
    lines.push("", "Agenda", draft.agendaText.trim());
  }

  if (summary) {
    lines.push("", "Summary", summary.detail.trim() || summary.title.trim());
  }

  if (decisions.length) {
    lines.push("", "Decisions");
    for (const decision of decisions) lines.push(`- ${decision.title.trim()}`);
  }

  if (actions.length) {
    lines.push("", "Action Items");
    for (const action of actions) {
      const meta = [
        action.ownerText ? `Owner: ${action.ownerText}` : null,
        action.dueText ? `Due: ${action.dueText}` : null,
      ]
        .filter(Boolean)
        .join(" · ");
      lines.push(`- ${action.title.trim()}${meta ? ` (${meta})` : ""}`);
    }
  }

  return lines.join("\n").trim();
}
