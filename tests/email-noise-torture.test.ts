import { describe, expect, it } from "vitest";
import type Anthropic from "@anthropic-ai/sdk";
import { analyzeEmailWithClaude, SYSTEM_PROMPT } from "../lib/anthropic-email-analysis";
import { EmailAnalysisSchema, type EmailAnalysis } from "../lib/inbox-intelligence-models";
import { buildWorkRecordDraftFromAnalysis } from "../lib/inbox-intelligence-work-record";
import { REFERENCE_DATA } from "../lib/reference-data";
import { WORK_RECORD_SCHEMA_VERSION, type WorkRecord } from "../lib/models";
import {
  autoReplyBody,
  buildEmail,
  CALENDAR_BLOCK,
  CONFIDENTIALITY_NOTICE,
  EXTERNAL_SENDER_WARNING,
  forwardedHeader,
  IMAGE_PLACEHOLDERS,
  LEGAL_SECURITY_FOOTER,
  SIGNATURE_BLOCK,
  SOCIAL_LINKS,
  TEAMS_BOILERPLATE,
  UNSUBSCRIBE_FOOTER,
  withQuotedThread,
  ZOOM_BOILERPLATE,
} from "./fixtures/email-noise";

// No test in this file ever calls the real Anthropic API. Every "model response" is a
// hand-authored fixture returned by a fake `{ messages: { parse } }` client, matching the
// established pattern in tests/inbox-intelligence-analysis.test.ts. Two different things are
// being tested, and each test says which:
//   (a) "outgoing request" assertions — proves the deterministic preprocessor and the
//       hardened system prompt actually reach the model, using captured request params.
//   (b) "pipeline fidelity" assertions — proves that IF the model returns the desired clean
//       extraction (as the hardened prompt instructs it to), nothing downstream (validation,
//       normalization, entity matching, Work Record mapping) corrupts or re-pollutes it.
// Neither proves live model behavior — see docs/AI_HANDOFF.md "Email Noise Torture Test"
// "Known limitations" for why that is out of reach without a real (forbidden) API call.

function baseAnalysis(overrides: Partial<EmailAnalysis> = {}): EmailAnalysis {
  return EmailAnalysisSchema.parse({
    summary: "Test summary",
    priority: "medium",
    needsAttention: false,
    actionItems: [],
    followUp: "",
    people: [],
    organizations: [],
    districts: [],
    projects: [],
    tags: [],
    suggestedWorkType: null,
    suggestedWorkRecord: { title: "Suggested title", description: "Suggested description" },
    ...overrides,
  });
}

type Captured = { params?: { system?: unknown; messages?: { content?: unknown }[] } };

function mockClient(mockAnalysis: EmailAnalysis, captured: Captured): Pick<Anthropic, "messages"> {
  return {
    messages: {
      parse: async (params: unknown) => {
        captured.params = params as Captured["params"];
        return { model: "claude-opus-5", usage: { input_tokens: 500, output_tokens: 150 }, parsed_output: mockAnalysis };
      },
    },
  } as unknown as Pick<Anthropic, "messages">;
}

async function analyze(rawEmail: string, mockAnalysis: EmailAnalysis) {
  const captured: Captured = {};
  const result = await analyzeEmailWithClaude(rawEmail, mockClient(mockAnalysis, captured));
  return { result, captured };
}

function outgoingText(captured: Captured): string {
  return String(captured.params?.messages?.[0]?.content ?? "");
}

const NEEDS_ATTENTION_FLAG = true; // shorthand used throughout — a genuinely current, real request

describe("System prompt hardening (content, not brittle exact-match)", () => {
  it("instructs prioritizing the current message over quoted/forwarded history", () => {
    expect(SYSTEM_PROMPT.toLowerCase()).toMatch(/quoted|forwarded/);
    expect(SYSTEM_PROMPT.toLowerCase()).toMatch(/newest|current/);
  });

  it("instructs ignoring signatures, disclaimers, and meeting-join boilerplate", () => {
    expect(SYSTEM_PROMPT.toLowerCase()).toMatch(/signature/);
    expect(SYSTEM_PROMPT.toLowerCase()).toMatch(/teams|zoom/);
  });

  it("instructs not treating technology/platform names as entities from boilerplate", () => {
    expect(SYSTEM_PROMPT.toLowerCase()).toMatch(/technology|platform|vendor/);
  });

  it("instructs deduplicating repeated entities", () => {
    expect(SYSTEM_PROMPT.toLowerCase()).toMatch(/more than once|duplicate/);
  });
});

describe("1. External sender banner ignored", () => {
  it("is stripped before the request reaches the model, and produces no junk", async () => {
    const email = buildEmail("Hi Sam, please send the STEELS agenda by Friday.", EXTERNAL_SENDER_WARNING);
    const mock = baseAnalysis({
      summary: "Sender asked for the STEELS agenda by Friday.",
      needsAttention: NEEDS_ATTENTION_FLAG,
      actionItems: [{ action: "Send the STEELS agenda", dueDate: "2026-09-04", owner: "me" }],
    });
    const { result, captured } = await analyze(email, mock);
    expect(outgoingText(captured)).not.toMatch(/caution/i);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.actionItems).toHaveLength(1);
    expect(result.analysis.tags.join(" ")).not.toMatch(/caution|external/i);
  });
});

describe("2. Confidentiality notice ignored", () => {
  it("produces no action/entity/tag from the disclaimer itself", async () => {
    const email = buildEmail("Please review the attached budget summary before Thursday's meeting.", CONFIDENTIALITY_NOTICE);
    const mock = baseAnalysis({
      summary: "Reviewer asked to review the budget summary before Thursday's meeting.",
      needsAttention: true,
      actionItems: [{ action: "Review the attached budget summary", dueDate: null, owner: "me" }],
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.actionItems.some((item) => /confidential|privileged/i.test(item.action))).toBe(false);
    expect(result.analysis.people).toEqual([]);
    expect(result.analysis.organizations).toEqual([]);
  });
});

describe("3 & 4. Signature contact details do not become actions or deadlines", () => {
  it("keeps phone/fax/address text out of action items and due dates", async () => {
    const email = buildEmail("Greg, can you send the revised roster by Monday?", SIGNATURE_BLOCK);
    const mock = baseAnalysis({
      summary: "Jordan asked Greg to send the revised roster by Monday.",
      needsAttention: true,
      actionItems: [{ action: "Send the revised roster", dueDate: "2026-09-07", owner: "me" }],
      people: ["Jordan Smith"],
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    for (const item of result.analysis.actionItems) {
      expect(item.action).not.toMatch(/570-555|fax|phone/i);
    }
    expect(result.analysis.actionItems.every((item) => item.dueDate === null || /^\d{4}-\d{2}-\d{2}$/.test(item.dueDate))).toBe(true);
    expect(result.analysis.projects).not.toContain("www.example.org");
  });
});

describe("5. Teams boilerplate ignored", () => {
  it("does not create a Microsoft/Teams organization or a join-meeting action", async () => {
    const email = buildEmail("Let's meet Tuesday at 2 to discuss the CSO equipment order.", TEAMS_BOILERPLATE);
    const mock = baseAnalysis({
      summary: "Proposed meeting Tuesday at 2 about the CSO equipment order.",
      actionItems: [],
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.organizations).not.toContain("Microsoft");
    expect(result.analysis.organizations).not.toContain("Microsoft Teams");
    expect(result.analysis.projects).not.toContain("Microsoft Teams");
    expect(result.analysis.actionItems.some((item) => /join.*meeting/i.test(item.action))).toBe(false);
  });
});

describe("6. Zoom boilerplate ignored", () => {
  it("does not create a Zoom organization or a join-meeting action", async () => {
    const email = buildEmail("Let's meet Tuesday at 2 to discuss the CSO equipment order.", ZOOM_BOILERPLATE);
    const mock = baseAnalysis({ summary: "Proposed meeting Tuesday at 2 about the CSO equipment order." });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.organizations).not.toContain("Zoom");
    expect(result.analysis.projects).not.toContain("Zoom");
  });
});

describe("7. Social links ignored", () => {
  it("does not create organizations/tags from footer social platform names", async () => {
    const email = buildEmail("The newsletter is out — nothing needed from you.", SOCIAL_LINKS);
    const mock = baseAnalysis({ summary: "FYI: the newsletter is out." });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    for (const name of ["Facebook", "Instagram", "LinkedIn", "YouTube"]) {
      expect(result.analysis.organizations).not.toContain(name);
      expect(result.analysis.tags).not.toContain(name);
    }
    expect(result.analysis.actionItems).toEqual([]);
  });
});

describe("8. Unsubscribe/footer ignored", () => {
  it("is stripped before the model sees it and never becomes an action", async () => {
    const email = buildEmail("The newsletter is out — nothing needed from you.", UNSUBSCRIBE_FOOTER);
    const mock = baseAnalysis({ summary: "FYI: the newsletter is out." });
    const { result, captured } = await analyze(email, mock);
    expect(outgoingText(captured)).not.toMatch(/unsubscribe/i);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.actionItems.some((item) => /unsubscribe/i.test(item.action))).toBe(false);
  });
});

describe("9. Image placeholders ignored", () => {
  it("is stripped before the model sees it and produces no intelligence", async () => {
    const email = buildEmail("See the attached flyer for details.", IMAGE_PLACEHOLDERS);
    const mock = baseAnalysis({ summary: "FYI: flyer attached." });
    const { result, captured } = await analyze(email, mock);
    expect(outgoingText(captured)).not.toMatch(/\[image\]|\[cid:|\[logo\]/i);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.actionItems).toEqual([]);
  });
});

describe("10. Forwarded header labels ignored", () => {
  it("does not treat From:/Sent:/To:/Subject: label text as an action or invented entity", async () => {
    const email = forwardedHeader("Pat Alvarez", "Monday, August 24, 2026", "Sam Rivera", "STEELS quarterly report");
    const mock = baseAnalysis({ summary: "Forwarded message about the STEELS quarterly report.", people: ["Pat Alvarez"] });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.actionItems.some((item) => /^(from|sent|to|subject):/i.test(item.action))).toBe(false);
    expect(result.analysis.people).not.toContain("Sent");
    expect(result.analysis.people).not.toContain("Subject");
  });
});

describe("11. Repeated signatures/entities deduplicated", () => {
  it("normalizeEmailAnalysis dedupes a person/organization the model returned more than once", async () => {
    const email = withQuotedThread(
      buildEmail("Any update on this?", SIGNATURE_BLOCK),
      "Jordan Smith",
      "Friday",
      buildEmail("Here's the update.", SIGNATURE_BLOCK),
    );
    // Deliberately construct a mock that has NOT deduplicated — proves normalizeEmailAnalysis
    // is the safety net, not merely a hope that the model complied.
    const mock = baseAnalysis({
      people: ["Jordan Smith", "Jordan Smith"],
      organizations: ["Example Area School District", "Example Area School District"],
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.people).toEqual(["Jordan Smith"]);
    expect(result.analysis.organizations).toEqual(["Example Area School District"]);
  });
});

describe("12 & 17. Current action and its real deadline survive heavy footer noise", () => {
  it("extracts one real action with its stated due date despite a full noise stack", async () => {
    const email = buildEmail(
      "Greg, please send the STEELS meeting agenda to the district team by Friday.",
      EXTERNAL_SENDER_WARNING,
      SIGNATURE_BLOCK,
      SOCIAL_LINKS,
      CONFIDENTIALITY_NOTICE,
      TEAMS_BOILERPLATE,
    );
    const mock = baseAnalysis({
      summary: "Jordan asked Greg to send the STEELS meeting agenda to the district team by Friday.",
      needsAttention: true,
      actionItems: [{ action: "Send the STEELS meeting agenda to the district team", dueDate: "2026-09-04", owner: "me" }],
      people: ["Jordan Smith"],
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.actionItems).toHaveLength(1);
    expect(result.analysis.actionItems[0].action).toMatch(/steels/i);
    expect(result.analysis.actionItems[0].dueDate).toBe("2026-09-04");
    expect(result.analysis.organizations).toEqual([]);
    expect(result.analysis.tags.some((tag) => /facebook|confidential|teams/i.test(tag))).toBe(false);
  });
});

describe("13 & 14. Current action beats quoted historical action; resolved requests are not resurrected", () => {
  it("does not surface an old quoted request as a current action when the new message closes it out", async () => {
    const email = withQuotedThread("Thanks — that takes care of it.", "Pat Alvarez", "Monday, August 24, 2026", "Please send the grant draft by August 20.");
    const mock = baseAnalysis({
      summary: "Sender confirmed the earlier request is resolved.",
      needsAttention: false,
      actionItems: [],
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.actionItems).toEqual([]);
    expect(result.analysis.actionItems.some((item) => /grant draft/i.test(item.action))).toBe(false);
  });
});

describe("15. Forward-with-new-instruction extracts the new instruction", () => {
  it("extracts the current review/invite decision, not the forwarded message's old requests", async () => {
    const email = withQuotedThread(
      "Can you review the forwarded note below and tell me whether we should invite them to the STEM Ecosystem meeting?",
      "Sam Rivera",
      "last week",
      "Please send the vendor quote by end of week and confirm the PO number.",
    );
    const mock = baseAnalysis({
      summary: "Asked to review a forwarded note and decide whether to invite the sender to the STEM Ecosystem meeting.",
      needsAttention: true,
      actionItems: [{ action: "Review the forwarded note and decide on inviting them to the STEM Ecosystem meeting", dueDate: null, owner: "me" }],
      projects: ["STEM Ecosystem"],
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.actionItems).toHaveLength(1);
    expect(result.analysis.actionItems[0].action).toMatch(/invite|review/i);
    expect(result.analysis.actionItems.some((item) => /vendor quote|po number/i.test(item.action))).toBe(false);
  });
});

describe("16. Meeting link boilerplate does not become work", () => {
  it("keeps meeting/topic intelligence but no conference-technology entity or action", async () => {
    const email = buildEmail("Let's meet Tuesday at 2 to discuss CSO equipment.", TEAMS_BOILERPLATE, ZOOM_BOILERPLATE);
    const mock = baseAnalysis({
      summary: "Proposed a Tuesday 2pm meeting about CSO equipment.",
      suggestedWorkType: "Internal planning",
      suggestedWorkRecord: { title: "CSO equipment planning meeting", description: "Discuss CSO equipment on Tuesday at 2pm." },
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.organizations).toEqual([]);
    expect(result.analysis.projects).not.toContain("Microsoft Teams");
    expect(result.analysis.projects).not.toContain("Zoom");
    expect(result.analysis.suggestedWorkRecord.title).toMatch(/cso equipment/i);
  });
});

describe("18. No false deadline from phone/fax/meeting IDs", () => {
  it("never emits a due date derived from a phone number, fax number, or meeting ID", async () => {
    const email = buildEmail("No action needed — just confirming receipt.", SIGNATURE_BLOCK, TEAMS_BOILERPLATE, ZOOM_BOILERPLATE);
    const mock = baseAnalysis({ summary: "Sender confirmed receipt; no action needed.", actionItems: [] });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    for (const item of result.analysis.actionItems) {
      if (item.dueDate) expect(item.dueDate).toMatch(/^\d{4}-\d{2}-\d{2}$/);
    }
    expect(result.analysis.actionItems).toEqual([]);
  });
});

describe("19. No false tags from boilerplate", () => {
  it("does not tag generic footer/platform words unless genuinely central to the work content", async () => {
    const email = buildEmail(
      "Reminder: submit your timesheet by Friday.",
      CONFIDENTIALITY_NOTICE,
      LEGAL_SECURITY_FOOTER,
      SOCIAL_LINKS,
      UNSUBSCRIBE_FOOTER,
    );
    const mock = baseAnalysis({
      summary: "Reminder to submit timesheet by Friday.",
      needsAttention: true,
      actionItems: [{ action: "Submit timesheet", dueDate: "2026-09-04", owner: "me" }],
      tags: ["timesheet"],
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    for (const banned of ["confidential", "security", "facebook", "unsubscribe", "microsoft", "zoom"]) {
      expect(result.analysis.tags.map((tag) => tag.toLowerCase())).not.toContain(banned);
    }
  });
});

describe("20. No fake Work Record suggestion from boilerplate", () => {
  it("never suggests reading a disclaimer, attending a platform, visiting a website, or unsubscribing as work", async () => {
    const email = buildEmail("FYI only — no action needed.", CONFIDENTIALITY_NOTICE, TEAMS_BOILERPLATE, SOCIAL_LINKS, UNSUBSCRIBE_FOOTER);
    const mock = baseAnalysis({
      summary: "FYI email; no action needed.",
      suggestedWorkType: null,
      suggestedWorkRecord: { title: "FYI email review", description: "No follow-up required." },
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    const suggestion = `${result.analysis.suggestedWorkRecord.title} ${result.analysis.suggestedWorkRecord.description}`.toLowerCase();
    for (const banned of ["confidentiality notice", "attend microsoft teams", "visit website", "unsubscribe", "call fax"]) {
      expect(suggestion).not.toContain(banned);
    }
    // Feeding the clean result into the existing, unmodified Work Record mapping produces no junk draft either.
    const baseRecord: WorkRecord = {
      appId: "draft", title: "", activityDate: "2026-08-29", activityType: "", description: "", detailedNotes: "",
      durationMinutes: 60, status: "complete", engagementScope: "none", projectIds: [], organizationIds: [], contactIds: [],
      categoryIds: [], reach: { educatorsLeaders: 0, studentsFamilies: 0, workforceCommunity: 0, other: 0 }, evidenceSummary: "",
      evidenceReferenceIds: [], output: "", outcome: "", nextStep: "", followUpNeeded: false, followUpDate: null,
      orbit: { reportable: false, primaryDeliverable: null, supportingDeliverables: [], stemPocMinutes: 0, tacMinutes: 0, evidence: "" },
      schemaVersion: WORK_RECORD_SCHEMA_VERSION, metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" }, isSample: false,
    };
    const draft = buildWorkRecordDraftFromAnalysis(result.analysis, REFERENCE_DATA, baseRecord);
    expect(draft.organizationIds).toEqual([]);
    expect(draft.projectIds).toEqual([]);
  });
});

describe("Combined torture scenarios", () => {
  it("SCENARIO 2 — waiting/response context survives noise without a junk action", async () => {
    const email = buildEmail(
      "We're waiting for the equipment vendor to send the revised quote. I'll forward it when it arrives.",
      SIGNATURE_BLOCK,
      LEGAL_SECURITY_FOOTER,
    );
    const mock = baseAnalysis({
      summary: "Sender is waiting on the equipment vendor's revised quote and will forward it when it arrives.",
      followUp: "Forward the vendor quote once it arrives.",
      actionItems: [],
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    for (const banned of [/call.*fax/i, /visit.*website/i, /unsubscribe/i, /open attachment/i]) {
      expect(result.analysis.actionItems.some((item) => banned.test(item.action))).toBe(false);
    }
  });

  it("SCENARIO 5 — meeting/topic intelligence survives a meeting-link overload without entity pollution", async () => {
    const email = buildEmail("Let's meet Tuesday at 2 to discuss CSO equipment.", TEAMS_BOILERPLATE, ZOOM_BOILERPLATE, CALENDAR_BLOCK);
    const mock = baseAnalysis({
      summary: "Proposed a Tuesday 2pm meeting to discuss CSO equipment.",
      suggestedWorkRecord: { title: "CSO equipment discussion", description: "Meeting Tuesday at 2pm about CSO equipment." },
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    expect(result.analysis.organizations).toEqual([]);
    expect(result.analysis.projects).not.toContain("Microsoft Teams");
    expect(result.analysis.projects).not.toContain("Zoom");
  });

  it("Auto-reply: boilerplate portion produces no junk entity while a genuine return date may be retained", async () => {
    const email = buildEmail(autoReplyBody("September 8, 2026"), SIGNATURE_BLOCK, CONFIDENTIALITY_NOTICE);
    const mock = baseAnalysis({
      summary: "Automatic reply: sender is out of office until September 8, 2026.",
      needsAttention: false,
      followUp: "Sender returns September 8, 2026 — resend if this is still needed then.",
      actionItems: [],
    });
    const { result } = await analyze(email, mock);
    expect(result.status).toBe("success");
    if (result.status !== "success") return;
    // No invented user obligation merely because an auto-reply exists.
    expect(result.analysis.actionItems).toEqual([]);
  });
});
