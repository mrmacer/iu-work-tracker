import type { InboxIntelligenceRecord } from "./inbox-intelligence-models";

export type InboxIntelligenceSummary = {
  needsAttentionCount: number;
  openActionCount: number;
};

export interface InboxIntelligenceProvider {
  list(): InboxIntelligenceRecord[];
  save(record: InboxIntelligenceRecord): void;
  summary(): InboxIntelligenceSummary;
}

/**
 * V1 persistence for Inbox Intelligence: in-memory only, scoped to one instance. Nothing
 * here survives a page reload — this is intentional per docs/AI_HANDOFF.md's Inbox
 * Intelligence scope and docs/INBOX_INTELLIGENCE_V1_REPORT.md "Persistence status": no
 * SharePoint list has been provisioned for this record type, so V1 does not pretend this
 * data is durable. See that report's "Proposed SharePoint persistence model" for the
 * intended first-class record shape a future phase would provision and swap in behind this
 * same interface.
 */
export class SessionInboxIntelligenceProvider implements InboxIntelligenceProvider {
  private records: InboxIntelligenceRecord[] = [];

  list(): InboxIntelligenceRecord[] {
    return [...this.records].sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  }

  save(record: InboxIntelligenceRecord): void {
    this.records = [record, ...this.records.filter((item) => item.appId !== record.appId)];
  }

  summary(): InboxIntelligenceSummary {
    return {
      needsAttentionCount: this.records.filter((record) => record.analysis.needsAttention).length,
      openActionCount: this.records.reduce((sum, record) => sum + record.analysis.actionItems.length, 0),
    };
  }
}
