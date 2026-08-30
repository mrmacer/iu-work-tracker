import type { AccountInfo } from "@azure/msal-browser";
import {
  createBrowserMicrosoftAuthController,
  InteractiveRedirectStartedError,
  type MicrosoftAuthController,
} from "./microsoft-auth";
import { readDevMicrosoftConfig } from "./microsoft-auth-config";
import type { InboxIntelligenceRecord } from "./inbox-intelligence-models";
import {
  createInboxIntelligenceItem,
  findInboxIntelligenceByAppId,
  listInboxIntelligenceItems,
  resolveInboxIntelligenceItem,
  SharePointInboxIntelligenceError,
  updateInboxIntelligenceItem,
  validateInboxIntelligenceRecord,
  type SharePointInboxIntelligenceConfig,
} from "./sharepoint-inbox-intelligence";
import type { ValidationIssue } from "./validation";

export type InboxIntelligenceResult<T> =
  | { status: "success"; value: T }
  | { status: "validation_error"; errors: ValidationIssue[] }
  | { status: "conflict"; current: InboxIntelligenceRecord; message: string }
  | { status: "network_error"; message: string }
  | { status: "persistence_error"; message: string };

/** Minimum operations only — no delete; the UX (open/waiting/resolved) never needs one. */
export interface InboxIntelligenceProvider {
  list(): Promise<InboxIntelligenceResult<InboxIntelligenceRecord[]>>;
  create(record: InboxIntelligenceRecord): Promise<InboxIntelligenceResult<InboxIntelligenceRecord>>;
  update(record: InboxIntelligenceRecord, expectedVersion: number): Promise<InboxIntelligenceResult<InboxIntelligenceRecord>>;
}

/**
 * Fallback persistence: in-memory only, scoped to one instance. Nothing here survives a
 * page reload or a failed SharePoint connection — this is intentional, mirroring
 * PrototypeFallbackProvider's role for Work Records in lib/data-provider.ts. Used whenever
 * DEV SharePoint configuration is absent, the user isn't signed in, or the SharePoint
 * provider fails to load for any reason.
 */
export class SessionInboxIntelligenceProvider implements InboxIntelligenceProvider {
  private records: InboxIntelligenceRecord[] = [];
  private sequence = 0;

  async list(): Promise<InboxIntelligenceResult<InboxIntelligenceRecord[]>> {
    return { status: "success", value: [...this.records].sort((a, b) => b.metadata.createdAt.localeCompare(a.metadata.createdAt)) };
  }

  async create(record: InboxIntelligenceRecord): Promise<InboxIntelligenceResult<InboxIntelligenceRecord>> {
    const existing = this.records.find((item) => item.appId === record.appId);
    if (existing) return { status: "conflict", current: existing, message: "A record with this application ID already exists." };
    const now = new Date().toISOString();
    const saved: InboxIntelligenceRecord = {
      ...record,
      metadata: { providerId: `memory:${++this.sequence}`, version: 1, createdAt: now, modifiedAt: now, syncState: "saved" },
    };
    this.records = [saved, ...this.records];
    return { status: "success", value: saved };
  }

  async update(record: InboxIntelligenceRecord, expectedVersion: number): Promise<InboxIntelligenceResult<InboxIntelligenceRecord>> {
    const index = this.records.findIndex((item) => item.appId === record.appId);
    if (index < 0) return { status: "persistence_error", message: "The record no longer exists." };
    const current = this.records[index];
    if (current.metadata.version !== expectedVersion) {
      return { status: "conflict", current, message: "This record was changed after you opened it." };
    }
    const saved: InboxIntelligenceRecord = {
      ...record,
      metadata: { ...current.metadata, version: current.metadata.version + 1, modifiedAt: new Date().toISOString() },
    };
    this.records[index] = saved;
    return { status: "success", value: saved };
  }
}

/**
 * Durable DEV SharePoint persistence for Inbox Intelligence, following the exact same
 * pattern as DelegatedSharePointDataProvider in lib/data-provider.ts: reuses the existing
 * delegated Microsoft sign-in, the same numeric-version + ETag/If-Match algorithm, and the
 * same "never trust the caller's version" re-read-before-write sequence.
 */
export class DelegatedSharePointInboxIntelligenceProvider implements InboxIntelligenceProvider {
  constructor(
    private readonly controller: MicrosoftAuthController,
    private readonly account: AccountInfo,
    private readonly config: SharePointInboxIntelligenceConfig,
  ) {}

  private token(): Promise<string> {
    return this.controller.acquireGraphToken(this.account);
  }

  async list(): Promise<InboxIntelligenceResult<InboxIntelligenceRecord[]>> {
    try {
      const token = await this.token();
      const records = await listInboxIntelligenceItems(this.config, token);
      records.sort((a, b) => b.metadata.createdAt.localeCompare(a.metadata.createdAt));
      return { status: "success", value: records };
    } catch (error) {
      return this.toErrorResult<InboxIntelligenceRecord[]>(error);
    }
  }

  async create(record: InboxIntelligenceRecord): Promise<InboxIntelligenceResult<InboxIntelligenceRecord>> {
    const issues = validateInboxIntelligenceRecord(record);
    if (issues.length) return { status: "validation_error", errors: issues };
    try {
      const token = await this.token();
      const existing = await findInboxIntelligenceByAppId(this.config, token, record.appId);
      if (existing) {
        return { status: "conflict", current: existing.record, message: "A record with this application ID already exists." };
      }
      const saved = await createInboxIntelligenceItem(this.config, token, record);
      return { status: "success", value: saved };
    } catch (error) {
      return this.toErrorResult<InboxIntelligenceRecord>(error);
    }
  }

  async update(record: InboxIntelligenceRecord, expectedVersion: number): Promise<InboxIntelligenceResult<InboxIntelligenceRecord>> {
    const issues = validateInboxIntelligenceRecord(record);
    if (issues.length) return { status: "validation_error", errors: issues };
    try {
      const token = await this.token();
      const resolved = await resolveInboxIntelligenceItem(this.config, token, record);
      if (!resolved) return { status: "persistence_error", message: "The record no longer exists." };
      if (resolved.record.metadata.version !== expectedVersion) {
        return { status: "conflict", current: resolved.record, message: "This record was changed after you opened it." };
      }
      const saved = await updateInboxIntelligenceItem(
        this.config,
        token,
        resolved.itemId,
        resolved.etag,
        record,
        expectedVersion + 1,
      );
      return { status: "success", value: saved };
    } catch (error) {
      if (error instanceof SharePointInboxIntelligenceError && error.kind === "conflict" && error.current) {
        return { status: "conflict", current: error.current, message: "This record was changed after you opened it." };
      }
      return this.toErrorResult<InboxIntelligenceRecord>(error);
    }
  }

  private toErrorResult<T>(error: unknown): InboxIntelligenceResult<T> {
    if (error instanceof InteractiveRedirectStartedError) {
      return { status: "network_error", message: "Microsoft sign-in confirmation is required. Finish signing in, then try again." };
    }
    if (error instanceof SharePointInboxIntelligenceError) {
      return error.kind === "auth"
        ? { status: "network_error", message: error.message }
        : { status: "persistence_error", message: error.message };
    }
    return { status: "network_error", message: "The DEV SharePoint data store could not be reached." };
  }
}

export type ActiveInboxIntelligenceProviderKind = "sharepoint" | "session";

/**
 * Provider selection mirrors selectDataProvider() in lib/data-provider.ts exactly: the
 * session-only provider is the default/fallback; SharePoint activates only when DEV
 * configuration is present AND a Microsoft account is already signed in (non-interactive —
 * this never triggers a sign-in prompt on its own).
 */
export async function selectInboxIntelligenceProvider(): Promise<{
  provider: InboxIntelligenceProvider;
  kind: ActiveInboxIntelligenceProviderKind;
}> {
  if (typeof window === "undefined") return { provider: new SessionInboxIntelligenceProvider(), kind: "session" };

  const config = readDevMicrosoftConfig();
  const siteId = process.env.NEXT_PUBLIC_SHAREPOINT_SITE_ID?.trim();
  const listId = process.env.NEXT_PUBLIC_SHAREPOINT_IU_INBOX_INTELLIGENCE_LIST_ID?.trim();
  if (config.status !== "enabled" || !siteId || !listId) {
    return { provider: new SessionInboxIntelligenceProvider(), kind: "session" };
  }

  try {
    const controller = createBrowserMicrosoftAuthController(config.value, window.location.origin);
    const account = await controller.initialize();
    if (!account) return { provider: new SessionInboxIntelligenceProvider(), kind: "session" };
    return {
      provider: new DelegatedSharePointInboxIntelligenceProvider(controller, account, { siteId, inboxIntelligenceListId: listId }),
      kind: "sharepoint",
    };
  } catch {
    return { provider: new SessionInboxIntelligenceProvider(), kind: "session" };
  }
}
