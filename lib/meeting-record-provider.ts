import type { AccountInfo } from "@azure/msal-browser";
import {
  createBrowserMicrosoftAuthController,
  InteractiveRedirectStartedError,
  type MicrosoftAuthController,
} from "./microsoft-auth";
import { readDevMicrosoftConfig } from "./microsoft-auth-config";
import { validateMeetingRecordShape, type MeetingRecord } from "./meeting-intelligence-models";
import {
  createMeetingRecordItem,
  findMeetingRecordByAppId,
  listMeetingRecordItems,
  resolveMeetingRecordItem,
  SharePointMeetingRecordsError,
  updateMeetingRecordItem,
  validateMeetingRecordSharePointLimits,
  type SharePointMeetingRecordConfig,
} from "./sharepoint-meeting-records";
import type { ValidationIssue } from "./validation";

export type MeetingRecordResult<T> =
  | { status: "success"; value: T }
  | { status: "validation_error"; errors: ValidationIssue[] }
  | { status: "conflict"; current: MeetingRecord; message: string }
  | { status: "network_error"; message: string }
  | { status: "persistence_error"; message: string };

/** Minimum operations only — no delete, matching InboxIntelligenceProvider (Patch 6B §24: create/read/update only). */
export interface MeetingRecordProvider {
  list(): Promise<MeetingRecordResult<MeetingRecord[]>>;
  create(record: MeetingRecord): Promise<MeetingRecordResult<MeetingRecord>>;
  update(record: MeetingRecord, expectedVersion: number): Promise<MeetingRecordResult<MeetingRecord>>;
}

/**
 * Fallback persistence: in-memory only, scoped to one instance. Nothing here survives a page
 * reload or a failed SharePoint connection — this is intentional, mirroring
 * SessionInboxIntelligenceProvider's role for Inbox Intelligence and PrototypeFallbackProvider's
 * role for Work Records. Used whenever DEV SharePoint configuration is absent, the user isn't
 * signed in, the Meeting Records list has not yet been provisioned, or the SharePoint provider
 * fails to load for any reason. A meeting "saved" here is NOT durable — the UI must say so.
 */
export class MemoryMeetingRecordProvider implements MeetingRecordProvider {
  private records: MeetingRecord[] = [];
  private sequence = 0;

  async list(): Promise<MeetingRecordResult<MeetingRecord[]>> {
    return { status: "success", value: [...this.records].sort((a, b) => b.metadata.createdAt.localeCompare(a.metadata.createdAt)) };
  }

  async create(record: MeetingRecord): Promise<MeetingRecordResult<MeetingRecord>> {
    const shapeIssues = validateMeetingRecordShape(record);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const existing = this.records.find((item) => item.appId === record.appId);
    if (existing) return { status: "conflict", current: existing, message: "A record with this application ID already exists." };
    const now = new Date().toISOString();
    const saved: MeetingRecord = {
      ...record,
      metadata: { providerId: `memory:${++this.sequence}`, version: 1, createdAt: now, modifiedAt: now, syncState: "saved" },
    };
    this.records = [saved, ...this.records];
    return { status: "success", value: saved };
  }

  async update(record: MeetingRecord, expectedVersion: number): Promise<MeetingRecordResult<MeetingRecord>> {
    const shapeIssues = validateMeetingRecordShape(record);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const index = this.records.findIndex((item) => item.appId === record.appId);
    if (index < 0) return { status: "persistence_error", message: "The record no longer exists." };
    const current = this.records[index];
    if (current.metadata.version !== expectedVersion) {
      return { status: "conflict", current, message: "This meeting was changed after you opened it. Your edits are still available." };
    }
    const saved: MeetingRecord = {
      ...record,
      metadata: { ...current.metadata, version: current.metadata.version + 1, modifiedAt: new Date().toISOString() },
    };
    this.records[index] = saved;
    return { status: "success", value: saved };
  }
}

/**
 * Durable DEV SharePoint persistence for Meeting Records, following the exact same pattern as
 * DelegatedSharePointInboxIntelligenceProvider: reuses the existing delegated Microsoft
 * sign-in, the same numeric-version + ETag/If-Match algorithm, and the same "never trust the
 * caller's version" re-read-before-write sequence.
 */
export class DelegatedSharePointMeetingRecordProvider implements MeetingRecordProvider {
  constructor(
    private readonly controller: MicrosoftAuthController,
    private readonly account: AccountInfo,
    private readonly config: SharePointMeetingRecordConfig,
  ) {}

  private token(): Promise<string> {
    return this.controller.acquireGraphToken(this.account);
  }

  async list(): Promise<MeetingRecordResult<MeetingRecord[]>> {
    try {
      const token = await this.token();
      const records = await listMeetingRecordItems(this.config, token);
      records.sort((a, b) => b.metadata.createdAt.localeCompare(a.metadata.createdAt));
      return { status: "success", value: records };
    } catch (error) {
      return this.toErrorResult<MeetingRecord[]>(error);
    }
  }

  async create(record: MeetingRecord): Promise<MeetingRecordResult<MeetingRecord>> {
    const shapeIssues = validateMeetingRecordShape(record);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const limitIssues = validateMeetingRecordSharePointLimits(record);
    if (limitIssues.length) return { status: "validation_error", errors: limitIssues };
    try {
      const token = await this.token();
      const existing = await findMeetingRecordByAppId(this.config, token, record.appId);
      if (existing) {
        return { status: "conflict", current: existing.record, message: "A record with this application ID already exists." };
      }
      const saved = await createMeetingRecordItem(this.config, token, record);
      return { status: "success", value: saved };
    } catch (error) {
      return this.toErrorResult<MeetingRecord>(error);
    }
  }

  async update(record: MeetingRecord, expectedVersion: number): Promise<MeetingRecordResult<MeetingRecord>> {
    const shapeIssues = validateMeetingRecordShape(record);
    if (shapeIssues.length) return { status: "validation_error", errors: shapeIssues };
    const limitIssues = validateMeetingRecordSharePointLimits(record);
    if (limitIssues.length) return { status: "validation_error", errors: limitIssues };
    try {
      const token = await this.token();
      const resolved = await resolveMeetingRecordItem(this.config, token, record);
      if (!resolved) return { status: "persistence_error", message: "The record no longer exists." };
      if (resolved.record.metadata.version !== expectedVersion) {
        return { status: "conflict", current: resolved.record, message: "This meeting was changed after you opened it. Your edits are still available." };
      }
      const saved = await updateMeetingRecordItem(
        this.config,
        token,
        resolved.itemId,
        resolved.etag,
        record,
        expectedVersion + 1,
      );
      return { status: "success", value: saved };
    } catch (error) {
      if (error instanceof SharePointMeetingRecordsError && error.kind === "conflict" && error.current) {
        return { status: "conflict", current: error.current, message: "This meeting was changed after you opened it. Your edits are still available." };
      }
      return this.toErrorResult<MeetingRecord>(error);
    }
  }

  private toErrorResult<T>(error: unknown): MeetingRecordResult<T> {
    if (error instanceof InteractiveRedirectStartedError) {
      return { status: "network_error", message: "Microsoft sign-in confirmation is required. Finish signing in, then try again." };
    }
    if (error instanceof SharePointMeetingRecordsError) {
      return error.kind === "auth"
        ? { status: "network_error", message: error.message }
        : { status: "persistence_error", message: error.message };
    }
    return { status: "network_error", message: "The DEV SharePoint data store could not be reached." };
  }
}

export type ActiveMeetingRecordProviderKind = "sharepoint" | "memory";

/**
 * Provider selection mirrors selectInboxIntelligenceProvider() / selectDataProvider() exactly:
 * the in-memory provider is the default/fallback; SharePoint activates only when DEV
 * configuration is present, a Microsoft account is already signed in (non-interactive — this
 * never triggers a sign-in prompt on its own), AND
 * NEXT_PUBLIC_SHAREPOINT_IU_MEETING_RECORDS_LIST_ID is configured. That variable is
 * intentionally unset until the IU_Meeting_Records list is actually provisioned (see the
 * Patch 6B SharePoint schema approval gate) — until then this always resolves to "memory",
 * exactly like every other resource before its list existed.
 */
export async function selectMeetingRecordProvider(): Promise<{
  provider: MeetingRecordProvider;
  kind: ActiveMeetingRecordProviderKind;
}> {
  if (typeof window === "undefined") return { provider: new MemoryMeetingRecordProvider(), kind: "memory" };

  const config = readDevMicrosoftConfig();
  const siteId = process.env.NEXT_PUBLIC_SHAREPOINT_SITE_ID?.trim();
  const listId = process.env.NEXT_PUBLIC_SHAREPOINT_IU_MEETING_RECORDS_LIST_ID?.trim();
  if (config.status !== "enabled" || !siteId || !listId) {
    return { provider: new MemoryMeetingRecordProvider(), kind: "memory" };
  }

  try {
    const controller = createBrowserMicrosoftAuthController(config.value, window.location.origin);
    const account = await controller.initialize();
    if (!account) return { provider: new MemoryMeetingRecordProvider(), kind: "memory" };
    return {
      provider: new DelegatedSharePointMeetingRecordProvider(controller, account, { siteId, meetingRecordsListId: listId }),
      kind: "sharepoint",
    };
  } catch {
    return { provider: new MemoryMeetingRecordProvider(), kind: "memory" };
  }
}
