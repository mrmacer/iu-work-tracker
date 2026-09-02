// @vitest-environment jsdom
//
// Patch 8B — exercises selectContactProvider()'s browser-context branches directly, mirroring
// tests/project-provider-selection.test.ts. Unlike Patch 7's Project list (which used a
// temporary safety-rail env var while IU_Projects's live schema was unverified), the live
// IU_Contacts schema was inspected and approved BEFORE this provider was written — so
// selectContactProvider() uses the existing NEXT_PUBLIC_SHAREPOINT_IU_CONTACTS_LIST_ID
// directly, with no second/temporary variable anywhere.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { DelegatedSharePointContactProvider, MemoryContactProvider, selectContactProvider } from "../lib/contact-provider";

const mocks = vi.hoisted(() => ({
  account: null as AccountInfo | null,
}));

vi.mock("../lib/microsoft-auth", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../lib/microsoft-auth")>();
  return {
    ...actual,
    createBrowserMicrosoftAuthController: vi.fn(() => ({
      initialize: vi.fn(async () => mocks.account),
    })),
  };
});

function stubMicrosoftConfig() {
  vi.stubEnv("NEXT_PUBLIC_MS_ENTRA_CLIENT_ID", "11111111-1111-4111-8111-111111111111");
  vi.stubEnv("NEXT_PUBLIC_MS_ENTRA_TENANT_ID", "3276761c-22db-462b-a930-172d155bd795");
  vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_HOSTNAME", "siu29.sharepoint.com");
  vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_SITE_PATH", "/sites/IUWorkTrackerDEV");
  vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_SITE_ID", "site-1");
}

beforeEach(() => {
  mocks.account = null;
});

afterEach(() => {
  vi.unstubAllEnvs();
});

describe("selectContactProvider — canonical env var", () => {
  it("resolves to the SharePoint provider using the existing NEXT_PUBLIC_SHAREPOINT_IU_CONTACTS_LIST_ID when configured and already signed in", async () => {
    stubMicrosoftConfig();
    vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_IU_CONTACTS_LIST_ID", "iu-contacts-list-id");
    mocks.account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
    const { provider, kind } = await selectContactProvider();
    expect(kind).toBe("sharepoint");
    expect(provider).toBeInstanceOf(DelegatedSharePointContactProvider);
  });

  it("does not silently use a durable fallback when configured but not signed in", async () => {
    stubMicrosoftConfig();
    vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_IU_CONTACTS_LIST_ID", "iu-contacts-list-id");
    mocks.account = null;
    const { provider, kind } = await selectContactProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryContactProvider);
  });

  it("uses the non-durable in-memory provider when Microsoft is not configured at all", async () => {
    const { provider, kind } = await selectContactProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryContactProvider);
  });
});
