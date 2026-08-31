// @vitest-environment jsdom
//
// Exercises selectDataProvider()'s browser-context branches directly (the existing
// "outside a browser context" test in tests/sharepoint-data-provider.test.ts runs under the
// default Node environment, so `window` is always undefined there and never reaches this
// logic). This is the deliberate no-silent-durable-fallback behavior from Patch 3A: SharePoint
// is the only durable production store; every other case — no config, or config present but
// not yet signed in — returns the same non-durable in-memory provider, never Cloudflare D1
// or any other database.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { DelegatedSharePointDataProvider, MemoryDataProvider, selectDataProvider } from "../lib/data-provider";

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

function stubSharePointConfig() {
  vi.stubEnv("NEXT_PUBLIC_MS_ENTRA_CLIENT_ID", "11111111-1111-4111-8111-111111111111");
  vi.stubEnv("NEXT_PUBLIC_MS_ENTRA_TENANT_ID", "3276761c-22db-462b-a930-172d155bd795");
  vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_HOSTNAME", "siu29.sharepoint.com");
  vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_SITE_PATH", "/sites/IUWorkTrackerDEV");
  vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_SITE_ID", "site-1");
  vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_IU_WORK_RECORDS_LIST_ID", "list-1");
}

beforeEach(() => {
  mocks.account = null;
});

afterEach(() => {
  vi.unstubAllEnvs();
});

describe("selectDataProvider — production provider behavior", () => {
  it("resolves to the SharePoint provider when configured and already signed in", async () => {
    stubSharePointConfig();
    mocks.account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
    const { provider, kind } = await selectDataProvider();
    expect(kind).toBe("sharepoint");
    expect(provider).toBeInstanceOf(DelegatedSharePointDataProvider);
  });

  it("does NOT silently use a durable fallback database when configured but not signed in", async () => {
    stubSharePointConfig();
    mocks.account = null;
    const { provider, kind } = await selectDataProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryDataProvider);
  });

  it("uses the non-durable in-memory provider when Microsoft is not configured at all", async () => {
    const { provider, kind } = await selectDataProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryDataProvider);
  });
});
