// @vitest-environment jsdom
//
// Patch 7B — exercises selectProjectProvider()'s browser-context branches directly, mirroring
// tests/data-provider-selection.test.ts exactly. IU_Projects is now the single authoritative
// durable-Projects list, configured through the one canonical
// NEXT_PUBLIC_SHAREPOINT_IU_PROJECTS_LIST_ID variable — the temporary
// NEXT_PUBLIC_SHAREPOINT_IU_DURABLE_PROJECTS_LIST_ID safety rail used while that list's live
// schema was unverified (Patch 7) has been fully retired and must have no effect.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { DelegatedSharePointProjectProvider, MemoryProjectProvider, selectProjectProvider } from "../lib/project-provider";

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

describe("selectProjectProvider — canonical env var only", () => {
  it("resolves to the SharePoint provider using NEXT_PUBLIC_SHAREPOINT_IU_PROJECTS_LIST_ID when configured and already signed in", async () => {
    stubMicrosoftConfig();
    vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_IU_PROJECTS_LIST_ID", "iu-projects-list-id");
    mocks.account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
    const { provider, kind } = await selectProjectProvider();
    expect(kind).toBe("sharepoint");
    expect(provider).toBeInstanceOf(DelegatedSharePointProjectProvider);
  });

  it("does NOT resolve to SharePoint using the retired NEXT_PUBLIC_SHAREPOINT_IU_DURABLE_PROJECTS_LIST_ID variable — it must be ignored entirely", async () => {
    stubMicrosoftConfig();
    vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_IU_DURABLE_PROJECTS_LIST_ID", "some-stale-value");
    mocks.account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
    const { provider, kind } = await selectProjectProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryProjectProvider);
  });

  it("does not silently use a durable fallback when configured but not signed in", async () => {
    stubMicrosoftConfig();
    vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_IU_PROJECTS_LIST_ID", "iu-projects-list-id");
    mocks.account = null;
    const { provider, kind } = await selectProjectProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryProjectProvider);
  });

  it("uses the non-durable in-memory provider when Microsoft is not configured at all", async () => {
    const { provider, kind } = await selectProjectProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryProjectProvider);
  });
});
