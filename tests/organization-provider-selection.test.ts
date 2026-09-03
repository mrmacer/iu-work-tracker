// @vitest-environment jsdom
//
// Patch 8E — exercises selectOrganizationProvider()'s browser-context branches directly,
// mirroring tests/project-provider-selection.test.ts / tests/contact-provider-selection.test.ts
// exactly. IU_Organizations is the single authoritative durable-Organizations list, configured
// through the one canonical NEXT_PUBLIC_SHAREPOINT_IU_ORGANIZATIONS_LIST_ID variable — no
// second Organization list environment variable exists or is checked.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { DelegatedSharePointOrganizationProvider, MemoryOrganizationProvider, selectOrganizationProvider } from "../lib/organization-provider";

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

describe("selectOrganizationProvider — canonical env var only", () => {
  it("resolves to the SharePoint provider using NEXT_PUBLIC_SHAREPOINT_IU_ORGANIZATIONS_LIST_ID when configured and already signed in", async () => {
    stubMicrosoftConfig();
    vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_IU_ORGANIZATIONS_LIST_ID", "iu-organizations-list-id");
    mocks.account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
    const { provider, kind } = await selectOrganizationProvider();
    expect(kind).toBe("sharepoint");
    expect(provider).toBeInstanceOf(DelegatedSharePointOrganizationProvider);
  });

  it("does not silently use a durable fallback when configured but not signed in", async () => {
    stubMicrosoftConfig();
    vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_IU_ORGANIZATIONS_LIST_ID", "iu-organizations-list-id");
    mocks.account = null;
    const { provider, kind } = await selectOrganizationProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryOrganizationProvider);
  });

  it("uses the non-durable in-memory provider when Microsoft is not configured at all", async () => {
    const { provider, kind } = await selectOrganizationProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryOrganizationProvider);
  });

  it("never activates SharePoint from an unrelated list-id env var — only the canonical Organizations variable counts", async () => {
    stubMicrosoftConfig();
    vi.stubEnv("NEXT_PUBLIC_SHAREPOINT_IU_PROJECTS_LIST_ID", "iu-projects-list-id");
    mocks.account = { username: "dev@example.edu", name: "Dev User" } as AccountInfo;
    const { provider, kind } = await selectOrganizationProvider();
    expect(kind).toBe("memory");
    expect(provider).toBeInstanceOf(MemoryOrganizationProvider);
  });
});
