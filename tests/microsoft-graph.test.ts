import { describe, expect, it, vi } from "vitest";
import {
  getGraphUser,
  getSharePointLists,
  GraphRequestError,
  resolveSharePointSite,
  runDevConnectionDiagnostic,
} from "../lib/microsoft-graph";

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

describe("Microsoft Graph DEV diagnostics", () => {
  it("loads the signed-in user through /me", async () => {
    const fetcher = vi.fn(async () =>
      jsonResponse({
        id: "user-id",
        displayName: "IU Dev User",
        userPrincipalName: "dev@iu29.example",
        mail: "dev@iu29.example",
      }),
    ) as unknown as typeof fetch;

    await expect(getGraphUser("token", fetcher)).resolves.toMatchObject({
      id: "user-id",
      displayName: "IU Dev User",
    });
    expect(fetcher).toHaveBeenCalledWith(
      "https://graph.microsoft.com/v1.0/me?$select=id,displayName,userPrincipalName,mail",
      expect.objectContaining({
        method: "GET",
        headers: expect.objectContaining({ Authorization: "Bearer token" }),
      }),
    );
  });

  it("resolves the configured SharePoint site and preserves its opaque ID", async () => {
    const fetcher = vi.fn(async () =>
      jsonResponse({
        id: "siu29.sharepoint.com,opaque-site,opaque-web",
        displayName: "IU Work Tracker DEV",
        webUrl: "https://siu29.sharepoint.com/sites/IUWorkTrackerDEV",
      }),
    ) as unknown as typeof fetch;

    const site = await resolveSharePointSite(
      "token",
      "siu29.sharepoint.com",
      "/sites/IUWorkTrackerDEV",
      fetcher,
    );
    expect(site.id).toBe("siu29.sharepoint.com,opaque-site,opaque-web");
    expect(fetcher).toHaveBeenCalledWith(
      "https://graph.microsoft.com/v1.0/sites/siu29.sharepoint.com:/sites/IUWorkTrackerDEV?$select=id,displayName,webUrl",
      expect.objectContaining({ method: "GET" }),
    );
  });

  it("reads existing lists without issuing a write request", async () => {
    const fetcher = vi.fn(async () =>
      jsonResponse({
        value: [
          { id: "list-1", name: "Documents", displayName: "Documents" },
        ],
      }),
    ) as unknown as typeof fetch;

    await expect(
      getSharePointLists("token", "opaque,site,id", fetcher),
    ).resolves.toHaveLength(1);
    expect(fetcher).toHaveBeenCalledWith(
      "https://graph.microsoft.com/v1.0/sites/opaque%2Csite%2Cid/lists?$select=id,name,displayName",
      expect.objectContaining({ method: "GET" }),
    );
  });

  it("runs /me, site resolution, and list read in sequence", async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({
          id: "user-id",
          displayName: "IU Dev User",
          userPrincipalName: "dev@iu29.example",
          mail: null,
        }),
      )
      .mockResolvedValueOnce(
        jsonResponse({
          id: "site-id",
          displayName: "IU Work Tracker DEV",
          webUrl: "https://siu29.sharepoint.com/sites/IUWorkTrackerDEV",
        }),
      )
      .mockResolvedValueOnce(
        jsonResponse({ value: [{ id: "1", name: "Docs", displayName: "Docs" }] }),
      ) as unknown as typeof fetch;

    const result = await runDevConnectionDiagnostic(
      "token",
      {
        sharePointHostname: "siu29.sharepoint.com",
        sharePointSitePath: "/sites/IUWorkTrackerDEV",
      },
      fetcher,
    );
    expect(result).toMatchObject({
      user: { id: "user-id" },
      site: { id: "site-id" },
      lists: [{ id: "1" }],
    });
    expect(fetcher).toHaveBeenCalledTimes(3);
  });

  it("handles an authentication failure without exposing the Graph body", async () => {
    const fetcher = vi.fn(async () =>
      jsonResponse({ error: { message: "sensitive detail" } }, 401),
    ) as unknown as typeof fetch;

    await expect(getGraphUser("expired", fetcher)).rejects.toEqual(
      new GraphRequestError(
        401,
        "Microsoft authentication has expired or is invalid.",
      ),
    );
  });

  it("handles a Graph permission failure explicitly", async () => {
    const fetcher = vi.fn(async () => jsonResponse({ error: "denied" }, 403)) as unknown as typeof fetch;

    await expect(
      resolveSharePointSite(
        "token",
        "siu29.sharepoint.com",
        "/sites/IUWorkTrackerDEV",
        fetcher,
      ),
    ).rejects.toEqual(
      new GraphRequestError(
        403,
        "The signed-in account is not authorized for this SharePoint resource.",
      ),
    );
  });
});
