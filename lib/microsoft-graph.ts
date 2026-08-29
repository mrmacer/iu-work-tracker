import { MICROSOFT_GRAPH_BASE_URL } from "./microsoft-auth-config";

export type MicrosoftGraphUser = {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail: string | null;
};

export type SharePointSite = {
  id: string;
  displayName: string;
  webUrl: string;
};

export type SharePointListSummary = {
  id: string;
  name: string;
  displayName: string;
};

export type DevConnectionDiagnostic = {
  user: MicrosoftGraphUser;
  site: SharePointSite;
  lists: SharePointListSummary[];
};

export class GraphRequestError extends Error {
  constructor(
    public readonly status: number,
    message: string,
  ) {
    super(message);
    this.name = "GraphRequestError";
  }
}

type FetchLike = typeof fetch;

async function graphGet<T>(
  path: string,
  accessToken: string,
  fetcher: FetchLike,
): Promise<T> {
  let response: Response;
  try {
    response = await fetcher(`${MICROSOFT_GRAPH_BASE_URL}${path}`, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
    });
  } catch {
    throw new GraphRequestError(0, "Microsoft Graph could not be reached.");
  }

  if (!response.ok) {
    const message =
      response.status === 401
        ? "Microsoft authentication has expired or is invalid."
        : response.status === 403
          ? "The signed-in account is not authorized for this SharePoint resource."
          : "Microsoft Graph returned an unexpected response.";
    throw new GraphRequestError(response.status, message);
  }

  try {
    return (await response.json()) as T;
  } catch {
    throw new GraphRequestError(
      response.status,
      "Microsoft Graph returned an unreadable response.",
    );
  }
}

function requireString(value: unknown, field: string): string {
  if (typeof value !== "string" || value.length === 0) {
    throw new GraphRequestError(200, `Microsoft Graph omitted ${field}.`);
  }
  return value;
}

export async function getGraphUser(
  accessToken: string,
  fetcher: FetchLike = fetch,
): Promise<MicrosoftGraphUser> {
  const value = await graphGet<Record<string, unknown>>(
    "/me?$select=id,displayName,userPrincipalName,mail",
    accessToken,
    fetcher,
  );
  return {
    id: requireString(value.id, "the user ID"),
    displayName: requireString(value.displayName, "the user display name"),
    userPrincipalName: requireString(
      value.userPrincipalName,
      "the user principal name",
    ),
    mail: typeof value.mail === "string" ? value.mail : null,
  };
}

export async function resolveSharePointSite(
  accessToken: string,
  hostname: string,
  sitePath: string,
  fetcher: FetchLike = fetch,
): Promise<SharePointSite> {
  const value = await graphGet<Record<string, unknown>>(
    `/sites/${encodeURIComponent(hostname)}:${sitePath}?$select=id,displayName,webUrl`,
    accessToken,
    fetcher,
  );
  return {
    id: requireString(value.id, "the SharePoint site ID"),
    displayName: requireString(value.displayName, "the SharePoint site name"),
    webUrl: requireString(value.webUrl, "the SharePoint site URL"),
  };
}

export async function getSharePointLists(
  accessToken: string,
  siteId: string,
  fetcher: FetchLike = fetch,
): Promise<SharePointListSummary[]> {
  const result = await graphGet<{ value?: unknown }>(
    `/sites/${encodeURIComponent(siteId)}/lists?$select=id,name,displayName`,
    accessToken,
    fetcher,
  );
  if (!Array.isArray(result.value)) {
    throw new GraphRequestError(200, "Microsoft Graph omitted the SharePoint lists.");
  }
  return result.value.map((item) => {
    const value = item as Record<string, unknown>;
    return {
      id: requireString(value.id, "a SharePoint list ID"),
      name: requireString(value.name, "a SharePoint list name"),
      displayName: requireString(value.displayName, "a SharePoint list display name"),
    };
  });
}

export async function runDevConnectionDiagnostic(
  accessToken: string,
  config: { sharePointHostname: string; sharePointSitePath: string },
  fetcher: FetchLike = fetch,
): Promise<DevConnectionDiagnostic> {
  const user = await getGraphUser(accessToken, fetcher);
  const site = await resolveSharePointSite(
    accessToken,
    config.sharePointHostname,
    config.sharePointSitePath,
    fetcher,
  );
  const lists = await getSharePointLists(accessToken, site.id, fetcher);
  return { user, site, lists };
}
