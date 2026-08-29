import {
  BrowserCacheLocation,
  type Configuration,
} from "@azure/msal-browser";

export const MICROSOFT_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
export const MICROSOFT_GRAPH_SCOPES = [
  "User.Read",
  "Sites.ReadWrite.All",
] as const;

export type DevMicrosoftPublicConfig = {
  clientId: string;
  tenantId: string;
  sharePointHostname: string;
  sharePointSitePath: string;
};

type PublicEnvironment = {
  NEXT_PUBLIC_MS_ENTRA_CLIENT_ID?: string;
  NEXT_PUBLIC_MS_ENTRA_TENANT_ID?: string;
  NEXT_PUBLIC_SHAREPOINT_HOSTNAME?: string;
  NEXT_PUBLIC_SHAREPOINT_SITE_PATH?: string;
};

export type DevMicrosoftConfigResult =
  | { status: "disabled" }
  | { status: "invalid"; message: string }
  | { status: "enabled"; value: DevMicrosoftPublicConfig };

const REQUIRED_ENVIRONMENT_KEYS = [
  "NEXT_PUBLIC_MS_ENTRA_CLIENT_ID",
  "NEXT_PUBLIC_MS_ENTRA_TENANT_ID",
  "NEXT_PUBLIC_SHAREPOINT_HOSTNAME",
  "NEXT_PUBLIC_SHAREPOINT_SITE_PATH",
] as const;

export function buildDevMicrosoftConfig(
  environment: PublicEnvironment,
): DevMicrosoftConfigResult {
  const values = REQUIRED_ENVIRONMENT_KEYS.map((key) =>
    environment[key]?.trim(),
  );

  if (values.every((value) => !value)) {
    return { status: "disabled" };
  }

  const missing = REQUIRED_ENVIRONMENT_KEYS.filter(
    (key) => !environment[key]?.trim(),
  );
  if (missing.length > 0) {
    return {
      status: "invalid",
      message: `Missing public DEV configuration: ${missing.join(", ")}`,
    };
  }

  const [clientId, tenantId, sharePointHostname, sharePointSitePath] = values as [
    string,
    string,
    string,
    string,
  ];

  if (sharePointSitePath !== "/sites/IUWorkTrackerDEV") {
    return {
      status: "invalid",
      message: "Delegated Microsoft access is restricted to the DEV site path.",
    };
  }

  if (!/^[0-9a-f-]{36}$/i.test(clientId) || !/^[0-9a-f-]{36}$/i.test(tenantId)) {
    return {
      status: "invalid",
      message: "The Entra client and tenant identifiers must be GUIDs.",
    };
  }

  if (!/^[a-z0-9.-]+\.sharepoint\.com$/i.test(sharePointHostname)) {
    return {
      status: "invalid",
      message: "The SharePoint hostname is invalid.",
    };
  }

  return {
    status: "enabled",
    value: { clientId, tenantId, sharePointHostname, sharePointSitePath },
  };
}

export function readDevMicrosoftConfig(): DevMicrosoftConfigResult {
  return buildDevMicrosoftConfig({
    NEXT_PUBLIC_MS_ENTRA_CLIENT_ID:
      process.env.NEXT_PUBLIC_MS_ENTRA_CLIENT_ID,
    NEXT_PUBLIC_MS_ENTRA_TENANT_ID:
      process.env.NEXT_PUBLIC_MS_ENTRA_TENANT_ID,
    NEXT_PUBLIC_SHAREPOINT_HOSTNAME:
      process.env.NEXT_PUBLIC_SHAREPOINT_HOSTNAME,
    NEXT_PUBLIC_SHAREPOINT_SITE_PATH:
      process.env.NEXT_PUBLIC_SHAREPOINT_SITE_PATH,
  });
}

export function createMsalConfiguration(
  config: DevMicrosoftPublicConfig,
  origin: string,
): Configuration {
  return {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
      redirectUri: `${origin}/`,
      postLogoutRedirectUri: `${origin}/`,
    },
    cache: {
      cacheLocation: BrowserCacheLocation.SessionStorage,
    },
  };
}
