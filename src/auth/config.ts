import type { Configuration } from "@azure/msal-browser";
import { LogLevel } from "@azure/msal-browser";

const STORAGE_KEY = "entraapp.authConfig";

export interface AuthConfig {
  clientId: string;
  tenantId: string;
}

// Default to the "Debble EntraApp" multi-tenant app registration.
// Override in Settings for tenants that register their own client.
const DEFAULT_CONFIG: AuthConfig = {
  clientId: "438ccf20-d1a2-4027-baf3-0104d018357e",
  tenantId: "organizations",
};

export function loadAuthConfig(): AuthConfig {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return DEFAULT_CONFIG;
    const parsed = JSON.parse(raw) as Partial<AuthConfig>;
    return {
      clientId: parsed.clientId || DEFAULT_CONFIG.clientId,
      tenantId: parsed.tenantId || DEFAULT_CONFIG.tenantId,
    };
  } catch {
    return DEFAULT_CONFIG;
  }
}

export function saveAuthConfig(config: AuthConfig): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
}

export function clearAuthConfig(): void {
  localStorage.removeItem(STORAGE_KEY);
}

export function isConfigured(config: AuthConfig): boolean {
  return /^[0-9a-f-]{36}$/i.test(config.clientId);
}

// Must match an SPA redirectUri on the app registration exactly. We preserve
// the trailing slash because Vite's dev server (and GitHub Pages) only serve
// the app at the base path *with* the slash — landing at "/EntraApp" hits
// Vite's "did you mean /EntraApp/" page and handleRedirectPromise never runs.
export function computeRedirectUri(): string {
  return window.location.origin + window.location.pathname;
}

export function buildMsalConfig(config: AuthConfig): Configuration {
  const redirectUri = computeRedirectUri();
  return {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId || "common"}`,
      redirectUri,
      postLogoutRedirectUri: redirectUri,
      navigateToLoginRequestUrl: false,
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
    },
    system: {
      loggerOptions: {
        logLevel: LogLevel.Warning,
        piiLoggingEnabled: false,
        loggerCallback: (_level, message) => {
          if (import.meta.env.DEV) console.debug(message);
        },
      },
    },
  };
}

// Incremental consent: sign-in only asks for User.Read so first-run users
// (including non-admins) get a normal user-consent screen. Every other scope
// is requested lazily by the feature that needs it; MSAL caches tokens per
// scope superset so repeat reads don't re-prompt. Admin consent is still
// required for the .All scopes below, but it is now scoped to the moment the
// feature is first used instead of forced at sign-in, and split read vs write
// so read-only admins never have to consent to the write scopes.
export const LOGIN_SCOPES = ["User.Read"];

export const SCOPES = {
  // Read-only app registration / service principal browsing.
  appsRead: ["Application.Read.All"],
  // Create / update / delete app registrations and service principals,
  // including credential operations.
  appsWrite: ["Application.ReadWrite.All"],
  // Principal-picker user search. User-consentable — no admin prompt.
  userSearch: ["User.ReadBasic.All"],
  // Principal-picker group search.
  groupSearch: ["Group.Read.All"],
  // Mixed-type principal resolution (users + groups + SPs) for grant lists.
  directoryRead: ["Directory.Read.All"],
  // Read existing oauth2PermissionGrants (delegated consents) for an SP.
  delegatedGrantsRead: ["DelegatedPermissionGrant.Read.All"],
  // Create / update / delete oauth2PermissionGrants.
  delegatedGrantsWrite: ["DelegatedPermissionGrant.ReadWrite.All"],
  // Read appRoleAssignments and appRoleAssignedTo on a service principal.
  appRolesRead: ["AppRoleAssignment.Read.All"],
  // Create / delete appRoleAssignments and appRoleAssignedTo.
  appRolesWrite: ["AppRoleAssignment.ReadWrite.All"],
  // /auditLogs/* and /reports/appCredentialSignInActivities.
  auditLogs: ["AuditLog.Read.All"],
  // findTenantInformationByTenantId for resolving external tenant names.
  crossTenant: ["CrossTenantInformation.ReadBasic.All"],
} as const;

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
export const GRAPH_BASE_BETA = "https://graph.microsoft.com/beta";
