import type { Configuration } from '@azure/msal-browser';
import { LogLevel } from '@azure/msal-browser';

const STORAGE_KEY = 'entraapp.authConfig';

export interface AuthConfig {
  clientId: string;
  tenantId: string;
}

// Default to the "Debble EntraApp" multi-tenant app registration.
// Override in Settings for tenants that register their own client.
const DEFAULT_CONFIG: AuthConfig = {
  clientId: '285e9c19-6642-4f20-af26-489b47636cc9',
  tenantId: 'organizations',
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

// Normalize to the URL MSAL will compare against the app registration's
// SPA redirectUris — strip the trailing slash so "/EntraApp/" and
// "/EntraApp" both reduce to the same canonical URI.
export function computeRedirectUri(): string {
  const raw = window.location.origin + window.location.pathname;
  return raw.length > 1 ? raw.replace(/\/$/, '') : raw;
}

export function buildMsalConfig(config: AuthConfig): Configuration {
  const redirectUri = computeRedirectUri();
  return {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId || 'common'}`,
      redirectUri,
      postLogoutRedirectUri: redirectUri,
      navigateToLoginRequestUrl: false,
    },
    cache: {
      cacheLocation: 'sessionStorage',
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

// Scopes must match those declared on the app registration (Debble EntraApp).
// Directory.AccessAsUser.All lets the app act as the signed-in user, which
// covers creating appRoleAssignments on service principals for admins;
// DelegatedPermissionGrant.ReadWrite.All is used for oauth2PermissionGrants.
export const GRAPH_SCOPES = [
  'User.Read',
  'User.Read.All',
  'Group.Read.All',
  'Application.ReadWrite.All',
  'Directory.AccessAsUser.All',
  'DelegatedPermissionGrant.ReadWrite.All',
  'AppRoleAssignment.ReadWrite.All',
  'CrossTenantInformation.ReadBasic.All',
  'AuditLog.Read.All',
];

export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
export const GRAPH_BASE_BETA = 'https://graph.microsoft.com/beta';
