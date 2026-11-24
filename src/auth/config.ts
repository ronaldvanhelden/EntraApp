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
  clientId: '438ccf20-d1a2-4027-baf3-0104d018357e',
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

// Must match an SPA redirectUri on the app registration exactly, e.g.
// https://entra.2cloud.nl/ in production. We preserve the trailing slash
// because Vite's dev server (and GitHub Pages) only serve the app at the base
// path *with* the slash — with a sub-path base, landing at "/EntraApp" hits
// Vite's "did you mean /EntraApp/" page and handleRedirectPromise never runs.
export function computeRedirectUri(): string {
  return window.location.origin + window.location.pathname;
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
  'DelegatedPermissionGrant.ReadWrite.All',
  'AppRoleAssignment.ReadWrite.All',
  'CrossTenantInformation.ReadBasic.All',
  'AuditLog.Read.All',
];

// Creating, editing and deleting users and groups. Kept out of GRAPH_SCOPES so
// sign-in keeps working on app registrations that don't declare them: they
// are requested (incremental consent) the first time such a write is made.
export const USER_WRITE_SCOPES = ['User.ReadWrite.All'];
export const GROUP_WRITE_SCOPES = ['Group.ReadWrite.All'];
// Setting or resetting a user's password: the permission Graph asks for to
// write passwordProfile. Requested the first time a password is reset.
export const USER_PASSWORD_SCOPES = ['User-PasswordProfile.ReadWrite.All'];
// Sending a guest their invitation again.
export const USER_INVITE_SCOPES = ['User.Invite.All'];
// Read-only extras on the user page, each tried silently and only asked for
// when the admin clicks: directory roles (active and PIM-eligible) and the
// Identity Protection risk state.
export const ROLE_READ_SCOPES = ['RoleManagement.Read.Directory'];
export const RISK_READ_SCOPES = ['IdentityRiskyUser.Read.All'];
// The SharePoint site behind a Microsoft 365 group. Also requested on demand:
// it is tried silently, and consent is only asked for when the user clicks.
export const SITE_READ_SCOPES = ['Sites.Read.All'];

export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
export const GRAPH_BASE_BETA = 'https://graph.microsoft.com/beta';
