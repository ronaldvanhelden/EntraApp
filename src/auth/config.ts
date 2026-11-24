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

// Permissions are asked for in steps, so nobody consents to more than they
// use:
//
//  1. Sign-in asks for LOGIN_SCOPES only.
//  2. Looking around needs READ_SCOPES. Right after sign-in the app checks
//     for them silently and, when they are missing, shows one screen with a
//     button that asks for them (a prompt on page load would be blocked).
//  3. Each kind of change asks for its own write permission the first time
//     it is made — always from a click, so the consent popup can open.
//
// All of them should be declared on the app registration (see the README).
export const LOGIN_SCOPES = ['User.Read'];

export const READ_SCOPES = [
  'User.Read',
  'User.Read.All',
  'Group.Read.All',
  'Application.Read.All',
  'DelegatedPermissionGrant.Read.All',
  'CrossTenantInformation.ReadBasic.All',
  'AuditLog.Read.All',
];

// App registrations and enterprise apps: create, edit, delete, credentials,
// owners, schema extensions and restoring deleted apps.
export const APP_WRITE_SCOPES = ['Application.ReadWrite.All'];
// Delegated permission grants (oauth2PermissionGrants).
export const GRANT_WRITE_SCOPES = ['DelegatedPermissionGrant.ReadWrite.All'];
// App role assignments: application permissions and user/group assignments.
export const ROLE_ASSIGN_SCOPES = ['AppRoleAssignment.ReadWrite.All'];

// The permission a Graph request needs beyond READ_SCOPES, judged from its
// method and path, or undefined when reading is enough. The Graph client asks
// this for every call, so no caller has to remember to pass a write token.
// Writes to users and groups are not routed here: their callers pin the token
// to USER_WRITE_SCOPES / GROUP_WRITE_SCOPES themselves.
export function scopesForRequest(method: string, path: string): string[] | undefined {
  if (method === 'GET') return undefined;
  const p = path.replace(/^https:\/\/graph\.microsoft\.com\/[^/]+/, '').split('?')[0];
  // POSTs that only read.
  if (/^\/(\$batch|directoryObjects\/getByIds)\b/.test(p)) return undefined;
  if (/\/appRoleAssign(edTo|ments)\b/.test(p)) return ROLE_ASSIGN_SCOPES;
  if (/^\/oauth2PermissionGrants\b/.test(p)) return GRANT_WRITE_SCOPES;
  if (/^\/(applications|servicePrincipals|schemaExtensions|directory\/deletedItems)\b/.test(p))
    return APP_WRITE_SCOPES;
  return undefined;
}

// Before reading and writing were split, tenants consented to the read-write
// permissions only. Those cover the read-only ones, so a scope set that isn't
// consented yet is retried in this older form before asking for consent.
const LEGACY_SCOPE: Record<string, string> = {
  'Application.Read.All': 'Application.ReadWrite.All',
  'DelegatedPermissionGrant.Read.All': 'DelegatedPermissionGrant.ReadWrite.All',
};

// The older equivalent of a scope set, or null when it has none.
export function legacyScopes(scopes: string[]): string[] | null {
  if (!scopes.some((s) => s in LEGACY_SCOPE)) return null;
  return scopes.map((s) => LEGACY_SCOPE[s] ?? s);
}

// Creating, editing and deleting users and groups: requested the first time
// such a write is made.
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
