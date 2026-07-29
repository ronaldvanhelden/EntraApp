// Tenant pinning via the URL.
//
// The app is served as a single static page with a HashRouter, so the query
// string (?tenant=…) is never touched by in-app navigation — a bookmark like
//
//   https://ronaldvanhelden.github.io/EntraApp/?tenant=debble.com#/applications
//
// keeps pointing at that tenant no matter where the user clicks. The value is
// used as the MSAL authority, so signing in through such a link acquires a
// token for that tenant (as a guest, if that's how the account is invited)
// instead of the signed-in user's home tenant.
//
// The override is mirrored into sessionStorage because MSAL's redirectUri is
// origin + pathname (no query), so a redirect sign-in would otherwise drop the
// parameter half-way through the flow. sessionStorage is per tab, which lines
// up with MSAL's own sessionStorage token cache: a second tab opened from the
// plain URL is back on the saved default tenant.

export const TENANT_PARAM = 'tenant';

const OVERRIDE_KEY = 'entraapp.tenantOverride';
const SHORTCUTS_KEY = 'entraapp.tenantShortcuts';

const GUID_RE = /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i;
const DOMAIN_RE = /^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)+$/i;
const META_TENANTS = ['common', 'organizations', 'consumers'];

// Accepts what Entra accepts as an authority segment: a tenant GUID, a
// verified domain (debble.com, contoso.onmicrosoft.com) or a meta tenant.
export function isValidTenantId(value: string): boolean {
  const v = value.trim();
  if (!v) return false;
  if (META_TENANTS.includes(v.toLowerCase())) return true;
  return GUID_RE.test(v) || DOMAIN_RE.test(v);
}

// A tenant that identifies one directory — as opposed to "common" /
// "organizations", which let the user pick at sign-in time.
export function isSpecificTenant(value: string): boolean {
  return isValidTenantId(value) && !META_TENANTS.includes(value.trim().toLowerCase());
}

function readFromUrl(): string | null {
  // Prefer ?tenant= before the hash (survives navigation), but also accept it
  // inside the hash (#/?tenant=…) since that's an easy shape to hand-type.
  const fromSearch = new URLSearchParams(window.location.search).get(TENANT_PARAM);
  if (fromSearch) return fromSearch;
  const hash = window.location.hash;
  const q = hash.indexOf('?');
  if (q >= 0) {
    return new URLSearchParams(hash.slice(q + 1)).get(TENANT_PARAM);
  }
  return null;
}

// Resolve the tenant this tab is pinned to, if any. Reading the URL also
// latches the value for the rest of the tab's lifetime.
export function resolveTenantOverride(): string | null {
  const fromUrl = readFromUrl();
  if (fromUrl && isValidTenantId(fromUrl)) {
    const value = fromUrl.trim();
    try {
      sessionStorage.setItem(OVERRIDE_KEY, value);
    } catch {
      /* private mode — the URL alone still carries us through popup sign-in */
    }
    return value;
  }
  try {
    const stored = sessionStorage.getItem(OVERRIDE_KEY);
    return stored && isValidTenantId(stored) ? stored : null;
  } catch {
    return null;
  }
}

export function clearTenantOverride(): void {
  try {
    sessionStorage.removeItem(OVERRIDE_KEY);
  } catch {
    /* ignore */
  }
}

// URL without ?tenant= — i.e. the app on its saved default tenant.
export function buildDefaultUrl(hashPath = '/'): string {
  return window.location.origin + window.location.pathname + normalizeHash(hashPath);
}

// Bookmarkable URL that pins the app to one tenant.
export function buildTenantUrl(tenantId: string, hashPath = '/'): string {
  const params = new URLSearchParams({ [TENANT_PARAM]: tenantId.trim() });
  return `${window.location.origin}${window.location.pathname}?${params}${normalizeHash(hashPath)}`;
}

function normalizeHash(hashPath: string): string {
  if (!hashPath || hashPath === '#') return '#/';
  const path = hashPath.startsWith('#') ? hashPath.slice(1) : hashPath;
  return `#${path.startsWith('/') ? path : `/${path}`}`;
}

// Drop the pin and reload on the clean URL. A reload is required because MSAL
// has to be rebuilt against the default authority.
export function exitTenantOverride(hashPath = '/'): void {
  clearTenantOverride();
  if (window.location.search) {
    // Dropping the query string is a full navigation, which is what rebuilds
    // MSAL against the default authority.
    window.location.replace(buildDefaultUrl(hashPath));
    return;
  }
  // Nothing to strip from the query (the pin came from sessionStorage or from
  // inside the hash): land on the target view and reload explicitly.
  window.location.hash = normalizeHash(hashPath);
  window.location.reload();
}

export interface TenantShortcut {
  label: string;
  tenantId: string;
}

export function loadTenantShortcuts(): TenantShortcut[] {
  try {
    const raw = localStorage.getItem(SHORTCUTS_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];
    return parsed
      .filter(
        (s): s is TenantShortcut =>
          typeof s === 'object' &&
          s !== null &&
          typeof (s as TenantShortcut).tenantId === 'string' &&
          isValidTenantId((s as TenantShortcut).tenantId),
      )
      .map((s) => ({ label: String(s.label ?? s.tenantId), tenantId: s.tenantId }));
  } catch {
    return [];
  }
}

export function saveTenantShortcuts(shortcuts: TenantShortcut[]): void {
  localStorage.setItem(SHORTCUTS_KEY, JSON.stringify(shortcuts));
}
