// Microsoft publishes a rich metadata catalog for every Graph permission, with
// a numeric `privilegeLevel` (1 = low, 4 = highest) plus short admin/user
// descriptions. We fetch it once per session, strip it down to the fields we
// render, and cache the reduced form in sessionStorage.
//
// This catalog only covers Microsoft Graph scopes (resourceAppId
// 00000003-0000-0000-c000-000000000000). For any other API the helpers return
// `null` so the UI can render without a privilege badge.

const CATALOG_URL =
  'https://raw.githubusercontent.com/microsoftgraph/microsoft-graph-devx-content/refs/heads/master/permissions/new/permissions.json';
const CACHE_KEY = 'entraapp.scopeCatalog.v1';
export const MS_GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';

export type ScopeKind = 'delegated' | 'application';

export interface ScopeMeta {
  delegated?: SchemeMeta;
  application?: SchemeMeta;
  // API paths this permission unlocks. Pre-filtered by scheme so the UI can
  // render the delegated vs application lists independently.
  pathSets?: {
    kinds: ScopeKind[];
    methods: string[];
    paths: string[];
  }[];
}
export interface SchemeMeta {
  privilegeLevel?: number;
  requiresAdminConsent: boolean;
  adminDescription?: string;
  userDescription?: string;
}

type RawScheme = {
  privilegeLevel?: number;
  requiresAdminConsent?: boolean;
  adminDescription?: string;
  userDescription?: string;
};
interface RawPathSet {
  schemeKeys?: string[];
  methods?: string[];
  paths?: Record<string, string>;
}
interface RawPermission {
  schemes?: Record<string, RawScheme>;
  pathSets?: RawPathSet[];
}
interface RawCatalog {
  permissions?: Record<string, RawPermission>;
}

type Catalog = Record<string, ScopeMeta>;

let memoryCache: Catalog | null = null;
let inFlight: Promise<Catalog> | null = null;

function reduce(raw: RawCatalog): Catalog {
  const out: Catalog = {};
  const perms = raw.permissions ?? {};
  for (const [name, entry] of Object.entries(perms)) {
    const schemes = entry.schemes ?? {};
    const meta: ScopeMeta = {};
    // DelegatedWork + DelegatedPersonal can coexist; DelegatedWork is the one
    // that matters in tenant-scoped administration.
    const delegated = schemes.DelegatedWork ?? schemes.DelegatedPersonal;
    if (delegated) {
      meta.delegated = {
        privilegeLevel: delegated.privilegeLevel,
        requiresAdminConsent: !!delegated.requiresAdminConsent,
        adminDescription: delegated.adminDescription,
        userDescription: delegated.userDescription,
      };
    }
    const app = schemes.Application;
    if (app) {
      meta.application = {
        privilegeLevel: app.privilegeLevel,
        requiresAdminConsent: !!app.requiresAdminConsent,
        adminDescription: app.adminDescription,
      };
    }
    if (entry.pathSets?.length) {
      meta.pathSets = entry.pathSets
        .map((p) => {
          const kinds: ScopeKind[] = [];
          if (p.schemeKeys?.includes('DelegatedWork')) kinds.push('delegated');
          if (p.schemeKeys?.includes('Application')) kinds.push('application');
          const paths = Object.keys(p.paths ?? {});
          return {
            kinds,
            methods: p.methods ?? [],
            paths,
          };
        })
        .filter((p) => p.kinds.length && p.paths.length);
    }
    out[name.toLowerCase()] = meta;
  }
  return out;
}

export async function loadScopeCatalog(): Promise<Catalog> {
  if (memoryCache) return memoryCache;
  if (inFlight) return inFlight;

  try {
    const cached = sessionStorage.getItem(CACHE_KEY);
    if (cached) {
      memoryCache = JSON.parse(cached) as Catalog;
      return memoryCache;
    }
  } catch {
    /* sessionStorage unavailable — fall through to fetch */
  }

  inFlight = (async () => {
    const res = await fetch(CATALOG_URL, {
      // The catalog doesn't need a Graph token; it's a public GitHub raw file.
      credentials: 'omit',
    });
    if (!res.ok) throw new Error(`Catalog fetch failed: HTTP ${res.status}`);
    const raw = (await res.json()) as RawCatalog;
    const reduced = reduce(raw);
    memoryCache = reduced;
    try {
      sessionStorage.setItem(CACHE_KEY, JSON.stringify(reduced));
    } catch {
      /* quota exceeded — memory cache is still valid */
    }
    return reduced;
  })();

  try {
    return await inFlight;
  } finally {
    inFlight = null;
  }
}

export function getScopeMeta(
  catalog: Catalog | null,
  resourceAppId: string | undefined,
  scope: string | undefined,
): ScopeMeta | null {
  if (!catalog || !scope) return null;
  if (resourceAppId && resourceAppId !== MS_GRAPH_APP_ID) return null;
  return catalog[scope.toLowerCase()] ?? null;
}

export function getPrivilegeLevel(
  meta: ScopeMeta | null,
  kind: ScopeKind,
): number | undefined {
  if (!meta) return undefined;
  return kind === 'delegated'
    ? meta.delegated?.privilegeLevel
    : meta.application?.privilegeLevel;
}

// Resource-Specific Consent scopes aren't marked in the catalog by scheme
// name — the three schemes are DelegatedWork / DelegatedPersonal / Application.
// Microsoft's naming convention for RSC is a suffix that names the container
// the consent applies to: `.Group`, `.Chat`, or `.Team`. Detection by suffix
// is fragile in theory but matches Microsoft's documented RSC list today.
export function isRscScope(scope: string): boolean {
  return /\.(Group|Chat|Team)$/i.test(scope);
}

export interface EndpointMatch {
  scope: string;
  meta: ScopeMeta;
  methods: string[];
  matchedPath: string;
  kinds: ScopeKind[];
  isRsc: boolean;
}

// Find every scope whose pathSets match the caller-supplied method + path.
// Path matching converts catalog templates like `/users/{id}/messages` into
// regexes where `{id}` matches any non-slash segment. User input gets
// normalized (strip query string, lowercase, drop trailing slash) before
// comparison.
export function findScopesForEndpoint(
  catalog: Catalog | null,
  method: string,
  path: string,
): EndpointMatch[] {
  if (!catalog) return [];
  const normMethod = method.toUpperCase();
  const normPath = normalizeEndpointPath(path);
  if (!normPath) return [];

  // Many Graph endpoints are catalogued only under `/users/{id}/...` while
  // users type the equivalent `/me/...` alias. Try both shapes so a query
  // for `/me/messages` still surfaces the `User.Read`-family scopes. (And
  // vice-versa for anyone typing the templated path.)
  const pathCandidates = new Set<string>([normPath]);
  if (normPath === '/me' || normPath.startsWith('/me/')) {
    pathCandidates.add('/users/{id}' + normPath.slice(3));
  }
  if (normPath === '/users/{id}' || normPath.startsWith('/users/{id}/')) {
    pathCandidates.add('/me' + normPath.slice('/users/{id}'.length));
  }

  const out: EndpointMatch[] = [];
  for (const [scope, meta] of Object.entries(catalog)) {
    if (!meta.pathSets?.length) continue;
    for (const set of meta.pathSets) {
      if (!set.methods.includes(normMethod)) continue;
      const matched = set.paths.find((p) =>
        [...pathCandidates].some((c) => pathMatches(p, c)),
      );
      if (!matched) continue;
      out.push({
        scope,
        meta,
        methods: set.methods,
        matchedPath: matched,
        kinds: set.kinds,
        isRsc: isRscScope(scope),
      });
      break; // one match per scope is enough
    }
  }
  // Sort: RSC last, then lowest privilege first, then alphabetically.
  return out.sort((a, b) => {
    if (a.isRsc !== b.isRsc) return a.isRsc ? 1 : -1;
    const ap =
      a.meta.delegated?.privilegeLevel ??
      a.meta.application?.privilegeLevel ??
      99;
    const bp =
      b.meta.delegated?.privilegeLevel ??
      b.meta.application?.privilegeLevel ??
      99;
    if (ap !== bp) return ap - bp;
    return a.scope.localeCompare(b.scope);
  });
}

function normalizeEndpointPath(raw: string): string {
  let p = (raw || '').trim();
  if (!p) return '';
  // Strip fully-qualified Graph URLs, version prefix, and query string so the
  // comparison always happens against the relative resource path.
  p = p.replace(/^https?:\/\/graph\.microsoft\.com/i, '');
  p = p.replace(/^\/(v1\.0|beta)(?=\/)/, '');
  p = p.split('?')[0];
  if (!p.startsWith('/')) p = '/' + p;
  if (p.length > 1 && p.endsWith('/')) p = p.slice(0, -1);
  return p.toLowerCase();
}

function pathMatches(template: string, input: string): boolean {
  // The catalog stores templates already lowercased with `{id}` segments.
  // Build a regex from each `{id}` → `[^/]+` while escaping everything else.
  const parts = template.split(/(\{[^}]+\})/g);
  const pattern = parts
    .map((p) =>
      /^\{[^}]+\}$/.test(p) ? '[^/]+' : p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
    )
    .join('');
  return new RegExp(`^${pattern}$`, 'i').test(input);
}

// Microsoft's Graph permissions reference groups privilegeLevel numerically.
// We label for humans: 1=Low, 2=Medium, 3=High, 4=Critical. Anything outside
// that range is rendered as a plain number.
export function privilegeLabel(level: number | undefined): {
  label: string;
  tone: 'low' | 'medium' | 'high' | 'critical' | 'unknown';
} {
  switch (level) {
    case 1:
      return { label: 'Low', tone: 'low' };
    case 2:
      return { label: 'Medium', tone: 'medium' };
    case 3:
      return { label: 'High', tone: 'high' };
    case 4:
      return { label: 'Critical', tone: 'critical' };
    case undefined:
      return { label: '—', tone: 'unknown' };
    default:
      return { label: `L${level}`, tone: 'unknown' };
  }
}
