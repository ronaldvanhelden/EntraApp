import { graph, graphAll } from './client';
import { FAVORITE_TAG } from '../lib/tags';
import type { Application, ServicePrincipalSignInActivity } from './types';

type TokenFn = () => Promise<string>;

// Number of objects in a directory collection. /$count returns a bare number
// and needs the eventual-consistency header (but not $count=true, which the
// `advanced` flag would add).
export function countCollection(token: TokenFn, collection: 'applications' | 'servicePrincipals') {
  return graph<number>(token, `/${collection}/$count`, {
    headers: { ConsistencyLevel: 'eventual' },
  });
}

// App registrations carrying the Favorite tag, tenant-wide.
export function listFavoriteApplications(token: TokenFn) {
  return graphAll<Application>(
    token,
    '/applications',
    {
      query: {
        $select: 'id,appId,displayName,info,tags',
        $filter: `tags/any(t:t eq '${FAVORITE_TAG.replace(/'/g, "''")}')`,
        $top: 100,
      },
      advanced: true,
    },
    2,
  );
}

// App registrations with no owners. The owners navigation property is
// expanded per row, so this is capped to keep the Overview responsive.
export async function listOwnerlessApplications(token: TokenFn) {
  const rows = await graphAll<Application & { owners?: { id: string }[] }>(
    token,
    '/applications',
    {
      query: {
        $select: 'id,appId,displayName,info',
        $expand: 'owners($select=id)',
        $top: 100,
      },
    },
    10,
  );
  return rows.filter((r) => !r.owners || r.owners.length === 0);
}

// Last sign-in per appId across the whole tenant, from the (beta)
// service principal sign-in activity report — one paged call instead of one
// per app. Any sign-in counts: delegated or app-only, as client or resource.
// Requires AuditLog.Read.All. Apps that never signed in have no entry.
export async function listLastSignInByAppId(token: TokenFn): Promise<Map<string, number>> {
  const rows = await graphAll<ServicePrincipalSignInActivity>(
    token,
    '/reports/servicePrincipalSignInActivities',
    { api: 'beta', query: { $top: 999 } },
    20,
  );
  const out = new Map<string, number>();
  for (const r of rows) {
    if (!r.appId) continue;
    const t = r.lastSignInActivity?.lastSignInDateTime;
    const ms = t ? new Date(t).getTime() : NaN;
    if (Number.isFinite(ms)) out.set(r.appId, Math.max(ms, out.get(r.appId) ?? 0));
  }
  return out;
}
