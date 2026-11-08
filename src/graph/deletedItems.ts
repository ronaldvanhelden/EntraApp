import { graph, graphAll, type Paged } from './client';
import { getMe } from './me';
import type { Application, ServicePrincipal } from './types';

type TokenFn = () => Promise<string>;

// Soft-deleted app registrations and service principals live in the directory
// recycle bin for 30 days, then are removed for good. Both object kinds are
// served from /directory/deletedItems, addressed by their OData type.
export const DELETED_RETENTION_DAYS = 30;

export function listDeletedApplications(token: TokenFn) {
  return graphAll<Application>(token, '/directory/deletedItems/microsoft.graph.application', {
    query: { $top: 100 },
  });
}

// Deleted app registrations the signed-in user owned. The recycle bin can't
// be filtered by owner with $filter; this action is the supported route and
// returns at most 1,000 objects in one response.
export async function listDeletedOwnedApplications(token: TokenFn) {
  const me = await getMe(token);
  const res = await graph<Paged<Application>>(
    token,
    '/directory/deletedItems/getUserOwnedObjects',
    {
      method: 'POST',
      body: { userId: me.id, type: 'Application' },
    },
  );
  return res.value;
}

export function listDeletedServicePrincipals(token: TokenFn) {
  return graphAll<ServicePrincipal>(
    token,
    '/directory/deletedItems/microsoft.graph.servicePrincipal',
    { query: { $top: 100 } },
  );
}

// Brings the object back with its original id, credentials and settings.
export function restoreDeletedItem(token: TokenFn, id: string) {
  return graph<{ id: string }>(token, `/directory/deletedItems/${id}/restore`, {
    method: 'POST',
  });
}

// Skips the rest of the retention period. Cannot be undone.
export function permanentlyDeleteItem(token: TokenFn, id: string) {
  return graph<void>(token, `/directory/deletedItems/${id}`, { method: 'DELETE' });
}

// Whole days until the object is purged (0 once the window has passed).
export function daysUntilPurge(deletedDateTime?: string | null): number | undefined {
  if (!deletedDateTime) return undefined;
  const purgeAt = new Date(deletedDateTime).getTime() + DELETED_RETENTION_DAYS * 864e5;
  if (!Number.isFinite(purgeAt)) return undefined;
  return Math.max(0, Math.ceil((purgeAt - Date.now()) / 864e5));
}
