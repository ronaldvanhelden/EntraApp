import { graph, graphAll } from './client';
import type { Application } from './types';

type TokenFn = () => Promise<string>;

export function listApplications(token: TokenFn, search?: string) {
  const filter = search
    ? `startswith(displayName,'${search.replace(/'/g, "''")}')`
    : undefined;
  return graphAll<Application>(token, '/applications', {
    query: {
      $select:
        'id,appId,displayName,createdDateTime,signInAudience,publisherDomain',
      $top: 100,
      $orderby: 'displayName',
      $filter: filter,
    },
  });
}

export function getApplication(token: TokenFn, id: string) {
  return graph<Application>(token, `/applications/${id}`);
}

export function getApplicationByAppId(token: TokenFn, appId: string) {
  return graphAll<Application>(token, '/applications', {
    query: { $filter: `appId eq '${appId}'`, $top: 1 },
  }).then((r) => r[0]);
}
