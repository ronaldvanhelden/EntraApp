import { graph, graphAll } from './client';
import type { Application } from './types';

type TokenFn = () => Promise<string>;

export interface CreateApplicationInput {
  displayName: string;
  signInAudience?: string;
}

export type UpdateApplicationPatch = Partial<
  Pick<Application, 'displayName' | 'signInAudience' | 'notes' | 'identifierUris'>
>;

export function listApplications(token: TokenFn, search?: string) {
  const filter = search
    ? `startswith(displayName,'${search.replace(/'/g, "''")}')`
    : undefined;
  return graphAll<Application>(token, '/applications', {
    query: {
      $select:
        'id,appId,displayName,createdDateTime,signInAudience,publisherDomain',
      $top: 100,
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

export function createApplication(
  token: TokenFn,
  input: CreateApplicationInput,
) {
  return graph<Application>(token, '/applications', {
    method: 'POST',
    body: input,
  });
}

export function updateApplication(
  token: TokenFn,
  id: string,
  patch: UpdateApplicationPatch,
) {
  return graph<void>(token, `/applications/${id}`, {
    method: 'PATCH',
    body: patch,
  });
}

export function deleteApplication(token: TokenFn, id: string) {
  return graph<void>(token, `/applications/${id}`, { method: 'DELETE' });
}
