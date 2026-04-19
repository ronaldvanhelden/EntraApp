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

// A bare hex-and-dashes token — treat as a (possibly partial) appId GUID
// rather than a display-name search term.
const GUID_LIKE = /^[0-9a-f-]+$/i;

export function listApplications(token: TokenFn, search?: string) {
  const q = search?.trim();
  const $select =
    'id,appId,displayName,createdDateTime,signInAudience,publisherDomain';

  if (!q) {
    return graphAll<Application>(token, '/applications', {
      query: { $select, $top: 100 },
    });
  }

  if (GUID_LIKE.test(q)) {
    return graphAll<Application>(
      token,
      '/applications',
      {
        query: {
          $select,
          $top: 100,
          $filter: `startswith(appId,'${q.replace(/'/g, "''")}')`,
        },
        advanced: true,
      },
      3,
    );
  }

  const s = q.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  return graphAll<Application>(
    token,
    '/applications',
    {
      query: {
        $select,
        $top: 100,
        $search: `"displayName:${s}"`,
      },
      advanced: true,
    },
    3,
  );
}

export function getApplication(token: TokenFn, id: string) {
  return graph<Application>(token, `/applications/${id}`, {
    query: {
      $select:
        'id,appId,displayName,createdDateTime,signInAudience,publisherDomain,' +
        'notes,identifierUris,requiredResourceAccess,' +
        'passwordCredentials,keyCredentials',
    },
  });
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
