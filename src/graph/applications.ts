import { graph, graphAll } from './client';
import type { Application, KeyCredential } from './types';

type TokenFn = () => Promise<string>;

export interface CreateApplicationInput {
  displayName: string;
  signInAudience?: string;
}

export type UpdateApplicationPatch = Partial<
  Pick<
    Application,
    | 'displayName'
    | 'signInAudience'
    | 'notes'
    | 'identifierUris'
    | 'requiredResourceAccess'
    | 'keyCredentials'
    | 'web'
    | 'spa'
    | 'publicClient'
    | 'isFallbackPublicClient'
  >
>;

// A bare hex-and-dashes token — treat as a (possibly partial) appId GUID
// rather than a display-name search term.
const GUID_LIKE = /^[0-9a-f-]+$/i;

export function listApplications(token: TokenFn, search?: string) {
  const q = search?.trim();
  // passwordCredentials and keyCredentials are inline properties so they ship
  // with the list response at no extra round-trip cost. Federated credentials
  // are a navigation property and require $expand — we include them so the
  // list can surface a FIC badge and expiry warnings per app.
  const $select =
    'id,appId,displayName,createdDateTime,signInAudience,publisherDomain,info,' +
    'passwordCredentials,keyCredentials';
  const $expand = 'federatedIdentityCredentials($select=id)';

  if (!q) {
    return graphAll<Application>(token, '/applications', {
      query: { $select, $expand, $top: 100 },
    });
  }

  if (GUID_LIKE.test(q)) {
    return graphAll<Application>(
      token,
      '/applications',
      {
        query: {
          $select,
          $expand,
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
        $expand,
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
        'id,appId,displayName,createdDateTime,signInAudience,publisherDomain,info,' +
        'notes,identifierUris,requiredResourceAccess,' +
        'passwordCredentials,keyCredentials,' +
        'web,spa,publicClient,isFallbackPublicClient',
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

export interface AddPasswordInput {
  displayName?: string;
  // ISO8601. If omitted, Graph defaults to ~2 years from now.
  endDateTime?: string;
}

// Uses the action endpoint so Entra generates the secret value and returns it
// once (in `secretText`). This value is not retrievable later.
export function addApplicationPassword(
  token: TokenFn,
  applicationObjectId: string,
  input: AddPasswordInput,
) {
  return graph<import('./types').PasswordCredential>(
    token,
    `/applications/${applicationObjectId}/addPassword`,
    {
      method: 'POST',
      body: { passwordCredential: input },
    },
  );
}

export function removeApplicationPassword(
  token: TokenFn,
  applicationObjectId: string,
  keyId: string,
) {
  return graph<void>(
    token,
    `/applications/${applicationObjectId}/removePassword`,
    {
      method: 'POST',
      body: { keyId },
    },
  );
}

export interface AddCertificateInput {
  displayName?: string;
  // base64(DER) X.509 certificate (public portion).
  keyBase64: string;
  // base64(SHA-1 of DER) — Graph's customKeyIdentifier. If omitted, Graph
  // computes it from key automatically, but we usually have it already.
  thumbprintBase64?: string;
  startDateTime?: string;
  endDateTime?: string;
}

// Certificates live on the Application's keyCredentials collection. There's
// no addKey shortcut without proof-of-possession, so we read-modify-write the
// full array via PATCH. Existing entries must be passed back intact (Graph
// preserves their stored public key by keyId match).
export function addApplicationCertificate(
  token: TokenFn,
  applicationObjectId: string,
  existing: KeyCredential[],
  input: AddCertificateInput,
): Promise<void> {
  const newKey: KeyCredential = {
    keyId: crypto.randomUUID(),
    type: 'AsymmetricX509Cert',
    usage: 'Verify',
    displayName: input.displayName ?? null,
    startDateTime: input.startDateTime,
    endDateTime: input.endDateTime,
    customKeyIdentifier: input.thumbprintBase64 ?? null,
    key: input.keyBase64,
  };
  return updateApplication(token, applicationObjectId, {
    keyCredentials: [...existing, newKey],
  });
}

export function removeApplicationCertificate(
  token: TokenFn,
  applicationObjectId: string,
  existing: KeyCredential[],
  keyId: string,
): Promise<void> {
  return updateApplication(token, applicationObjectId, {
    keyCredentials: existing.filter((k) => k.keyId !== keyId),
  });
}
