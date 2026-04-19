import { graph, graphAll, GraphError } from './client';
import type {
  Application,
  DirectoryObjectLite,
  KeyCredential,
} from './types';

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

// /applications/{id}/createdOnBehalfOf returns the directory object (usually a
// user) that created the app. $expand of this nav property is rejected on
// v1.0 ("Unsupported link query on Application property"), so we fetch it
// separately. 404 is a normal response for apps created via Graph with an
// app-only token — Entra doesn't track a creator in that case.
export async function getApplicationCreator(
  token: TokenFn,
  applicationObjectId: string,
): Promise<DirectoryObjectLite | null> {
  // Primary source: the nav property. Usually only populated for apps
  // created through the portal with an interactive signed-in user.
  try {
    const direct = await graph<DirectoryObjectLite>(
      token,
      `/applications/${applicationObjectId}/createdOnBehalfOf`,
      { query: { $select: 'id,displayName,userPrincipalName' } },
    );
    if (direct?.id) return direct;
  } catch (e) {
    if (!(e instanceof GraphError) || (e.status !== 404 && e.status !== 403))
      throw e;
  }

  // Fallback: mine the directory audit log for the creation event. This
  // works for apps created through Graph or automation, as long as the
  // caller has AuditLog.Read.All.
  try {
    const safeId = applicationObjectId.replace(/'/g, "''");
    const page = await graph<{
      value?: Array<{
        initiatedBy?: {
          user?: {
            id?: string;
            displayName?: string;
            userPrincipalName?: string;
          };
          app?: {
            appId?: string;
            displayName?: string;
          };
        };
      }>;
    }>(token, '/auditLogs/directoryAudits', {
      query: {
        $filter:
          `activityDisplayName eq 'Add application' and ` +
          `targetResources/any(t:t/id eq '${safeId}')`,
        $top: 1,
        $orderby: 'activityDateTime desc',
      },
    });
    const event = page.value?.[0];
    const user = event?.initiatedBy?.user;
    if (user?.id) {
      return {
        id: user.id,
        displayName: user.displayName ?? null,
        userPrincipalName: user.userPrincipalName ?? null,
      };
    }
    const initApp = event?.initiatedBy?.app;
    if (initApp?.appId) {
      return {
        id: initApp.appId,
        displayName: initApp.displayName
          ? `${initApp.displayName} (app)`
          : 'Service principal',
        userPrincipalName: null,
      };
    }
  } catch (e) {
    // AuditLog.Read.All missing or audits pruned — fall through to null.
    if (!(e instanceof GraphError) || (e.status !== 403 && e.status !== 404))
      throw e;
  }

  return null;
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
