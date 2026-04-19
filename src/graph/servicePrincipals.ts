import { graph, graphAll } from './client';
import type {
  AppRoleAssignment,
  OAuth2PermissionGrant,
  ServicePrincipal,
} from './types';

type TokenFn = () => Promise<string>;

// A bare hex-and-dashes token — treat as a (possibly partial) appId GUID
// rather than a display-name search term.
const GUID_LIKE = /^[0-9a-f-]+$/i;

export function listServicePrincipals(token: TokenFn, search?: string) {
  const q = search?.trim();
  const $select =
    'id,appId,displayName,servicePrincipalType,accountEnabled,tags,publisherName,appOwnerOrganizationId';

  if (!q) {
    return graphAll<ServicePrincipal>(token, '/servicePrincipals', {
      query: { $select, $top: 100 },
    });
  }

  if (GUID_LIKE.test(q)) {
    return graphAll<ServicePrincipal>(
      token,
      '/servicePrincipals',
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
  return graphAll<ServicePrincipal>(
    token,
    '/servicePrincipals',
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

export function getServicePrincipal(token: TokenFn, id: string) {
  return graph<ServicePrincipal>(token, `/servicePrincipals/${id}`);
}

export function getServicePrincipalByAppId(token: TokenFn, appId: string) {
  return graphAll<ServicePrincipal>(token, '/servicePrincipals', {
    query: { $filter: `appId eq '${appId}'`, $top: 1 },
  }).then((r) => r[0]);
}

// Delegated permission grants (oauth2PermissionGrants) for a given client SP
export function listOAuth2GrantsForClient(token: TokenFn, clientSpId: string) {
  return graphAll<OAuth2PermissionGrant>(token, '/oauth2PermissionGrants', {
    query: { $filter: `clientId eq '${clientSpId}'`, $top: 100 },
  });
}

export function createOAuth2Grant(
  token: TokenFn,
  grant: Omit<OAuth2PermissionGrant, 'id'>,
) {
  return graph<OAuth2PermissionGrant>(token, '/oauth2PermissionGrants', {
    method: 'POST',
    body: grant,
  });
}

export function updateOAuth2Grant(
  token: TokenFn,
  id: string,
  patch: Partial<Pick<OAuth2PermissionGrant, 'scope'>>,
) {
  return graph<void>(token, `/oauth2PermissionGrants/${id}`, {
    method: 'PATCH',
    body: patch,
  });
}

export function deleteOAuth2Grant(token: TokenFn, id: string) {
  return graph<void>(token, `/oauth2PermissionGrants/${id}`, {
    method: 'DELETE',
  });
}

// App-only permission assignments (appRoleAssignments) — assignments TO the SP
export function listAppRoleAssignments(token: TokenFn, spId: string) {
  return graphAll<AppRoleAssignment>(
    token,
    `/servicePrincipals/${spId}/appRoleAssignments`,
    { query: { $top: 100 } },
  );
}

export function createAppRoleAssignment(
  token: TokenFn,
  principalSpId: string,
  body: { appRoleId: string; principalId: string; resourceId: string },
) {
  return graph<AppRoleAssignment>(
    token,
    `/servicePrincipals/${principalSpId}/appRoleAssignments`,
    { method: 'POST', body },
  );
}

export function deleteAppRoleAssignment(
  token: TokenFn,
  principalSpId: string,
  assignmentId: string,
) {
  return graph<void>(
    token,
    `/servicePrincipals/${principalSpId}/appRoleAssignments/${assignmentId}`,
    { method: 'DELETE' },
  );
}

// SP lifecycle
export function createServicePrincipalFromAppId(token: TokenFn, appId: string) {
  return graph<ServicePrincipal>(token, '/servicePrincipals', {
    method: 'POST',
    body: { appId },
  });
}

export function updateServicePrincipal(
  token: TokenFn,
  id: string,
  patch: Partial<Pick<ServicePrincipal, 'accountEnabled' | 'tags'>>,
) {
  return graph<void>(token, `/servicePrincipals/${id}`, {
    method: 'PATCH',
    body: patch,
  });
}

export function deleteServicePrincipal(token: TokenFn, id: string) {
  return graph<void>(token, `/servicePrincipals/${id}`, { method: 'DELETE' });
}

// Assignments TO this SP — who (users/groups/SPs) has been granted one of the
// appRoles that this SP exposes. Distinct from appRoleAssignments (roles this
// SP holds on OTHER resources).
export function listAppRoleAssignedTo(token: TokenFn, spId: string) {
  return graphAll<AppRoleAssignment>(
    token,
    `/servicePrincipals/${spId}/appRoleAssignedTo`,
    { query: { $top: 100 } },
  );
}

export function createAppRoleAssignedTo(
  token: TokenFn,
  resourceSpId: string,
  body: { appRoleId: string; principalId: string; resourceId: string },
) {
  return graph<AppRoleAssignment>(
    token,
    `/servicePrincipals/${resourceSpId}/appRoleAssignedTo`,
    { method: 'POST', body },
  );
}

export function deleteAppRoleAssignedTo(
  token: TokenFn,
  resourceSpId: string,
  assignmentId: string,
) {
  return graph<void>(
    token,
    `/servicePrincipals/${resourceSpId}/appRoleAssignedTo/${assignmentId}`,
    { method: 'DELETE' },
  );
}
