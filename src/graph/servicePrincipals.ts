import { graph, graphAll } from './client';
import type {
  AppRoleAssignment,
  OAuth2PermissionGrant,
  ServicePrincipal,
} from './types';

type TokenFn = () => Promise<string>;

export function listServicePrincipals(token: TokenFn, search?: string) {
  const filter = search
    ? `startswith(displayName,'${search.replace(/'/g, "''")}')`
    : undefined;
  return graphAll<ServicePrincipal>(token, '/servicePrincipals', {
    query: {
      $select:
        'id,appId,displayName,servicePrincipalType,accountEnabled,tags,publisherName,appOwnerOrganizationId',
      $top: 100,
      $orderby: 'displayName',
      $filter: filter,
    },
  });
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
