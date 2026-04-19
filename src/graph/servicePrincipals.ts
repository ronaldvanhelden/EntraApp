import { graph, graphAll, GraphError } from './client';
import type {
  AppRoleAssignment,
  DirectoryObjectLite,
  OAuth2PermissionGrant,
  ServicePrincipal,
  ServicePrincipalSignInActivity,
} from './types';

type TokenFn = () => Promise<string>;

// A bare hex-and-dashes token — treat as a (possibly partial) appId GUID
// rather than a display-name search term.
const GUID_LIKE = /^[0-9a-f-]+$/i;

export function listServicePrincipals(token: TokenFn, search?: string) {
  const q = search?.trim();
  const $select =
    'id,appId,displayName,servicePrincipalType,accountEnabled,tags,publisherName,appOwnerOrganizationId,info';

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
  // Default projection omits credential collections — explicitly request them
  // so the Overview card can surface secrets/certs on SPs that manage their
  // own (SAML signing certs, federated SaaS apps, etc).
  return graph<ServicePrincipal>(token, `/servicePrincipals/${id}`, {
    query: {
      $select:
        'id,appId,displayName,servicePrincipalType,accountEnabled,appRoles,' +
        'oauth2PermissionScopes,tags,publisherName,appOwnerOrganizationId,' +
        'info,keyCredentials,passwordCredentials,createdDateTime',
    },
  });
}

// Mines the directory audit log for the SP's creation event. Works for SPs
// created via portal, Graph, automation — anything that leaves an audit
// trail. Requires AuditLog.Read.All. Returns a best-effort
// { who, when } — either field can be missing if the audit was pruned.
export interface SpCreationRecord {
  when?: string;
  who: DirectoryObjectLite | null;
}

export async function getServicePrincipalCreation(
  token: TokenFn,
  applicationObjectId: string,
): Promise<SpCreationRecord | null> {
  const safeId = applicationObjectId.replace(/'/g, "''");
  try {
    // Microsoft emits the activity under several names depending on how the
    // SP was created. "Add service principal" covers interactive + Graph.
    const page = await graph<{
      value?: Array<{
        activityDateTime?: string;
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
          `activityDisplayName eq 'Add service principal' and ` +
          `targetResources/any(t:t/id eq '${safeId}')`,
        $top: 1,
        $orderby: 'activityDateTime desc',
      },
    });
    const event = page.value?.[0];
    if (!event) return null;

    let who: DirectoryObjectLite | null = null;
    const user = event.initiatedBy?.user;
    if (user?.id) {
      who = {
        id: user.id,
        displayName: user.displayName ?? null,
        userPrincipalName: user.userPrincipalName ?? null,
      };
    } else {
      const initApp = event.initiatedBy?.app;
      if (initApp?.appId) {
        who = {
          id: initApp.appId,
          displayName: initApp.displayName
            ? `${initApp.displayName} (app)`
            : 'Service principal',
          userPrincipalName: null,
        };
      }
    }
    return { when: event.activityDateTime, who };
  } catch (e) {
    if (e instanceof GraphError && (e.status === 403 || e.status === 404))
      return null;
    throw e;
  }
}

export function getServicePrincipalByAppId(token: TokenFn, appId: string) {
  return graphAll<ServicePrincipal>(token, '/servicePrincipals', {
    query: { $filter: `appId eq '${appId}'`, $top: 1 },
  }).then((r) => r[0]);
}

// /beta report. Requires AuditLog.Read.All. Keyed by the app registration's
// appId (not the service principal object id).
export function getServicePrincipalSignInActivity(
  token: TokenFn,
  appId: string,
) {
  return graph<ServicePrincipalSignInActivity>(
    token,
    `/reports/servicePrincipalSignInActivities/${appId}`,
    { api: 'beta' },
  );
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
