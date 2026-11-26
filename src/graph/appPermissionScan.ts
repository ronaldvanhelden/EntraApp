import { graph, graphAll } from './client';
import {
  SCANNED_RESOURCES,
  getDangerousPermission,
  type DangerousPermission,
} from '../lib/dangerousPermissions';
import type { AppRoleAssignment, ServicePrincipal } from './types';

type TokenFn = () => Promise<string>;

// Tenant-wide view of who holds application (app-only) permissions.
//
// Graph has no "all app role assignments" collection, and asking every
// service principal for its own would be one call per principal. The other
// direction is cheap: an API's appRoleAssignedTo lists everyone holding one of
// its roles, so a handful of APIs covers the permissions that matter.
// Application.Read.All is enough for all of it.

export interface HeldPermission {
  assignmentId: string;
  resourceAppId: string;
  resourceName: string;
  // The role's value (e.g. Mail.Send); the role id when the API no longer
  // defines the role.
  value: string;
  displayName?: string;
  grantedDateTime?: string;
  risk: DangerousPermission | null;
}

export interface PermissionHolder {
  id: string;
  displayName: string;
  // Missing when the directory lookup failed; the name above still comes
  // from the assignment.
  sp?: ServicePrincipal;
  permissions: HeldPermission[];
}

export interface ScannedResource {
  appId: string;
  name: string;
  // 'absent': the API has no service principal in this tenant.
  state: 'scanned' | 'absent' | 'failed';
  error?: string;
}

export interface AppPermissionScan {
  holders: PermissionHolder[];
  resources: ScannedResource[];
}

// Enough pages for 20,000 assignments on one API.
const MAX_PAGES = 200;

export async function scanAppPermissions(token: TokenFn): Promise<AppPermissionScan> {
  const apis = await graphAll<ServicePrincipal>(token, '/servicePrincipals', {
    query: {
      $filter: SCANNED_RESOURCES.map((r) => `appId eq '${r.appId}'`).join(' or '),
      $select: 'id,appId,displayName,appRoles',
    },
  });

  const holders = new Map<string, PermissionHolder>();
  const resources = await Promise.all(
    SCANNED_RESOURCES.map(async ({ appId, name }): Promise<ScannedResource> => {
      const api = apis.find((a) => a.appId?.toLowerCase() === appId);
      if (!api) return { appId, name, state: 'absent' };
      let assignments: AppRoleAssignment[];
      try {
        assignments = await graphAll<AppRoleAssignment>(
          token,
          `/servicePrincipals/${api.id}/appRoleAssignedTo`,
          { query: { $top: 100 } },
          MAX_PAGES,
        );
      } catch (e: unknown) {
        return {
          appId,
          name,
          state: 'failed',
          error: e instanceof Error ? e.message : String(e),
        };
      }
      const roles = new Map((api.appRoles ?? []).map((r) => [r.id, r]));
      for (const a of assignments) {
        // Users and groups hold app roles too (that is being assigned to an
        // app); only a service principal holding one is an app-only permission.
        if (a.principalType !== 'ServicePrincipal') continue;
        const role = roles.get(a.appRoleId);
        const value = role?.value || a.appRoleId;
        let holder = holders.get(a.principalId);
        if (!holder) {
          holder = {
            id: a.principalId,
            displayName: a.principalDisplayName || a.principalId,
            permissions: [],
          };
          holders.set(a.principalId, holder);
        }
        holder.permissions.push({
          assignmentId: a.id,
          resourceAppId: appId,
          resourceName: name,
          value,
          displayName: role?.displayName,
          grantedDateTime: a.createdDateTime,
          risk: getDangerousPermission(appId, role?.value),
        });
      }
      return { appId, name, state: 'scanned' };
    }),
  );

  const details = await getServicePrincipalsByIds(token, [...holders.keys()]);
  for (const holder of holders.values()) {
    const sp = details.get(holder.id);
    if (sp) {
      holder.sp = sp;
      holder.displayName = sp.displayName || holder.displayName;
    }
  }
  return { holders: [...holders.values()], resources };
}

// Type, home tenant, enabled state and logo for a set of service principals.
// Best-effort: a batch that fails leaves its principals without details.
async function getServicePrincipalsByIds(
  token: TokenFn,
  ids: string[],
): Promise<Map<string, ServicePrincipal>> {
  const out = new Map<string, ServicePrincipal>();
  // getByIds takes up to 1000 ids.
  const CHUNK = 1000;
  for (let i = 0; i < ids.length; i += CHUNK) {
    try {
      const res = await graph<{ value: ServicePrincipal[] }>(token, '/directoryObjects/getByIds', {
        method: 'POST',
        body: { ids: ids.slice(i, i + CHUNK), types: ['servicePrincipal'] },
      });
      for (const sp of res.value ?? []) out.set(sp.id, sp);
    } catch {
      /* the rows render from the assignment alone */
    }
  }
  return out;
}
