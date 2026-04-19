import { graph, GraphError } from './client';

type TokenFn = () => Promise<string>;

export interface DirectoryAuditInitiator {
  user?: {
    id?: string;
    displayName?: string;
    userPrincipalName?: string;
  };
  app?: {
    appId?: string;
    displayName?: string;
    servicePrincipalId?: string;
  };
}

export interface DirectoryAuditTargetResource {
  id?: string;
  displayName?: string;
  type?: string;
  userPrincipalName?: string;
}

export interface DirectoryAuditEvent {
  id: string;
  activityDateTime?: string;
  activityDisplayName?: string;
  category?: string;
  result?: string;
  initiatedBy?: DirectoryAuditInitiator;
  targetResources?: DirectoryAuditTargetResource[];
}

// Fetches the most recent directory-audit events matching any of the given
// activity display names. Graph retains directory audits for ~30 days.
//
// Requires AuditLog.Read.All. Returns [] (not throws) on 403/404 so the UI
// can render a gentle empty state instead of an error banner when the caller
// lacks the permission or audits have been pruned.
export async function listRecentDirectoryAudits(
  token: TokenFn,
  activities: string[],
  top = 25,
): Promise<DirectoryAuditEvent[]> {
  if (activities.length === 0) return [];
  const filter = activities
    .map((a) => `activityDisplayName eq '${a.replace(/'/g, "''")}'`)
    .join(' or ');
  try {
    const page = await graph<{ value?: DirectoryAuditEvent[] }>(
      token,
      '/auditLogs/directoryAudits',
      {
        query: {
          $filter: filter,
          $top: top,
          $orderby: 'activityDateTime desc',
        },
      },
    );
    return page.value ?? [];
  } catch (e) {
    if (e instanceof GraphError && (e.status === 403 || e.status === 404))
      return [];
    throw e;
  }
}
