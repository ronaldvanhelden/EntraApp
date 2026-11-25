import { graph, graphAll, GraphError } from './client';

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
  // What the event changed on this target. Values are JSON-encoded strings.
  modifiedProperties?: DirectoryAuditModifiedProperty[];
}

export interface DirectoryAuditModifiedProperty {
  displayName?: string;
  oldValue?: string | null;
  newValue?: string | null;
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
    if (e instanceof GraphError && (e.status === 403 || e.status === 404)) return [];
    throw e;
  }
}

const OBJECT_AUDIT_TOP = 100;
const OBJECT_AUDIT_PAGES = 2;

export interface ObjectAudits {
  events: DirectoryAuditEvent[];
  // True when an object had more events than were fetched.
  more: boolean;
}

// Every audit event that names one of the given objects as a target, newest
// first: who changed it, when, and what the old and new values were. One
// request per object (the log's filter doesn't reliably take an "or" of
// lambdas); an event that names several of them is listed once.
//
// Requires AuditLog.Read.All and a role that can read audit logs. Returns null
// when the log can't be read.
export async function listObjectAudits(
  token: TokenFn,
  objectIds: string[],
): Promise<ObjectAudits | null> {
  try {
    const lists = await Promise.all(
      objectIds.map((id) =>
        graphAll<DirectoryAuditEvent>(
          token,
          '/auditLogs/directoryAudits',
          {
            query: {
              $filter: `targetResources/any(t:t/id eq '${id}')`,
              $top: OBJECT_AUDIT_TOP,
              $orderby: 'activityDateTime desc',
            },
          },
          OBJECT_AUDIT_PAGES,
        ),
      ),
    );
    const seen = new Set<string>();
    const events = lists
      .flat()
      .filter((e) => !seen.has(e.id) && seen.add(e.id))
      .sort((a, b) => (b.activityDateTime ?? '').localeCompare(a.activityDateTime ?? ''));
    return {
      events,
      more: lists.some((l) => l.length >= OBJECT_AUDIT_TOP * OBJECT_AUDIT_PAGES),
    };
  } catch (e) {
    if (e instanceof GraphError && (e.status === 403 || e.status === 404)) return null;
    throw e;
  }
}
