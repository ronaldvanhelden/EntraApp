import { graph, graphAll } from './client';

type TokenFn = () => Promise<string>;

export interface DirectoryUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail?: string;
}

export interface DirectoryGroup {
  id: string;
  displayName: string;
  mailNickname?: string;
  mail?: string;
  securityEnabled?: boolean;
}

export type PrincipalKind = 'user' | 'group' | 'sp';

export interface PrincipalRef {
  id: string;
  displayName: string;
  kind: PrincipalKind;
  subtitle?: string;
}

// Escape a user query for use inside a $search="…" phrase. Double quotes and
// backslashes are the two literal characters that must be escaped.
function escapeSearch(q: string): string {
  return q.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

export function searchUsers(token: TokenFn, q: string) {
  const s = escapeSearch(q);
  const search =
    `"displayName:${s}" OR "userPrincipalName:${s}" ` +
    `OR "mail:${s}" OR "givenName:${s}" OR "surname:${s}"`;
  return graphAll<DirectoryUser>(
    token,
    '/users',
    {
      query: {
        $select: 'id,displayName,userPrincipalName,mail',
        $top: 50,
        $search: search,
      },
      advanced: true,
    },
    2,
  );
}

export function searchGroups(token: TokenFn, q: string) {
  const s = escapeSearch(q);
  const search = `"displayName:${s}" OR "description:${s}" OR "mail:${s}"`;
  return graphAll<DirectoryGroup>(
    token,
    '/groups',
    {
      query: {
        $select: 'id,displayName,mailNickname,mail,securityEnabled',
        $top: 50,
        $search: search,
      },
      advanced: true,
    },
    2,
  );
}

interface DirectoryObjectResolved {
  id: string;
  '@odata.type'?: string;
  displayName?: string;
  userPrincipalName?: string;
  mail?: string;
  appId?: string;
}

// Resolve a mixed set of principal object IDs to displayName + kind in one call.
export async function resolveDirectoryObjects(
  token: TokenFn,
  ids: string[],
): Promise<Record<string, PrincipalRef>> {
  if (!ids.length) return {};
  const unique = Array.from(new Set(ids));
  const out: Record<string, PrincipalRef> = {};
  // getByIds accepts up to 1000; chunk defensively.
  const CHUNK = 500;
  for (let i = 0; i < unique.length; i += CHUNK) {
    const batch = unique.slice(i, i + CHUNK);
    try {
      const res = await graph<{ value: DirectoryObjectResolved[] }>(
        token,
        '/directoryObjects/getByIds',
        {
          method: 'POST',
          body: {
            ids: batch,
            types: ['user', 'group', 'servicePrincipal'],
          },
        },
      );
      for (const obj of res.value ?? []) {
        const t = (obj['@odata.type'] ?? '').toLowerCase();
        const kind: PrincipalKind = t.includes('user')
          ? 'user'
          : t.includes('group')
            ? 'group'
            : 'sp';
        out[obj.id] = {
          id: obj.id,
          displayName: obj.displayName ?? obj.id,
          kind,
          subtitle: obj.userPrincipalName ?? obj.mail ?? obj.appId,
        };
      }
    } catch {
      /* fall through — caller renders raw id */
    }
  }
  return out;
}
