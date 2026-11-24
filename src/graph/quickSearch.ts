import { graph, GraphError, type Paged } from './client';
import { buildQuery } from '../lib/searchQuery';
import { toStoredTag } from '../lib/tags';
import type { Application, ServicePrincipal } from './types';

type TokenFn = () => Promise<string>;

// Type-ahead search for the Overview: one small page per query instead of the
// full, credential-laden projection the list pages load. Matching follows
// those lists — display name (substring), appId (prefix) and, for app
// registrations, unique name and tag label (prefix). Users and groups match
// on name and mail (users also on UPN), or on a complete object id.

export interface QuickSearchResult<T> {
  items: T[];
  // More matches exist than were fetched.
  more: boolean;
}

// How much to fetch: a single small page for the type-ahead, a few big ones
// for the search results page.
export interface SearchDepth {
  top: number;
  pages: number;
}
const QUICK: SearchDepth = { top: 15, pages: 1 };
export const FULL_SEARCH: SearchDepth = { top: 100, pages: 3 };

const GUID_LIKE = /^[0-9a-f-]+$/i;
const FULL_GUID = /^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i;
const APP_SELECT = 'id,appId,displayName,uniqueName,info,tags,signInAudience,createdDateTime';
const SP_SELECT =
  'id,appId,displayName,servicePrincipalType,accountEnabled,publisherName,appOwnerOrganizationId,info';

// The name, department and job title fields are only there for the query
// language (searchQuery.ts), which filters on them.
const USER_SELECT =
  'id,displayName,userPrincipalName,mail,userType,accountEnabled,createdDateTime,' +
  'givenName,surname,department,jobTitle';
const GROUP_SELECT =
  'id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,createdDateTime,' +
  'mailNickname';

const odata = (v: string) => v.replace(/'/g, "''");
const searchTerms = (q: string, props: string[]) => {
  const s = q.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  return props.map((p) => `"${p}:${s}"`).join(' OR ');
};
const searchTerm = (q: string) => searchTerms(q, ['displayName']);

// Users and groups: an empty query lists alphabetically (the Users / Groups
// pages), a complete object id looks that one up, anything else searches.
// $orderby can't be combined with $search, so matches are ranked client-side.
// `filter` narrows any of the three (e.g. to guests).
function principalQuery(
  select: string,
  q: string,
  props: string[],
  filter?: string,
): Record<string, string> {
  const narrowed: Record<string, string> = filter ? { $filter: filter } : {};
  if (!q) return { $select: select, $orderby: 'displayName', $count: 'true', ...narrowed };
  if (FULL_GUID.test(q))
    return { $select: select, $filter: `id eq '${q}'${filter ? ` and ${filter}` : ''}` };
  return { $select: select, $search: searchTerms(q, props), ...narrowed };
}

async function page<T>(
  token: TokenFn,
  path: string,
  query: Record<string, string | number>,
  depth: SearchDepth,
): Promise<QuickSearchResult<T>> {
  const items: T[] = [];
  let next: string | undefined;
  for (let i = 0; i < depth.pages; i++) {
    const res: Paged<T> = await graph<Paged<T>>(
      token,
      next ?? path,
      next ? { advanced: true } : { query: { ...query, $top: depth.top }, advanced: true },
    );
    items.push(...res.value);
    next = res['@odata.nextLink'];
    if (!next) break;
  }
  return { items, more: Boolean(next) };
}

// Names starting with the query first, then names containing it, then the
// rest (matched on appId, unique name or tag) — alphabetical within each.
function rank<T extends { displayName?: string }>(items: T[], q: string): T[] {
  const needle = q.toLowerCase();
  const score = (i: T) => {
    const n = (i.displayName ?? '').toLowerCase();
    return n.startsWith(needle) ? 0 : n.includes(needle) ? 1 : 2;
  };
  return [...items].sort(
    (a, b) => score(a) - score(b) || (a.displayName ?? '').localeCompare(b.displayName ?? ''),
  );
}

export async function quickSearchApplications(
  token: TokenFn,
  q: string,
  depth: SearchDepth = QUICK,
): Promise<QuickSearchResult<Application>> {
  const path = '/applications';
  const structured = buildQuery(q, 'app');
  if (structured === 'none') return { items: [], more: false };
  if (structured) {
    const r = await page<Application>(
      token,
      path,
      { $select: APP_SELECT, ...structured.params() },
      depth,
    );
    return { ...r, items: rank(r.items.filter(structured.keep), structured.text) };
  }
  if (GUID_LIKE.test(q)) {
    const r = await page<Application>(
      token,
      path,
      { $select: APP_SELECT, $filter: `startswith(appId,'${odata(q)}')` },
      depth,
    );
    return { ...r, items: rank(r.items, q) };
  }
  // $search and $filter can't be OR-ed in one request, so they run side by
  // side. The filter half is a bonus: if a tenant rejects it, names still work.
  const [byName, byOther] = await Promise.all([
    page<Application>(token, path, { $select: APP_SELECT, $search: searchTerm(q) }, depth),
    page<Application>(
      token,
      path,
      {
        $select: APP_SELECT,
        $filter:
          `startswith(uniqueName,'${odata(q)}') or ` +
          `tags/any(t:startswith(t,'${odata(toStoredTag(q))}'))`,
      },
      depth,
    ).catch(() => ({ items: [] as Application[], more: false })),
  ]);
  const seen = new Set<string>();
  const items = [...byName.items, ...byOther.items].filter((a) =>
    seen.has(a.id) ? false : (seen.add(a.id), true),
  );
  return { items: rank(items, q), more: byName.more || byOther.more };
}

export async function quickSearchServicePrincipals(
  token: TokenFn,
  q: string,
  depth: SearchDepth = QUICK,
): Promise<QuickSearchResult<ServicePrincipal>> {
  const structured = buildQuery(q, 'sp');
  if (structured === 'none') return { items: [], more: false };
  const r = await page<ServicePrincipal>(
    token,
    '/servicePrincipals',
    structured
      ? { $select: SP_SELECT, ...structured.params() }
      : GUID_LIKE.test(q)
        ? { $select: SP_SELECT, $filter: `startswith(appId,'${odata(q)}')` }
        : { $select: SP_SELECT, $search: searchTerm(q) },
    depth,
  );
  return structured
    ? { ...r, items: rank(r.items.filter(structured.keep), structured.text) }
    : { ...r, items: rank(r.items, q) };
}

export interface QuickUser {
  id: string;
  createdDateTime?: string | null;
  displayName?: string;
  userPrincipalName?: string;
  mail?: string;
  userType?: string;
  accountEnabled?: boolean;
}

export interface QuickGroup {
  id: string;
  createdDateTime?: string | null;
  displayName?: string;
  description?: string;
  mail?: string;
  mailEnabled?: boolean;
  securityEnabled?: boolean;
  groupTypes?: string[];
}

// 'shared' finds shared mailboxes from what the directory knows about them.
// Entra has no "shared mailbox" flag — that lives in Exchange — but the
// account behind one is a member that can't sign in, has a mail address and
// carries no licence. Room and equipment mailboxes look the same from here,
// and a shared mailbox that was given a licence (for an archive or more than
// 50 GB) is missed.
export type UserKindFilter = 'Member' | 'Guest' | 'shared';
const USER_KIND_FILTER: Record<UserKindFilter, string> = {
  Member: "userType eq 'Member'",
  Guest: "userType eq 'Guest'",
  shared:
    "accountEnabled eq false and userType eq 'Member' and mail ne null " +
    'and assignedLicenses/$count eq 0',
};

// An external account keeps "#EXT#" in its sign-in name even when its
// userType was switched to Member, so it counts as a guest here either way.
export const isGuest = (u: { userType?: string; userPrincipalName?: string }) =>
  u.userType === 'Guest' || Boolean(u.userPrincipalName?.toUpperCase().includes('#EXT#'));

// Graph can't filter on "UPN contains", but an external UPN always has the
// shape alias_domain#EXT#@<tenant>.onmicrosoft.com — so with the tenant's
// onmicrosoft.com domains known, "ends with #EXT#@domain" finds them all.
// Without those domains the filter falls back to userType alone.
function userKindFilter(kind: UserKindFilter, extDomains: string[]): string {
  const ext = extDomains.map((d) => `endswith(userPrincipalName,'#EXT#@${odata(d)}')`);
  if (kind === 'shared' || ext.length === 0) return USER_KIND_FILTER[kind];
  return kind === 'Guest'
    ? `(userType eq 'Guest' or ${ext.join(' or ')})`
    : `userType eq 'Member' and ${ext.map((e) => `not(${e})`).join(' and ')}`;
}

export async function quickSearchUsers(
  token: TokenFn,
  q: string,
  depth: SearchDepth = QUICK,
  kind?: UserKindFilter,
  // The tenant's *.onmicrosoft.com domains; lets Member / Guest sort
  // "#EXT#" accounts under guests.
  extDomains: string[] = [],
): Promise<QuickSearchResult<QuickUser>> {
  const structured = buildQuery(q, 'user');
  if (structured === 'none') return { items: [], more: false };
  const run = (filter?: string) =>
    page<QuickUser>(
      token,
      '/users',
      structured
        ? { $select: USER_SELECT, ...structured.params(filter, true) }
        : principalQuery(USER_SELECT, q, ['displayName', 'userPrincipalName', 'mail'], filter),
      depth,
    );
  let r: QuickSearchResult<QuickUser>;
  try {
    r = await run(kind && userKindFilter(kind, extDomains));
  } catch (e) {
    // A tenant that rejects the "#EXT#" filter still gets the plain split.
    if (!(e instanceof GraphError && e.status === 400) || !kind || !extDomains.length) throw e;
    r = await run(USER_KIND_FILTER[kind]);
  }
  if (structured) return { ...r, items: rank(r.items.filter(structured.keep), structured.text) };
  return { ...r, items: rank(r.items, q) };
}

// 'unified' is a Microsoft 365 group. Those are mail-enabled and can be
// security-enabled as well, so the other two leave them out to keep the three
// apart. A mail-enabled security group counts as security, not distribution.
export type GroupKindFilter = 'security' | 'unified' | 'distribution';
const GROUP_KIND_FILTER: Record<GroupKindFilter, string> = {
  security: "securityEnabled eq true and NOT groupTypes/any(t:t eq 'Unified')",
  unified: "groupTypes/any(t:t eq 'Unified')",
  distribution:
    "mailEnabled eq true and securityEnabled eq false and NOT groupTypes/any(t:t eq 'Unified')",
};

export async function quickSearchGroups(
  token: TokenFn,
  q: string,
  depth: SearchDepth = QUICK,
  kind?: GroupKindFilter,
): Promise<QuickSearchResult<QuickGroup>> {
  const filter = kind && GROUP_KIND_FILTER[kind];
  const structured = buildQuery(q, 'group');
  if (structured === 'none') return { items: [], more: false };
  const r = await page<QuickGroup>(
    token,
    '/groups',
    structured
      ? { $select: GROUP_SELECT, ...structured.params(filter, true) }
      : principalQuery(GROUP_SELECT, q, ['displayName', 'mail'], filter),
    depth,
  );
  return structured
    ? { ...r, items: rank(r.items.filter(structured.keep), structured.text) }
    : { ...r, items: rank(r.items, q) };
}

interface BatchResponse<B> {
  responses: { id: string; status: number; body?: B }[];
}

// One GET per id, sent as JSON batches of 20 (Graph's limit), a few batches
// at a time. Ids whose batch failed outright are left out of the result.
async function batchGet<B>(
  token: TokenFn,
  ids: string[],
  url: (id: string) => string,
  headers?: Record<string, string>,
): Promise<Map<string, { status: number; body?: B }>> {
  const out = new Map<string, { status: number; body?: B }>();
  const chunks: string[][] = [];
  for (let i = 0; i < ids.length; i += 20) chunks.push(ids.slice(i, i + 20));
  let next = 0;
  const lane = async () => {
    while (next < chunks.length) {
      const chunk = chunks[next++];
      const res = await graph<BatchResponse<B>>(token, '/$batch', {
        method: 'POST',
        body: {
          requests: chunk.map((id) => ({ id, method: 'GET', url: url(id), headers })),
        },
      }).catch(() => null);
      for (const r of res?.responses ?? []) out.set(r.id, { status: r.status, body: r.body });
    }
  };
  await Promise.all(Array.from({ length: Math.min(4, chunks.length) }, lane));
  return out;
}

// Direct member count per group (people, devices, apps and nested groups each
// count once; members of nested groups don't). Graph has no way to return
// this with the list itself. Groups whose count couldn't be read are left out
// of the map.
export async function countGroupMembers(
  token: TokenFn,
  groupIds: string[],
): Promise<Map<string, number>> {
  const res = await batchGet<{ '@odata.count'?: number }>(
    token,
    groupIds,
    (id) => `/groups/${id}/members?$count=true&$top=1&$select=id`,
    { ConsistencyLevel: 'eventual' },
  );
  const out = new Map<string, number>();
  for (const [id, r] of res) {
    const n = r.body?.['@odata.count'];
    if (r.status === 200 && typeof n === 'number') out.set(id, n);
  }
  return out;
}

// The SharePoint team site of a Microsoft 365 group: its URL, or the HTTP
// status that came back instead (403 — the signed-in user may not open the
// site; 404 — it hasn't been provisioned).
export interface GroupSite {
  url?: string;
  status: number;
}

// Needs a token with Sites.Read.All. The token is taken up front so that a
// missing consent surfaces as an error rather than as a list of failures.
export async function listGroupSites(
  token: TokenFn,
  groupIds: string[],
): Promise<Map<string, GroupSite>> {
  await token();
  const res = await batchGet<{ webUrl?: string }>(
    token,
    groupIds,
    (id) => `/groups/${id}/sites/root?$select=webUrl`,
  );
  const out = new Map<string, GroupSite>();
  for (const [id, r] of res) {
    out.set(id, { status: r.status, url: r.status === 200 ? r.body?.webUrl : undefined });
  }
  return out;
}

interface SignInActivityBody {
  signInActivity?: {
    lastSignInDateTime?: string | null;
    lastNonInteractiveSignInDateTime?: string | null;
    lastSuccessfulSignInDateTime?: string | null;
  };
}

// Most recent sign-in per user — interactive or not, whichever is latest —
// or null for a user who never signed in. Graph caps a user list that selects
// signInActivity at 120 rows and fails it outright on tenants without Entra
// ID P1, so it's read per user in batches after the list has loaded. Returns
// null when the tenant or the account can't read it at all (needs
// AuditLog.Read.All and Entra ID P1/P2).
export async function listLastSignIns(
  token: TokenFn,
  userIds: string[],
): Promise<Map<string, string | null> | null> {
  const res = await batchGet<SignInActivityBody>(
    token,
    userIds,
    (id) => `/users/${id}?$select=signInActivity`,
  );
  const out = new Map<string, string | null>();
  let denied = 0;
  for (const [id, r] of res) {
    if (r.status === 403) denied++;
    if (r.status !== 200) continue;
    const a = r.body?.signInActivity;
    const times = [
      a?.lastSignInDateTime,
      a?.lastNonInteractiveSignInDateTime,
      a?.lastSuccessfulSignInDateTime,
    ].filter((t): t is string => Boolean(t));
    // ISO timestamps in UTC sort as text.
    out.set(id, times.length > 0 ? times.sort().at(-1)! : null);
  }
  return out.size === 0 && denied > 0 ? null : out;
}

interface IdList {
  value?: { id: string }[];
  '@odata.count'?: number;
  '@odata.nextLink'?: string;
}

export interface UserGroups {
  // Groups the user is a direct member of.
  direct: string[];
  // The exact number of those, also when there are more than one page holds.
  directCount: number;
  // Direct groups plus every group those are nested in.
  all: string[];
}

const EVENTUAL = { ConsistencyLevel: 'eventual' };

// Group memberships per user, direct and through nesting. Users whose
// memberships couldn't be read are left out of the map.
export async function listUserGroupIds(
  token: TokenFn,
  userIds: string[],
): Promise<Map<string, UserGroups>> {
  const groups = (kind: 'memberOf' | 'transitiveMemberOf') => (id: string) =>
    `/users/${id}/${kind}/microsoft.graph.group?$select=id&$top=999&$count=true`;
  // One after the other: both at once would be a burst of hundreds of
  // directory reads, which Graph throttles.
  const direct = await batchGet<IdList>(token, userIds, groups('memberOf'), EVENTUAL);
  const all = await batchGet<IdList>(token, userIds, groups('transitiveMemberOf'), EVENTUAL);
  const out = new Map<string, UserGroups>();
  for (const id of userIds) {
    const d = direct.get(id);
    if (d?.status !== 200) continue;
    const directIds = (d.body?.value ?? []).map((g) => g.id);
    const a = all.get(id);
    out.set(id, {
      direct: directIds,
      directCount: d.body?.['@odata.count'] ?? directIds.length,
      all: a?.status === 200 ? (a.body?.value ?? []).map((g) => g.id) : directIds,
    });
  }
  return out;
}

export interface UserApps {
  // Enterprise apps the user holds a role on: assigned to the user, or to a
  // group the user is a direct member of.
  total: number;
  direct: number;
  viaGroups: number;
  // Further apps assigned only to a group the user is in through nesting.
  // Entra does not pass an app assignment down to nested groups, so these
  // don't give the user access and aren't part of `total`.
  nestedOnly: number;
  // A list was longer than one page: the numbers are a lower bound.
  more: boolean;
}

interface AssignmentList {
  value?: { resourceId?: string }[];
  '@odata.nextLink'?: string;
}

// Counts distinct enterprise apps per user. Assignments are read once per
// user and once per distinct group, then combined. Users whose own
// assignments couldn't be read are left out of the map.
export async function countUserApps(
  token: TokenFn,
  userIds: string[],
  groupsByUser: Map<string, UserGroups>,
): Promise<Map<string, UserApps>> {
  const own = await batchGet<AssignmentList>(
    token,
    userIds,
    (id) => `/users/${id}/appRoleAssignments?$select=resourceId`,
  );
  const groupIds = new Set<string>();
  for (const g of groupsByUser.values()) for (const id of g.all) groupIds.add(id);
  const byGroup = await batchGet<AssignmentList>(
    token,
    [...groupIds],
    (id) => `/groups/${id}/appRoleAssignments?$select=resourceId`,
  );
  const apps = (r?: { status: number; body?: AssignmentList }) =>
    r?.status === 200
      ? (r.body?.value ?? []).map((a) => a.resourceId).filter((x): x is string => Boolean(x))
      : [];
  const paged = (r?: { body?: AssignmentList }) => Boolean(r?.body?.['@odata.nextLink']);

  const out = new Map<string, UserApps>();
  for (const id of userIds) {
    const mine = own.get(id);
    if (mine?.status !== 200) continue;
    const groups = groupsByUser.get(id);
    const direct = new Set(apps(mine));
    const effective = new Set(direct);
    let more = paged(mine);
    for (const g of groups?.direct ?? []) {
      for (const a of apps(byGroup.get(g))) effective.add(a);
      more ||= paged(byGroup.get(g));
    }
    const nested = new Set<string>();
    const directGroups = new Set(groups?.direct ?? []);
    for (const g of groups?.all ?? []) {
      if (directGroups.has(g)) continue;
      for (const a of apps(byGroup.get(g))) if (!effective.has(a)) nested.add(a);
    }
    out.set(id, {
      total: effective.size,
      direct: direct.size,
      viaGroups: effective.size - direct.size,
      nestedOnly: nested.size,
      more,
    });
  }
  return out;
}
