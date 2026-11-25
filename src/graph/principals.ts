import { GRAPH_BASE } from '../auth/config';
import { graph, graphAll, GraphError } from './client';
import type { DirectoryAuditEvent } from './directoryAudits';
import type { QuickGroup, QuickUser } from './quickSearch';
import type { AppRoleAssignment } from './types';

type TokenFn = () => Promise<string>;

// Create / read / update / delete for users and groups. Writes need a token
// with User.ReadWrite.All or Group.ReadWrite.All (see USER_WRITE_SCOPES and
// GROUP_WRITE_SCOPES) and a directory role that allows the change.

const USER_DETAIL_SELECT =
  'id,displayName,givenName,surname,userPrincipalName,mail,userType,accountEnabled,' +
  'jobTitle,department,onPremisesSyncEnabled,createdDateTime,lastPasswordChangeDateTime,' +
  'companyName,officeLocation,mobilePhone,businessPhones,employeeId,employeeType,' +
  'streetAddress,city,state,postalCode,country,usageLocation,proxyAddresses,' +
  'onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesDomainName,' +
  'externalUserState,externalUserStateChangeDateTime,identities';

export interface UserDetail extends QuickUser {
  givenName?: string | null;
  surname?: string | null;
  jobTitle?: string | null;
  department?: string | null;
  // Synced from on-premises AD: most properties are read-only in the cloud.
  onPremisesSyncEnabled?: boolean | null;
  createdDateTime?: string | null;
  lastPasswordChangeDateTime?: string | null;
  companyName?: string | null;
  officeLocation?: string | null;
  mobilePhone?: string | null;
  employeeId?: string | null;
  // Free text such as Employee, Contractor or Consultant.
  employeeType?: string | null;
  // Graph accepts a single number here.
  businessPhones?: string[] | null;
  streetAddress?: string | null;
  city?: string | null;
  state?: string | null;
  postalCode?: string | null;
  country?: string | null;
  // Two-letter country code; licences can't be assigned without one.
  usageLocation?: string | null;
  // Every address of the mailbox: "SMTP:" marks the primary, "smtp:" aliases.
  proxyAddresses?: string[] | null;
  onPremisesLastSyncDateTime?: string | null;
  onPremisesSamAccountName?: string | null;
  onPremisesDomainName?: string | null;
  // Guests: PendingAcceptance until the invitation is redeemed, then Accepted.
  externalUserState?: string | null;
  externalUserStateChangeDateTime?: string | null;
  // How the account signs in; for guests the issuer is their home identity
  // provider (a tenant domain, "google.com", "mail" for one-time passcode, …).
  identities?: { signInType?: string; issuer?: string; issuerAssignedId?: string | null }[] | null;
}

export function getUser(token: TokenFn, id: string) {
  return graph<UserDetail>(token, `/users/${id}`, { query: { $select: USER_DETAIL_SELECT } });
}

export interface NewUser {
  displayName: string;
  userPrincipalName: string;
  mailNickname: string;
  givenName?: string;
  surname?: string;
  accountEnabled: boolean;
  password: string;
}

export function createUser(token: TokenFn, u: NewUser) {
  const { password, ...rest } = u;
  return graph<UserDetail>(token, '/users', {
    method: 'POST',
    body: {
      ...rest,
      passwordProfile: { password, forceChangePasswordNextSignIn: true },
    },
  });
}

export type UserPatch = Partial<
  Pick<
    UserDetail,
    | 'displayName'
    | 'givenName'
    | 'surname'
    | 'jobTitle'
    | 'department'
    | 'accountEnabled'
    | 'companyName'
    | 'officeLocation'
    | 'mobilePhone'
    | 'businessPhones'
    | 'employeeId'
    | 'employeeType'
    | 'streetAddress'
    | 'city'
    | 'state'
    | 'postalCode'
    | 'country'
    | 'usageLocation'
  >
>;

export function updateUser(token: TokenFn, id: string, patch: UserPatch) {
  return graph<void>(token, `/users/${id}`, { method: 'PATCH', body: patch });
}

// Replaces the user's password. Needs a token with
// User-PasswordProfile.ReadWrite.All and a role that may reset this user's
// password (User Administrator for ordinary users; resetting an admin takes
// Privileged Authentication Administrator or Global Administrator).
export function setUserPassword(
  token: TokenFn,
  id: string,
  password: string,
  forceChangePasswordNextSignIn: boolean,
) {
  return graph<void>(token, `/users/${id}`, {
    method: 'PATCH',
    body: { passwordProfile: { password, forceChangePasswordNextSignIn } },
  });
}

// Soft delete: the user sits in the directory recycle bin for 30 days.
export function deleteUser(token: TokenFn, id: string) {
  return graph<void>(token, `/users/${id}`, { method: 'DELETE' });
}

export type NewGroupKind = 'security' | 'm365';

export interface NewGroup {
  displayName: string;
  description?: string;
  mailNickname: string;
  kind: NewGroupKind;
}

export function createGroup(token: TokenFn, g: NewGroup) {
  const m365 = g.kind === 'm365';
  return graph<QuickGroup>(token, '/groups', {
    method: 'POST',
    body: {
      displayName: g.displayName,
      ...(g.description ? { description: g.description } : {}),
      mailNickname: g.mailNickname,
      mailEnabled: m365,
      securityEnabled: !m365,
      groupTypes: m365 ? ['Unified'] : [],
    },
  });
}

export function updateGroup(
  token: TokenFn,
  id: string,
  patch: { displayName?: string; description?: string | null },
) {
  return graph<void>(token, `/groups/${id}`, { method: 'PATCH', body: patch });
}

// Microsoft 365 groups are soft-deleted (30 days); every other kind is gone
// for good.
export function deleteGroup(token: TokenFn, id: string) {
  return graph<void>(token, `/groups/${id}`, { method: 'DELETE' });
}

interface VerifiedDomain {
  name: string;
  isDefault?: boolean;
}

// Domains a new user's sign-in name can use, default first.
export async function listVerifiedDomains(token: TokenFn): Promise<string[]> {
  const res = await graph<{ value: { verifiedDomains?: VerifiedDomain[] }[] }>(
    token,
    '/organization',
    { query: { $select: 'verifiedDomains' } },
  );
  const domains = res.value[0]?.verifiedDomains ?? [];
  return [...domains]
    .sort((a, b) => Number(b.isDefault ?? false) - Number(a.isDefault ?? false))
    .map((d) => d.name);
}

// Temporary password for a new user: 16 characters with every class Entra's
// complexity rules ask for. The user must change it at first sign-in.
export function generatePassword(): string {
  const sets = ['ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnpqrstuvwxyz', '23456789', '!@#$%*-_+='];
  const all = sets.join('');
  // Rejection sampling keeps the choice uniform.
  const rand = (n: number) => {
    const limit = 256 - (256 % n);
    const b = new Uint8Array(1);
    do crypto.getRandomValues(b);
    while (b[0] >= limit);
    return b[0] % n;
  };
  const pick = (chars: string) => chars[rand(chars.length)];
  const out = sets.map(pick);
  while (out.length < 16) out.push(pick(all));
  // Shuffle so the guaranteed classes aren't always up front.
  for (let i = out.length - 1; i > 0; i--) {
    const j = rand(i + 1);
    [out[i], out[j]] = [out[j], out[i]];
  }
  return out.join('');
}

// ------------------------------------------------------------- Detail pages

const GROUP_DETAIL_SELECT =
  'id,displayName,description,mail,mailNickname,mailEnabled,securityEnabled,groupTypes,' +
  'membershipRule,onPremisesSyncEnabled,createdDateTime';

export interface GroupDetail extends QuickGroup {
  mailNickname?: string | null;
  // Set on dynamic groups: the rule that decides who is a member.
  membershipRule?: string | null;
  onPremisesSyncEnabled?: boolean | null;
  createdDateTime?: string | null;
}

export function getGroup(token: TokenFn, id: string) {
  return graph<GroupDetail>(token, `/groups/${id}`, { query: { $select: GROUP_DETAIL_SELECT } });
}

// A user, group, device or service principal as it appears in a member or
// owner list.
export interface DirectoryMember {
  id: string;
  '@odata.type'?: string;
  displayName?: string;
  userPrincipalName?: string;
  mail?: string;
  appId?: string;
}

// Lists on a detail page stop after this many entries; `more` says so.
const RELATED_TOP = 100;
const RELATED_PAGES = 3;

export interface Related<T> {
  items: T[];
  more: boolean;
}

async function related<T>(token: TokenFn, path: string, select?: string): Promise<Related<T>> {
  const items = await graphAll<T>(
    token,
    path,
    { query: { ...(select ? { $select: select } : {}), $top: RELATED_TOP } },
    RELATED_PAGES,
  );
  return { items, more: items.length >= RELATED_TOP * RELATED_PAGES };
}

const MEMBER_SELECT = 'id,displayName,userPrincipalName,mail,appId';

// Groups the user is a direct member of.
export function listUserGroups(token: TokenFn, userId: string) {
  return related<QuickGroup>(
    token,
    `/users/${userId}/memberOf/microsoft.graph.group`,
    'id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes',
  );
}

// Makes a user (or any directory object) a direct member of the group. Needs
// Group.ReadWrite.All and a role that manages the group, or ownership of it.
// Graph refuses groups it doesn't manage: dynamic groups, groups synced from
// on-premises AD, and distribution or mail-enabled security groups (Exchange).
export function addGroupMember(token: TokenFn, groupId: string, memberId: string) {
  return graph<void>(token, `/groups/${groupId}/members/$ref`, {
    method: 'POST',
    body: { '@odata.id': `${GRAPH_BASE}/directoryObjects/${memberId}` },
  });
}

// Takes a direct member out of the group. Same permissions and the same
// unmanaged kinds of group as addGroupMember.
export function removeGroupMember(token: TokenFn, groupId: string, memberId: string) {
  return graph<void>(token, `/groups/${groupId}/members/${memberId}/$ref`, { method: 'DELETE' });
}

export function listGroupMembers(token: TokenFn, groupId: string) {
  return related<DirectoryMember>(token, `/groups/${groupId}/members`, MEMBER_SELECT);
}

export function listGroupOwners(token: TokenFn, groupId: string) {
  return related<DirectoryMember>(token, `/groups/${groupId}/owners`, MEMBER_SELECT);
}

// Enterprise apps the user or group is assigned to (directly, not through a
// group they are in).
export function listAppAssignments(token: TokenFn, kind: 'users' | 'groups', id: string) {
  return related<AppRoleAssignment>(token, `/${kind}/${id}/appRoleAssignments`);
}

// ------------------------------------------------------------ Audit details

export interface UserSignInActivity {
  // Interactive sign-in (the user typed credentials or was prompted).
  lastSignInDateTime?: string | null;
  // Token refreshes and other sign-ins a client made on the user's behalf.
  lastNonInteractiveSignInDateTime?: string | null;
  lastSuccessfulSignInDateTime?: string | null;
}

const denied = (e: unknown) => e instanceof GraphError && (e.status === 403 || e.status === 404);

// Needs AuditLog.Read.All and an Entra ID P1/P2 tenant; null when either is
// missing, so the page can say "unavailable" instead of failing.
export async function getUserSignInActivity(
  token: TokenFn,
  id: string,
): Promise<UserSignInActivity | null> {
  try {
    const res = await graph<{ signInActivity?: UserSignInActivity }>(token, `/users/${id}`, {
      query: { $select: 'signInActivity' },
    });
    return res.signInActivity ?? {};
  } catch (e) {
    if (denied(e)) return null;
    throw e;
  }
}

// Who created an object and who touched it last. Graph keeps no such
// properties on users and groups, so both come from the directory audit log —
// which only reaches back about 30 days (7 on free tenants). Older events are
// simply absent.
export interface AuditTrail {
  created?: DirectoryAuditEvent;
  // Most recent successful change of any kind, the creation itself excluded.
  modified?: DirectoryAuditEvent;
}

export async function getAuditTrail(
  token: TokenFn,
  id: string,
  addActivity: 'Add user' | 'Add group',
): Promise<AuditTrail | null> {
  const target = `targetResources/any(t:t/id eq '${id}')`;
  const audits = (filter: string, top: number) =>
    graph<{ value?: DirectoryAuditEvent[] }>(token, '/auditLogs/directoryAudits', {
      query: { $filter: filter, $top: top, $orderby: 'activityDateTime desc' },
    }).then((r) => r.value ?? []);
  try {
    const [recent, added] = await Promise.all([
      audits(target, 25),
      audits(`activityDisplayName eq '${addActivity}' and ${target}`, 1),
    ]);
    const created = added[0];
    const modified = recent.find((e) => e.result !== 'failure' && e.id !== created?.id);
    return { created, modified };
  } catch (e) {
    if (denied(e)) return null;
    throw e;
  }
}
