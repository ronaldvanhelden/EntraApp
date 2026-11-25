import { graph, graphAll, GraphError, type GraphTokenFn } from './client';
import { resolveDirectoryObjects, type PrincipalKind } from './directoryObjects';

// Privileged Identity Management for Microsoft Entra roles: what someone is
// eligible for or holds, activating an eligible role, and assigning or
// removing roles as an admin. Everything here needs Entra ID P2 (or ID
// Governance); without it Graph answers with a licence error.
//
// Reads take RoleManagement.Read.Directory, activating your own role
// RoleAssignmentSchedule.ReadWrite.Directory, and assigning or removing also
// RoleEligibilitySchedule.ReadWrite.Directory — see the PIM_* scope sets.

const BASE = '/roleManagement/directory';

export interface RoleDefinition {
  id: string;
  displayName?: string;
  description?: string;
  isBuiltIn?: boolean;
}

export async function listRoleDefinitions(token: GraphTokenFn): Promise<RoleDefinition[]> {
  const defs = await graphAll<RoleDefinition>(token, `${BASE}/roleDefinitions`, {
    query: { $select: 'id,displayName,description,isBuiltIn' },
  });
  return defs.sort((a, b) => (a.displayName ?? '').localeCompare(b.displayName ?? ''));
}

// One role someone is eligible for or holds, as PIM sees it right now.
export interface RoleSchedule {
  key: string;
  // Eligible: may activate it. Active: holds it now.
  kind: 'eligible' | 'active';
  roleDefinitionId: string;
  roleName: string;
  roleDescription?: string;
  // Who the schedule is on: a user, a role-assignable group or an app.
  principalId: string;
  principalName: string;
  principalKind?: PrincipalKind;
  // "/" is the whole directory; otherwise an administrative unit or object.
  directoryScopeId: string;
  // Direct, or Group when it reaches the person through a group.
  memberType?: string;
  // Active only: Assigned by an admin, or Activated from an eligibility.
  assignmentType?: string;
  startDateTime?: string | null;
  // Missing or null: no end date.
  endDateTime?: string | null;
  // The group the role comes through. Absent when it is on the principal's
  // own account; null when it is indirect but the group isn't known.
  via?: { id: string; name: string } | null;
  // Eligible rows in the tenant-wide list: when the principal last activated
  // this role. null: no activation in PIM's request history; undefined: the
  // history couldn't be read (or the row isn't a user's eligibility).
  lastActivated?: string | null;
}

interface ScheduleInstance {
  id: string;
  principalId: string;
  roleDefinitionId: string;
  directoryScopeId?: string;
  memberType?: string;
  assignmentType?: string;
  startDateTime?: string | null;
  endDateTime?: string | null;
  roleDefinition?: RoleDefinition;
}

const ELIGIBLE = `${BASE}/roleEligibilityScheduleInstances`;
const ACTIVE = `${BASE}/roleAssignmentScheduleInstances`;

function toRows(
  kind: RoleSchedule['kind'],
  instances: ScheduleInstance[],
  roles: Map<string, RoleDefinition>,
  names: Record<string, { displayName: string; kind: PrincipalKind }>,
): RoleSchedule[] {
  return instances.map((i) => {
    const role = i.roleDefinition ?? roles.get(i.roleDefinitionId);
    return {
      key: `${kind}-${i.id}`,
      kind,
      roleDefinitionId: i.roleDefinitionId,
      roleName: role?.displayName ?? i.roleDefinitionId,
      roleDescription: role?.description,
      principalId: i.principalId,
      principalName: names[i.principalId]?.displayName ?? i.principalId,
      principalKind: names[i.principalId]?.kind,
      directoryScopeId: i.directoryScopeId ?? '/',
      memberType: i.memberType,
      assignmentType: i.assignmentType,
      startDateTime: i.startDateTime,
      endDateTime: i.endDateTime,
    };
  });
}

const byRole = (a: RoleSchedule, b: RoleSchedule) =>
  a.roleName.localeCompare(b.roleName) ||
  a.principalName.localeCompare(b.principalName) ||
  a.kind.localeCompare(b.kind);

// The signed-in admin's own roles: what they can activate and what is active.
// Each row also says which group the role comes through, if it isn't on the
// admin's own account. Graph can report that in more than one way — the
// schedule names the group as its principal, or it names the user (with or
// without flagging the membership as indirect) — so the answer is taken from
// who actually holds the role: if I'm not on that list myself, the group of
// mine that is on it is how the role reaches me.
export async function listMyRoles(token: GraphTokenFn): Promise<RoleSchedule[]> {
  const mine = (path: string) =>
    graphAll<ScheduleInstance>(token, `${path}/filterByCurrentUser(on='principal')`, {
      query: { $expand: 'roleDefinition' },
    });
  const [eligible, active, me] = await Promise.all([
    mine(ELIGIBLE),
    mine(ACTIVE),
    graph<{ id: string }>(token, '/me', { query: { $select: 'id' } }),
  ]);
  const none = new Map<string, RoleDefinition>();
  const names = await resolveDirectoryObjects(
    token,
    [...eligible, ...active].map((i) => i.principalId),
  );
  const rows = [
    ...toRows('eligible', eligible, none, names),
    ...toRows('active', active, none, names),
  ];

  for (const row of rows) {
    if (row.principalId !== me.id) row.via = { id: row.principalId, name: row.principalName };
  }
  // Rows that name me as the principal are checked too, whatever they say
  // about being direct: the role's own list of holders settles it.
  const indirect = rows.filter((r) => !r.via);
  if (indirect.length > 0) {
    const groups = await graphAll<{ id: string; displayName?: string }>(
      token,
      '/me/transitiveMemberOf/microsoft.graph.group',
      { advanced: true, query: { $select: 'id,displayName', $top: 999 } },
    ).catch(() => []);
    const groupName = new Map(groups.map((g) => [g.id, g.displayName ?? g.id]));
    // Everyone holding the same role: one of my groups is among them.
    const holders = new Map<string, Promise<ScheduleInstance[]>>();
    const holdersOf = (row: RoleSchedule) => {
      const key = `${row.kind}|${row.roleDefinitionId}`;
      let found = holders.get(key);
      if (!found) {
        found = graphAll<ScheduleInstance>(token, row.kind === 'eligible' ? ELIGIBLE : ACTIVE, {
          query: { $filter: `roleDefinitionId eq '${row.roleDefinitionId}'` },
        }).catch(() => []);
        holders.set(key, found);
      }
      return found;
    };
    await Promise.all(
      indirect.map(async (row) => {
        const sameScope = (i: ScheduleInstance) =>
          (i.directoryScopeId ?? '/') === row.directoryScopeId;
        const all = (await holdersOf(row)).filter(sameScope);
        const flagged = row.memberType !== undefined && row.memberType !== 'Direct';
        // On my own account as well: that is how I hold it.
        if (!flagged && all.some((i) => i.principalId === me.id)) return;
        const match = all.find((i) => groupName.has(i.principalId));
        if (match) row.via = { id: match.principalId, name: groupName.get(match.principalId)! };
        // Known to be indirect, but the group couldn't be worked out.
        else if (flagged) row.via = null;
      }),
    );
  }
  return rows.sort(byRole);
}

interface ActivationRequest {
  principalId?: string;
  roleDefinitionId?: string;
  directoryScopeId?: string;
  action?: string;
  status?: string;
  createdDateTime?: string;
}

const activationKey = (principalId: string, roleDefinitionId: string, scope: string) =>
  `${principalId}|${roleDefinitionId}|${scope}`;

// When each person last activated each role, from PIM's own request history
// (the audit log only reaches back 30 days). How far back that history goes
// is up to the service, so a missing entry means "not in the history", not
// necessarily "never".
async function listLastActivations(token: GraphTokenFn): Promise<Map<string, string>> {
  const path = `${BASE}/roleAssignmentScheduleRequests`;
  const select = 'principalId,roleDefinitionId,directoryScopeId,action,status,createdDateTime';
  const requests = await graphAll<ActivationRequest>(
    token,
    path,
    { query: { $filter: "action eq 'selfActivate'", $select: select } },
    40,
  ).catch(() =>
    // Without the filter, where it isn't supported.
    graphAll<ActivationRequest>(token, path, { query: { $select: select } }, 40),
  );
  const last = new Map<string, string>();
  for (const r of requests) {
    if (r.action !== 'selfActivate' || !r.createdDateTime) continue;
    if (/fail|denied|cancel|pending/i.test(r.status ?? '')) continue;
    if (!r.principalId || !r.roleDefinitionId) continue;
    const key = activationKey(r.principalId, r.roleDefinitionId, r.directoryScopeId ?? '/');
    if ((last.get(key) ?? '') < r.createdDateTime) last.set(key, r.createdDateTime);
  }
  return last;
}

// Every eligible and active role in the tenant, or only those on one
// principal (directly — roles reached through a group sit on the group).
export async function listRoleSchedules(
  token: GraphTokenFn,
  principalId?: string,
): Promise<RoleSchedule[]> {
  const query = principalId ? { $filter: `principalId eq '${principalId}'` } : {};
  const all = (path: string) => graphAll<ScheduleInstance>(token, path, { query }, 20);
  const [eligible, active, defs, activations] = await Promise.all([
    all(ELIGIBLE),
    all(ACTIVE),
    listRoleDefinitions(token),
    listLastActivations(token).catch(() => null),
  ]);
  const roles = new Map(defs.map((d) => [d.id, d]));
  const names = await resolveDirectoryObjects(
    token,
    [...eligible, ...active].map((i) => i.principalId),
  );
  const direct = [
    ...toRows('eligible', eligible, roles, names),
    ...toRows('active', active, roles, names),
  ];

  // A group's row says nothing about who actually gets the role, so each
  // member is listed as well, pointing back at the group. Left out when one
  // principal was asked for.
  const viaGroup: RoleSchedule[] = [];
  if (!principalId) {
    const groupIds = [
      ...new Set(direct.filter((r) => r.principalKind === 'group').map((r) => r.principalId)),
    ].slice(0, MAX_ROLE_GROUPS);
    const members = new Map(
      await Promise.all(groupIds.map(async (id) => [id, await listGroupUsers(token, id)] as const)),
    );
    for (const row of direct) {
      for (const user of members.get(row.principalId) ?? []) {
        viaGroup.push({
          ...row,
          key: `${row.key}-${user.id}`,
          principalId: user.id,
          principalName: user.displayName ?? user.userPrincipalName ?? user.id,
          principalKind: 'user',
          via: { id: row.principalId, name: row.principalName },
        });
      }
    }
  }

  return [...direct, ...viaGroup]
    .map((row) => ({
      ...row,
      // Only a user activates; a group's eligibility is used by its members.
      lastActivated:
        row.kind === 'eligible' && activations && row.principalKind === 'user'
          ? (activations.get(
              activationKey(row.principalId, row.roleDefinitionId, row.directoryScopeId),
            ) ?? null)
          : undefined,
    }))
    .sort(byRole);
}

// Role-assignable groups are few; the cap only guards against a surprise.
const MAX_ROLE_GROUPS = 60;

interface GroupUser {
  id: string;
  '@odata.type'?: string;
  displayName?: string;
  userPrincipalName?: string;
}

// The users in a group, nested groups included. Empty when it can't be read.
async function listGroupUsers(token: GraphTokenFn, groupId: string): Promise<GroupUser[]> {
  try {
    const all = await graphAll<GroupUser>(
      token,
      `/groups/${groupId}/transitiveMembers`,
      { query: { $select: 'id,displayName,userPrincipalName', $top: 999 } },
      5,
    );
    return all.filter((m) => (m['@odata.type'] ?? '').toLowerCase().endsWith('.user'));
  } catch {
    return [];
  }
}

// ------------------------------------------------------------- Activation

// What the role's PIM policy asks of someone activating it.
export interface ActivationRules {
  maxHours: number;
  justification: boolean;
  ticket: boolean;
  mfa: boolean;
  approval: boolean;
}

const DEFAULT_RULES: ActivationRules = {
  maxHours: 8,
  justification: true,
  ticket: false,
  mfa: false,
  approval: false,
};

// "PT8H", "PT30M", "P1D" → hours.
function hours(iso?: string): number | null {
  const m = /^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?)?$/.exec(iso ?? '');
  if (!m) return null;
  const total = Number(m[1] ?? 0) * 24 + Number(m[2] ?? 0) + Number(m[3] ?? 0) / 60;
  return total > 0 ? total : null;
}

interface PolicyRule {
  id?: string;
  maximumDuration?: string;
  enabledRules?: string[];
  setting?: { isApprovalRequired?: boolean };
}

// Best effort: when the policy can't be read the defaults apply and Graph
// still enforces the real rules when the request is sent.
export async function getActivationRules(
  token: GraphTokenFn,
  roleDefinitionId: string,
): Promise<ActivationRules> {
  try {
    const res = await graph<{ value?: { policy?: { rules?: PolicyRule[] } }[] }>(
      token,
      '/policies/roleManagementPolicyAssignments',
      {
        query: {
          $filter: `scopeId eq '/' and scopeType eq 'DirectoryRole' and roleDefinitionId eq '${roleDefinitionId}'`,
          $expand: 'policy($expand=rules)',
        },
      },
    );
    const rules = res.value?.[0]?.policy?.rules ?? [];
    const rule = (id: string) => rules.find((r) => r.id === id);
    const enabled = rule('Enablement_EndUser_Assignment')?.enabledRules ?? [];
    return {
      maxHours:
        hours(rule('Expiration_EndUser_Assignment')?.maximumDuration) ?? DEFAULT_RULES.maxHours,
      justification: enabled.includes('Justification'),
      ticket: enabled.includes('Ticketing'),
      mfa: enabled.includes('MultiFactorAuthentication'),
      approval: rule('Approval_EndUser_Assignment')?.setting?.isApprovalRequired === true,
    };
  } catch {
    return DEFAULT_RULES;
  }
}

export interface RoleRequestResult {
  id?: string;
  // Provisioned, Granted, PendingApproval, PendingProvisioning, …
  status?: string;
}

interface RoleTarget {
  principalId: string;
  roleDefinitionId: string;
  directoryScopeId: string;
}

const durationIso = (h: number) => `PT${Math.round(h * 60)}M`;

// Turn an eligible role on for a limited time. Needs a token with
// RoleAssignmentSchedule.ReadWrite.Directory.
export function activateRole(
  token: GraphTokenFn,
  target: RoleTarget,
  options: { hours: number; justification: string; ticketNumber?: string },
) {
  return graph<RoleRequestResult>(token, `${BASE}/roleAssignmentScheduleRequests`, {
    method: 'POST',
    body: {
      action: 'selfActivate',
      ...target,
      justification: options.justification,
      scheduleInfo: {
        startDateTime: new Date().toISOString(),
        expiration: { type: 'afterDuration', duration: durationIso(options.hours) },
      },
      ...(options.ticketNumber ? { ticketInfo: { ticketNumber: options.ticketNumber } } : {}),
    },
  });
}

// End an activation early. Entra refuses this in the first five minutes.
export function deactivateRole(token: GraphTokenFn, target: RoleTarget) {
  return graph<RoleRequestResult>(token, `${BASE}/roleAssignmentScheduleRequests`, {
    method: 'POST',
    body: { action: 'selfDeactivate', ...target },
  });
}

// Activation can demand a stronger sign-in than the current token carries
// (MFA, or a Conditional Access authentication context). Graph then fails
// the request and says what is missing; the caller gets a new token with it
// and tries again. null: the error is about something else.
export interface StepUp {
  // A claims challenge to pass to MSAL as-is.
  claims?: string;
  // Plain MFA was asked for, without a claims challenge.
  mfa?: boolean;
}

export function stepUpFor(error: unknown): StepUp | null {
  if (!(error instanceof GraphError)) return null;
  const text = error.message ?? '';
  const at = text.indexOf('claims=');
  if (at >= 0) {
    try {
      return { claims: decodeURIComponent(text.slice(at + 7).split('&')[0]) };
    } catch {
      return null;
    }
  }
  if (/MfaRule/i.test(text)) return { mfa: true };
  return null;
}

// ------------------------------------------------------- Admin assignment

// How long an assignment lasts: for good, or until a date.
export type AssignmentEnd = { permanent: true } | { permanent: false; endDateTime: string };

function scheduleInfo(end: AssignmentEnd) {
  return {
    startDateTime: new Date().toISOString(),
    expiration: end.permanent
      ? { type: 'noExpiration' }
      : { type: 'afterDateTime', endDateTime: end.endDateTime },
  };
}

const requestPath = (kind: RoleSchedule['kind']) =>
  `${BASE}/${kind === 'eligible' ? 'roleEligibilityScheduleRequests' : 'roleAssignmentScheduleRequests'}`;

// Make a principal eligible for a role, or give it the role outright. Needs
// RoleEligibilitySchedule.ReadWrite.Directory (eligible) or
// RoleAssignmentSchedule.ReadWrite.Directory (active), and a role that may
// manage role assignments (Privileged Role Administrator).
export function assignRole(
  token: GraphTokenFn,
  kind: RoleSchedule['kind'],
  target: RoleTarget,
  end: AssignmentEnd,
  justification: string,
) {
  return graph<RoleRequestResult>(token, requestPath(kind), {
    method: 'POST',
    body: { action: 'adminAssign', ...target, justification, scheduleInfo: scheduleInfo(end) },
  });
}

// Take an eligibility or an admin-made assignment away again.
export function removeRole(
  token: GraphTokenFn,
  row: Pick<RoleSchedule, 'kind' | 'principalId' | 'roleDefinitionId' | 'directoryScopeId'>,
) {
  return graph<RoleRequestResult>(token, requestPath(row.kind), {
    method: 'POST',
    body: {
      action: 'adminRemove',
      principalId: row.principalId,
      roleDefinitionId: row.roleDefinitionId,
      directoryScopeId: row.directoryScopeId,
    },
  });
}

// A licence error reads badly raw; say what is actually missing.
export function pimError(error: unknown): string {
  const text = error instanceof Error ? error.message : String(error);
  if (/AadPremiumLicenseRequired|license/i.test(text))
    return 'Privileged Identity Management needs a Microsoft Entra ID P2 or ID Governance licence in this tenant.';
  return text;
}
