import { GRAPH_BASE } from '../auth/config';
import { graph, graphAll, GraphError } from './client';
import { resolveDirectoryObjects } from './directoryObjects';
import type { DirectoryMember, Related } from './principals';
import type { OAuth2PermissionGrant, SignIn } from './types';

type TokenFn = () => Promise<string>;

// Everything the user detail page shows beyond the user object itself: what
// the user owns, has consented to, is licensed for and may administer, how
// they sign in, and the two session-level actions (revoke sessions, resend a
// guest invitation). Each call stands alone so one missing permission or
// licence only empties its own tab.

const denied = (e: unknown) => e instanceof GraphError && (e.status === 403 || e.status === 404);
const TOP = 100;
const PAGES = 3;

async function list<T>(token: TokenFn, path: string, select?: string): Promise<Related<T>> {
  const items = await graphAll<T>(
    token,
    path,
    { query: { ...(select ? { $select: select } : {}), $top: TOP } },
    PAGES,
  );
  return { items, more: items.length >= TOP * PAGES };
}

// ---------------------------------------------------------------- Manager

export interface UserManager {
  id: string;
  displayName?: string;
  userPrincipalName?: string;
}

// null when the user has no manager (Graph answers 404).
export async function getUserManager(token: TokenFn, id: string): Promise<UserManager | null> {
  try {
    return await graph<UserManager>(token, `/users/${id}/manager`, {
      query: { $select: 'id,displayName,userPrincipalName' },
    });
  } catch (e) {
    if (denied(e)) return null;
    throw e;
  }
}

// Sets or (with null) clears the user's manager. Needs User.ReadWrite.All.
export function setUserManager(token: TokenFn, id: string, managerId: string | null) {
  return managerId
    ? graph<void>(token, `/users/${id}/manager/$ref`, {
        method: 'PUT',
        body: { '@odata.id': `${GRAPH_BASE}/users/${managerId}` },
      })
    : graph<void>(token, `/users/${id}/manager/$ref`, { method: 'DELETE' });
}

export interface DirectReport extends DirectoryMember {
  jobTitle?: string | null;
}

export function listDirectReports(token: TokenFn, id: string) {
  return list<DirectReport>(
    token,
    `/users/${id}/directReports`,
    'id,displayName,userPrincipalName,mail,jobTitle',
  );
}

// ------------------------------------------------------------------- Owns

// App registrations, enterprise apps and groups the user is an owner of.
export function listOwnedObjects(token: TokenFn, id: string) {
  return list<DirectoryMember>(
    token,
    `/users/${id}/ownedObjects`,
    'id,displayName,appId,mail,userPrincipalName',
  );
}

// --------------------------------------------------------------- Consents

export interface UserConsent {
  grant: OAuth2PermissionGrant;
  clientName: string;
  resourceName: string;
  scopes: string[];
}

// Delegated permissions this user personally consented to (admin consent for
// all users is not listed: it isn't tied to the user).
export async function listUserConsents(token: TokenFn, id: string): Promise<Related<UserConsent>> {
  const grants = await graphAll<OAuth2PermissionGrant>(
    token,
    '/oauth2PermissionGrants',
    { query: { $filter: `principalId eq '${id}'`, $top: TOP } },
    PAGES,
  );
  const names = await resolveDirectoryObjects(
    token,
    grants.flatMap((g) => [g.clientId, g.resourceId]),
  );
  const items = grants
    .map((grant) => ({
      grant,
      clientName: names[grant.clientId]?.displayName ?? grant.clientId,
      resourceName: names[grant.resourceId]?.displayName ?? grant.resourceId,
      scopes: grant.scope.split(' ').filter(Boolean),
    }))
    .sort((a, b) => a.clientName.localeCompare(b.clientName));
  return { items, more: grants.length >= TOP * PAGES };
}

// ---------------------------------------------------------------- Licenses

// Product names for the SKUs seen most; anything else shows its part number.
const SKU_NAMES: Record<string, string> = {
  AAD_PREMIUM: 'Microsoft Entra ID P1',
  AAD_PREMIUM_P2: 'Microsoft Entra ID P2',
  EMS: 'Enterprise Mobility + Security E3',
  EMSPREMIUM: 'Enterprise Mobility + Security E5',
  ENTERPRISEPACK: 'Office 365 E3',
  ENTERPRISEPREMIUM: 'Office 365 E5',
  STANDARDPACK: 'Office 365 E1',
  SPE_E3: 'Microsoft 365 E3',
  SPE_E5: 'Microsoft 365 E5',
  SPE_F1: 'Microsoft 365 F3',
  SPB: 'Microsoft 365 Business Premium',
  O365_BUSINESS_PREMIUM: 'Microsoft 365 Business Standard',
  O365_BUSINESS_ESSENTIALS: 'Microsoft 365 Business Basic',
  EXCHANGESTANDARD: 'Exchange Online (Plan 1)',
  EXCHANGEENTERPRISE: 'Exchange Online (Plan 2)',
  POWER_BI_STANDARD: 'Power BI (free)',
  POWER_BI_PRO: 'Power BI Pro',
  FLOW_FREE: 'Power Automate Free',
  TEAMS_EXPLORATORY: 'Microsoft Teams Exploratory',
  Microsoft_365_Copilot: 'Microsoft 365 Copilot',
  VISIOCLIENT: 'Visio Plan 2',
  PROJECTPROFESSIONAL: 'Project Plan 3',
  INTUNE_A: 'Microsoft Intune Plan 1',
};

export interface UserLicense {
  skuId: string;
  partNumber: string;
  name: string;
  // Service plans switched on / in the product.
  plansEnabled: number;
  plansTotal: number;
  direct: boolean;
  groups: { id: string; name: string }[];
  // Set when the assignment didn't go through (e.g. no licences left).
  error?: string;
}

interface LicenseDetail {
  skuId: string;
  skuPartNumber?: string;
  servicePlans?: { provisioningStatus?: string; appliesTo?: string }[];
}
interface LicenseAssignmentState {
  skuId: string;
  assignedByGroup?: string | null;
  state?: string;
  error?: string;
}

export async function listUserLicenses(token: TokenFn, id: string): Promise<Related<UserLicense>> {
  const [details, user] = await Promise.all([
    graphAll<LicenseDetail>(token, `/users/${id}/licenseDetails`),
    graph<{ licenseAssignmentStates?: LicenseAssignmentState[] }>(token, `/users/${id}`, {
      query: { $select: 'licenseAssignmentStates' },
    }).catch(() => ({ licenseAssignmentStates: undefined })),
  ]);
  const states = user.licenseAssignmentStates ?? [];
  const groupNames = await resolveDirectoryObjects(
    token,
    states.flatMap((s) => (s.assignedByGroup ? [s.assignedByGroup] : [])),
  );
  const items = details
    .map((d): UserLicense => {
      const mine = states.filter((s) => s.skuId === d.skuId);
      const plans = (d.servicePlans ?? []).filter((p) => p.appliesTo !== 'Company');
      const failed = mine.find((s) => s.error && s.error !== 'None');
      const partNumber = d.skuPartNumber ?? d.skuId;
      return {
        skuId: d.skuId,
        partNumber,
        name: SKU_NAMES[partNumber] ?? partNumber,
        plansEnabled: plans.filter((p) => p.provisioningStatus !== 'Disabled').length,
        plansTotal: plans.length,
        // Without assignment states (older tenants) the source is unknown;
        // count it as direct rather than inventing a group.
        direct: mine.length === 0 || mine.some((s) => !s.assignedByGroup),
        groups: mine.flatMap((s) =>
          s.assignedByGroup
            ? [
                {
                  id: s.assignedByGroup,
                  name: groupNames[s.assignedByGroup]?.displayName ?? s.assignedByGroup,
                },
              ]
            : [],
        ),
        error: failed?.error,
      };
    })
    .sort((a, b) => a.name.localeCompare(b.name));
  return { items, more: false };
}

// ------------------------------------------------------------------- Roles

export interface UserRole {
  key: string;
  roleDefinitionId: string;
  name: string;
  description?: string;
  // Active now, or eligible: can be activated through PIM.
  state: 'active' | 'eligible';
  // "/" is the whole directory; anything else is an administrative unit or
  // a single object.
  scope: string;
  // The role-assignable group the role comes through, if not assigned directly.
  viaGroup?: { id: string; name: string };
  // When an eligibility ends; null/undefined is permanent.
  until?: string | null;
}

interface RoleAssignment {
  id: string;
  roleDefinitionId: string;
  principalId: string;
  directoryScopeId?: string;
}
interface RoleDefinition {
  id: string;
  displayName?: string;
  description?: string;
}
interface EligibilityInstance extends RoleAssignment {
  endDateTime?: string | null;
  roleDefinition?: RoleDefinition;
}

// Needs a token with RoleManagement.Read.Directory. Eligible roles are only
// returned on tenants with Entra ID P2 (PIM); elsewhere that part is empty.
export async function listUserRoles(token: TokenFn, id: string): Promise<Related<UserRole>> {
  const filter = `principalId eq '${id}'`;
  const active = graphAll<RoleAssignment>(
    token,
    '/roleManagement/directory/transitiveRoleAssignments',
    {
      advanced: true,
      query: { $filter: filter },
    },
  ).catch(() =>
    // Direct assignments only, where the transitive listing isn't available.
    graphAll<RoleAssignment>(token, '/roleManagement/directory/roleAssignments', {
      query: { $filter: filter },
    }),
  );
  const eligible = graphAll<EligibilityInstance>(
    token,
    '/roleManagement/directory/roleEligibilityScheduleInstances',
    { query: { $filter: filter, $expand: 'roleDefinition' } },
  ).catch(() => [] as EligibilityInstance[]);
  const definitions = graphAll<RoleDefinition>(token, '/roleManagement/directory/roleDefinitions', {
    query: { $select: 'id,displayName,description' },
  });
  const [act, eli, defs] = await Promise.all([active, eligible, definitions]);
  const byId = new Map(defs.map((d) => [d.id, d]));
  const groups = await resolveDirectoryObjects(
    token,
    act.filter((a) => a.principalId !== id).map((a) => a.principalId),
  );
  const row = (
    a: RoleAssignment,
    state: UserRole['state'],
    def: RoleDefinition | undefined,
    until?: string | null,
  ): UserRole => ({
    key: `${state}-${a.id}`,
    roleDefinitionId: a.roleDefinitionId,
    name: def?.displayName ?? a.roleDefinitionId,
    description: def?.description,
    state,
    scope: a.directoryScopeId ?? '/',
    viaGroup:
      a.principalId !== id
        ? { id: a.principalId, name: groups[a.principalId]?.displayName ?? a.principalId }
        : undefined,
    until,
  });
  const items = [
    ...act.map((a) => row(a, 'active', byId.get(a.roleDefinitionId))),
    ...eli.map((e) =>
      row(e, 'eligible', e.roleDefinition ?? byId.get(e.roleDefinitionId), e.endDateTime),
    ),
  ].sort((a, b) => a.name.localeCompare(b.name) || a.state.localeCompare(b.state));
  return { items, more: false };
}

// ----------------------------------------------------------------- Devices

export interface UserDevice {
  id: string;
  displayName?: string;
  operatingSystem?: string | null;
  operatingSystemVersion?: string | null;
  // AzureAd (joined), ServerAd (hybrid joined) or Workplace (registered).
  trustType?: string | null;
  isCompliant?: boolean | null;
  isManaged?: boolean | null;
  accountEnabled?: boolean | null;
  approximateLastSignInDateTime?: string | null;
}

const DEVICE_SELECT =
  'id,displayName,operatingSystem,operatingSystemVersion,trustType,isCompliant,isManaged,' +
  'accountEnabled,approximateLastSignInDateTime';

// Devices the user registered or owns, most recently used first.
export async function listUserDevices(token: TokenFn, id: string): Promise<Related<UserDevice>> {
  const [registered, owned] = await Promise.all([
    list<UserDevice>(token, `/users/${id}/registeredDevices`, DEVICE_SELECT),
    list<UserDevice>(token, `/users/${id}/ownedDevices`, DEVICE_SELECT),
  ]);
  const byId = new Map<string, UserDevice>();
  for (const d of [...registered.items, ...owned.items]) byId.set(d.id, d);
  const items = [...byId.values()].sort((a, b) =>
    (b.approximateLastSignInDateTime ?? '').localeCompare(a.approximateLastSignInDateTime ?? ''),
  );
  return { items, more: registered.more || owned.more };
}

// ---------------------------------------------------------------- Sign-ins

// The latest interactive sign-ins. Needs AuditLog.Read.All, a role that can
// read the sign-in log and Entra ID P1/P2.
export async function listUserSignIns(token: TokenFn, id: string): Promise<Related<SignIn>> {
  try {
    const res = await graph<{ value?: SignIn[] }>(token, '/auditLogs/signIns', {
      query: { $filter: `userId eq '${id}'`, $top: 25, $orderby: 'createdDateTime desc' },
    });
    return { items: res.value ?? [], more: (res.value?.length ?? 0) >= 25 };
  } catch (e) {
    if (denied(e))
      throw new Error(
        'The sign-in log can’t be read: it needs an Entra ID P1 or P2 licence and a role such as Reports Reader or Security Reader.',
      );
    throw e;
  }
}

// --------------------------------------------------------- Sign-in methods

export interface RegistrationDetails {
  isMfaRegistered?: boolean;
  isMfaCapable?: boolean;
  isPasswordlessCapable?: boolean;
  isSsprRegistered?: boolean;
  isSsprEnabled?: boolean;
  isAdmin?: boolean;
  methodsRegistered?: string[];
  defaultMfaMethod?: string | null;
  userPreferredMethodForSecondaryAuthentication?: string | null;
  lastUpdatedDateTime?: string | null;
}

// From the authentication-methods usage report (AuditLog.Read.All, Entra ID
// P1/P2); null when it can't be read. The report trails reality by some hours.
export async function getRegistrationDetails(
  token: TokenFn,
  id: string,
): Promise<RegistrationDetails | null> {
  try {
    return await graph<RegistrationDetails>(
      token,
      `/reports/authenticationMethods/userRegistrationDetails/${id}`,
    );
  } catch (e) {
    if (denied(e)) return null;
    throw e;
  }
}

// -------------------------------------------------------------------- Risk

export interface UserRisk {
  // none, low, medium, high, hidden.
  riskLevel?: string;
  // none, atRisk, confirmedCompromised, remediated, dismissed, confirmedSafe.
  riskState?: string;
  riskDetail?: string;
  riskLastUpdatedDateTime?: string | null;
}

// Needs a token with IdentityRiskyUser.Read.All and Entra ID P2. null when
// the tenant or the signed-in admin isn't allowed to read it; a user who was
// never flagged simply has no record.
export async function getUserRisk(token: TokenFn, id: string): Promise<UserRisk | null> {
  try {
    return await graph<UserRisk>(token, `/identityProtection/riskyUsers/${id}`, {
      query: { $select: 'riskLevel,riskState,riskDetail,riskLastUpdatedDateTime' },
    });
  } catch (e) {
    if (e instanceof GraphError && e.status === 404)
      return { riskLevel: 'none', riskState: 'none' };
    if (denied(e)) return null;
    throw e;
  }
}

// ----------------------------------------------------------------- Actions

// Invalidates the user's refresh tokens and session cookies: every app asks
// them to sign in again (access tokens already issued run out within the
// hour). Needs User.ReadWrite.All.
export function revokeSignInSessions(token: TokenFn, id: string) {
  return graph<unknown>(token, `/users/${id}/revokeSignInSessions`, { method: 'POST' });
}

// Sends the invitation mail to a guest again. Needs User.Invite.All.
export function resendInvitation(token: TokenFn, email: string, tenantId?: string) {
  return graph<unknown>(token, '/invitations', {
    method: 'POST',
    body: {
      invitedUserEmailAddress: email,
      inviteRedirectUrl: `https://myapplications.microsoft.com/${tenantId ? `?tenantid=${tenantId}` : ''}`,
      sendInvitationMessage: true,
    },
  });
}
