// Application (app-only) permissions that deserve a second look, per API.
//
// Microsoft's own privilegeLevel (see graph/scopeCatalog.ts) covers Microsoft
// Graph only and rates most read-everything permissions the same as the ones
// that lead to Global Administrator, so it can't single those out. This list
// does, in two tiers:
//
//   critical — the holder can make itself (or anything it controls) a Global
//              Administrator, or sign in as any user. Owning the app's
//              credentials is owning the tenant.
//   high     — tenant-wide access to content (mail, files, chats), or write
//              access that changes who can do what without reaching Global
//              Administrator on its own.
//
// An app-only permission applies to the whole tenant and isn't bound by
// Conditional Access for users or by MFA, which is why these are worth
// reviewing even on apps that look dormant.

export type RiskTier = 'critical' | 'high';

export interface DangerousPermission {
  tier: RiskTier;
  // What the holder can do with it, as one sentence.
  why: string;
}

export const MS_GRAPH = '00000003-0000-0000-c000-000000000000';
export const AAD_GRAPH = '00000002-0000-0000-c000-000000000000';
export const EXCHANGE_ONLINE = '00000002-0000-0ff1-ce00-000000000000';
export const SHAREPOINT_ONLINE = '00000003-0000-0ff1-ce00-000000000000';

// The APIs whose app-only grants are scanned, in the order they are listed.
export const SCANNED_RESOURCES: { appId: string; name: string }[] = [
  { appId: MS_GRAPH, name: 'Microsoft Graph' },
  { appId: EXCHANGE_ONLINE, name: 'Exchange Online' },
  { appId: SHAREPOINT_ONLINE, name: 'SharePoint Online' },
  { appId: AAD_GRAPH, name: 'Azure AD Graph (legacy)' },
];

const critical = (why: string): DangerousPermission => ({ tier: 'critical', why });
const high = (why: string): DangerousPermission => ({ tier: 'high', why });

const MAIL_READ = 'Read the mail in every mailbox.';
const MAIL_WRITE = 'Read, change and delete the mail in every mailbox.';
const MAIL_SEND = 'Send mail as any user.';
const SITES_READ = 'Read every SharePoint site and the files in it.';
const SITES_WRITE = 'Read and change every SharePoint site and the files in it.';
const SITES_FULL =
  'Full control of every SharePoint site, including its permissions and the files in it.';

const GRAPH: Record<string, DangerousPermission> = {
  // --- Routes to Global Administrator
  'RoleManagement.ReadWrite.Directory': critical(
    'Assign any directory role, Global Administrator included, to itself or anyone else.',
  ),
  'AppRoleAssignment.ReadWrite.All': critical(
    'Grant itself any application permission without admin consent, including the one that assigns directory roles.',
  ),
  'Application.ReadWrite.All': critical(
    'Add a secret or certificate to any application and sign in as it, taking over that application’s permissions.',
  ),
  'RoleEligibilitySchedule.ReadWrite.Directory': critical(
    'Make anyone eligible for any directory role through PIM.',
  ),
  'RoleAssignmentSchedule.ReadWrite.Directory': critical(
    'Create active assignments for any directory role through PIM.',
  ),
  'PrivilegedAccess.ReadWrite.AzureADGroup': critical(
    'Add members and owners to role-assignable groups through PIM for Groups.',
  ),
  'PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup': critical(
    'Add active members and owners to role-assignable groups through PIM for Groups.',
  ),
  'PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup': critical(
    'Make anyone eligible for membership of role-assignable groups through PIM for Groups.',
  ),
  'UserAuthenticationMethod.ReadWrite.All': critical(
    'Add a sign-in method (such as a Temporary Access Pass) to any user, administrators included, and sign in as them.',
  ),
  'Policy.ReadWrite.PermissionGrant': critical(
    'Change the consent policies so that it, or any user, may consent to any permission.',
  ),
  'Domain.ReadWrite.All': critical(
    'Add or change a federated domain and issue sign-in tokens for any user.',
  ),
  'Organization.ReadWrite.All': critical(
    'Change tenant-wide settings, including the certificate authorities trusted for certificate-based sign-in.',
  ),

  // --- Changes who can do what
  'Directory.ReadWrite.All': high(
    'Change most directory objects: users, groups and their members, devices and settings.',
  ),
  'Application.ReadWrite.OwnedBy': high(
    'Create applications and add credentials to the ones it owns.',
  ),
  'DelegatedPermissionGrant.ReadWrite.All': high(
    'Grant any delegated permission to any application on behalf of every user.',
  ),
  'User.ReadWrite.All': high('Create, change and delete users.'),
  'User-PasswordProfile.ReadWrite.All': high('Reset the password of users.'),
  'User.EnableDisableAccount.All': high('Block or unblock sign-in for any user.'),
  'User.DeleteRestore.All': high('Delete and restore users.'),
  'Group.ReadWrite.All': high(
    'Change every group and its members, and read group mail and files. Groups often hand out access elsewhere.',
  ),
  'GroupMember.ReadWrite.All': high(
    'Add anyone to any group that isn’t role-assignable. Groups often hand out access elsewhere.',
  ),
  'AdministrativeUnit.ReadWrite.All': high(
    'Change administrative units and who is in them, moving users in or out of an admin’s reach.',
  ),
  'EntitlementManagement.ReadWrite.All': high(
    'Change access packages and assign them, handing out the groups, apps and roles they carry.',
  ),
  'Policy.ReadWrite.ConditionalAccess': high(
    'Change or remove Conditional Access policies, MFA requirements included.',
  ),
  'Policy.ReadWrite.AuthenticationMethod': high(
    'Change which sign-in methods the tenant allows, and for whom.',
  ),
  'Policy.ReadWrite.Authorization': high(
    'Change the tenant’s authorization policy, such as what ordinary users and guests may do.',
  ),
  'Policy.ReadWrite.CrossTenantAccess': high(
    'Change which other tenants are trusted, and whether their MFA is accepted.',
  ),
  'Policy.ReadWrite.SecurityDefaults': high('Turn security defaults off.'),
  'IdentityProvider.ReadWrite.All': high('Add or change the identity providers the tenant trusts.'),
  'Domain-InternalFederation.ReadWrite.All': high('Change the federation settings of a domain.'),
  'Synchronization.ReadWrite.All': high(
    'Change provisioning and directory synchronization jobs, which write to users and groups.',
  ),
  'Device.ReadWrite.All': high('Change and delete every device in the directory.'),
  'DeviceManagementConfiguration.ReadWrite.All': high(
    'Change Intune configuration and push scripts to managed devices, admin workstations included.',
  ),
  'DeviceManagementScripts.ReadWrite.All': high(
    'Push scripts to Intune-managed devices, admin workstations included.',
  ),
  'DeviceManagementApps.ReadWrite.All': high('Deploy apps to Intune-managed devices.'),
  'DeviceManagementRBAC.ReadWrite.All': high('Change Intune roles and who holds them.'),
  'DeviceManagementManagedDevices.PrivilegedOperations.All': high(
    'Wipe, reset the passcode of, or remote-lock any managed device.',
  ),

  // --- Tenant-wide content
  'Mail.Read': high(MAIL_READ),
  'Mail.ReadWrite': high(MAIL_WRITE),
  'Mail.Send': high(MAIL_SEND),
  'MailboxSettings.ReadWrite': high(
    'Change the settings of every mailbox, such as rules that forward mail elsewhere.',
  ),
  'Files.Read.All': high('Read every file in OneDrive and SharePoint.'),
  'Files.ReadWrite.All': high('Read, change and delete every file in OneDrive and SharePoint.'),
  'Sites.Read.All': high(SITES_READ),
  'Sites.ReadWrite.All': high(SITES_WRITE),
  'Sites.Manage.All': high('Create, change and delete lists and libraries in every site.'),
  'Sites.FullControl.All': high(SITES_FULL),
  'Chat.Read.All': high('Read every Teams chat message.'),
  'Chat.ReadWrite.All': high('Read and change every Teams chat message.'),
  'ChannelMessage.Read.All': high('Read every Teams channel message.'),
  'Calendars.ReadWrite': high('Read and change every calendar.'),
  'Contacts.ReadWrite': high('Read and change the contacts of every user.'),
  'Notes.ReadWrite.All': high('Read and change every OneNote notebook.'),
  'eDiscovery.ReadWrite.All': high(
    'Run eDiscovery searches over mail, files and chats, and export the results.',
  ),
};

const EXCHANGE: Record<string, DangerousPermission> = {
  full_access_as_app: high(
    'Full access to every mailbox through Exchange Web Services: read, send, delete.',
  ),
  'Exchange.ManageAsApp': high(
    'Run Exchange Online PowerShell as the app. What it can change follows from the Exchange or directory role it holds.',
  ),
  'Mail.Read': high(MAIL_READ),
  'Mail.ReadWrite': high(MAIL_WRITE),
  'Mail.Send': high(MAIL_SEND),
  'IMAP.AccessAsApp': high(
    'Read and change mail over IMAP in the mailboxes it is given access to.',
  ),
  'POP.AccessAsApp': high('Read mail over POP in the mailboxes it is given access to.'),
};

const SHAREPOINT: Record<string, DangerousPermission> = {
  'Sites.Read.All': high(SITES_READ),
  'Sites.ReadWrite.All': high(SITES_WRITE),
  'Sites.Manage.All': high('Create, change and delete lists and libraries in every site.'),
  'Sites.FullControl.All': high(SITES_FULL),
};

// The retired Azure AD Graph API. Grants on it linger on older apps; the same
// names carry the same reach as on Microsoft Graph.
const AAD: Record<string, DangerousPermission> = {
  'Application.ReadWrite.All': GRAPH['Application.ReadWrite.All'],
  'Domain.ReadWrite.All': GRAPH['Domain.ReadWrite.All'],
  'Directory.ReadWrite.All': GRAPH['Directory.ReadWrite.All'],
  'Application.ReadWrite.OwnedBy': GRAPH['Application.ReadWrite.OwnedBy'],
};

const lower = (m: Record<string, DangerousPermission>) =>
  Object.fromEntries(Object.entries(m).map(([k, v]) => [k.toLowerCase(), v]));

const BY_RESOURCE: Record<string, Record<string, DangerousPermission>> = {
  [MS_GRAPH]: lower(GRAPH),
  [EXCHANGE_ONLINE]: lower(EXCHANGE),
  [SHAREPOINT_ONLINE]: lower(SHAREPOINT),
  [AAD_GRAPH]: lower(AAD),
};

// The rating of an application permission, or null when it isn't on the list.
export function getDangerousPermission(
  resourceAppId: string | undefined,
  value: string | undefined,
): DangerousPermission | null {
  if (!resourceAppId || !value) return null;
  return BY_RESOURCE[resourceAppId.toLowerCase()]?.[value.toLowerCase()] ?? null;
}

// Tenants Microsoft publishes its own applications from.
const MICROSOFT_TENANTS = new Set([
  'f8cdef31-a31e-4b4a-93e4-5f571e91255a', // Microsoft Services
  '72f988bf-86f1-41af-91ab-2d7cd011db47', // Microsoft
]);

export function isMicrosoftTenant(tenantId: string | undefined | null): boolean {
  return Boolean(tenantId && MICROSOFT_TENANTS.has(tenantId.toLowerCase()));
}
