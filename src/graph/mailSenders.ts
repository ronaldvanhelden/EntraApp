import { graph, graphAll } from './client';
import { resolveDirectoryObjects } from './directoryObjects';
import {
  createOAuth2Grant,
  deleteAppRoleAssignedTo,
  deleteOAuth2Grant,
  listOAuth2GrantsForClient,
  updateOAuth2Grant,
} from './servicePrincipals';
import type { AppRole, AppRoleAssignment, OAuth2PermissionGrant } from './types';

type TokenFn = () => Promise<string>;

// Which enterprise apps hold a permission that lets them send mail as a given
// user. Two routes exist, and both are read from the directory:
//
//  - delegated: the app was consented a "send mail" scope, for everyone
//    (admin consent) or for this user. It can then send as the user while the
//    user is signed in to it.
//  - app-only: the app holds a "send mail" application permission. It needs
//    no user at all and can send as any mailbox in the tenant.
//
// What the directory does NOT show: Exchange can narrow an app-only permission
// to certain mailboxes (application access policies / RBAC for applications),
// and Send As / Send on Behalf rights between mailboxes live in Exchange too.

const RESOURCES = [
  { appId: '00000003-0000-0000-c000-000000000000', name: 'Microsoft Graph' },
  { appId: '00000002-0000-0ff1-ce00-000000000000', name: 'Exchange Online' },
];

// Delegated scopes that include sending mail as the signed-in user.
const SEND_SCOPES = new Set(
  ['Mail.Send', 'Mail.Send.Shared', 'SMTP.Send', 'EWS.AccessAsUser.All', 'full_access_as_user'].map(
    (s) => s.toLowerCase(),
  ),
);
// Application permissions that include sending mail as any user.
const SEND_ROLES = new Set(['Mail.Send', 'full_access_as_app'].map((s) => s.toLowerCase()));

export interface MailSender {
  key: string;
  // Object id of the enterprise app (service principal).
  spId: string;
  name: string;
  mode: 'app' | 'delegated';
  // e.g. "Mail.Send (Microsoft Graph)".
  permissions: string[];
  // Delegated only: consented by an admin for every user, or for this user.
  consent?: 'admin' | 'user';
  // The directory objects behind this row; what removing it deletes or edits.
  sources: MailSenderSource[];
}

export type MailSenderSource =
  // An app role assignment on the resource (Graph / Exchange Online).
  | { kind: 'role'; resourceId: string; assignmentId: string }
  // A delegated grant; only its send-mail scopes are taken out.
  | { kind: 'grant'; id: string; scope: string };

interface Resource {
  id: string;
  name: string;
  appRoles: AppRole[];
}

async function loadResource(token: TokenFn, appId: string, name: string): Promise<Resource | null> {
  try {
    const sp = await graph<{ id: string; appRoles?: AppRole[] }>(
      token,
      `/servicePrincipals(appId='${appId}')`,
      { query: { $select: 'id,appRoles' } },
    );
    return { id: sp.id, name, appRoles: sp.appRoles ?? [] };
  } catch {
    // Not present in this tenant (Exchange Online without Exchange).
    return null;
  }
}

export async function listMailSenders(token: TokenFn, userId: string): Promise<MailSender[]> {
  const resources = (
    await Promise.all(RESOURCES.map((r) => loadResource(token, r.appId, r.name)))
  ).filter((r): r is Resource => r !== null);
  const byId = new Map(resources.map((r) => [r.id, r]));

  const grants = (filter: string) =>
    graphAll<OAuth2PermissionGrant>(token, '/oauth2PermissionGrants', {
      query: { $filter: filter, $top: 999 },
    });
  const [forUser, forAll, ...assigned] = await Promise.all([
    grants(`principalId eq '${userId}'`),
    grants("consentType eq 'AllPrincipals'"),
    ...resources.map(async (r) => {
      const roles = new Map(
        r.appRoles
          .filter((role) => SEND_ROLES.has((role.value ?? '').toLowerCase()))
          .map((role) => [role.id, role.value]),
      );
      if (roles.size === 0) return [];
      const all = await graphAll<AppRoleAssignment>(
        token,
        `/servicePrincipals/${r.id}/appRoleAssignedTo`,
        { query: { $top: 999 } },
        20,
      );
      return all
        .filter((a) => roles.has(a.appRoleId))
        .map((a) => ({ a, resourceId: r.id, permission: `${roles.get(a.appRoleId)} (${r.name})` }));
    }),
  ]);

  const out = new Map<string, MailSender>();
  const add = (
    spId: string,
    mode: MailSender['mode'],
    name: string,
    permissions: string[],
    source: MailSenderSource,
    consent?: 'admin' | 'user',
  ) => {
    // Admin consent and consent for this user stay separate rows: removing
    // one affects everybody, the other only this user.
    const key = `${mode}-${consent ?? ''}-${spId}`;
    const row = out.get(key) ?? { key, spId, name, mode, permissions: [], consent, sources: [] };
    for (const p of permissions) if (!row.permissions.includes(p)) row.permissions.push(p);
    row.sources.push(source);
    out.set(key, row);
  };

  for (const { a, resourceId, permission } of assigned.flat()) {
    add(a.principalId, 'app', a.principalDisplayName ?? a.principalId, [permission], {
      kind: 'role',
      resourceId,
      assignmentId: a.id,
    });
  }

  const delegated = [...forUser, ...forAll].flatMap((g) => {
    const resource = byId.get(g.resourceId);
    if (!resource) return [];
    const scopes = (g.scope ?? '').split(/\s+/).filter((s) => SEND_SCOPES.has(s.toLowerCase()));
    return scopes.length ? [{ g, permissions: scopes.map((s) => `${s} (${resource.name})`) }] : [];
  });
  const names = await resolveDirectoryObjects(token, [
    ...new Set(delegated.map((d) => d.g.clientId)),
  ]);
  for (const { g, permissions } of delegated) {
    add(
      g.clientId,
      'delegated',
      names[g.clientId]?.displayName ?? g.clientId,
      permissions,
      { kind: 'grant', id: g.id, scope: g.scope ?? '' },
      g.consentType === 'AllPrincipals' ? 'admin' : 'user',
    );
  }

  // The stronger route first, then by name.
  return [...out.values()].sort(
    (a, b) => Number(b.mode === 'app') - Number(a.mode === 'app') || a.name.localeCompare(b.name),
  );
}

// Takes the send-mail permission behind a row away. A delegated grant keeps
// its other scopes; it is deleted only when nothing else is left in it.
export async function removeMailSender(token: TokenFn, sender: MailSender): Promise<void> {
  for (const src of sender.sources) {
    if (src.kind === 'role') {
      await deleteAppRoleAssignedTo(token, src.resourceId, src.assignmentId);
      continue;
    }
    const rest = src.scope.split(/\s+/).filter((s) => s && !SEND_SCOPES.has(s.toLowerCase()));
    if (rest.length === 0) await deleteOAuth2Grant(token, src.id);
    else await updateOAuth2Grant(token, src.id, { scope: rest.join(' ') });
  }
}

// Consents, on this one user's behalf, to the app sending mail as them
// (delegated Mail.Send on Microsoft Graph). The app can use it only while
// the user is signed in to it.
export async function allowMailSend(
  token: TokenFn,
  userId: string,
  clientSpId: string,
): Promise<void> {
  const graphSp = await loadResource(token, RESOURCES[0].appId, RESOURCES[0].name);
  if (!graphSp) throw new Error('Microsoft Graph was not found in this tenant.');
  const existing = (await listOAuth2GrantsForClient(token, clientSpId)).find(
    (g) => g.consentType === 'Principal' && g.principalId === userId && g.resourceId === graphSp.id,
  );
  if (!existing) {
    await createOAuth2Grant(token, {
      clientId: clientSpId,
      consentType: 'Principal',
      principalId: userId,
      resourceId: graphSp.id,
      scope: 'Mail.Send',
    });
    return;
  }
  const scopes = (existing.scope ?? '').split(/\s+/).filter(Boolean);
  if (scopes.some((s) => s.toLowerCase() === 'mail.send')) return;
  await updateOAuth2Grant(token, existing.id, { scope: [...scopes, 'Mail.Send'].join(' ') });
}
