import { Fragment, useEffect, useState, type ReactNode } from 'react';
import { Link, useLocation, useNavigate, useParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import { useGraphToken } from '../auth/useGraphToken';
import { useCurrentTenantId } from '../auth/useCurrentTenantId';
import {
  GROUP_WRITE_SCOPES,
  PIM_ASSIGN_SCOPES,
  RISK_READ_SCOPES,
  ROLE_READ_SCOPES,
  USER_INVITE_SCOPES,
  USER_WRITE_SCOPES,
} from '../auth/config';
import {
  deleteGroup,
  deleteUser,
  getAuditTrail,
  getGroup,
  getUser,
  getUserSignInActivity,
  listAppAssignments,
  listGroupMembers,
  listGroupOwners,
  addGroupMember,
  removeGroupMember,
  listUserGroups,
  updateGroup,
  updateUser,
  type AuditTrail,
  type DirectoryMember,
  type GroupDetail as GroupData,
  type Related,
  type UserDetail as UserData,
  type UserPatch,
  type UserSignInActivity,
} from '../graph/principals';
import type { DirectoryAuditEvent } from '../graph/directoryAudits';
import {
  allowMailSend,
  listMailSenders,
  removeMailSender,
  type MailSender,
} from '../graph/mailSenders';
import {
  getRegistrationDetails,
  getUserManager,
  getUserRisk,
  listDirectReports,
  listOwnedObjects,
  listUserConsents,
  listUserDevices,
  listUserLicenses,
  listUserRoles,
  listUserSignIns,
  resendInvitation,
  revokeSignInSessions,
  setUserManager,
  type RegistrationDetails,
  type UserManager,
  type UserRisk,
} from '../graph/userInsights';
import { deleteOAuth2Grant } from '../graph/servicePrincipals';
import type { PrincipalRef } from '../graph/directoryObjects';
import { isGuest } from '../graph/quickSearch';
import type { AppRoleAssignment } from '../graph/types';
import { entraGroupUrl, entraUserUrl } from '../lib/entraLinks';
import { ChainLinkIcon, ExternalLinkIcon } from '../components/ButtonIcons';
import { useConfirm, type ConfirmOptions } from '../components/ConfirmDialog';
import { ConsentGate } from '../components/ConsentGate';
import { CopyButton } from '../components/CopyButton';
import { AssignRole } from '../components/PimDialogs';
import { removeRole } from '../graph/pim';
import { groupKind } from '../components/DirectorySearch';
import { GroupSiteLink, useGroupSites } from '../components/GroupSites';
import { Identicon } from '../components/Identicon';
import { Modal } from '../components/Modal';
import { PrincipalPicker } from '../components/PrincipalPicker';
import { RemoveButton } from '../components/RemoveButton';
import { Check, Field, ResetPassword } from '../components/PrincipalEditors';

// Detail pages for a user (/users/:id) and a group (/groups/:id), kept to one
// screen where possible: editable properties beside the read-only details
// (ids, who created and last changed the object, last sign-in), and below
// them one tabbed card showing how the object ties into the directory —
// groups, members, owners and the enterprise apps it is assigned to; for a
// user also roles, owned objects, consents, licences, devices, sign-ins and
// sign-in methods. Anything beyond that is a click away in Entra.

type TokenFn = () => Promise<string>;

const message = (e: unknown) => (e instanceof Error ? e.message : String(e));
const dash = <span className="muted">—</span>;
const when = (iso?: string | null) => (iso ? new Date(iso).toLocaleString() : dash);

// Back to the list this page was opened from (search results keep their
// filters that way), or to the plain list.
function BackLink({ to, label }: { to: string; label: string }) {
  const nav = useNavigate();
  const location = useLocation();
  const fromList = (location.state as { fromList?: string } | null)?.fromList;
  return (
    <button className="ghost back-link" onClick={() => nav(fromList ?? to)}>
      ← {fromList?.startsWith('/search') ? 'Search results' : label}
    </button>
  );
}

function EntraButton({ href }: { href: string }) {
  return (
    <a href={href} target="_blank" rel="noreferrer" style={{ textDecoration: 'none' }}>
      <button type="button" className="btn-with-icon">
        <ExternalLinkIcon />
        Open in Entra
      </button>
    </a>
  );
}

function IdRow({
  label,
  value,
  sync,
}: {
  label: string;
  value?: string | null;
  // Key in SHAREPOINT_PROFILE when the value is synced to SharePoint.
  sync?: string;
}) {
  return (
    <>
      <div className="k">
        {label}
        {sync && <SharePointSync property={sync} />}
      </div>
      <div className={value ? 'mono' : undefined}>
        {value ? (
          <>
            {value}
            <CopyButton value={value} label={label.toLowerCase()} />
          </>
        ) : (
          dash
        )}
      </div>
    </>
  );
}

// A read-only list of related objects. Each loads on its own, so a slow or
// forbidden call doesn't hold up the rest of the page, and reports how many
// it found for the tab label.
function RelatedPanel<T>({
  load,
  deps,
  columns,
  row,
  empty,
  rowKey,
  to,
  onCount,
  note,
  remove,
}: {
  // Adds a bin button to each row: asks first, then drops the row.
  remove?: {
    label: string;
    // Rows this returns false for get no button.
    canRemove?: (item: T) => boolean;
    confirm: (item: T) => ConfirmOptions;
    run: (item: T) => Promise<void>;
  };
  // Shown above the list once it has loaded: what the list does and doesn't say.
  note?: ReactNode;
  load: () => Promise<Related<T>>;
  deps: unknown[];
  columns: string[];
  row: (item: T) => ReactNode[];
  empty: string;
  rowKey: (item: T) => string;
  // Where a click on the row goes; rows without a target don't react.
  to?: (item: T) => string | undefined;
  onCount: (label: string) => void;
}) {
  const nav = useNavigate();
  const { confirm, confirmDialog } = useConfirm();
  const [data, setData] = useState<Related<T> | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const removeRow = async (item: T) => {
    if (!remove || !(await confirm(remove.confirm(item)))) return;
    const key = rowKey(item);
    setBusy(key);
    setError(null);
    try {
      await remove.run(item);
      setData((prev) => {
        if (!prev) return prev;
        const items = prev.items.filter((i) => rowKey(i) !== key);
        onCount(`${items.length}${prev.more ? '+' : ''}`);
        return { ...prev, items };
      });
    } catch (e: unknown) {
      setError(message(e));
    } finally {
      setBusy(null);
    }
  };
  useEffect(() => {
    let cancelled = false;
    setData(null);
    setError(null);
    onCount('');
    load()
      .then((r) => {
        if (cancelled) return;
        setData(r);
        onCount(`${r.items.length}${r.more ? '+' : ''}`);
      })
      .catch((e: unknown) => !cancelled && setError(message(e)));
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);

  if (error && !data) return <div className="error related-note">{error}</div>;
  if (!data)
    return (
      <div className="related-note">
        <span className="spinner" />
      </div>
    );
  const noteRow = note && (
    <div className="muted related-note" style={{ fontSize: 12, paddingBottom: 0 }}>
      {note}
    </div>
  );
  if (data.items.length === 0)
    return (
      <>
        {noteRow}
        <div className="muted related-note">{empty}</div>
      </>
    );
  return (
    <>
      {noteRow}
      {error && <div className="error related-note">{error}</div>}
      <table className="table">
        <thead>
          <tr>
            {columns.map((c) => (
              <th key={c}>{c}</th>
            ))}
            {remove && <th></th>}
          </tr>
        </thead>
        <tbody>
          {data.items.map((item) => {
            const target = to?.(item);
            return (
              <tr
                key={rowKey(item)}
                className={target ? undefined : 'static'}
                onClick={target ? () => nav(target) : undefined}
              >
                {row(item).map((cell, i) => (
                  <td key={columns[i]}>{cell}</td>
                ))}
                {remove && (
                  <td className="row-actions" onClick={(e) => e.stopPropagation()}>
                    {remove.canRemove?.(item) !== false && (
                      <RemoveButton
                        label={remove.label}
                        busy={busy === rowKey(item)}
                        onClick={() => void removeRow(item)}
                      />
                    )}
                  </td>
                )}
              </tr>
            );
          })}
        </tbody>
      </table>
      {data.more && (
        <div className="muted related-note" style={{ fontSize: 12 }}>
          Showing the first {data.items.length} — the full list is in the Entra admin center.
        </div>
      )}
      {confirmDialog}
    </>
  );
}

interface RelatedTab {
  key: string;
  title: string;
  // Only loads once the tab is opened (no count in the label until then):
  // for lists that are slow or that few visits need.
  lazy?: boolean;
  panel: (onCount: (label: string) => void) => ReactNode;
}

// One card for all related lists. Panels stay mounted (hidden when not
// selected) so the counts in the tab labels fill in without a click; lazy
// ones join in when first opened.
function RelatedTabs({ tabs }: { tabs: RelatedTab[] }) {
  const [active, setActiveKey] = useState(tabs[0].key);
  const [opened, setOpened] = useState<string[]>([]);
  const setActive = (key: string) => {
    setActiveKey(key);
    setOpened((prev) => (prev.includes(key) ? prev : [...prev, key]));
  };
  const [counts, setCounts] = useState<Record<string, string>>({});
  return (
    <div className="card related-card">
      <div className="tabs">
        {tabs.map((t) => (
          <button
            key={t.key}
            className={active === t.key ? 'active' : ''}
            onClick={() => setActive(t.key)}
          >
            {t.title}
            {counts[t.key] && <span className="muted"> ({counts[t.key]})</span>}
          </button>
        ))}
      </div>
      {tabs.map((t) => (
        <div key={t.key} hidden={active !== t.key}>
          {(!t.lazy || opened.includes(t.key)) &&
            t.panel((label) =>
              setCounts((prev) => (prev[t.key] === label ? prev : { ...prev, [t.key]: label })),
            )}
        </div>
      ))}
    </div>
  );
}

function NameCell({ id, name, to }: { id: string; name: string; to?: string }) {
  return (
    <div className="row" style={{ gap: 8, alignItems: 'center' }}>
      <Identicon id={id} size={20} title={name} />
      {to ? (
        <Link to={to} style={{ color: 'inherit' }} onClick={(e) => e.stopPropagation()}>
          {name}
        </Link>
      ) : (
        <span>{name}</span>
      )}
    </div>
  );
}

function memberTarget(m: DirectoryMember): [kind: string, to: string | undefined] {
  const type = (m['@odata.type'] ?? '').toLowerCase();
  return type.includes('user')
    ? ['User', `/users/${m.id}`]
    : type.includes('group')
      ? ['Group', `/groups/${m.id}`]
      : type.includes('serviceprincipal')
        ? ['Enterprise app', `/enterprise-apps/${m.id}`]
        : type.endsWith('.application')
          ? ['App registration', `/applications/${m.id}`]
          : [type.split('.').pop() || 'Object', undefined];
}

function memberCells(m: DirectoryMember): ReactNode[] {
  const [kind, to] = memberTarget(m);
  return [
    <NameCell id={m.id} name={m.displayName ?? m.id} to={to} />,
    <span className="muted">{kind}</span>,
    m.userPrincipalName ?? m.mail ?? m.appId ?? dash,
  ];
}

function appAssignmentsTab(token: TokenFn, kind: 'users' | 'groups', id: string): RelatedTab {
  return {
    key: 'apps',
    title: 'Enterprise apps',
    panel: (onCount) => (
      <RelatedPanel<AppRoleAssignment>
        load={() => listAppAssignments(token, kind, id)}
        deps={[token, kind, id]}
        onCount={onCount}
        columns={['Assigned to', 'Since']}
        rowKey={(a) => a.id}
        to={(a) => `/enterprise-apps/${a.resourceId}`}
        row={(a) => [
          <NameCell
            id={a.resourceId}
            name={a.resourceDisplayName ?? a.resourceId}
            to={`/enterprise-apps/${a.resourceId}`}
          />,
          when(a.createdDateTime),
        ]}
        empty={
          kind === 'users'
            ? 'Not assigned to any enterprise app directly — access can still come through a group.'
            : 'This group is not assigned to any enterprise app.'
        }
      />
    ),
  };
}

// -------------------------------------------------------------------- Roles

// The user's admin roles — held or eligible, directly or through a group —
// with assigning a new one and removing a direct one (both through PIM).
function UserRolesPanel({
  token,
  userId,
  userName,
  onCount,
}: {
  token: TokenFn;
  userId: string;
  userName: string;
  onCount: (label: string) => void;
}) {
  const assignToken = useGraphToken(PIM_ASSIGN_SCOPES);
  const [assigning, setAssigning] = useState(false);
  const [version, setVersion] = useState(0);
  return (
    <>
      <div className="row related-note" style={{ justifyContent: 'space-between', gap: 12 }}>
        <div className="muted" style={{ fontSize: 12 }}>
          Eligible roles have to be activated before use; the user does that under{' '}
          <Link to="/pim">Privileged roles</Link>.
        </div>
        <button style={{ flexShrink: 0 }} onClick={() => setAssigning(true)}>
          + Assign role
        </button>
      </div>
      <RelatedPanel
        load={() => listUserRoles(token, userId)}
        deps={[token, userId, version]}
        onCount={onCount}
        columns={['Admin role', 'Status', 'Scope', 'Assigned']}
        rowKey={(r) => r.key}
        row={(r) => [
          <span title={r.description}>{r.name}</span>,
          r.state === 'active' ? (
            <span className="badge granted">Active</span>
          ) : (
            <span className="badge pending" title="Can be activated through PIM">
              Eligible
              {r.until ? ` until ${new Date(r.until).toLocaleDateString()}` : ''}
            </span>
          ),
          r.scope === '/' ? 'Whole directory' : <span className="mono">{r.scope}</span>,
          r.viaGroup ? (
            <>
              Through <Link to={`/groups/${r.viaGroup.id}`}>{r.viaGroup.name}</Link>
            </>
          ) : (
            'Directly'
          ),
        ]}
        remove={{
          label: 'Remove this role',
          // A role that comes through a group is removed on the group.
          canRemove: (r) => !r.viaGroup,
          confirm: (r) => ({
            title: `Remove ${r.name} from ${userName}?`,
            message:
              r.state === 'eligible'
                ? `${userName} can no longer activate this role.`
                : `${userName} loses this role straight away. A role that is only active because it was activated can't be removed here — remove the eligibility instead.`,
            confirmLabel: 'Remove role',
          }),
          run: (r) =>
            removeRole(assignToken, {
              kind: r.state,
              principalId: userId,
              roleDefinitionId: r.roleDefinitionId,
              directoryScopeId: r.scope,
            }).then(() => undefined),
        }}
        empty="No admin roles."
      />
      {assigning && (
        <AssignRole
          principal={{ id: userId, displayName: userName, kind: 'user' }}
          onClose={() => setAssigning(false)}
          onDone={() => {
            setAssigning(false);
            // PIM applies the assignment a moment after accepting it.
            window.setTimeout(() => setVersion((v) => v + 1), 2000);
          }}
        />
      )}
    </>
  );
}

// ----------------------------------------------------------------- Security

const IDENTITY_PROVIDERS: Record<string, string> = {
  ExternalAzureAD: 'Another Microsoft Entra tenant',
  MicrosoftAccount: 'Microsoft account',
  mail: 'Email one-time passcode',
  'google.com': 'Google',
  'facebook.com': 'Facebook',
};

const DEVICE_TRUST: Record<string, string> = {
  AzureAd: 'Entra joined',
  ServerAd: 'Hybrid joined',
  Workplace: 'Registered',
};
const CA_STATUS: Record<string, string> = {
  success: 'Passed',
  failure: 'Blocked',
  notApplied: 'Not applied',
};

const METHOD_NAMES: Record<string, string> = {
  microsoftAuthenticatorPush: 'Authenticator app (push)',
  microsoftAuthenticatorPasswordless: 'Authenticator app (passwordless)',
  softwareOneTimePasscode: 'Authenticator app code',
  hardwareOneTimePasscode: 'Hardware token',
  mobilePhone: 'Phone (SMS or call)',
  alternateMobilePhone: 'Alternate phone',
  officePhone: 'Office phone',
  email: 'Email',
  fido2SecurityKey: 'FIDO2 security key',
  passKeyDeviceBound: 'Passkey (device-bound)',
  passKeyDeviceBoundAuthenticator: 'Passkey (Authenticator)',
  passKeyDeviceBoundWindowsHello: 'Passkey (Windows Hello)',
  windowsHelloForBusiness: 'Windows Hello for Business',
  temporaryAccessPass: 'Temporary Access Pass',
  securityQuestion: 'Security questions',
  macOsSecureEnclaveKey: 'macOS Platform SSO',
  x509CertificateSingleFactor: 'Certificate',
  x509CertificateMultiFactor: 'Certificate (multifactor)',
};
const methodName = (m?: string | null) => (m && m !== 'none' ? (METHOD_NAMES[m] ?? m) : null);

const yesNo = (v: boolean | undefined, good = true) =>
  v === undefined ? (
    dash
  ) : (
    <span className={`badge ${v === good ? 'granted' : 'expired'}`}>{v ? 'Yes' : 'No'}</span>
  );

const RISK_BADGE: Record<string, string> = { high: 'expired', medium: 'pending', low: 'pending' };

function RiskValue({ token, userId }: { token: TokenFn; userId: string }) {
  const [risk, setRisk] = useState<UserRisk | null | undefined>(undefined);
  useEffect(() => {
    let cancelled = false;
    setRisk(undefined);
    getUserRisk(token, userId)
      .then((r) => !cancelled && setRisk(r))
      .catch(() => !cancelled && setRisk(null));
    return () => {
      cancelled = true;
    };
  }, [token, userId]);
  if (risk === undefined) return spinner;
  if (risk === null)
    return (
      <span className="muted" title="Needs Entra ID P2 and a role such as Security Reader">
        Unavailable
      </span>
    );
  const level = risk.riskLevel ?? 'none';
  if (level === 'none' && (risk.riskState ?? 'none') === 'none')
    return <span className="badge granted">No risk recorded</span>;
  return (
    <>
      <span className={`badge ${RISK_BADGE[level] ?? 'granted'}`}>{level}</span>{' '}
      <span className="muted">
        {risk.riskState}
        {risk.riskDetail && risk.riskDetail !== 'none' && ` · ${risk.riskDetail}`}
        {risk.riskLastUpdatedDateTime &&
          ` · ${new Date(risk.riskLastUpdatedDateTime).toLocaleDateString()}`}
      </span>
    </>
  );
}

// How the user proves who they are (registered methods, MFA) and whether
// Identity Protection has flagged the account.
function SecurityPanel({ userId }: { userId: string }) {
  const token = useGraphToken();
  const [reg, setReg] = useState<RegistrationDetails | null | undefined>(undefined);
  useEffect(() => {
    let cancelled = false;
    setReg(undefined);
    getRegistrationDetails(token, userId)
      .then((r) => !cancelled && setReg(r))
      .catch(() => !cancelled && setReg(null));
    return () => {
      cancelled = true;
    };
  }, [token, userId]);

  const methods = reg?.methodsRegistered ?? [];
  return (
    <div className="related-note">
      <div className="kv compact">
        {reg === undefined ? (
          <>
            <div className="k">Sign-in methods</div>
            <div>{spinner}</div>
          </>
        ) : reg === null ? (
          <>
            <div className="k">Sign-in methods</div>
            <div className="muted">
              Unavailable — the registration report needs Entra ID P1 or P2 and a role such as
              Reports Reader.
            </div>
          </>
        ) : (
          <>
            <div className="k">Registered for MFA</div>
            <div>{yesNo(reg.isMfaRegistered)}</div>
            <div className="k" title="Registered and allowed by policy to use a strong method">
              Can do MFA
            </div>
            <div>{yesNo(reg.isMfaCapable)}</div>
            <div className="k">Sign-in methods</div>
            <div>
              {methods.length > 0 ? (
                methods.map((m) => methodName(m) ?? m).join(', ')
              ) : (
                <span className="muted">Password only</span>
              )}
            </div>
            <div className="k">Default MFA method</div>
            <div>
              {methodName(reg.userPreferredMethodForSecondaryAuthentication) ??
                methodName(reg.defaultMfaMethod) ??
                dash}
            </div>
            <div className="k">Passwordless</div>
            <div>{yesNo(reg.isPasswordlessCapable)}</div>
            <div className="k">Self-service reset</div>
            <div>{yesNo(reg.isSsprRegistered)}</div>
            <div className="k" title="The report is refreshed a few times a day">
              Report updated
            </div>
            <div>{when(reg.lastUpdatedDateTime)}</div>
          </>
        )}
        <div className="k">Risk</div>
        <div>
          <ConsentGate scopes={RISK_READ_SCOPES} what="the risk state" inline>
            {(riskToken) => <RiskValue token={riskToken} userId={userId} />}
          </ConsentGate>
        </div>
      </div>
    </div>
  );
}

// ------------------------------------------------------------ Mail senders

const RBAC_DOCS = 'https://learn.microsoft.com/exchange/permissions-exo/application-rbac';

// Apps that may send mail as the user, with the two things the directory lets
// us change: take a permission away, and consent for this user alone.
// Narrowing an app-only permission to certain mailboxes is an Exchange
// setting with no Graph API, so that part only points to the instructions.
function MailSendersPanel({
  userId,
  userName,
  onCount,
}: {
  userId: string;
  userName: string;
  onCount: (label: string) => void;
}) {
  const token = useGraphToken();
  const nav = useNavigate();
  const { confirm, confirmDialog } = useConfirm();
  const [rows, setRows] = useState<MailSender[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [adding, setAdding] = useState(false);
  const [reload, setReload] = useState(0);

  useEffect(() => {
    let cancelled = false;
    setError(null);
    listMailSenders(token, userId)
      .then((list) => {
        if (cancelled) return;
        setRows(list);
        onCount(String(list.length));
      })
      .catch((e: unknown) => !cancelled && setError(message(e)));
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [token, userId, reload]);

  const remove = async (m: MailSender) => {
    const everyone = m.mode === 'app' || m.consent === 'admin';
    const ok = await confirm({
      title: `Stop ${m.name} sending mail?`,
      message: (
        <div style={{ fontSize: 14 }}>
          <p style={{ marginTop: 0 }}>
            Removes <span className="mono">{m.permissions.join(', ')}</span> from{' '}
            <strong>{m.name}</strong>.
          </p>
          <p style={{ marginBottom: 0 }}>
            {everyone ? (
              <strong>
                This permission is tenant-wide: the app loses it for every user, not only for{' '}
                {userName}. Anything in the app that relies on it stops working.
              </strong>
            ) : (
              <>Only the consent given for {userName} is removed; other users are not affected.</>
            )}
          </p>
        </div>
      ),
      confirmLabel: everyone ? 'Remove for everyone' : 'Remove',
    });
    if (!ok) return;
    setBusy(m.key);
    setError(null);
    try {
      await removeMailSender(token, m);
      setRows((prev) => {
        const next = (prev ?? []).filter((r) => r.key !== m.key);
        onCount(String(next.length));
        return next;
      });
    } catch (e: unknown) {
      setError(`Couldn't remove the permission: ${message(e)}`);
      // Part of it may have gone through; show what is really there.
      setReload((n) => n + 1);
    } finally {
      setBusy(null);
    }
  };

  return (
    <>
      <div className="row related-note" style={{ justifyContent: 'space-between', gap: 12 }}>
        <div className="muted" style={{ fontSize: 12 }}>
          Apps holding a send-mail permission in Microsoft Graph or Exchange Online. Limiting an
          app-only permission to certain mailboxes is done in Exchange (
          <a href={RBAC_DOCS} target="_blank" rel="noreferrer">
            RBAC for Applications
          </a>
          ), as are Send As / Send on Behalf rights — neither can be seen or changed here.
        </div>
        <button style={{ flexShrink: 0 }} onClick={() => setAdding(true)}>
          + Allow an app
        </button>
      </div>
      {error && <div className="error related-note">{error}</div>}
      {!rows && !error && (
        <div className="related-note">
          <span className="spinner" />
        </div>
      )}
      {rows && rows.length === 0 && (
        <div className="muted related-note">No app has a permission to send mail as this user.</div>
      )}
      {rows && rows.length > 0 && (
        <table className="table">
          <thead>
            <tr>
              <th>App</th>
              <th>When</th>
              <th>Permission</th>
              <th>Granted</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {rows.map((m) => (
              <tr key={m.key} onClick={() => nav(`/enterprise-apps/${m.spId}`)}>
                <td>
                  <NameCell id={m.spId} name={m.name} to={`/enterprise-apps/${m.spId}`} />
                </td>
                <td>
                  {m.mode === 'app' ? (
                    <span className="badge app" title="Application permission: no user involved">
                      Any time, on its own
                    </span>
                  ) : (
                    <span
                      className="badge delegated"
                      title="Delegated permission: the app acts as the signed-in user"
                    >
                      While the user is signed in
                    </span>
                  )}
                </td>
                <td className="mono">{m.permissions.join(', ')}</td>
                <td>
                  {m.mode === 'app'
                    ? 'For every mailbox'
                    : m.consent === 'admin'
                      ? 'For all users (admin consent)'
                      : 'For this user only'}
                </td>
                <td className="row-actions" onClick={(e) => e.stopPropagation()}>
                  <RemoveButton
                    label={
                      m.mode === 'app' || m.consent === 'admin'
                        ? 'Remove this permission for everyone'
                        : 'Remove the consent for this user'
                    }
                    busy={busy === m.key}
                    onClick={() => void remove(m)}
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      {adding && (
        <AllowMailApp
          userId={userId}
          userName={userName}
          onClose={() => setAdding(false)}
          onAdded={() => {
            setAdding(false);
            setReload((n) => n + 1);
          }}
        />
      )}
      {confirmDialog}
    </>
  );
}

function AllowMailApp({
  userId,
  userName,
  onClose,
  onAdded,
}: {
  userId: string;
  userName: string;
  onClose: () => void;
  onAdded: () => void;
}) {
  const token = useGraphToken();
  const [app, setApp] = useState<PrincipalRef | null>(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const submit = async () => {
    if (!app) return;
    setSaving(true);
    setError(null);
    try {
      await allowMailSend(token, userId, app.id);
      onAdded();
    } catch (e: unknown) {
      setError(message(e));
      setSaving(false);
    }
  };

  return (
    <Modal
      title="Allow an app to send mail"
      onClose={() => !saving && onClose()}
      footer={
        <>
          <button disabled={saving} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={saving || !app} onClick={() => void submit()}>
            {saving ? 'Allowing…' : 'Allow'}
          </button>
        </>
      }
    >
      <p style={{ marginTop: 0, fontSize: 14 }}>
        Consents, on behalf of <strong>{userName}</strong> only, to the app sending mail as them
        (delegated <span className="mono">Mail.Send</span> on Microsoft Graph). The app can use it
        while {userName} is signed in to it — not on its own.
      </p>
      <PrincipalPicker
        kinds={['sp']}
        selected={app}
        onChange={setApp}
        placeholder="Search enterprise apps (min 2 chars)…"
        autoFocus
      />
      {error && (
        <p className="error" style={{ marginTop: 12 }}>
          {error}
        </p>
      )}
    </Modal>
  );
}

// ----------------------------------------------------------- Audit details

// undefined = still loading, null = the audit log can't be read.
function useAuditTrail(id: string, addActivity: 'Add user' | 'Add group') {
  const token = useGraphToken();
  const [trail, setTrail] = useState<AuditTrail | null | undefined>(undefined);
  useEffect(() => {
    let cancelled = false;
    setTrail(undefined);
    getAuditTrail(token, id, addActivity)
      .then((t) => !cancelled && setTrail(t))
      .catch(() => !cancelled && setTrail(null));
    return () => {
      cancelled = true;
    };
  }, [token, id, addActivity]);
  return trail;
}

const spinner = <span className="spinner" />;
const unavailable = (
  <span className="muted" title="Needs AuditLog.Read.All and a role that can read audit logs">
    Unavailable
  </span>
);
const notLogged = (
  <span className="muted" title="The audit log only goes back about 30 days">
    Not in the audit log
  </span>
);

// Who did it: a user or an app, linked to its page here when it has one.
function Actor({ event }: { event: DirectoryAuditEvent }) {
  const { user, app } = event.initiatedBy ?? {};
  if (user) {
    const name = user.displayName || user.userPrincipalName || user.id || 'Unknown user';
    return user.id ? <Link to={`/users/${user.id}`}>{name}</Link> : <>{name}</>;
  }
  if (app) {
    const name = app.displayName || app.appId || 'Unknown app';
    return app.servicePrincipalId ? (
      <Link to={`/enterprise-apps/${app.servicePrincipalId}`}>{name}</Link>
    ) : (
      <>{name}</>
    );
  }
  return dash;
}

// Created / created by / last modified / modified by. `createdAt` is the
// object's own property and outlives the audit log; the rest is only known
// while the event is still in the log.
function AuditRows({
  trail,
  createdAt,
}: {
  trail: AuditTrail | null | undefined;
  createdAt?: string | null;
}) {
  const pending = trail === undefined;
  const from = (
    event: DirectoryAuditEvent | undefined,
    cell: (e: DirectoryAuditEvent) => ReactNode,
  ) => (pending ? spinner : trail === null ? unavailable : event ? cell(event) : notLogged);
  return (
    <>
      <div className="k">Created</div>
      <div>{when(createdAt ?? trail?.created?.activityDateTime)}</div>
      <div className="k">Created by</div>
      <div>
        {from(trail?.created, (e) => (
          <Actor event={e} />
        ))}
      </div>
      <div className="k">Last modified</div>
      <div>
        {from(trail?.modified, (e) => (
          <>
            {when(e.activityDateTime)}
            {e.activityDisplayName && <span className="muted"> · {e.activityDisplayName}</span>}
          </>
        ))}
      </div>
      <div className="k">Modified by</div>
      <div>
        {from(trail?.modified, (e) => (
          <Actor event={e} />
        ))}
      </div>
    </>
  );
}

function SignInRows({ activity }: { activity: UserSignInActivity | null | undefined }) {
  const cell = (iso?: string | null) =>
    activity === undefined ? (
      spinner
    ) : activity === null ? (
      <span className="muted" title="Needs AuditLog.Read.All and an Entra ID P1 or P2 tenant">
        Unavailable
      </span>
    ) : iso ? (
      when(iso)
    ) : (
      <span className="muted">Never</span>
    );
  return (
    <>
      <div className="k">Last sign-in</div>
      <div>{cell(activity?.lastSignInDateTime)}</div>
      <div className="k" title="Token refreshes and other sign-ins made by a client for the user">
        Last background sign-in
      </div>
      <div>{cell(activity?.lastNonInteractiveSignInDateTime)}</div>
    </>
  );
}

function SaveBar({
  dirty,
  valid,
  saving,
  saved,
  error,
  onSave,
  onReset,
}: {
  dirty: boolean;
  valid: boolean;
  saving: boolean;
  saved: boolean;
  error: string | null;
  onSave: () => void;
  onReset: () => void;
}) {
  return (
    <>
      <div className="row" style={{ gap: 8, marginTop: 4 }}>
        <button className="primary" disabled={!dirty || !valid || saving} onClick={onSave}>
          {saving ? 'Saving…' : 'Save changes'}
        </button>
        <button disabled={!dirty || saving} onClick={onReset}>
          Discard
        </button>
        {saved && !dirty && <span className="muted">Saved</span>}
      </div>
      {error && (
        <p className="error" style={{ marginBottom: 0 }}>
          {error}
        </p>
      )}
    </>
  );
}

// The groups a user is a direct member of, with a button to add another.
function UserGroupsPanel({
  token,
  userId,
  userName,
  onCount,
}: {
  token: TokenFn;
  userId: string;
  userName: string;
  onCount: (label: string) => void;
}) {
  const [adding, setAdding] = useState(false);
  const [version, setVersion] = useState(0);
  // Ids of the listed groups, so the picker can say "already a member".
  const [memberOf, setMemberOf] = useState<string[]>([]);
  return (
    <>
      <div className="row related-note" style={{ justifyContent: 'space-between', gap: 12 }}>
        <div className="muted" style={{ fontSize: 12 }}>
          Groups the user is a direct member of.
        </div>
        <button style={{ flexShrink: 0 }} onClick={() => setAdding(true)}>
          + Add to group
        </button>
      </div>
      <RelatedPanel
        load={async () => {
          const groups = await listUserGroups(token, userId);
          setMemberOf(groups.items.map((g) => g.id));
          return groups;
        }}
        deps={[token, userId, version]}
        onCount={onCount}
        columns={['Member of', 'Type', 'Mail']}
        rowKey={(g) => g.id}
        to={(g) => `/groups/${g.id}`}
        row={(g) => [
          <NameCell id={g.id} name={g.displayName ?? g.id} to={`/groups/${g.id}`} />,
          <span className="muted">{groupKind(g)}</span>,
          g.mail ?? dash,
        ]}
        empty="Not a member of any group."
      />
      {adding && (
        <AddToGroup
          userId={userId}
          userName={userName}
          memberOf={memberOf}
          onClose={() => setAdding(false)}
          onAdded={() => {
            setAdding(false);
            setVersion((v) => v + 1);
          }}
        />
      )}
    </>
  );
}

function AddToGroup({
  userId,
  userName,
  memberOf,
  onClose,
  onAdded,
}: {
  userId: string;
  userName: string;
  memberOf: string[];
  onClose: () => void;
  onAdded: () => void;
}) {
  const token = useGraphToken(GROUP_WRITE_SCOPES);
  const [picked, setPicked] = useState<PrincipalRef | null>(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const already = picked !== null && memberOf.includes(picked.id);

  const submit = async () => {
    if (!picked) return;
    setSaving(true);
    setError(null);
    try {
      await addGroupMember(token, picked.id, userId);
      onAdded();
    } catch (e: unknown) {
      setError(message(e));
      setSaving(false);
    }
  };

  return (
    <Modal
      title={`Add ${userName} to a group`}
      onClose={() => !saving && onClose()}
      footer={
        <>
          <button disabled={saving} onClick={onClose}>
            Cancel
          </button>
          <button
            className="primary"
            disabled={saving || !picked || already}
            onClick={() => void submit()}
          >
            {saving ? 'Adding…' : 'Add to group'}
          </button>
        </>
      }
    >
      <p style={{ marginTop: 0, fontSize: 14 }}>
        Members of dynamic groups, groups synced from on-premises AD and distribution lists can't be
        changed here.
      </p>
      <PrincipalPicker
        kinds={['group']}
        selected={picked}
        onChange={(p) => {
          setPicked(p);
          setError(null);
        }}
        placeholder="Search groups (min 2 chars)…"
        autoFocus
      />
      {already && (
        <p className="error" style={{ marginTop: 12 }}>
          {userName} is already a member of this group.
        </p>
      )}
      {error && (
        <p className="error" style={{ marginTop: 12 }}>
          {error}
        </p>
      )}
    </Modal>
  );
}

// Why Graph won't take a new member for this group, if it won't.
function membersLocked(group: GroupData): string | undefined {
  const types = group.groupTypes ?? [];
  if (group.onPremisesSyncEnabled)
    return 'This group is synced from on-premises AD; change its members there.';
  if (group.membershipRule || types.includes('DynamicMembership'))
    return 'Members of a dynamic group follow from its membership rule.';
  if (group.mailEnabled && !types.includes('Unified'))
    return 'Distribution lists and mail-enabled security groups are managed in Exchange.';
  return undefined;
}

// The direct members of a group, with buttons to add one and to remove one.
function GroupMembersPanel({
  token,
  group,
  onCount,
}: {
  token: TokenFn;
  group: GroupData;
  onCount: (label: string) => void;
}) {
  const [adding, setAdding] = useState(false);
  const [version, setVersion] = useState(0);
  // Ids of the listed members, so the picker can say "already a member".
  const [members, setMembers] = useState<string[]>([]);
  const locked = membersLocked(group);
  const writeToken = useGraphToken(GROUP_WRITE_SCOPES);
  return (
    <>
      <div className="row related-note" style={{ justifyContent: 'space-between', gap: 12 }}>
        <div className="muted" style={{ fontSize: 12 }}>
          {locked ?? 'Direct members of this group.'}
        </div>
        <button
          style={{ flexShrink: 0 }}
          disabled={!!locked}
          title={locked}
          onClick={() => setAdding(true)}
        >
          + Add member
        </button>
      </div>
      <RelatedPanel
        load={async () => {
          const list = await listGroupMembers(token, group.id);
          setMembers(list.items.map((m) => m.id));
          return list;
        }}
        deps={[token, group.id, version]}
        onCount={onCount}
        columns={['Name', 'Type', 'Sign-in name / ID']}
        rowKey={(m) => m.id}
        to={(m) => memberTarget(m)[1]}
        row={memberCells}
        empty="This group has no members."
        remove={
          locked
            ? undefined
            : {
                label: 'Remove from group',
                confirm: (m) => ({
                  title: `Remove ${m.displayName ?? 'this member'} from the group?`,
                  message: (
                    <p style={{ margin: 0, fontSize: 14 }}>
                      <strong>{m.userPrincipalName ?? m.displayName ?? m.id}</strong> loses whatever{' '}
                      <strong>{group.displayName ?? 'this group'}</strong> gives access to: apps,
                      licences, sites and roles assigned to the group.
                    </p>
                  ),
                  confirmLabel: 'Remove member',
                }),
                run: async (m) => {
                  await removeGroupMember(writeToken, group.id, m.id);
                  setMembers((prev) => prev.filter((id) => id !== m.id));
                },
              }
        }
      />
      {adding && (
        <AddGroupMember
          group={group}
          members={members}
          onClose={() => setAdding(false)}
          onAdded={() => {
            setAdding(false);
            setVersion((v) => v + 1);
          }}
        />
      )}
    </>
  );
}

function AddGroupMember({
  group,
  members,
  onClose,
  onAdded,
}: {
  group: GroupData;
  members: string[];
  onClose: () => void;
  onAdded: () => void;
}) {
  const token = useGraphToken(GROUP_WRITE_SCOPES);
  const [picked, setPicked] = useState<PrincipalRef | null>(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  // Microsoft 365 groups only take users; security groups also take groups
  // and service principals.
  const unified = (group.groupTypes ?? []).includes('Unified');
  const problem = !picked
    ? null
    : picked.id === group.id
      ? "A group can't be a member of itself."
      : members.includes(picked.id)
        ? `${picked.displayName ?? 'This one'} is already a member of this group.`
        : null;

  const submit = async () => {
    if (!picked) return;
    setSaving(true);
    setError(null);
    try {
      await addGroupMember(token, group.id, picked.id);
      onAdded();
    } catch (e: unknown) {
      setError(message(e));
      setSaving(false);
    }
  };

  return (
    <Modal
      title={`Add a member to ${group.displayName ?? 'this group'}`}
      onClose={() => !saving && onClose()}
      footer={
        <>
          <button disabled={saving} onClick={onClose}>
            Cancel
          </button>
          <button
            className="primary"
            disabled={saving || !picked || problem !== null}
            onClick={() => void submit()}
          >
            {saving ? 'Adding…' : 'Add member'}
          </button>
        </>
      }
    >
      <p style={{ marginTop: 0, fontSize: 14 }}>
        {unified
          ? 'A Microsoft 365 group can only have users as members.'
          : 'A security group can have users, other groups and service principals as members.'}
      </p>
      <PrincipalPicker
        kinds={unified ? ['user'] : ['user', 'group', 'sp']}
        selected={picked}
        onChange={(p) => {
          setPicked(p);
          setError(null);
        }}
        placeholder={unified ? 'Search users (min 2 chars)…' : 'Search (min 2 chars)…'}
        autoFocus
      />
      {(problem ?? error) && (
        <p className="error" style={{ marginTop: 12 }}>
          {problem ?? error}
        </p>
      )}
    </Modal>
  );
}

// Pick another manager for the user, or take the current one away.
function ChangeManager({
  userId,
  userName,
  current,
  onClose,
  onChanged,
}: {
  userId: string;
  userName: string;
  current: UserManager | null;
  onClose: () => void;
  onChanged: (manager: UserManager | null) => void;
}) {
  const token = useGraphToken(USER_WRITE_SCOPES);
  const [picked, setPicked] = useState<PrincipalRef | null>(null);
  const [saving, setSaving] = useState<'set' | 'clear' | null>(null);
  const [error, setError] = useState<string | null>(null);
  const self = picked?.id === userId;

  const apply = async (next: PrincipalRef | null) => {
    setSaving(next ? 'set' : 'clear');
    setError(null);
    try {
      await setUserManager(token, userId, next?.id ?? null);
      onChanged(
        next
          ? { id: next.id, displayName: next.displayName, userPrincipalName: next.subtitle }
          : null,
      );
    } catch (e: unknown) {
      setError(message(e));
      setSaving(null);
    }
  };

  return (
    <Modal
      title={`Manager of ${userName}`}
      onClose={() => !saving && onClose()}
      footer={
        <>
          {current && (
            <button
              className="danger"
              style={{ marginRight: 'auto' }}
              disabled={saving !== null}
              onClick={() => void apply(null)}
            >
              {saving === 'clear' ? 'Removing…' : 'Remove manager'}
            </button>
          )}
          <button disabled={saving !== null} onClick={onClose}>
            Cancel
          </button>
          <button
            className="primary"
            disabled={saving !== null || !picked || self || picked.id === current?.id}
            onClick={() => void apply(picked)}
          >
            {saving === 'set' ? 'Saving…' : 'Set manager'}
          </button>
        </>
      }
    >
      <p style={{ marginTop: 0, fontSize: 14 }}>
        {current ? (
          <>
            Current manager:{' '}
            <strong>{current.displayName ?? current.userPrincipalName ?? current.id}</strong>.
          </>
        ) : (
          'No manager set.'
        )}{' '}
        The manager is also synced to the SharePoint user profile and shown in the organisation
        chart in Teams and Outlook.
      </p>
      <PrincipalPicker
        kinds={['user']}
        selected={picked}
        onChange={setPicked}
        placeholder="Search users (min 2 chars)…"
        autoFocus
      />
      {self && (
        <p className="error" style={{ marginTop: 12 }}>
          A user can't be their own manager.
        </p>
      )}
      {error && (
        <p className="error" style={{ marginTop: 12 }}>
          {error}
        </p>
      )}
    </Modal>
  );
}

// ---------------------------------------------------------------------- User

const USER_TEXT_FIELDS = [
  ['displayName', 'Display name'],
  ['givenName', 'First name'],
  ['surname', 'Last name'],
  ['jobTitle', 'Job title'],
  ['department', 'Department'],
] as const;
// Behind "More properties": organisation, contact and address.
const USER_MORE_FIELDS = [
  ['companyName', 'Company'],
  ['officeLocation', 'Office'],
  ['employeeId', 'Employee ID'],
  ['employeeType', 'Employee type'],
  ['mobilePhone', 'Mobile phone'],
  ['businessPhone', 'Business phone'],
  ['streetAddress', 'Street address'],
  ['city', 'City'],
  ['state', 'State or province'],
  ['postalCode', 'Postal code'],
  ['country', 'Country or region'],
  ['usageLocation', 'Usage location'],
] as const;
const USER_FIELDS = [...USER_TEXT_FIELDS, ...USER_MORE_FIELDS];
type UserTextField = (typeof USER_FIELDS)[number][0];
type UserForm = Record<UserTextField, string> & { enabled: boolean };

// Properties Microsoft copies from Entra ID into the SharePoint user profile
// (the fixed set of the User Profile Application sync), with the name the
// property has there: what SharePoint shows, and the internal name used in
// search, people web parts and the profile API.
// https://learn.microsoft.com/sharepoint/user-profile-sync
const SHAREPOINT_PROFILE: Record<string, [display: string, internal: string]> = {
  displayName: ['Name', 'PreferredName'],
  givenName: ['FirstName', 'FirstName'],
  surname: ['LastName', 'LastName'],
  jobTitle: ['Title / Job Title', 'Title, SPS-JobTitle'],
  department: ['Department', 'Department, SPS-Department'],
  officeLocation: ['Office', 'Office'],
  businessPhone: ['Work phone', 'WorkPhone'],
  userPrincipalName: ['Account Name / User Name', 'AccountName, UserName, SPS-UserPrincipalName'],
  mail: ['Work Email', 'WorkEmail'],
  manager: ['Manager', 'Manager'],
};

// Chain icon plus the SharePoint name, shown after a property's label.
function SharePointSync({ property }: { property: string }) {
  const names = SHAREPOINT_PROFILE[property];
  if (!names) return null;
  return (
    <em
      className="sp-sync"
      title={`Synced to the SharePoint user profile as "${names[0]}" (internal name: ${names[1]})`}
    >
      <ChainLinkIcon />
      {names[0]}
    </em>
  );
}

// The form's "businessPhone" is the one number Graph keeps in businessPhones.
const userValue = (u: UserData, key: UserTextField) =>
  (key === 'businessPhone' ? u.businessPhones?.[0] : u[key]) ?? '';

const userForm = (u: UserData): UserForm => ({
  displayName: u.displayName ?? '',
  givenName: u.givenName ?? '',
  surname: u.surname ?? '',
  jobTitle: u.jobTitle ?? '',
  department: u.department ?? '',
  companyName: u.companyName ?? '',
  officeLocation: u.officeLocation ?? '',
  employeeId: u.employeeId ?? '',
  employeeType: u.employeeType ?? '',
  mobilePhone: u.mobilePhone ?? '',
  businessPhone: u.businessPhones?.[0] ?? '',
  streetAddress: u.streetAddress ?? '',
  city: u.city ?? '',
  state: u.state ?? '',
  postalCode: u.postalCode ?? '',
  country: u.country ?? '',
  usageLocation: u.usageLocation ?? '',
  enabled: u.accountEnabled !== false,
});

export function UserDetail() {
  const { id = '' } = useParams();
  const token = useGraphToken();
  const writeToken = useGraphToken(USER_WRITE_SCOPES);
  const tenantId = useCurrentTenantId();
  const nav = useNavigate();
  const { instance, accounts } = useMsal();
  const isMe = (instance.getActiveAccount() ?? accounts[0])?.localAccountId === id;
  const { confirm, confirmDialog } = useConfirm();
  const [user, setUser] = useState<UserData | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [form, setForm] = useState<UserForm | null>(null);
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [deleting, setDeleting] = useState(false);
  const trail = useAuditTrail(id, 'Add user');
  const [resetting, setResetting] = useState(false);
  const [signIn, setSignIn] = useState<UserSignInActivity | null | undefined>(undefined);
  const inviteToken = useGraphToken(USER_INVITE_SCOPES);
  // undefined while loading, null when there is none.
  const [manager, setManager] = useState<UserManager | null | undefined>(undefined);
  const [acting, setActing] = useState<'revoke' | 'invite' | null>(null);
  const [pickingManager, setPickingManager] = useState(false);
  // null: follow the data — open when any of the extra properties is filled in.
  const [showMore, setShowMore] = useState<boolean | null>(null);
  const [actionResult, setActionResult] = useState<{ ok: boolean; text: string } | null>(null);

  useEffect(() => {
    let cancelled = false;
    setManager(undefined);
    setActionResult(null);
    getUserManager(token, id)
      .then((m) => !cancelled && setManager(m))
      .catch(() => !cancelled && setManager(null));
    return () => {
      cancelled = true;
    };
  }, [token, id]);

  useEffect(() => {
    let cancelled = false;
    setSignIn(undefined);
    getUserSignInActivity(token, id)
      .then((a) => !cancelled && setSignIn(a))
      .catch(() => !cancelled && setSignIn(null));
    return () => {
      cancelled = true;
    };
  }, [token, id]);

  useEffect(() => {
    let cancelled = false;
    setUser(null);
    setForm(null);
    setError(null);
    setSaved(false);
    setSaveError(null);
    getUser(token, id)
      .then((u) => {
        if (cancelled) return;
        setUser(u);
        setForm(userForm(u));
      })
      .catch((e: unknown) => !cancelled && setError(message(e)));
    return () => {
      cancelled = true;
    };
  }, [token, id]);

  if (error)
    return (
      <>
        <BackLink to="/users" label="Users" />
        <div className="card error">{error}</div>
      </>
    );
  if (!user || !form)
    return (
      <div className="center">
        <span className="spinner" />
      </div>
    );

  // Only what actually changed is sent; an emptied field is cleared.
  const patch: UserPatch = {};
  for (const [key] of USER_FIELDS) {
    const next = key === 'usageLocation' ? form[key].trim().toUpperCase() : form[key].trim();
    if (next === userValue(user, key)) continue;
    if (key === 'displayName') patch.displayName = next;
    else if (key === 'businessPhone') patch.businessPhones = next ? [next] : [];
    else patch[key] = next || null;
  }
  // A two-letter country code such as NL, or empty.
  const usageLocationOk = /^([A-Za-z]{2})?$/.test(form.usageLocation.trim());
  const moreOpen = showMore ?? USER_MORE_FIELDS.some(([key]) => userValue(user, key) !== '');
  const textField = ([key, label]: (typeof USER_FIELDS)[number]) => (
    <Field
      key={key}
      label={
        <>
          {label}
          <SharePointSync property={key} />
        </>
      }
    >
      <input
        value={form[key]}
        placeholder={key === 'usageLocation' ? 'Two-letter country code, e.g. NL' : undefined}
        maxLength={key === 'usageLocation' ? 2 : undefined}
        onChange={(e) => {
          setSaved(false);
          setForm({ ...form, [key]: e.target.value });
        }}
      />
    </Field>
  );
  if (form.enabled !== (user.accountEnabled !== false)) patch.accountEnabled = form.enabled;
  const dirty = Object.keys(patch).length > 0;
  const name = user.displayName ?? user.id;

  const save = async () => {
    setSaving(true);
    setSaveError(null);
    try {
      await updateUser(writeToken, user.id, patch);
      const next = { ...user, ...patch };
      setUser(next);
      setForm(userForm(next));
      setSaved(true);
    } catch (e: unknown) {
      setSaveError(message(e));
    } finally {
      setSaving(false);
    }
  };

  // Why this user's password can't be set here, if it can't.
  const passwordBlocked = isMe
    ? 'This is your own account — change your password from your account settings'
    : isGuest(user)
      ? 'Guests sign in with their home organisation — there is no password to reset here'
      : undefined;

  const remove = async () => {
    const ok = await confirm({
      title: `Delete ${name}?`,
      message: (
        <p style={{ margin: 0, fontSize: 14 }}>
          <strong>{user.userPrincipalName ?? user.id}</strong> will no longer be able to sign in and
          loses access to every app. Deleted users can be restored from the Entra admin center for
          30 days.
        </p>
      ),
      confirmLabel: 'Delete user',
    });
    if (!ok) return;
    setDeleting(true);
    setSaveError(null);
    try {
      await deleteUser(writeToken, user.id);
      nav('/users');
    } catch (e: unknown) {
      setSaveError(`Couldn't delete this user: ${message(e)}`);
      setDeleting(false);
    }
  };

  const revoke = async () => {
    const ok = await confirm({
      title: `Sign ${name} out everywhere?`,
      message: (
        <p style={{ margin: 0, fontSize: 14 }}>
          Ends every session of <strong>{user.userPrincipalName ?? name}</strong>: all apps and
          devices ask them to sign in again, which can take up to an hour to reach apps that are
          already open. Their password and sign-in methods stay as they are.
        </p>
      ),
      confirmLabel: 'Revoke sessions',
    });
    if (!ok) return;
    setActing('revoke');
    setActionResult(null);
    try {
      await revokeSignInSessions(writeToken, user.id);
      setActionResult({ ok: true, text: 'All sessions revoked — the user has to sign in again.' });
    } catch (e: unknown) {
      setActionResult({ ok: false, text: `Couldn't revoke the sessions: ${message(e)}` });
    } finally {
      setActing(null);
    }
  };

  const pendingInvite = user.externalUserState === 'PendingAcceptance' ? user.mail : undefined;
  const reinvite = async () => {
    if (!pendingInvite) return;
    const ok = await confirm({
      title: 'Send the invitation again?',
      message: (
        <p style={{ margin: 0, fontSize: 14 }}>
          Microsoft sends a new invitation email to <strong>{pendingInvite}</strong>.
        </p>
      ),
      confirmLabel: 'Send invitation',
    });
    if (!ok) return;
    setActing('invite');
    setActionResult(null);
    try {
      await resendInvitation(inviteToken, pendingInvite, tenantId);
      setActionResult({ ok: true, text: `Invitation sent to ${pendingInvite}.` });
    } catch (e: unknown) {
      setActionResult({ ok: false, text: `Couldn't send the invitation: ${message(e)}` });
    } finally {
      setActing(null);
    }
  };

  const aliases = (user.proxyAddresses ?? [])
    .filter((a) => a.startsWith('smtp:'))
    .map((a) => a.slice(5));
  const homeIdentity = isGuest(user)
    ? user.identities?.find((i) => i.signInType === 'federated')?.issuer
    : undefined;
  return (
    <>
      <BackLink to="/users" label="Users" />
      <div className="page-header detail-header">
        <div className="row" style={{ gap: 10, alignItems: 'center' }}>
          <Identicon id={user.id} size={36} title={name} />
          <div>
            <h1 style={{ margin: 0 }}>
              {name} {isGuest(user) && <span className="badge pending">Guest</span>}{' '}
              {user.accountEnabled === false && <span className="badge expired">Disabled</span>}
            </h1>
            <div className="subtitle mono">
              {user.userPrincipalName}
              {user.userPrincipalName && (
                <CopyButton value={user.userPrincipalName} label="user principal name" />
              )}
              <SharePointSync property="userPrincipalName" />
            </div>
          </div>
        </div>
        <div className="row">
          <EntraButton href={entraUserUrl(user.id, tenantId)} />
          <button
            disabled={Boolean(passwordBlocked)}
            title={passwordBlocked ?? 'Set a new password or generate a temporary one'}
            onClick={() => setResetting(true)}
          >
            Reset password
          </button>
          <button
            disabled={acting !== null}
            title="Sign the user out of every app and device"
            onClick={() => void revoke()}
          >
            {acting === 'revoke' ? 'Revoking…' : 'Revoke sessions'}
          </button>
          {pendingInvite && (
            <button
              disabled={acting !== null}
              title="The guest hasn't accepted the invitation yet"
              onClick={() => void reinvite()}
            >
              {acting === 'invite' ? 'Sending…' : 'Resend invitation'}
            </button>
          )}
          <button
            className="danger"
            disabled={deleting || isMe}
            title={isMe ? "This is your own account — you can't delete it here" : undefined}
            onClick={() => void remove()}
          >
            {deleting ? 'Deleting…' : 'Delete user'}
          </button>
        </div>
      </div>

      {actionResult && (
        <p className={actionResult.ok ? 'muted' : 'error'} style={{ margin: '0 0 12px' }}>
          {actionResult.text}
        </p>
      )}

      <div className="detail-grid">
        <div className="card">
          <h3>Properties</h3>
          {user.onPremisesSyncEnabled && (
            <p className="muted" style={{ fontSize: 12, marginTop: 0 }}>
              Synced from on-premises Active Directory; most properties can only be changed there.
            </p>
          )}
          <div className="form-grid">
            {USER_TEXT_FIELDS.map(textField)}
            {moreOpen && USER_MORE_FIELDS.map(textField)}
            <Check
              checked={form.enabled}
              onChange={(enabled) => {
                setSaved(false);
                setForm({ ...form, enabled });
              }}
            >
              Account enabled (can sign in)
            </Check>
          </div>
          <button
            className="ghost"
            style={{ padding: 0, marginBottom: 8 }}
            onClick={() => setShowMore(!moreOpen)}
          >
            {moreOpen ? '▾ Fewer properties' : '▸ More properties (organisation, contact, address)'}
          </button>
          {!usageLocationOk && (
            <p className="error" style={{ marginTop: 0 }}>
              Usage location must be a two-letter country code, such as NL.
            </p>
          )}
          <SaveBar
            dirty={dirty}
            valid={form.displayName.trim().length > 0 && usageLocationOk}
            saving={saving}
            saved={saved}
            error={saveError}
            onSave={() => void save()}
            onReset={() => setForm(userForm(user))}
          />
        </div>

        <div className="card">
          <h3>Details</h3>
          <div className="kv compact">
            <IdRow label="Object ID" value={user.id} />
            <IdRow label="Mail" value={user.mail} sync="mail" />
            <div className="k">Type</div>
            <div>
              {user.userType ?? 'Member'}
              <span className="muted">
                {' · '}
                {user.onPremisesSyncEnabled ? 'synced from on-premises AD' : 'cloud'}
              </span>
            </div>
            {aliases.length > 0 && (
              <>
                <div className="k">Mail aliases</div>
                <div className="mono">{aliases.join(', ')}</div>
              </>
            )}
            {isGuest(user) && (
              <>
                <div className="k">Invitation</div>
                <div>
                  {user.externalUserState === 'PendingAcceptance' ? (
                    <span className="badge pending">Not accepted yet</span>
                  ) : user.externalUserState === 'Accepted' ? (
                    'Accepted'
                  ) : (
                    (user.externalUserState ?? dash)
                  )}
                  {user.externalUserStateChangeDateTime && (
                    <span className="muted">
                      {' · '}
                      {new Date(user.externalUserStateChangeDateTime).toLocaleDateString()}
                    </span>
                  )}
                </div>
                {homeIdentity && (
                  <>
                    <div className="k">Signs in with</div>
                    <div>{IDENTITY_PROVIDERS[homeIdentity] ?? homeIdentity}</div>
                  </>
                )}
              </>
            )}
            {user.onPremisesSyncEnabled && (
              <>
                <div className="k">Last AD sync</div>
                <div>{when(user.onPremisesLastSyncDateTime)}</div>
                <div className="k">AD account</div>
                <div className="mono">
                  {user.onPremisesSamAccountName
                    ? `${user.onPremisesDomainName ? `${user.onPremisesDomainName}\\` : ''}${user.onPremisesSamAccountName}`
                    : dash}
                </div>
              </>
            )}
            <div className="k">
              Manager
              <SharePointSync property="manager" />
            </div>
            <div>
              {manager === undefined ? (
                spinner
              ) : (
                <>
                  {manager ? (
                    <Link to={`/users/${manager.id}`}>
                      {manager.displayName ?? manager.userPrincipalName ?? manager.id}
                    </Link>
                  ) : (
                    dash
                  )}{' '}
                  <button className="site-consent" onClick={() => setPickingManager(true)}>
                    {manager ? 'Change' : 'Set'}
                  </button>
                </>
              )}
            </div>
            <AuditRows trail={trail} createdAt={user.createdDateTime} />
            <SignInRows activity={signIn} />
            <div className="k">Password changed</div>
            <div>{when(user.lastPasswordChangeDateTime)}</div>
          </div>
        </div>
      </div>

      <RelatedTabs
        key={user.id}
        tabs={[
          {
            key: 'groups',
            title: 'Groups',
            panel: (onCount) => (
              <UserGroupsPanel token={token} userId={user.id} userName={name} onCount={onCount} />
            ),
          },
          {
            key: 'roles',
            title: 'Roles',
            panel: (onCount) => (
              <ConsentGate scopes={ROLE_READ_SCOPES} what="directory roles">
                {(roleToken) => (
                  <UserRolesPanel
                    token={roleToken}
                    userId={user.id}
                    userName={name}
                    onCount={onCount}
                  />
                )}
              </ConsentGate>
            ),
          },
          appAssignmentsTab(token, 'users', user.id),
          {
            key: 'owns',
            title: 'Owns',
            panel: (onCount) => (
              <RelatedPanel
                load={() => listOwnedObjects(token, user.id)}
                deps={[token, user.id]}
                onCount={onCount}
                columns={['Name', 'Type', 'App ID / mail']}
                rowKey={(m) => m.id}
                to={(m) => memberTarget(m)[1]}
                row={memberCells}
                empty="Doesn't own any app registration, enterprise app or group."
              />
            ),
          },
          {
            key: 'consents',
            title: 'Consents',
            panel: (onCount) => (
              <RelatedPanel
                load={() => listUserConsents(token, user.id)}
                deps={[token, user.id]}
                onCount={onCount}
                note="Permissions this user gave apps to act on their behalf. Admin consent for all users is not listed here."
                columns={['App', 'May access', 'Permissions']}
                rowKey={(c) => c.grant.id}
                to={(c) => `/enterprise-apps/${c.grant.clientId}`}
                row={(c) => [
                  <NameCell
                    id={c.grant.clientId}
                    name={c.clientName}
                    to={`/enterprise-apps/${c.grant.clientId}`}
                  />,
                  c.resourceName,
                  <span className="mono">{c.scopes.join(', ')}</span>,
                ]}
                remove={{
                  label: 'Withdraw this consent',
                  confirm: (c) => ({
                    title: `Withdraw the consent for ${c.clientName}?`,
                    message: (
                      <p style={{ margin: 0, fontSize: 14 }}>
                        <strong>{c.clientName}</strong> loses{' '}
                        <span className="mono">{c.scopes.join(', ')}</span> on {c.resourceName} for{' '}
                        {name}. Other users are not affected; the app asks {name} for consent again
                        the next time it needs it.
                      </p>
                    ),
                    confirmLabel: 'Withdraw consent',
                  }),
                  run: (c) => deleteOAuth2Grant(token, c.grant.id),
                }}
                empty="This user hasn't consented to any app themselves."
              />
            ),
          },
          {
            key: 'mail',
            title: 'Can send mail as',
            panel: (onCount) => (
              <MailSendersPanel userId={user.id} userName={name} onCount={onCount} />
            ),
          },
          {
            key: 'licenses',
            title: 'Licenses',
            panel: (onCount) => (
              <RelatedPanel
                load={() => listUserLicenses(token, user.id)}
                deps={[token, user.id]}
                onCount={onCount}
                columns={['License', 'Services on', 'Assigned']}
                rowKey={(l) => l.skuId}
                row={(l) => [
                  <span title={l.partNumber}>
                    {l.name}{' '}
                    {l.error && (
                      <span className="badge expired" title={l.error}>
                        Error
                      </span>
                    )}
                  </span>,
                  `${l.plansEnabled} of ${l.plansTotal}`,
                  <>
                    {l.direct && 'Directly'}
                    {l.groups.map((g, i) => (
                      <Fragment key={g.id}>
                        {(l.direct || i > 0) && ', '}
                        {i === 0 && 'through '}
                        <Link to={`/groups/${g.id}`}>{g.name}</Link>
                      </Fragment>
                    ))}
                  </>,
                ]}
                empty="No licenses assigned."
              />
            ),
          },
          {
            key: 'devices',
            title: 'Devices',
            panel: (onCount) => (
              <RelatedPanel
                load={() => listUserDevices(token, user.id)}
                deps={[token, user.id]}
                onCount={onCount}
                columns={['Device', 'System', 'Join type', 'Compliant', 'Last active']}
                rowKey={(d) => d.id}
                row={(d) => [
                  <>
                    {d.displayName ?? d.id}{' '}
                    {d.accountEnabled === false && <span className="badge expired">Disabled</span>}
                  </>,
                  [d.operatingSystem, d.operatingSystemVersion].filter(Boolean).join(' ') || dash,
                  d.trustType ? (DEVICE_TRUST[d.trustType] ?? d.trustType) : dash,
                  d.isCompliant == null ? (
                    <span className="muted">{d.isManaged ? 'Unknown' : 'Not managed'}</span>
                  ) : (
                    yesNo(d.isCompliant)
                  ),
                  when(d.approximateLastSignInDateTime),
                ]}
                empty="No registered or owned devices."
              />
            ),
          },
          {
            key: 'reports',
            title: 'Direct reports',
            panel: (onCount) => (
              <RelatedPanel
                load={() => listDirectReports(token, user.id)}
                deps={[token, user.id]}
                onCount={onCount}
                columns={['Name', 'Job title', 'Sign-in name']}
                rowKey={(m) => m.id}
                to={(m) => `/users/${m.id}`}
                row={(m) => [
                  <NameCell id={m.id} name={m.displayName ?? m.id} to={`/users/${m.id}`} />,
                  m.jobTitle ?? dash,
                  m.userPrincipalName ?? m.mail ?? dash,
                ]}
                empty="Nobody reports to this user."
              />
            ),
          },
          {
            key: 'signins',
            title: 'Sign-ins',
            lazy: true,
            panel: (onCount) => (
              <RelatedPanel
                load={() => listUserSignIns(token, user.id)}
                deps={[token, user.id]}
                onCount={onCount}
                note="The latest interactive sign-ins. Token refreshes in the background are not listed."
                columns={['When', 'App', 'Result', 'From', 'Conditional Access']}
                rowKey={(s) => s.id}
                row={(s) => [
                  when(s.createdDateTime),
                  s.appDisplayName ?? s.appId ?? dash,
                  s.status?.errorCode === 0 ? (
                    <span className="badge granted">Success</span>
                  ) : (
                    <span className="badge expired" title={s.status?.failureReason}>
                      Failed{s.status?.errorCode ? ` (${s.status.errorCode})` : ''}
                    </span>
                  ),
                  <>
                    {[s.location?.city, s.location?.countryOrRegion].filter(Boolean).join(', ') ||
                      dash}
                    {s.ipAddress && <span className="muted mono"> · {s.ipAddress}</span>}
                  </>,
                  CA_STATUS[s.conditionalAccessStatus ?? ''] ?? s.conditionalAccessStatus ?? dash,
                ]}
                empty="No interactive sign-ins in the log (it goes back 30 days, 7 on free tenants)."
              />
            ),
          },
          {
            key: 'security',
            title: 'Security',
            lazy: true,
            panel: () => <SecurityPanel userId={user.id} />,
          },
        ]}
      />
      {pickingManager && (
        <ChangeManager
          userId={user.id}
          userName={name}
          current={manager ?? null}
          onClose={() => setPickingManager(false)}
          onChanged={(m) => {
            setManager(m);
            setPickingManager(false);
          }}
        />
      )}
      {resetting && (
        <ResetPassword
          user={user}
          onClose={() => setResetting(false)}
          onDone={() => setUser({ ...user, lastPasswordChangeDateTime: new Date().toISOString() })}
        />
      )}
      {confirmDialog}
    </>
  );
}

// --------------------------------------------------------------------- Group

interface GroupForm {
  displayName: string;
  description: string;
}
const groupForm = (g: GroupData): GroupForm => ({
  displayName: g.displayName ?? '',
  description: g.description ?? '',
});

export function GroupDetail() {
  const { id = '' } = useParams();
  const token = useGraphToken();
  const writeToken = useGraphToken(GROUP_WRITE_SCOPES);
  const tenantId = useCurrentTenantId();
  const nav = useNavigate();
  const { confirm, confirmDialog } = useConfirm();
  const [group, setGroup] = useState<GroupData | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [form, setForm] = useState<GroupForm | null>(null);
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [deleting, setDeleting] = useState(false);
  const trail = useAuditTrail(id, 'Add group');
  const { sites, needsConsent, request: requestSites, grant: grantSites } = useGroupSites();
  // Only Microsoft 365 groups have a SharePoint site.
  const unifiedId = group?.groupTypes?.includes('Unified') ? group.id : undefined;
  useEffect(() => {
    if (unifiedId) requestSites([unifiedId]);
  }, [unifiedId, requestSites]);

  useEffect(() => {
    let cancelled = false;
    setGroup(null);
    setForm(null);
    setError(null);
    setSaved(false);
    setSaveError(null);
    getGroup(token, id)
      .then((g) => {
        if (cancelled) return;
        setGroup(g);
        setForm(groupForm(g));
      })
      .catch((e: unknown) => !cancelled && setError(message(e)));
    return () => {
      cancelled = true;
    };
  }, [token, id]);

  if (error)
    return (
      <>
        <BackLink to="/groups" label="Groups" />
        <div className="card error">{error}</div>
      </>
    );
  if (!group || !form)
    return (
      <div className="center">
        <span className="spinner" />
      </div>
    );

  const displayName = form.displayName.trim();
  const description = form.description.trim();
  const patch: { displayName?: string; description?: string | null } = {};
  if (displayName !== (group.displayName ?? '')) patch.displayName = displayName;
  if (description !== (group.description ?? '')) patch.description = description || null;
  const dirty = Object.keys(patch).length > 0;
  const name = group.displayName ?? group.id;
  const restorable = group.groupTypes?.includes('Unified');
  const dynamic = group.groupTypes?.includes('DynamicMembership');

  const save = async () => {
    setSaving(true);
    setSaveError(null);
    try {
      await updateGroup(writeToken, group.id, patch);
      const next = {
        ...group,
        ...(patch.displayName !== undefined ? { displayName: patch.displayName } : {}),
        ...(patch.description !== undefined ? { description: patch.description ?? undefined } : {}),
      };
      setGroup(next);
      setForm(groupForm(next));
      setSaved(true);
    } catch (e: unknown) {
      setSaveError(message(e));
    } finally {
      setSaving(false);
    }
  };

  const remove = async () => {
    const ok = await confirm({
      title: `Delete ${name}?`,
      message: (
        <p style={{ margin: 0, fontSize: 14 }}>
          Everyone who gets access to apps or resources through this group loses it.{' '}
          {restorable ? (
            'Microsoft 365 groups can be restored from the Entra admin center for 30 days.'
          ) : (
            <strong>This is permanent — only Microsoft 365 groups can be restored.</strong>
          )}
        </p>
      ),
      confirmLabel: 'Delete group',
    });
    if (!ok) return;
    setDeleting(true);
    setSaveError(null);
    try {
      await deleteGroup(writeToken, group.id);
      nav('/groups');
    } catch (e: unknown) {
      setSaveError(`Couldn't delete this group: ${message(e)}`);
      setDeleting(false);
    }
  };

  return (
    <>
      <BackLink to="/groups" label="Groups" />
      <div className="page-header detail-header">
        <div className="row" style={{ gap: 10, alignItems: 'center' }}>
          <Identicon id={group.id} size={36} title={name} />
          <div>
            <h1 style={{ margin: 0 }}>{name}</h1>
            <div className="subtitle">
              {groupKind(group)}
              {group.mail && ` · ${group.mail}`}
            </div>
          </div>
        </div>
        <div className="row">
          <EntraButton href={entraGroupUrl(group.id, tenantId)} />
          <button
            className="danger"
            disabled={deleting}
            title={
              restorable
                ? 'Can be restored from the Entra admin center for 30 days'
                : 'Permanent — only Microsoft 365 groups can be restored'
            }
            onClick={() => void remove()}
          >
            {deleting ? 'Deleting…' : 'Delete group'}
          </button>
        </div>
      </div>

      <div className="detail-grid">
        <div className="card">
          <h3>Properties</h3>
          {group.onPremisesSyncEnabled && (
            <p className="muted" style={{ fontSize: 12, marginTop: 0 }}>
              Synced from on-premises Active Directory; it can only be changed there.
            </p>
          )}
          <Field label="Display name">
            <input
              value={form.displayName}
              onChange={(e) => {
                setSaved(false);
                setForm({ ...form, displayName: e.target.value });
              }}
            />
          </Field>
          <Field label="Description">
            <textarea
              rows={3}
              value={form.description}
              onChange={(e) => {
                setSaved(false);
                setForm({ ...form, description: e.target.value });
              }}
            />
          </Field>
          <SaveBar
            dirty={dirty}
            valid={displayName.length > 0}
            saving={saving}
            saved={saved}
            error={saveError}
            onSave={() => void save()}
            onReset={() => setForm(groupForm(group))}
          />
        </div>

        <div className="card">
          <h3>Details</h3>
          <div className="kv compact">
            <IdRow label="Object ID" value={group.id} />
            <IdRow label="Mail nickname" value={group.mailNickname} />
            <div className="k">Type</div>
            <div>
              {groupKind(group)}
              <span className="muted">
                {' · '}
                {group.onPremisesSyncEnabled ? 'synced from on-premises AD' : 'cloud'}
              </span>
            </div>
            {restorable && (
              <>
                <div className="k">SharePoint site</div>
                <div>
                  <GroupSiteLink
                    site={sites.get(group.id)}
                    needsConsent={needsConsent}
                    onGrant={grantSites}
                    full
                  />
                </div>
              </>
            )}
            {dynamic && (
              <>
                <div className="k">Membership rule</div>
                <div className="mono">{group.membershipRule ?? dash}</div>
              </>
            )}
            <AuditRows trail={trail} createdAt={group.createdDateTime} />
          </div>
        </div>
      </div>

      <RelatedTabs
        key={group.id}
        tabs={[
          {
            key: 'members',
            title: 'Members',
            panel: (onCount) => <GroupMembersPanel token={token} group={group} onCount={onCount} />,
          },
          {
            key: 'owners',
            title: 'Owners',
            panel: (onCount) => (
              <RelatedPanel
                load={() => listGroupOwners(token, group.id)}
                deps={[token, group.id]}
                onCount={onCount}
                columns={['Name', 'Type', 'Sign-in name / ID']}
                rowKey={(m) => m.id}
                to={(m) => memberTarget(m)[1]}
                row={memberCells}
                empty="This group has no owners."
              />
            ),
          },
          appAssignmentsTab(token, 'groups', group.id),
        ]}
      />
      {confirmDialog}
    </>
  );
}
