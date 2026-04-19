import { useEffect, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { useCurrentTenantId } from '../auth/useCurrentTenantId';
import {
  deleteServicePrincipal,
  getServicePrincipal,
  updateServicePrincipal,
} from '../graph/servicePrincipals';
import { getApplicationByAppId } from '../graph/applications';
import {
  listAppOnlySignInsForApp,
  listUserSignInsForApp,
} from '../graph/signIns';
import {
  findTenantInformation,
  type TenantInformation,
} from '../graph/organization';
import type {
  AppCredentialSignInActivity,
  KeyCredential,
  PasswordCredential,
  ServicePrincipal,
  SignIn,
} from '../graph/types';
import {
  buildCredentialActivityMap,
  listAppCredentialSignInActivities,
} from '../graph/credentials';
import { PermissionsManager } from '../components/PermissionsManager';
import { AssignedUsersManager } from '../components/AssignedUsersManager';
import { SignInAuditTab } from '../components/SignInAuditTab';
import { AppIcon } from '../components/AppIcon';
import { CopyButton } from '../components/CopyButton';
import { Modal } from '../components/Modal';

type Tab = 'overview' | 'permissions' | 'assignments' | 'audit';

export function EnterpriseAppDetail() {
  const { id = '' } = useParams();
  const token = useGraphToken();
  const nav = useNavigate();
  const currentTenantId = useCurrentTenantId();
  const [sp, setSp] = useState<ServicePrincipal | null>(null);
  // Directory-object id of the underlying app registration, resolved only
  // when the SP lives in this tenant. null = not yet resolved or not
  // applicable; undefined isn't used. Drives the "View app registration"
  // link in the header.
  const [appRegId, setAppRegId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [tab, setTab] = useState<Tab>('overview');
  const [togglingEnabled, setTogglingEnabled] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [deleting, setDeleting] = useState(false);
  // The /reports/servicePrincipalSignInActivities endpoint lags heavily in
  // many tenants (often empty while /auditLogs/signIns already has 100+
  // events). Query the authoritative sign-in log directly for the most
  // recent events per flow, matching what the Audit tab shows.
  const [lastDelegated, setLastDelegated] = useState<SignIn | null | 'loading'>(
    'loading',
  );
  const [lastAppOnly, setLastAppOnly] = useState<SignIn | null | 'loading'>(
    'loading',
  );
  const [signInError, setSignInError] = useState<string | null>(null);
  const [credActivity, setCredActivity] = useState<
    Map<string, AppCredentialSignInActivity>
  >(new Map());

  useEffect(() => {
    setSp(null);
    setError(null);
    setAppRegId(null);
    setLastDelegated('loading');
    setLastAppOnly('loading');
    setSignInError(null);
    setCredActivity(new Map());
    getServicePrincipal(token, id)
      .then((full) => {
        setSp(full);

        listUserSignInsForApp(token, full.appId, 1)
          .then((rows) => setLastDelegated(rows[0] ?? null))
          .catch((e: unknown) => {
            setLastDelegated(null);
            setSignInError(e instanceof Error ? e.message : String(e));
          });
        listAppOnlySignInsForApp(token, full.appId, 1)
          .then((rows) => setLastAppOnly(rows[0] ?? null))
          .catch((e: unknown) => {
            setLastAppOnly(null);
            setSignInError(e instanceof Error ? e.message : String(e));
          });

        listAppCredentialSignInActivities(token, full.appId)
          .then((rows) => setCredActivity(buildCredentialActivityMap(rows)))
          .catch(() => {
            /* AuditLog.Read.All absent — credentials render without "last used" */
          });
      })
      .catch((e) => setError(e.message));
  }, [token, id]);

  // Resolve the underlying app registration's directory-object id so the
  // header can link to /applications/:id. Only attempted when the SP is
  // registered in this tenant — foreign-tenant SPs won't have a readable
  // application object. Silently gives up on failure (e.g. missing
  // Application.Read.All).
  useEffect(() => {
    setAppRegId(null);
    if (!sp) return;
    if (sp.servicePrincipalType !== 'Application') return;
    const owner = sp.appOwnerOrganizationId;
    if (!owner || !currentTenantId) return;
    if (owner.toLowerCase() !== currentTenantId.toLowerCase()) return;
    let cancelled = false;
    getApplicationByAppId(token, sp.appId)
      .then((app) => {
        if (!cancelled && app) setAppRegId(app.id);
      })
      .catch(() => {});
    return () => {
      cancelled = true;
    };
  }, [token, sp, currentTenantId]);

  const toggleEnabled = async () => {
    if (!sp) return;
    setTogglingEnabled(true);
    setError(null);
    try {
      const next = !sp.accountEnabled;
      await updateServicePrincipal(token, sp.id, { accountEnabled: next });
      setSp({ ...sp, accountEnabled: next });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setTogglingEnabled(false);
    }
  };

  const doDelete = async () => {
    if (!sp) return;
    setDeleting(true);
    setError(null);
    try {
      await deleteServicePrincipal(token, sp.id);
      nav('/enterprise-apps');
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setDeleting(false);
    }
  };

  if (error) return <div className="card error">{error}</div>;
  if (!sp)
    return (
      <div className="center">
        <span className="spinner" />
      </div>
    );

  return (
    <>
      <div className="page-header">
        <div className="row" style={{ gap: 12, alignItems: 'center' }}>
          <AppIcon
            id={sp.appId || sp.id}
            logoUrl={sp.info?.logoUrl}
            size={48}
            title={sp.displayName}
          />
          <div>
            <h1 style={{ margin: 0 }}>{sp.displayName}</h1>
            <div className="subtitle mono">
              {sp.appId}
              <CopyButton value={sp.appId} label="application (client) ID" />
            </div>
          </div>
        </div>
        {appRegId && (
          <div className="row">
            <Link
              to={`/applications/${appRegId}`}
              style={{ textDecoration: 'none' }}
            >
              <button>View app registration →</button>
            </Link>
          </div>
        )}
      </div>

      <div className="tabs">
        <button
          className={tab === 'overview' ? 'active' : ''}
          onClick={() => setTab('overview')}
        >
          Overview
        </button>
        <button
          className={tab === 'permissions' ? 'active' : ''}
          onClick={() => setTab('permissions')}
        >
          API permissions
        </button>
        <button
          className={tab === 'assignments' ? 'active' : ''}
          onClick={() => setTab('assignments')}
        >
          Users and groups
        </button>
        <button
          className={tab === 'audit' ? 'active' : ''}
          onClick={() => setTab('audit')}
        >
          Audit
        </button>
      </div>

      {tab === 'overview' && (
        <>
          <div className="card">
            <h3>Details</h3>
            <div className="kv">
              <div className="k">Object ID</div>
              <div className="mono">
                {sp.id}
                <CopyButton value={sp.id} label="object ID" />
              </div>
              <div className="k">App ID</div>
              <div className="mono">
                {sp.appId}
                <CopyButton value={sp.appId} label="app ID" />
              </div>
              <div className="k">Type</div>
              <div>{sp.servicePrincipalType ?? '—'}</div>
              <div className="k">Publisher</div>
              <div>{sp.publisherName ?? '—'}</div>
              <div className="k">Home tenant</div>
              <div>
                <HomeTenantCell ownerTenantId={sp.appOwnerOrganizationId} />
              </div>
              <div className="k">Enabled</div>
              <div className="row">
                {sp.accountEnabled ? (
                  <span className="badge granted">Enabled</span>
                ) : (
                  <span className="badge">Disabled</span>
                )}
                <button
                  onClick={toggleEnabled}
                  disabled={togglingEnabled}
                  style={{ marginLeft: 8 }}
                >
                  {togglingEnabled
                    ? '…'
                    : sp.accountEnabled
                      ? 'Disable'
                      : 'Enable'}
                </button>
              </div>
              <div className="k">Tags</div>
              <div className="muted">{sp.tags?.join(', ') || '—'}</div>
              <div className="k">Created</div>
              <div>
                {sp.createdDateTime
                  ? new Date(sp.createdDateTime).toLocaleString()
                  : <span className="muted">—</span>}
              </div>
              <div className="k">Last delegated sign-in</div>
              <div>
                <LastSignInCell signIn={lastDelegated} error={signInError} />
              </div>
              <div className="k">Last app-only sign-in</div>
              <div>
                <LastSignInCell signIn={lastAppOnly} error={signInError} />
              </div>
            </div>
          </div>

          <SpCredentialsCard sp={sp} credActivity={credActivity} />

          <div className="card">
            <h3>Danger zone</h3>
            <div className="row" style={{ justifyContent: 'space-between' }}>
              <div className="muted" style={{ fontSize: 13 }}>
                Deletes the service principal. Users lose access immediately;
                the underlying app registration (if any) is not affected.
              </div>
              <button
                className="danger"
                onClick={() => setConfirmDelete(true)}
              >
                Delete enterprise app
              </button>
            </div>
          </div>
        </>
      )}

      {tab === 'permissions' && <PermissionsManager clientSp={sp} />}
      {tab === 'assignments' && <AssignedUsersManager sp={sp} />}
      {tab === 'audit' && <SignInAuditTab sp={sp} />}

      {confirmDelete && (
        <Modal
          title="Delete enterprise app"
          onClose={() => !deleting && setConfirmDelete(false)}
          footer={
            <>
              <button
                disabled={deleting}
                onClick={() => setConfirmDelete(false)}
              >
                Cancel
              </button>
              <button className="danger" disabled={deleting} onClick={doDelete}>
                {deleting ? 'Deleting…' : 'Delete'}
              </button>
            </>
          }
        >
          <p>
            Delete <strong>{sp.displayName}</strong>? This removes the service
            principal and all its role assignments and grants. This cannot be
            undone.
          </p>
          <p className="mono muted" style={{ fontSize: 12 }}>
            {sp.id}
          </p>
        </Modal>
      )}
    </>
  );
}

function HomeTenantCell({ ownerTenantId }: { ownerTenantId?: string }) {
  const token = useGraphToken();
  const currentTenantId = useCurrentTenantId();
  const [info, setInfo] = useState<TenantInformation | null>(null);
  const [lookupFailed, setLookupFailed] = useState(false);

  const isHome =
    !ownerTenantId ||
    (currentTenantId && ownerTenantId.toLowerCase() === currentTenantId.toLowerCase());

  useEffect(() => {
    if (isHome || !ownerTenantId) return;
    let cancelled = false;
    setInfo(null);
    setLookupFailed(false);
    findTenantInformation(token, ownerTenantId)
      .then((r) => {
        if (!cancelled) setInfo(r);
      })
      .catch(() => {
        if (!cancelled) setLookupFailed(true);
      });
    return () => {
      cancelled = true;
    };
  }, [token, ownerTenantId, isHome]);

  if (!ownerTenantId) {
    return <span className="muted">—</span>;
  }

  if (isHome) {
    return (
      <div>
        <span className="badge granted">This tenant</span>
        <div className="mono muted" style={{ fontSize: 11, marginTop: 2 }}>
          {ownerTenantId}
        </div>
      </div>
    );
  }

  return (
    <div>
      <div className="row" style={{ gap: 8 }}>
        <span className="badge">External tenant</span>
        {info?.displayName ? (
          <span style={{ fontWeight: 500 }}>{info.displayName}</span>
        ) : lookupFailed ? (
          <span className="muted" style={{ fontSize: 12 }}>
            (name unavailable)
          </span>
        ) : (
          <span className="spinner" />
        )}
      </div>
      {info?.defaultDomainName && (
        <div className="muted" style={{ fontSize: 12, marginTop: 2 }}>
          {info.defaultDomainName}
        </div>
      )}
      <div className="mono muted" style={{ fontSize: 11, marginTop: 2 }}>
        {ownerTenantId}
      </div>
    </div>
  );
}

// Reads directly from /auditLogs/signIns (via signIns.ts). Mirrors what the
// Audit tab shows — the older /reports/servicePrincipalSignInActivities
// endpoint often lags for hours/days or returns empty aggregate rows, which
// made this card misleadingly claim "no sign-ins" while the actual logs were
// full.
function LastSignInCell({
  signIn,
  error,
}: {
  signIn: SignIn | null | 'loading';
  error: string | null;
}) {
  if (signIn === 'loading') {
    return (
      <span className="muted">
        <span className="spinner" />
      </span>
    );
  }
  if (!signIn) {
    if (error) {
      return (
        <span className="muted" style={{ fontSize: 13 }}>
          Unavailable — {error}
        </span>
      );
    }
    return <span className="muted">No sign-ins in the last 30 days</span>;
  }
  const when = signIn.createdDateTime;
  const actor =
    signIn.userDisplayName ??
    signIn.userPrincipalName ??
    signIn.servicePrincipalName ??
    signIn.appDisplayName;
  return (
    <>
      <div title={when}>
        {when ? new Date(when).toLocaleString() : '—'}
      </div>
      {actor && (
        <div className="muted" style={{ fontSize: 12 }}>
          by {actor}
          {signIn.resourceDisplayName ? ` → ${signIn.resourceDisplayName}` : ''}
        </div>
      )}
    </>
  );
}

function SpCredentialsCard({
  sp,
  credActivity,
}: {
  sp: ServicePrincipal;
  credActivity: Map<string, AppCredentialSignInActivity>;
}) {
  const secrets = sp.passwordCredentials ?? [];
  const rawCerts = sp.keyCredentials ?? [];
  // keyCredentials are only meaningful on Legacy-type service principals
  // (the old password-SSO / legacy SAML bucket). For Application,
  // ManagedIdentity, and SocialIdp types they're typically empty or internal
  // noise — hide the Certificates section there so the UI doesn't mislead.
  const showCerts =
    sp.servicePrincipalType === 'Legacy' && rawCerts.length > 0;
  const certs = showCerts ? rawCerts : [];
  if (secrets.length === 0 && !showCerts) {
    return (
      <div className="card">
        <h3>Credentials</h3>
        <div className="muted" style={{ fontSize: 13 }}>
          This service principal has no secrets or certificates of its own.
          First-party Microsoft apps and most multi-tenant apps fall in this
          bucket — credentials live on the app registration in the home
          tenant. SAML-federated apps and some SaaS apps keep signing certs
          here.
        </div>
      </div>
    );
  }
  return (
    <div className="card">
      <h3>Credentials</h3>

      {secrets.length > 0 && (
        <>
          <h4 style={{ marginTop: 8, marginBottom: 8 }}>
            Client secrets ({secrets.length})
          </h4>
          <table className="table">
            <thead>
              <tr>
                <th>Description</th>
                <th>Secret ID</th>
                <th>Hint</th>
                <th>Created</th>
                <th>Expires</th>
                <th>Last used</th>
              </tr>
            </thead>
            <tbody>
              {secrets.map((s) => (
                <SpSecretRow
                  key={s.keyId}
                  secret={s}
                  activity={credActivity.get(s.keyId)}
                />
              ))}
            </tbody>
          </table>
        </>
      )}

      {showCerts && (
        <>
          <h4 style={{ marginTop: 20, marginBottom: 8 }}>
            Certificates ({certs.length})
          </h4>
          <table className="table">
            <thead>
              <tr>
                <th>Name</th>
                <th>Thumbprint</th>
                <th>Type</th>
                <th>Usage</th>
                <th>Created</th>
                <th>Expires</th>
                <th>Last used</th>
              </tr>
            </thead>
            <tbody>
              {certs.map((c) => (
                <SpCertRow
                  key={c.keyId}
                  cert={c}
                  activity={credActivity.get(c.keyId)}
                />
              ))}
            </tbody>
          </table>
        </>
      )}
    </div>
  );
}

function SpSecretRow({
  secret,
  activity,
}: {
  secret: PasswordCredential;
  activity?: AppCredentialSignInActivity;
}) {
  const last = activity?.signInActivity?.lastSignInDateTime;
  return (
    <tr>
      <td>{secret.displayName || <span className="muted">—</span>}</td>
      <td className="mono" title={secret.keyId} style={{ fontSize: 12 }}>
        {shortenId(secret.keyId)}
        <CopyButton value={secret.keyId} label="keyId" />
      </td>
      <td className="mono">
        {secret.hint ? `${secret.hint}…` : <span className="muted">—</span>}
      </td>
      <td>{formatDateTime(secret.startDateTime)}</td>
      <td>
        <div className="row" style={{ gap: 6 }}>
          <span>{formatDate(secret.endDateTime)}</span>
          <ExpiryBadge end={secret.endDateTime} />
        </div>
      </td>
      <td>
        {last ? (
          <span title={last}>{new Date(last).toLocaleString()}</span>
        ) : (
          <span className="muted">—</span>
        )}
      </td>
    </tr>
  );
}

function SpCertRow({
  cert,
  activity,
}: {
  cert: KeyCredential;
  activity?: AppCredentialSignInActivity;
}) {
  const last = activity?.signInActivity?.lastSignInDateTime;
  const thumb = formatThumbprintHex(cert.customKeyIdentifier);
  return (
    <tr>
      <td>{cert.displayName || <span className="muted">—</span>}</td>
      <td
        className="mono"
        title={thumb || cert.keyId}
        style={{ fontSize: 12 }}
      >
        {thumb ? shortenId(thumb) : shortenId(cert.keyId)}
        <CopyButton value={thumb || cert.keyId} label="thumbprint" />
      </td>
      <td>{cert.type ?? <span className="muted">—</span>}</td>
      <td>{cert.usage ?? <span className="muted">—</span>}</td>
      <td>{formatDateTime(cert.startDateTime)}</td>
      <td>
        <div className="row" style={{ gap: 6 }}>
          <span>{formatDate(cert.endDateTime)}</span>
          <ExpiryBadge end={cert.endDateTime} />
        </div>
      </td>
      <td>
        {last ? (
          <span title={last}>{new Date(last).toLocaleString()}</span>
        ) : (
          <span className="muted">—</span>
        )}
      </td>
    </tr>
  );
}

function formatDate(d?: string): string {
  if (!d) return '—';
  const t = new Date(d).getTime();
  return Number.isNaN(t) ? '—' : new Date(t).toLocaleDateString();
}
function formatDateTime(d?: string): string {
  if (!d) return '—';
  const t = new Date(d).getTime();
  return Number.isNaN(t) ? '—' : new Date(t).toLocaleString();
}
function formatThumbprintHex(b64?: string | null): string {
  if (!b64) return '';
  try {
    const bin = atob(b64);
    let hex = '';
    for (let i = 0; i < bin.length; i++) {
      hex += bin.charCodeAt(i).toString(16).padStart(2, '0');
    }
    return hex.toUpperCase();
  } catch {
    return b64;
  }
}
function shortenId(s: string): string {
  return s.length > 14 ? `${s.slice(0, 8)}…${s.slice(-4)}` : s;
}
function ExpiryBadge({ end }: { end?: string }) {
  if (!end) return null;
  const endMs = new Date(end).getTime();
  if (Number.isNaN(endMs)) return null;
  const now = Date.now();
  if (endMs < now) return <span className="badge expired">Expired</span>;
  if (endMs - now < 30 * 24 * 60 * 60 * 1000)
    return <span className="badge pending">Expires soon</span>;
  return <span className="badge granted">Active</span>;
}
