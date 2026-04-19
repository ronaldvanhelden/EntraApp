import { useEffect, useState, type ReactNode } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import {
  addApplicationCertificate,
  addApplicationPassword,
  deleteApplication,
  getApplication,
  removeApplicationCertificate,
  removeApplicationPassword,
  updateApplication,
  type AddPasswordInput,
  type UpdateApplicationPatch,
} from '../graph/applications';
import {
  AddCertificateModal,
  type CertSubmitPayload,
} from '../components/AddCertificateModal';
import {
  deleteServicePrincipal,
  getServicePrincipalByAppId,
} from '../graph/servicePrincipals';
import {
  buildCredentialActivityMap,
  listAppCredentialSignInActivities,
  listFederatedIdentityCredentials,
} from '../graph/credentials';
import type {
  AppCredentialSignInActivity,
  Application,
  FederatedIdentityCredential,
  KeyCredential,
  PasswordCredential,
  RequiredResourceAccess,
  ServicePrincipal,
} from '../graph/types';
import { Modal } from '../components/Modal';
import { AppIcon } from '../components/AppIcon';
import { CopyButton } from '../components/CopyButton';
import { PrivilegeBadge, useScopeCatalog } from '../components/PrivilegeBadge';
import { ScopeDetailsModal } from '../components/ScopeDetailsModal';
import { AddRequiredPermissionModal } from '../components/AddRequiredPermissionModal';
import { getSignIn } from '../graph/signIns';
import type { SignIn } from '../graph/types';

const AUDIENCES = [
  'AzureADMyOrg',
  'AzureADMultipleOrgs',
  'AzureADandPersonalMicrosoftAccount',
  'PersonalMicrosoftAccount',
];

export function ApplicationDetail() {
  const { id = '' } = useParams();
  const token = useGraphToken();
  const nav = useNavigate();
  const [app, setApp] = useState<Application | null>(null);
  const [sp, setSp] = useState<ServicePrincipal | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [editing, setEditing] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [fic, setFic] = useState<FederatedIdentityCredential[] | null>(null);
  const [activityByKeyId, setActivityByKeyId] = useState<
    Map<string, AppCredentialSignInActivity>
  >(new Map());
  const [activityError, setActivityError] = useState<string | null>(null);
  const [resourceSps, setResourceSps] = useState<
    Record<string, ServicePrincipal | null>
  >({});

  useEffect(() => {
    setApp(null);
    setSp(null);
    setError(null);
    setEditing(false);
    setFic(null);
    setActivityByKeyId(new Map());
    setActivityError(null);
    setResourceSps({});
    getApplication(token, id)
      .then(async (a) => {
        setApp(a);
        getServicePrincipalByAppId(token, a.appId)
          .then((maybe) => setSp(maybe ?? null))
          .catch(() => {
            /* no SP yet — fine */
          });
        listFederatedIdentityCredentials(token, a.id)
          .then(setFic)
          .catch(() => setFic([]));
        listAppCredentialSignInActivities(token, a.appId)
          .then((rows) => setActivityByKeyId(buildCredentialActivityMap(rows)))
          .catch((e: unknown) =>
            setActivityError(e instanceof Error ? e.message : String(e)),
          );
      })
      .catch((e) => setError(e.message));
  }, [token, id]);

  useEffect(() => {
    if (!app?.requiredResourceAccess?.length) return;
    const toFetch = Array.from(
      new Set(app.requiredResourceAccess.map((r) => r.resourceAppId)),
    ).filter((appId) => !(appId in resourceSps));
    toFetch.forEach((appId) => {
      getServicePrincipalByAppId(token, appId)
        .then((resSp) =>
          setResourceSps((prev) => ({ ...prev, [appId]: resSp ?? null })),
        )
        .catch(() => setResourceSps((prev) => ({ ...prev, [appId]: null })));
    });
  }, [app, token, resourceSps]);

  if (error) return <div className="card error">{error}</div>;
  if (!app)
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
            id={app.appId || app.id}
            logoUrl={app.info?.logoUrl}
            size={48}
            title={app.displayName}
          />
          <div>
            <h1 style={{ margin: 0 }}>{app.displayName}</h1>
            <div className="subtitle mono">
              {app.appId}
              <CopyButton value={app.appId} label="application (client) ID" />
            </div>
          </div>
        </div>
        <div className="row">
          {!editing && (
            <button onClick={() => setEditing(true)}>Edit</button>
          )}
          {sp && (
            <Link
              to={`/enterprise-apps/${sp.id}`}
              style={{ textDecoration: 'none' }}
            >
              <button className="primary">Manage enterprise app →</button>
            </Link>
          )}
        </div>
      </div>

      {editing ? (
        <EditCard
          app={app}
          onCancel={() => setEditing(false)}
          onSaved={(updated) => {
            setApp({ ...app, ...updated });
            setEditing(false);
          }}
        />
      ) : (
        <div className="card">
          <h3>Details</h3>
          <div className="kv">
            <div className="k">Object ID</div>
            <div className="mono">
              {app.id}
              <CopyButton value={app.id} label="object ID" />
            </div>
            <div className="k">Application (client) ID</div>
            <div className="mono">
              {app.appId}
              <CopyButton value={app.appId} label="application (client) ID" />
            </div>
            <div className="k">Display name</div>
            <div>{app.displayName}</div>
            <div className="k">Sign-in audience</div>
            <div>{app.signInAudience}</div>
            <div className="k">Publisher domain</div>
            <div>{app.publisherDomain ?? '—'}</div>
            <div className="k">Identifier URIs</div>
            <div>
              {app.identifierUris?.length
                ? app.identifierUris.map((u) => (
                    <div key={u} className="mono">
                      {u}
                    </div>
                  ))
                : '—'}
            </div>
            <div className="k">Notes</div>
            <div className="muted">{app.notes || '—'}</div>
            <div className="k">Created</div>
            <div>
              {app.createdDateTime
                ? new Date(app.createdDateTime).toLocaleString()
                : '—'}
            </div>
            <div className="k">OAuth flows</div>
            <div>
              <OAuthFlowSummary app={app} />
            </div>
          </div>
        </div>
      )}

      <CredentialsCard
        app={app}
        fic={fic}
        activityByKeyId={activityByKeyId}
        activityError={activityError}
        onAppChange={setApp}
      />

      <AuthenticationCard app={app} onAppChange={setApp} />

      <RequiredResourceAccessCard
        app={app}
        resourceSps={resourceSps}
        onChange={setApp}
      />

      <div className="card">
        <h3>Danger zone</h3>
        <div className="row" style={{ justifyContent: 'space-between' }}>
          <div className="muted" style={{ fontSize: 13 }}>
            Deleting the app registration removes the application object. The
            service principal is separate and must be deleted explicitly.
          </div>
          <button className="danger" onClick={() => setConfirmDelete(true)}>
            Delete app registration
          </button>
        </div>
      </div>

      {confirmDelete && (
        <DeleteModal
          app={app}
          sp={sp}
          onClose={() => setConfirmDelete(false)}
          onDeleted={() => nav('/applications')}
        />
      )}
    </>
  );
}

function EditCard({
  app,
  onCancel,
  onSaved,
}: {
  app: Application;
  onCancel: () => void;
  onSaved: (patch: Partial<Application>) => void;
}) {
  const token = useGraphToken();
  const [displayName, setDisplayName] = useState(app.displayName);
  const [audience, setAudience] = useState(app.signInAudience ?? AUDIENCES[0]);
  const [notes, setNotes] = useState(app.notes ?? '');
  const [identifierUris, setIdentifierUris] = useState(
    (app.identifierUris ?? []).join('\n'),
  );
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const save = async () => {
    setSaving(true);
    setError(null);
    try {
      const uris = identifierUris
        .split(/\r?\n/)
        .map((s) => s.trim())
        .filter(Boolean);
      const patch = {
        displayName: displayName.trim(),
        signInAudience: audience,
        notes: notes.trim() || undefined,
        identifierUris: uris,
      };
      await updateApplication(token, app.id, patch);
      onSaved(patch);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="card">
      <h3>Edit details</h3>
      <label className="field">
        <span>Display name</span>
        <input
          value={displayName}
          onChange={(e) => setDisplayName(e.target.value)}
        />
      </label>
      <label className="field">
        <span>Sign-in audience</span>
        <select value={audience} onChange={(e) => setAudience(e.target.value)}>
          {AUDIENCES.map((a) => (
            <option key={a} value={a}>
              {a}
            </option>
          ))}
        </select>
      </label>
      <label className="field">
        <span>Identifier URIs (one per line)</span>
        <textarea
          rows={3}
          value={identifierUris}
          onChange={(e) => setIdentifierUris(e.target.value)}
          placeholder="api://..."
        />
      </label>
      <label className="field">
        <span>Notes</span>
        <textarea
          rows={3}
          value={notes}
          onChange={(e) => setNotes(e.target.value)}
        />
      </label>
      {error && <p className="error">{error}</p>}
      <div className="row" style={{ justifyContent: 'flex-end', gap: 8 }}>
        <button disabled={saving} onClick={onCancel}>
          Cancel
        </button>
        <button
          className="primary"
          disabled={saving || !displayName.trim()}
          onClick={save}
        >
          {saving ? 'Saving…' : 'Save'}
        </button>
      </div>
    </div>
  );
}

function DeleteModal({
  app,
  sp,
  onClose,
  onDeleted,
}: {
  app: Application;
  sp: ServicePrincipal | null;
  onClose: () => void;
  onDeleted: () => void;
}) {
  const token = useGraphToken();
  const [alsoDeleteSp, setAlsoDeleteSp] = useState(Boolean(sp));
  const [deleting, setDeleting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const submit = async () => {
    setDeleting(true);
    setError(null);
    try {
      if (alsoDeleteSp && sp) {
        try {
          await deleteServicePrincipal(token, sp.id);
        } catch (e: unknown) {
          setError(
            `Failed to delete service principal: ${
              e instanceof Error ? e.message : String(e)
            }`,
          );
          setDeleting(false);
          return;
        }
      }
      await deleteApplication(token, app.id);
      onDeleted();
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setDeleting(false);
    }
  };

  return (
    <Modal
      title="Delete app registration"
      onClose={() => !deleting && onClose()}
      footer={
        <>
          <button disabled={deleting} onClick={onClose}>
            Cancel
          </button>
          <button className="danger" disabled={deleting} onClick={submit}>
            {deleting ? 'Deleting…' : 'Delete'}
          </button>
        </>
      }
    >
      <p>
        Delete <strong>{app.displayName}</strong>?
      </p>
      <p className="mono muted" style={{ fontSize: 12 }}>
        {app.appId}
      </p>
      {sp && (
        <label className="row" style={{ gap: 6, marginTop: 12 }}>
          <input
            type="checkbox"
            style={{ width: 'auto' }}
            checked={alsoDeleteSp}
            onChange={(e) => setAlsoDeleteSp(e.target.checked)}
          />
          Also delete the enterprise app (service principal) in this tenant
        </label>
      )}
      {error && <p className="error" style={{ marginTop: 12 }}>{error}</p>}
    </Modal>
  );
}

function expiryStatus(
  end?: string,
): 'expired' | 'soon' | 'active' | null {
  if (!end) return null;
  const ms = new Date(end).getTime();
  if (Number.isNaN(ms)) return null;
  const now = Date.now();
  if (ms < now) return 'expired';
  if (ms - now < 30 * 24 * 60 * 60 * 1000) return 'soon';
  return 'active';
}

function ExpiryBadge({ end }: { end?: string }) {
  const status = expiryStatus(end);
  if (!status) return null;
  if (status === 'expired')
    return <span className="badge expired">Expired</span>;
  if (status === 'soon')
    return <span className="badge pending">Expires soon</span>;
  return <span className="badge granted">Active</span>;
}

function formatDate(d?: string) {
  if (!d) return '—';
  const ms = new Date(d).getTime();
  if (Number.isNaN(ms)) return '—';
  return new Date(ms).toLocaleDateString();
}

function formatDateTime(d?: string) {
  if (!d) return '—';
  const ms = new Date(d).getTime();
  if (Number.isNaN(ms)) return '—';
  return new Date(ms).toLocaleString();
}

function formatThumbprint(b64?: string | null) {
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

function shortenKeyId(keyId: string) {
  return keyId.length > 13 ? `${keyId.slice(0, 8)}…${keyId.slice(-4)}` : keyId;
}

function LastUsedCell({
  activity,
}: {
  activity?: AppCredentialSignInActivity;
}) {
  if (!activity) return <span className="muted">—</span>;
  const last = activity.signInActivity?.lastSignInDateTime;
  if (!last) return <span className="muted">Never</span>;
  return <span title={last}>{formatDateTime(last)}</span>;
}

function ResourceCell({
  activity,
}: {
  activity?: AppCredentialSignInActivity;
}) {
  const name = activity?.signInActivity?.resourceDisplayName;
  return name ? <span>{name}</span> : <span className="muted">—</span>;
}

// Compact read-only summary of which OAuth 2.0 grant types the app is wired
// for — shown inline in the Details card. For toggling + full redirect URI
// editing the user drops into the dedicated AuthenticationCard below.
function OAuthFlowSummary({ app }: { app: Application }) {
  const flows: Array<{ label: string; title: string }> = [];
  const web = app.web;
  const spa = app.spa;
  const publicClient = app.publicClient;

  if ((web?.redirectUris?.length ?? 0) > 0) {
    flows.push({
      label: 'Auth code (Web)',
      title: 'Confidential client — authorization code + client secret/cert',
    });
  }
  if ((spa?.redirectUris?.length ?? 0) > 0) {
    flows.push({
      label: 'Auth code + PKCE (SPA)',
      title: 'Single-page app — authorization code + PKCE, no client secret',
    });
  }
  if ((publicClient?.redirectUris?.length ?? 0) > 0) {
    flows.push({
      label: 'Auth code (Public)',
      title: 'Native / public client — authorization code on a mobile or desktop app',
    });
  }
  if (app.isFallbackPublicClient) {
    flows.push({
      label: 'Device code / ROPC',
      title: 'isFallbackPublicClient = true enables device-code and ROPC flows',
    });
  }
  if (web?.implicitGrantSettings?.enableIdTokenIssuance) {
    flows.push({
      label: 'Implicit (ID token)',
      title: 'Legacy implicit flow issuing ID tokens',
    });
  }
  if (web?.implicitGrantSettings?.enableAccessTokenIssuance) {
    flows.push({
      label: 'Implicit (access token)',
      title: 'Legacy implicit flow issuing access tokens',
    });
  }
  // Client credentials is always available for confidential clients that
  // hold at least one secret or certificate.
  const hasCred =
    (app.passwordCredentials?.length ?? 0) > 0 ||
    (app.keyCredentials?.length ?? 0) > 0;
  if (hasCred) {
    flows.push({
      label: 'Client credentials',
      title:
        'App-only: uses a client secret, certificate, or FIC to obtain tokens without a user',
    });
  }

  if (flows.length === 0) {
    return (
      <span className="muted">
        No flows configured — add a redirect URI or credential in the
        Authentication section below.
      </span>
    );
  }
  return (
    <div className="row wrap" style={{ gap: 6 }}>
      {flows.map((f) => (
        <span key={f.label} className="badge granted" title={f.title}>
          {f.label}
        </span>
      ))}
    </div>
  );
}

function AuthenticationCard({
  app,
  onAppChange,
}: {
  app: Application;
  onAppChange: (next: Application) => void;
}) {
  const token = useGraphToken();
  const [saving, setSaving] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [editPlatform, setEditPlatform] = useState<null | 'web' | 'spa' | 'publicClient'>(null);

  const spaUris = app.spa?.redirectUris ?? [];
  const webUris = app.web?.redirectUris ?? [];
  const publicUris = app.publicClient?.redirectUris ?? [];
  const idTokenIssuance = !!app.web?.implicitGrantSettings?.enableIdTokenIssuance;
  const accessTokenIssuance =
    !!app.web?.implicitGrantSettings?.enableAccessTokenIssuance;
  const fallbackPublic = !!app.isFallbackPublicClient;

  const patch = async (
    body: UpdateApplicationPatch,
    local: Partial<Application>,
    key: string,
  ) => {
    setSaving(key);
    setError(null);
    try {
      await updateApplication(token, app.id, body);
      onAppChange({ ...app, ...local });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(null);
    }
  };

  const toggleFallbackPublic = () => {
    const next = !fallbackPublic;
    patch({ isFallbackPublicClient: next }, { isFallbackPublicClient: next }, 'fallback');
  };

  const toggleIdToken = () => {
    const next = !idTokenIssuance;
    const web = {
      ...app.web,
      implicitGrantSettings: {
        ...app.web?.implicitGrantSettings,
        enableIdTokenIssuance: next,
        enableAccessTokenIssuance: accessTokenIssuance,
      },
    };
    patch({ web }, { web }, 'idtoken');
  };

  const toggleAccessToken = () => {
    const next = !accessTokenIssuance;
    const web = {
      ...app.web,
      implicitGrantSettings: {
        ...app.web?.implicitGrantSettings,
        enableIdTokenIssuance: idTokenIssuance,
        enableAccessTokenIssuance: next,
      },
    };
    patch({ web }, { web }, 'accesstoken');
  };

  const saveRedirectUris = async (
    platform: 'web' | 'spa' | 'publicClient',
    uris: string[],
  ) => {
    const local: Partial<Application> = {};
    const body: UpdateApplicationPatch = {};
    if (platform === 'web') {
      body.web = { ...app.web, redirectUris: uris };
      local.web = body.web;
    } else if (platform === 'spa') {
      body.spa = { ...app.spa, redirectUris: uris };
      local.spa = body.spa;
    } else {
      body.publicClient = { ...app.publicClient, redirectUris: uris };
      local.publicClient = body.publicClient;
    }
    await patch(body, local, `uris:${platform}`);
    setEditPlatform(null);
  };

  return (
    <div className="card">
      <h3>Authentication</h3>
      <div className="row wrap" style={{ gap: 8, marginBottom: 12 }}>
        <FlowBadge
          active={webUris.length > 0}
          label={`Web (${webUris.length})`}
          hint="Confidential client — redirect URIs live on web.redirectUris. Uses client secret or certificate."
        />
        <FlowBadge
          active={spaUris.length > 0}
          label={`SPA (${spaUris.length})`}
          hint="Single-page app — Authorization Code + PKCE. No client secret."
        />
        <FlowBadge
          active={publicUris.length > 0}
          label={`Mobile/Desktop (${publicUris.length})`}
          hint="Native public client redirect URIs (e.g. MSAL mobile/desktop apps)."
        />
        <FlowBadge
          active={fallbackPublic}
          label="Allow public client flows"
          hint="Enables device-code and ROPC flows. Required for apps that can't protect a secret."
        />
        <FlowBadge
          active={idTokenIssuance}
          label="Implicit ID token"
          hint="Legacy implicit flow for ID tokens. Prefer Auth Code + PKCE."
        />
        <FlowBadge
          active={accessTokenIssuance}
          label="Implicit access token"
          hint="Legacy implicit flow for access tokens. Prefer Auth Code + PKCE."
        />
      </div>

      {error && <p className="error">{error}</p>}

      <RedirectUriBlock
        title="Web"
        subtitle="Authorization code flow with client secret / certificate."
        uris={webUris}
        onEdit={() => setEditPlatform('web')}
        disabled={saving !== null}
      />
      <RedirectUriBlock
        title="Single-page application (SPA)"
        subtitle="Authorization code + PKCE. No client secret."
        uris={spaUris}
        onEdit={() => setEditPlatform('spa')}
        disabled={saving !== null}
      />
      <RedirectUriBlock
        title="Mobile / desktop (public client)"
        subtitle="Native MSAL apps and custom URI schemes."
        uris={publicUris}
        onEdit={() => setEditPlatform('publicClient')}
        disabled={saving !== null}
      />

      <h4 style={{ marginTop: 20, marginBottom: 8 }}>Advanced settings</h4>
      <div className="kv">
        <div className="k">Allow public client flows</div>
        <div>
          <label className="row" style={{ gap: 6 }}>
            <input
              type="checkbox"
              style={{ width: 'auto' }}
              checked={fallbackPublic}
              disabled={saving === 'fallback'}
              onChange={toggleFallbackPublic}
            />
            {fallbackPublic ? 'Yes' : 'No'}{' '}
            <span className="muted" style={{ fontSize: 12 }}>
              (device code, ROPC)
            </span>
          </label>
        </div>
        <div className="k">Implicit grant — ID tokens</div>
        <div>
          <label className="row" style={{ gap: 6 }}>
            <input
              type="checkbox"
              style={{ width: 'auto' }}
              checked={idTokenIssuance}
              disabled={saving === 'idtoken'}
              onChange={toggleIdToken}
            />
            {idTokenIssuance ? 'Issued' : 'Disabled'}
          </label>
        </div>
        <div className="k">Implicit grant — access tokens</div>
        <div>
          <label className="row" style={{ gap: 6 }}>
            <input
              type="checkbox"
              style={{ width: 'auto' }}
              checked={accessTokenIssuance}
              disabled={saving === 'accesstoken'}
              onChange={toggleAccessToken}
            />
            {accessTokenIssuance ? 'Issued' : 'Disabled'}
          </label>
        </div>
      </div>

      {editPlatform && (
        <RedirectUriEditModal
          title={
            editPlatform === 'web'
              ? 'Edit Web redirect URIs'
              : editPlatform === 'spa'
                ? 'Edit SPA redirect URIs'
                : 'Edit public-client redirect URIs'
          }
          initial={
            editPlatform === 'web'
              ? webUris
              : editPlatform === 'spa'
                ? spaUris
                : publicUris
          }
          onClose={() => setEditPlatform(null)}
          onSave={(uris) => saveRedirectUris(editPlatform, uris)}
        />
      )}
    </div>
  );
}

function FlowBadge({
  active,
  label,
  hint,
}: {
  active: boolean;
  label: string;
  hint?: string;
}) {
  return (
    <span
      className={`badge ${active ? 'granted' : ''}`}
      title={hint}
      style={{ opacity: active ? 1 : 0.55 }}
    >
      {active ? '✓ ' : '○ '}
      {label}
    </span>
  );
}

function RedirectUriBlock({
  title,
  subtitle,
  uris,
  onEdit,
  disabled,
}: {
  title: string;
  subtitle: string;
  uris: string[];
  onEdit: () => void;
  disabled?: boolean;
}) {
  return (
    <div style={{ marginTop: 12 }}>
      <div className="row" style={{ justifyContent: 'space-between' }}>
        <div>
          <div style={{ fontWeight: 600 }}>{title}</div>
          <div className="muted" style={{ fontSize: 12 }}>
            {subtitle}
          </div>
        </div>
        <button onClick={onEdit} disabled={disabled}>
          Edit
        </button>
      </div>
      {uris.length === 0 ? (
        <div className="muted" style={{ fontSize: 13, marginTop: 4 }}>
          No redirect URIs configured.
        </div>
      ) : (
        <ul
          className="mono"
          style={{
            fontSize: 12,
            margin: '6px 0 0',
            paddingLeft: 20,
          }}
        >
          {uris.map((u) => (
            <li key={u}>
              {u}
              <CopyButton value={u} label="redirect URI" />
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

function RedirectUriEditModal({
  title,
  initial,
  onClose,
  onSave,
}: {
  title: string;
  initial: string[];
  onClose: () => void;
  onSave: (uris: string[]) => Promise<void>;
}) {
  const [text, setText] = useState(initial.join('\n'));
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const save = async () => {
    const uris = text
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);
    setSaving(true);
    setError(null);
    try {
      await onSave(uris);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setSaving(false);
    }
  };

  return (
    <Modal
      title={title}
      onClose={() => !saving && onClose()}
      footer={
        <>
          <button disabled={saving} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={saving} onClick={save}>
            {saving ? 'Saving…' : 'Save'}
          </button>
        </>
      }
    >
      <p className="muted" style={{ fontSize: 13, marginTop: 0 }}>
        One URI per line. Must be an HTTPS URL (or <code>http://localhost</code>),
        except for public-client URIs which also allow custom schemes.
      </p>
      <textarea
        value={text}
        onChange={(e) => setText(e.target.value)}
        className="mono"
        rows={Math.max(6, Math.min(16, text.split('\n').length + 2))}
        style={{ width: '100%', fontSize: 12 }}
        placeholder="https://contoso.example.com/auth/callback"
        autoFocus
      />
      {error && <p className="error">{error}</p>}
    </Modal>
  );
}

function RequiredResourceAccessCard({
  app,
  resourceSps,
  onChange,
}: {
  app: Application;
  resourceSps: Record<string, ServicePrincipal | null>;
  onChange: (next: Application) => void;
}) {
  const token = useGraphToken();
  const [showAdd, setShowAdd] = useState(false);
  const [working, setWorking] = useState<string | null>(null);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set());
  const catalog = useScopeCatalog();
  const [detailsFor, setDetailsFor] = useState<{
    scope: string;
    kind: 'delegated' | 'application';
    resourceAppId: string;
    resourceName?: string;
    title?: string;
    description?: string;
  } | null>(null);

  const groups: RequiredResourceAccess[] = app.requiredResourceAccess ?? [];
  const toggleCollapsed = (key: string) =>
    setCollapsed((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });

  const commit = async (next: RequiredResourceAccess[], workingKey: string) => {
    setSaveError(null);
    setWorking(workingKey);
    try {
      await updateApplication(token, app.id, { requiredResourceAccess: next });
      onChange({ ...app, requiredResourceAccess: next });
    } catch (e: unknown) {
      setSaveError(e instanceof Error ? e.message : String(e));
      throw e;
    } finally {
      setWorking(null);
    }
  };

  const addEntry = async (entry: RequiredResourceAccess) => {
    const idx = groups.findIndex((g) => g.resourceAppId === entry.resourceAppId);
    const merged =
      idx === -1
        ? [...groups, entry]
        : groups.map((g, i) =>
            i === idx
              ? {
                  resourceAppId: g.resourceAppId,
                  resourceAccess: mergeAccess(g.resourceAccess, entry.resourceAccess),
                }
              : g,
          );
    await commit(merged, `add:${entry.resourceAppId}`);
    setShowAdd(false);
  };

  const removePermission = (resourceAppId: string, id: string) => {
    const next = groups
      .map((g) =>
        g.resourceAppId === resourceAppId
          ? {
              ...g,
              resourceAccess: g.resourceAccess.filter((a) => a.id !== id),
            }
          : g,
      )
      .filter((g) => g.resourceAccess.length > 0);
    void commit(next, `perm:${resourceAppId}:${id}`).catch(() => {});
  };

  const removeResource = (resourceAppId: string) => {
    const next = groups.filter((g) => g.resourceAppId !== resourceAppId);
    void commit(next, `res:${resourceAppId}`).catch(() => {});
  };

  return (
    <div className="card">
      <div className="row" style={{ justifyContent: 'space-between' }}>
        <div>
          <h3 style={{ margin: 0 }}>Required resource access (manifest)</h3>
          <div className="muted" style={{ fontSize: 12, marginTop: 4 }}>
            Permissions declared in the application manifest. To take effect,
            they must be consented/granted on the enterprise app.
          </div>
        </div>
        <button className="primary" onClick={() => setShowAdd(true)}>
          + Add a required permission
        </button>
      </div>

      {saveError && (
        <p className="error" style={{ marginTop: 12 }}>{saveError}</p>
      )}

      {groups.length === 0 ? (
        <div className="muted" style={{ marginTop: 12 }}>
          No required resource access declared.
        </div>
      ) : (
        <div style={{ marginTop: 12 }}>
          {groups.map((g) => {
            const resSp = resourceSps[g.resourceAppId];
            const resolving = resSp === undefined;
            const resName =
              resSp?.displayName ?? (resolving ? '…' : g.resourceAppId);
            const isCollapsed = collapsed.has(g.resourceAppId);
            return (
              <div
                key={g.resourceAppId}
                className={`group-box${isCollapsed ? ' collapsed' : ''}`}
              >
                <div
                  className="group-header"
                  onClick={() => toggleCollapsed(g.resourceAppId)}
                >
                  <div className="row" style={{ gap: 10, alignItems: 'center' }}>
                    <span className="chevron" aria-hidden>
                      ▾
                    </span>
                    <AppIcon
                      id={g.resourceAppId}
                      logoUrl={resSp?.info?.logoUrl}
                      size={24}
                      title={resName}
                    />
                    <div>
                      <div className="group-title">{resName}</div>
                      <div className="group-subtitle">
                        {g.resourceAppId}
                        <CopyButton
                          value={g.resourceAppId}
                          label="resource appId"
                        />
                        {' · '}
                        {g.resourceAccess.length} permission
                        {g.resourceAccess.length === 1 ? '' : 's'}
                      </div>
                    </div>
                  </div>
                  <button
                    className="danger"
                    disabled={working === `res:${g.resourceAppId}`}
                    onClick={(e) => {
                      e.stopPropagation();
                      removeResource(g.resourceAppId);
                    }}
                  >
                    {working === `res:${g.resourceAppId}`
                      ? '…'
                      : 'Remove resource'}
                  </button>
                </div>
                {!isCollapsed && (
                <div className="group-body">
                <table className="table">
                  <thead>
                    <tr>
                      <th>Permission</th>
                      <th>Type</th>
                      <th>Description</th>
                      <th style={{ width: 100 }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {g.resourceAccess.map((ra) => {
                      const role =
                        ra.type === 'Role'
                          ? resSp?.appRoles?.find((x) => x.id === ra.id)
                          : undefined;
                      const scope =
                        ra.type === 'Scope'
                          ? resSp?.oauth2PermissionScopes?.find(
                              (x) => x.id === ra.id,
                            )
                          : undefined;
                      const permValue = role?.value ?? scope?.value;
                      const permTitle =
                        role?.displayName ?? scope?.adminConsentDisplayName;
                      const permDesc =
                        role?.description ?? scope?.adminConsentDescription;
                      const key = `perm:${g.resourceAppId}:${ra.id}`;
                      const permKind =
                        ra.type === 'Role' ? 'application' : 'delegated';
                      return (
                        <tr key={ra.id}>
                          <td>
                            {permValue ? (
                              <button
                                type="button"
                                className="link"
                                onClick={() =>
                                  setDetailsFor({
                                    scope: permValue,
                                    kind: permKind,
                                    resourceAppId: g.resourceAppId,
                                    resourceName: resSp?.displayName,
                                    title: permTitle,
                                    description: permDesc,
                                  })
                                }
                                title="Show the URLs this permission unlocks"
                              >
                                <span
                                  className="mono"
                                  style={{ fontWeight: 600 }}
                                >
                                  {permValue}
                                </span>
                              </button>
                            ) : (
                              <div className="mono" style={{ fontWeight: 600 }}>
                                {resolving ? '…' : ra.id}
                              </div>
                            )}
                            {permTitle && (
                              <div className="muted" style={{ fontSize: 12 }}>
                                {permTitle}
                              </div>
                            )}
                            {!permValue && !resolving && (
                              <div className="mono muted" style={{ fontSize: 11 }}>
                                {ra.id}
                              </div>
                            )}
                            <PrivilegeBadge
                              catalog={catalog}
                              resourceAppId={g.resourceAppId}
                              scope={permValue}
                              kind={permKind}
                            />
                          </td>
                          <td>
                            {ra.type === 'Role' ? (
                              <span className="badge app">Application</span>
                            ) : (
                              <span className="badge delegated">Delegated</span>
                            )}
                          </td>
                          <td
                            className="muted"
                            style={{ maxWidth: 380, fontSize: 12 }}
                          >
                            {permDesc ?? ''}
                          </td>
                          <td style={{ textAlign: 'right' }}>
                            <button
                              className="danger"
                              disabled={working === key}
                              onClick={() =>
                                removePermission(g.resourceAppId, ra.id)
                              }
                            >
                              {working === key ? '…' : 'Remove'}
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {showAdd && (
        <AddRequiredPermissionModal
          existing={groups}
          onClose={() => setShowAdd(false)}
          onSubmit={addEntry}
        />
      )}
      {detailsFor && (
        <ScopeDetailsModal
          scope={detailsFor.scope}
          kind={detailsFor.kind}
          resourceName={detailsFor.resourceName}
          catalog={catalog}
          fallbackTitle={detailsFor.title}
          fallbackDescription={detailsFor.description}
          onClose={() => setDetailsFor(null)}
        />
      )}
    </div>
  );
}

function mergeAccess(
  a: RequiredResourceAccess['resourceAccess'],
  b: RequiredResourceAccess['resourceAccess'],
): RequiredResourceAccess['resourceAccess'] {
  const seen = new Set(a.map((x) => `${x.type}:${x.id}`));
  const merged = [...a];
  for (const x of b) {
    const k = `${x.type}:${x.id}`;
    if (!seen.has(k)) {
      seen.add(k);
      merged.push(x);
    }
  }
  return merged;
}

function CredentialsCard({
  app,
  fic,
  activityByKeyId,
  activityError,
  onAppChange,
}: {
  app: Application;
  fic: FederatedIdentityCredential[] | null;
  activityByKeyId: Map<string, AppCredentialSignInActivity>;
  activityError: string | null;
  onAppChange: (next: Application) => void;
}) {
  const token = useGraphToken();
  const secrets: PasswordCredential[] = app.passwordCredentials ?? [];
  const certs: KeyCredential[] = app.keyCredentials ?? [];
  const [selectedSecret, setSelectedSecret] = useState<PasswordCredential | null>(
    null,
  );
  const [selectedCert, setSelectedCert] = useState<KeyCredential | null>(null);
  const [showAddSecret, setShowAddSecret] = useState(false);
  const [showAddCert, setShowAddCert] = useState(false);
  const [newSecret, setNewSecret] = useState<PasswordCredential | null>(null);
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set());
  const [bulkBusy, setBulkBusy] = useState<'secrets' | 'certs' | null>(null);
  const [bulkError, setBulkError] = useState<string | null>(null);

  const expiredSecrets = secrets.filter(
    (s) => expiryStatus(s.endDateTime) === 'expired',
  );
  const expiredCerts = certs.filter(
    (c) => expiryStatus(c.endDateTime) === 'expired',
  );
  const toggleCollapsed = (key: string) =>
    setCollapsed((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });

  const handleAddSecret = async (input: AddPasswordInput) => {
    const created = await addApplicationPassword(token, app.id, input);
    onAppChange({
      ...app,
      passwordCredentials: [...secrets, stripSecretText(created)],
    });
    setShowAddSecret(false);
    setNewSecret(created);
  };

  const handleRemoveSecret = async (keyId: string) => {
    await removeApplicationPassword(token, app.id, keyId);
    onAppChange({
      ...app,
      passwordCredentials: secrets.filter((s) => s.keyId !== keyId),
    });
    setSelectedSecret(null);
  };

  const handleAddCert = async (payload: CertSubmitPayload) => {
    // We intentionally do NOT forward startDateTime / endDateTime to Graph.
    // When `key` is a DER certificate Graph extracts validity from the cert
    // itself; passing our own values can produce "Key credential end date is
    // invalid" if there's any drift (milliseconds, timezone, etc).
    await addApplicationCertificate(token, app.id, certs, {
      displayName: payload.displayName,
      keyBase64: payload.keyBase64,
      thumbprintBase64: payload.thumbprintBase64,
    });
    // Graph doesn't return the new entry; fall back to refreshing the app so
    // the UI reflects the server-assigned keyId and parsed validity.
    const refreshed = await getApplication(token, app.id);
    onAppChange(refreshed);
  };

  const handleRemoveCert = async (keyId: string) => {
    await removeApplicationCertificate(token, app.id, certs, keyId);
    onAppChange({
      ...app,
      keyCredentials: certs.filter((c) => c.keyId !== keyId),
    });
    setSelectedCert(null);
  };

  const handleRemoveExpiredSecrets = async () => {
    if (expiredSecrets.length === 0) return;
    const n = expiredSecrets.length;
    if (
      !window.confirm(
        `Remove ${n} expired client secret${n === 1 ? '' : 's'}?`,
      )
    )
      return;
    setBulkBusy('secrets');
    setBulkError(null);
    try {
      // Remove one at a time so a mid-loop failure still leaves already-removed
      // secrets gone in Graph. removePassword is cheap.
      for (const s of expiredSecrets) {
        await removeApplicationPassword(token, app.id, s.keyId);
      }
      const removedIds = new Set(expiredSecrets.map((s) => s.keyId));
      onAppChange({
        ...app,
        passwordCredentials: secrets.filter((s) => !removedIds.has(s.keyId)),
      });
    } catch (e) {
      setBulkError(e instanceof Error ? e.message : String(e));
    } finally {
      setBulkBusy(null);
    }
  };

  const handleRemoveExpiredCerts = async () => {
    if (expiredCerts.length === 0) return;
    const n = expiredCerts.length;
    if (
      !window.confirm(
        `Remove ${n} expired certificate${n === 1 ? '' : 's'}?`,
      )
    )
      return;
    setBulkBusy('certs');
    setBulkError(null);
    try {
      const kept = certs.filter(
        (c) => expiryStatus(c.endDateTime) !== 'expired',
      );
      await updateApplication(token, app.id, { keyCredentials: kept });
      onAppChange({ ...app, keyCredentials: kept });
    } catch (e) {
      setBulkError(e instanceof Error ? e.message : String(e));
    } finally {
      setBulkBusy(null);
    }
  };

  return (
    <div className="card">
      <h3>Credentials</h3>
      {activityError && (
        <p className="muted" style={{ fontSize: 13, marginTop: 0 }}>
          Last-used info unavailable — grant{' '}
          <span className="mono">AuditLog.Read.All</span> to enable.
        </p>
      )}
      {bulkError && (
        <p className="error" style={{ marginTop: 0 }}>{bulkError}</p>
      )}

      <CredentialSection
        id="secrets"
        title="Client secrets"
        count={secrets.length}
        collapsed={collapsed.has('secrets')}
        onToggle={() => toggleCollapsed('secrets')}
        action={
          <div className="row" style={{ gap: 8 }}>
            <button
              className="danger"
              onClick={handleRemoveExpiredSecrets}
              disabled={
                expiredSecrets.length === 0 || bulkBusy === 'secrets'
              }
              title={
                expiredSecrets.length === 0
                  ? 'No expired client secrets'
                  : `Remove ${expiredSecrets.length} expired`
              }
            >
              {bulkBusy === 'secrets'
                ? 'Removing…'
                : `Remove expired${
                    expiredSecrets.length > 0
                      ? ` (${expiredSecrets.length})`
                      : ''
                  }`}
            </button>
            <button className="primary" onClick={() => setShowAddSecret(true)}>
              + New client secret
            </button>
          </div>
        }
        emptyLabel="No client secrets configured."
        isEmpty={secrets.length === 0}
      >
        <table className="table">
          <thead>
            <tr>
              <th>Description</th>
              <th>Secret ID</th>
              <th>Hint</th>
              <th>Expires</th>
              <th>Last used</th>
              <th>Resource</th>
            </tr>
          </thead>
          <tbody>
            {secrets.map((s) => {
              const activity = activityByKeyId.get(s.keyId);
              return (
                <tr
                  key={s.keyId}
                  onClick={() => setSelectedSecret(s)}
                  style={{ cursor: 'pointer' }}
                >
                  <td>{s.displayName || <span className="muted">—</span>}</td>
                  <td className="mono" title={s.keyId}>
                    {shortenKeyId(s.keyId)}
                  </td>
                  <td className="mono">
                    {s.hint ? `${s.hint}…` : <span className="muted">—</span>}
                  </td>
                  <td>
                    <div className="row" style={{ gap: 6 }}>
                      <span>{formatDate(s.endDateTime)}</span>
                      <ExpiryBadge end={s.endDateTime} />
                    </div>
                  </td>
                  <td>
                    <LastUsedCell activity={activity} />
                  </td>
                  <td>
                    <ResourceCell activity={activity} />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </CredentialSection>

      <CredentialSection
        id="certs"
        title="Certificates"
        count={certs.length}
        collapsed={collapsed.has('certs')}
        onToggle={() => toggleCollapsed('certs')}
        action={
          <div className="row" style={{ gap: 8 }}>
            <button
              className="danger"
              onClick={handleRemoveExpiredCerts}
              disabled={expiredCerts.length === 0 || bulkBusy === 'certs'}
              title={
                expiredCerts.length === 0
                  ? 'No expired certificates'
                  : `Remove ${expiredCerts.length} expired`
              }
            >
              {bulkBusy === 'certs'
                ? 'Removing…'
                : `Remove expired${
                    expiredCerts.length > 0
                      ? ` (${expiredCerts.length})`
                      : ''
                  }`}
            </button>
            <button className="primary" onClick={() => setShowAddCert(true)}>
              + New certificate
            </button>
          </div>
        }
        emptyLabel="No certificates configured."
        isEmpty={certs.length === 0}
      >
        <table className="table">
          <thead>
            <tr>
              <th>Name</th>
              <th>Thumbprint</th>
              <th>Type</th>
              <th>Usage</th>
              <th>Expires</th>
              <th>Last used</th>
              <th>Resource</th>
            </tr>
          </thead>
          <tbody>
            {certs.map((c) => {
              const activity = activityByKeyId.get(c.keyId);
              const thumb = formatThumbprint(c.customKeyIdentifier);
              return (
                <tr
                  key={c.keyId}
                  onClick={() => setSelectedCert(c)}
                  style={{ cursor: 'pointer' }}
                >
                  <td>{c.displayName || <span className="muted">—</span>}</td>
                  <td
                    className="mono"
                    title={thumb || c.keyId}
                    style={{ fontSize: 12 }}
                  >
                    {thumb ? shortenKeyId(thumb) : shortenKeyId(c.keyId)}
                  </td>
                  <td>{c.type ?? <span className="muted">—</span>}</td>
                  <td>{c.usage ?? <span className="muted">—</span>}</td>
                  <td>
                    <div className="row" style={{ gap: 6 }}>
                      <span>{formatDate(c.endDateTime)}</span>
                      <ExpiryBadge end={c.endDateTime} />
                    </div>
                  </td>
                  <td>
                    <LastUsedCell activity={activity} />
                  </td>
                  <td>
                    <ResourceCell activity={activity} />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </CredentialSection>

      <CredentialSection
        id="fic"
        title="Federated credentials"
        count={fic?.length ?? 0}
        collapsed={collapsed.has('fic')}
        onToggle={() => toggleCollapsed('fic')}
        isLoading={fic === null}
        emptyLabel="No federated credentials configured."
        isEmpty={fic !== null && fic.length === 0}
      >
        {fic && fic.length > 0 && (
          <table className="table">
            <thead>
              <tr>
                <th>Name</th>
                <th>Issuer</th>
                <th>Subject</th>
                <th>Audiences</th>
              </tr>
            </thead>
            <tbody>
              {fic.map((f) => (
                <tr key={f.id}>
                  <td>{f.name}</td>
                  <td className="mono" style={{ fontSize: 12 }}>
                    {f.issuer}
                  </td>
                  <td className="mono" style={{ fontSize: 12 }}>
                    {f.subject}
                  </td>
                  <td className="mono" style={{ fontSize: 12 }}>
                    {f.audiences.join(', ')}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </CredentialSection>

      {selectedSecret && (
        <SecretAuditModal
          secret={selectedSecret}
          activity={activityByKeyId.get(selectedSecret.keyId)}
          activityError={activityError}
          onClose={() => setSelectedSecret(null)}
          onDelete={() => handleRemoveSecret(selectedSecret.keyId)}
        />
      )}
      {selectedCert && (
        <CertificateAuditModal
          cert={selectedCert}
          activity={activityByKeyId.get(selectedCert.keyId)}
          activityError={activityError}
          onClose={() => setSelectedCert(null)}
          onDelete={() => handleRemoveCert(selectedCert.keyId)}
        />
      )}
      {showAddSecret && (
        <AddClientSecretModal
          onClose={() => setShowAddSecret(false)}
          onSubmit={handleAddSecret}
        />
      )}
      {showAddCert && (
        <AddCertificateModal
          onClose={() => setShowAddCert(false)}
          onSubmit={handleAddCert}
        />
      )}
      {newSecret && (
        <NewSecretValueModal
          secret={newSecret}
          onClose={() => setNewSecret(null)}
        />
      )}
    </div>
  );
}

function CredentialSection({
  id,
  title,
  count,
  collapsed,
  onToggle,
  action,
  children,
  isLoading,
  isEmpty,
  emptyLabel,
}: {
  id: string;
  title: string;
  count: number;
  collapsed: boolean;
  onToggle: () => void;
  action?: ReactNode;
  children?: ReactNode;
  isLoading?: boolean;
  isEmpty?: boolean;
  emptyLabel?: string;
}) {
  return (
    <div className={`group-box${collapsed ? ' collapsed' : ''}`}>
      <div className="group-header" onClick={onToggle}>
        <div className="row" style={{ gap: 10, alignItems: 'center' }}>
          <span className="chevron" aria-hidden>
            ▾
          </span>
          <div>
            <div className="group-title">{title}</div>
            <div className="group-subtitle">
              {count} {count === 1 ? 'entry' : 'entries'}
            </div>
          </div>
        </div>
        {action && (
          <div onClick={(e) => e.stopPropagation()}>{action}</div>
        )}
      </div>
      {!collapsed && (
        <div className="group-body" data-section={id}>
          {isLoading ? (
            <div className="muted" style={{ padding: 12 }}>
              <span className="spinner" /> Loading…
            </div>
          ) : isEmpty ? (
            <div className="muted" style={{ padding: 12 }}>
              {emptyLabel ?? 'Nothing here yet.'}
            </div>
          ) : (
            children
          )}
        </div>
      )}
    </div>
  );
}

function AuditRow({
  label,
  value,
  mono,
}: {
  label: string;
  value?: ReactNode;
  mono?: boolean;
}) {
  const empty = value === null || value === undefined || value === '';
  return (
    <>
      <div className="k">{label}</div>
      <div className={mono ? 'mono' : undefined}>
        {empty ? <span className="muted">—</span> : value}
      </div>
    </>
  );
}

function SecretAuditModal({
  secret,
  activity,
  activityError,
  onClose,
  onDelete,
}: {
  secret: PasswordCredential;
  activity?: AppCredentialSignInActivity;
  activityError: string | null;
  onClose: () => void;
  onDelete: () => Promise<void>;
}) {
  const [confirming, setConfirming] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [deleteError, setDeleteError] = useState<string | null>(null);

  const doDelete = async () => {
    setDeleting(true);
    setDeleteError(null);
    try {
      await onDelete();
    } catch (e: unknown) {
      setDeleteError(e instanceof Error ? e.message : String(e));
      setDeleting(false);
    }
  };

  return (
    <Modal
      title="Client secret details"
      onClose={() => !deleting && onClose()}
      footer={
        confirming ? (
          <>
            <button disabled={deleting} onClick={() => setConfirming(false)}>
              Cancel
            </button>
            <button className="danger" disabled={deleting} onClick={doDelete}>
              {deleting ? 'Deleting…' : 'Confirm delete'}
            </button>
          </>
        ) : (
          <>
            <button disabled={deleting} onClick={onClose}>
              Close
            </button>
            <button className="danger" onClick={() => setConfirming(true)}>
              Delete secret
            </button>
          </>
        )
      }
    >
      <div className="kv">
        <AuditRow label="Description" value={secret.displayName} />
        <AuditRow label="Secret ID (keyId)" value={secret.keyId} mono />
        <AuditRow label="Hint" value={secret.hint ? `${secret.hint}…` : null} mono />
        <AuditRow label="Created" value={formatDateTime(secret.startDateTime)} />
        <AuditRow
          label="Expires"
          value={
            <div className="row" style={{ gap: 6 }}>
              <span>{formatDateTime(secret.endDateTime)}</span>
              <ExpiryBadge end={secret.endDateTime} />
            </div>
          }
        />
      </div>
      {confirming && (
        <p className="error" style={{ marginTop: 12 }}>
          Deleting this secret invalidates it immediately. Anything still using
          it will fail to authenticate.
        </p>
      )}
      {deleteError && <p className="error">{deleteError}</p>}
      <SignInEventSection activity={activity} activityError={activityError} />
    </Modal>
  );
}

function CertificateAuditModal({
  cert,
  activity,
  activityError,
  onClose,
  onDelete,
}: {
  cert: KeyCredential;
  activity?: AppCredentialSignInActivity;
  activityError: string | null;
  onClose: () => void;
  onDelete: () => Promise<void>;
}) {
  const thumb = formatThumbprint(cert.customKeyIdentifier);
  const [confirming, setConfirming] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [deleteError, setDeleteError] = useState<string | null>(null);

  const doDelete = async () => {
    setDeleting(true);
    setDeleteError(null);
    try {
      await onDelete();
    } catch (e: unknown) {
      setDeleteError(e instanceof Error ? e.message : String(e));
      setDeleting(false);
    }
  };

  return (
    <Modal
      title="Certificate details"
      onClose={() => !deleting && onClose()}
      footer={
        confirming ? (
          <>
            <button disabled={deleting} onClick={() => setConfirming(false)}>
              Cancel
            </button>
            <button className="danger" disabled={deleting} onClick={doDelete}>
              {deleting ? 'Deleting…' : 'Confirm delete'}
            </button>
          </>
        ) : (
          <>
            <button disabled={deleting} onClick={onClose}>
              Close
            </button>
            <button className="danger" onClick={() => setConfirming(true)}>
              Delete certificate
            </button>
          </>
        )
      }
    >
      <div className="kv">
        <AuditRow label="Name" value={cert.displayName} />
        <AuditRow label="Key ID" value={cert.keyId} mono />
        <AuditRow label="Thumbprint" value={thumb} mono />
        <AuditRow label="Type" value={cert.type} />
        <AuditRow label="Usage" value={cert.usage} />
        <AuditRow label="Created" value={formatDateTime(cert.startDateTime)} />
        <AuditRow
          label="Expires"
          value={
            <div className="row" style={{ gap: 6 }}>
              <span>{formatDateTime(cert.endDateTime)}</span>
              <ExpiryBadge end={cert.endDateTime} />
            </div>
          }
        />
      </div>
      {confirming && (
        <p className="error" style={{ marginTop: 12 }}>
          Deleting this certificate invalidates it immediately. Anything using
          it will fail to authenticate.
        </p>
      )}
      {deleteError && <p className="error">{deleteError}</p>}
      <SignInEventSection activity={activity} activityError={activityError} />
    </Modal>
  );
}

function SignInEventSection({
  activity,
  activityError,
}: {
  activity?: AppCredentialSignInActivity;
  activityError: string | null;
}) {
  const token = useGraphToken();
  const si = activity?.signInActivity;
  const requestId = si?.requestId;
  const [event, setEvent] = useState<SignIn | null>(null);
  const [eventError, setEventError] = useState<string | null>(null);

  useEffect(() => {
    setEvent(null);
    setEventError(null);
    if (!requestId) return;
    let cancelled = false;
    getSignIn(token, requestId)
      .then((r) => {
        if (!cancelled) setEvent(r);
      })
      .catch((e: unknown) => {
        if (!cancelled)
          setEventError(e instanceof Error ? e.message : String(e));
      });
    return () => {
      cancelled = true;
    };
  }, [token, requestId]);

  if (activityError) {
    return (
      <>
        <h4 style={{ marginTop: 20, marginBottom: 8 }}>Sign-in activity</h4>
        <p className="muted" style={{ fontSize: 13 }}>
          Requires <span className="mono">AuditLog.Read.All</span>.
        </p>
      </>
    );
  }
  if (!activity) {
    return (
      <>
        <h4 style={{ marginTop: 20, marginBottom: 8 }}>Sign-in activity</h4>
        <div className="muted">No sign-in recorded for this credential.</div>
      </>
    );
  }

  const actorName =
    event?.userDisplayName ??
    event?.userPrincipalName ??
    event?.servicePrincipalName ??
    event?.appDisplayName;
  const actorSub =
    event?.userPrincipalName && event?.userDisplayName
      ? event.userPrincipalName
      : event?.servicePrincipalId;

  return (
    <>
      <h4 style={{ marginTop: 20, marginBottom: 8 }}>Last sign-in activity</h4>
      <div className="kv">
        <AuditRow
          label="When"
          value={
            si?.lastSignInDateTime
              ? formatDateTime(si.lastSignInDateTime)
              : 'Never'
          }
        />
        <AuditRow label="Key type" value={activity.keyType} />
        <AuditRow label="Key usage" value={activity.keyUsage} />
        {requestId ? (
          <>
            <AuditRow
              label="Who"
              value={
                event ? (
                  <>
                    <div>{actorName ?? '—'}</div>
                    {actorSub && (
                      <div className="muted mono" style={{ fontSize: 11 }}>
                        {actorSub}
                      </div>
                    )}
                  </>
                ) : eventError ? (
                  <span className="muted" style={{ fontSize: 12 }}>
                    (unavailable)
                  </span>
                ) : (
                  <span className="spinner" />
                )
              }
            />
            <AuditRow
              label="Flow"
              value={
                event ? (
                  <>
                    {isAppOnlyEvent(event) ? 'App-only' : 'Delegated'}
                    {event.authenticationProtocol && (
                      <div className="muted" style={{ fontSize: 11 }}>
                        {event.authenticationProtocol}
                      </div>
                    )}
                  </>
                ) : undefined
              }
            />
            <AuditRow label="From IP" value={event?.ipAddress} mono />
            <AuditRow label="Client" value={event?.clientAppUsed} />
            <AuditRow
              label="Resource"
              value={event?.resourceDisplayName ?? si?.resourceDisplayName}
            />
            <AuditRow
              label="Resource ID"
              value={event?.resourceId ?? si?.resourceId}
              mono
            />
            <AuditRow label="Request ID" value={requestId} mono />
            <AuditRow
              label="Status"
              value={
                !event ? undefined : event.status?.errorCode ? (
                  <>
                    <span className="badge expired">Failed</span>
                    <div className="muted" style={{ fontSize: 11 }}>
                      {event.status.failureReason ??
                        `code ${event.status.errorCode}`}
                    </div>
                  </>
                ) : (
                  <span className="badge granted">Success</span>
                )
              }
            />
          </>
        ) : (
          <>
            <AuditRow
              label="Target resource"
              value={si?.resourceDisplayName}
            />
            <AuditRow label="Resource ID" value={si?.resourceId} mono />
          </>
        )}
      </div>
    </>
  );
}

function isAppOnlyEvent(s: SignIn): boolean {
  return (s.signInEventTypes ?? []).some(
    (t) => t === 'servicePrincipal' || t === 'managedIdentity',
  );
}

function stripSecretText(c: PasswordCredential): PasswordCredential {
  const { secretText: _s, ...rest } = c;
  return rest as PasswordCredential;
}

const SECRET_DURATIONS: Array<{ label: string; months: number }> = [
  { label: '3 months', months: 3 },
  { label: '6 months', months: 6 },
  { label: '12 months', months: 12 },
  { label: '24 months (max)', months: 24 },
];

function AddClientSecretModal({
  onClose,
  onSubmit,
}: {
  onClose: () => void;
  onSubmit: (input: AddPasswordInput) => Promise<void>;
}) {
  const [displayName, setDisplayName] = useState('');
  const [months, setMonths] = useState(6);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const submit = async () => {
    setSubmitting(true);
    setError(null);
    try {
      const end = new Date();
      end.setMonth(end.getMonth() + months);
      await onSubmit({
        displayName: displayName.trim() || undefined,
        endDateTime: end.toISOString(),
      });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setSubmitting(false);
    }
  };

  return (
    <Modal
      title="New client secret"
      onClose={() => !submitting && onClose()}
      footer={
        <>
          <button disabled={submitting} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={submitting} onClick={submit}>
            {submitting ? 'Creating…' : 'Add'}
          </button>
        </>
      }
    >
      <label className="field">
        <span>Description (optional)</span>
        <input
          autoFocus
          value={displayName}
          onChange={(e) => setDisplayName(e.target.value)}
          placeholder="e.g. prod backend"
          maxLength={120}
        />
      </label>
      <label className="field">
        <span>Expires in</span>
        <select
          value={months}
          onChange={(e) => setMonths(Number(e.target.value))}
        >
          {SECRET_DURATIONS.map((d) => (
            <option key={d.months} value={d.months}>
              {d.label}
            </option>
          ))}
        </select>
      </label>
      <p className="muted" style={{ fontSize: 12 }}>
        Entra caps client secret lifetime at 24 months. Microsoft recommends
        certificates for production workloads.
      </p>
      {error && <p className="error">{error}</p>}
    </Modal>
  );
}

function NewSecretValueModal({
  secret,
  onClose,
}: {
  secret: PasswordCredential;
  onClose: () => void;
}) {
  const [copied, setCopied] = useState(false);
  const value = secret.secretText ?? '';
  const copy = async () => {
    try {
      await navigator.clipboard.writeText(value);
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    } catch {
      /* clipboard blocked — user can still select manually */
    }
  };
  return (
    <Modal
      title="Copy your client secret"
      onClose={onClose}
      footer={
        <button className="primary" onClick={onClose}>
          Done
        </button>
      }
    >
      <p className="error">
        This is the only time the secret value will be shown. Copy it now and
        store it somewhere safe.
      </p>
      <div className="kv">
        <AuditRow label="Description" value={secret.displayName} />
        <AuditRow label="Secret ID" value={secret.keyId} mono />
        <AuditRow label="Expires" value={formatDateTime(secret.endDateTime)} />
      </div>
      <div style={{ marginTop: 16 }}>
        <div className="muted" style={{ fontSize: 12, marginBottom: 4 }}>
          Secret value
        </div>
        <div className="row" style={{ gap: 6 }}>
          <input readOnly value={value} className="mono" style={{ flex: 1 }} />
          <button onClick={copy}>{copied ? 'Copied' : 'Copy'}</button>
        </div>
      </div>
    </Modal>
  );
}
