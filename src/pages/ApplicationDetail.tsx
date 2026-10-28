import { useCallback, useEffect, useState, type ReactNode } from 'react';
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
import { AddCertificateModal, type CertSubmitPayload } from '../components/AddCertificateModal';
import {
  createAppRoleAssignment,
  createOAuth2Grant,
  createServicePrincipalFromAppId,
  deleteServicePrincipal,
  getServicePrincipalByAppId,
  listAppRoleAssignments,
  listOAuth2GrantsForClient,
  updateOAuth2Grant,
} from '../graph/servicePrincipals';
import {
  listMyActiveRoleTemplateIds,
  PRIVILEGED_CONSENT_RESOURCES,
  ROLE_APP_ADMIN,
  ROLE_CLOUD_APP_ADMIN,
  ROLE_GLOBAL_ADMIN,
  ROLE_PRIVILEGED_ROLE_ADMIN,
} from '../graph/directoryRoles';
import {
  buildCredentialActivityMap,
  listAppCredentialSignInActivities,
  listFederatedIdentityCredentials,
} from '../graph/credentials';
import type {
  AppCredentialSignInActivity,
  AppRoleAssignment,
  Application,
  FederatedIdentityCredential,
  KeyCredential,
  OAuth2PermissionGrant,
  PasswordCredential,
  RequiredResourceAccess,
  ServicePrincipal,
} from '../graph/types';
import { Modal } from '../components/Modal';
import { RemoveButton } from '../components/RemoveButton';
import { useConfirm } from '../components/ConfirmDialog';
import { AppIcon } from '../components/AppIcon';
import { CopyButton } from '../components/CopyButton';
import { TagEditor } from '../components/TagEditor';
import { FavoriteStar } from '../components/TagChip';
import { isFavorite, withFavorite } from '../lib/tags';
import { PrivilegeBadge, useScopeCatalog } from '../components/PrivilegeBadge';
import { ScopeDetailsModal } from '../components/ScopeDetailsModal';
import { AddRequiredPermissionModal } from '../components/AddRequiredPermissionModal';
import { getSignIn, listAppOnlySignInsForCredential } from '../graph/signIns';
import { GraphError } from '../graph/client';
import type { SignIn } from '../graph/types';

const AUDIENCES = [
  'AzureADMyOrg',
  'AzureADMultipleOrgs',
  'AzureADandPersonalMicrosoftAccount',
  'PersonalMicrosoftAccount',
];

type Tab = 'overview' | 'credentials' | 'authentication' | 'permissions';

export function ApplicationDetail() {
  const { id = '' } = useParams();
  const token = useGraphToken();
  const nav = useNavigate();
  const [app, setApp] = useState<Application | null>(null);
  const [sp, setSp] = useState<ServicePrincipal | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [editing, setEditing] = useState(false);
  const [tab, setTab] = useState<Tab>('overview');
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [fic, setFic] = useState<FederatedIdentityCredential[] | null>(null);
  const [activityByKeyId, setActivityByKeyId] = useState<Map<string, AppCredentialSignInActivity>>(
    new Map(),
  );
  const [activityError, setActivityError] = useState<string | null>(null);
  const [resourceSps, setResourceSps] = useState<Record<string, ServicePrincipal | null>>({});

  useEffect(() => {
    setApp(null);
    setSp(null);
    setError(null);
    setEditing(false);
    setTab('overview');
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
          .catch((e: unknown) => setActivityError(e instanceof Error ? e.message : String(e)));
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
        .then((resSp) => setResourceSps((prev) => ({ ...prev, [appId]: resSp ?? null })))
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
            <div className="row" style={{ gap: 6, alignItems: 'center' }}>
              <h1 style={{ margin: 0 }}>{app.displayName}</h1>
              <FavoriteToggle app={app} onAppChange={setApp} />
            </div>
            <div className="subtitle mono">
              {app.appId}
              <CopyButton value={app.appId} label="application (client) ID" />
            </div>
          </div>
        </div>
        <div className="row">
          {tab === 'overview' && !editing && <button onClick={() => setEditing(true)}>Edit</button>}
          {sp && (
            <Link to={`/enterprise-apps/${sp.id}`} style={{ textDecoration: 'none' }}>
              <button className="primary">Manage enterprise app →</button>
            </Link>
          )}
        </div>
      </div>

      <div className="tabs">
        <button className={tab === 'overview' ? 'active' : ''} onClick={() => setTab('overview')}>
          Overview
        </button>
        <button
          className={tab === 'permissions' ? 'active' : ''}
          onClick={() => setTab('permissions')}
        >
          API permissions
        </button>
        <button
          className={tab === 'credentials' ? 'active' : ''}
          onClick={() => setTab('credentials')}
        >
          Credentials
        </button>
        <button
          className={tab === 'authentication' ? 'active' : ''}
          onClick={() => setTab('authentication')}
        >
          Authentication
        </button>
      </div>

      {tab === 'overview' && editing && (
        <EditCard
          app={app}
          onCancel={() => setEditing(false)}
          onSaved={(updated) => {
            setApp({ ...app, ...updated });
            setEditing(false);
          }}
        />
      )}

      {tab === 'overview' && !editing && (
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
            <div className="k">Unique name</div>
            <div className="mono">
              {app.uniqueName ? (
                <>
                  {app.uniqueName}
                  <CopyButton value={app.uniqueName} label="unique name" />
                </>
              ) : (
                <span className="muted">—</span>
              )}
            </div>
            <div className="k">Tags</div>
            <div>
              <TagsRow app={app} onAppChange={setApp} />
            </div>
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
            <div className="k">Redirect URIs</div>
            <div>
              <RedirectUriSummary app={app} />
            </div>
            <div className="k">Notes</div>
            <div className="muted">{app.notes || '—'}</div>
            <div className="k">Created</div>
            <div>{app.createdDateTime ? new Date(app.createdDateTime).toLocaleString() : '—'}</div>
            <div className="k">OAuth flows</div>
            <div>
              <OAuthFlowSummary app={app} />
            </div>
          </div>
        </div>
      )}

      {tab === 'overview' && (
        <div className="card card-danger">
          <h3>Danger zone</h3>
          <div className="row" style={{ justifyContent: 'space-between' }}>
            <div className="muted" style={{ fontSize: 13 }}>
              Deleting the app registration removes the application object. The service principal is
              separate and must be deleted explicitly.
            </div>
            <RemoveButton label="Delete app registration" onClick={() => setConfirmDelete(true)} />
          </div>
        </div>
      )}

      {tab === 'permissions' && (
        <RequiredResourceAccessCard
          app={app}
          sp={sp}
          onSpChange={setSp}
          resourceSps={resourceSps}
          onChange={setApp}
        />
      )}

      {tab === 'credentials' && (
        <CredentialsCard
          app={app}
          fic={fic}
          activityByKeyId={activityByKeyId}
          activityError={activityError}
          onAppChange={setApp}
        />
      )}

      {tab === 'authentication' && <AuthenticationCard app={app} onAppChange={setApp} />}

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
  const [identifierUris, setIdentifierUris] = useState((app.identifierUris ?? []).join('\n'));
  // Graph accepts uniqueName only while it is null; it is immutable afterwards,
  // so the field is offered exactly once.
  const canSetUniqueName = !app.uniqueName;
  const [uniqueName, setUniqueName] = useState('');
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
      const trimmedUniqueName = uniqueName.trim();
      const patch = {
        displayName: displayName.trim(),
        signInAudience: audience,
        notes: notes.trim() || undefined,
        identifierUris: uris,
        ...(canSetUniqueName && trimmedUniqueName ? { uniqueName: trimmedUniqueName } : {}),
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
        <input value={displayName} onChange={(e) => setDisplayName(e.target.value)} />
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
      {canSetUniqueName ? (
        <label className="field">
          <span>Unique name (optional alternate key — can only be set once)</span>
          <input
            value={uniqueName}
            onChange={(e) => setUniqueName(e.target.value)}
            placeholder="e.g. contoso-payments-api"
          />
        </label>
      ) : (
        <div className="field">
          <span
            style={{
              display: 'block',
              fontSize: 12,
              color: 'var(--text-muted)',
              marginBottom: 4,
            }}
          >
            Unique name (immutable)
          </span>
          <div className="mono">{app.uniqueName}</div>
        </div>
      )}
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
        <textarea rows={3} value={notes} onChange={(e) => setNotes(e.target.value)} />
      </label>
      {error && <p className="error">{error}</p>}
      <div className="row" style={{ justifyContent: 'flex-end', gap: 8 }}>
        <button disabled={saving} onClick={onCancel}>
          Cancel
        </button>
        <button className="primary" disabled={saving || !displayName.trim()} onClick={save}>
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
            `Failed to delete service principal: ${e instanceof Error ? e.message : String(e)}`,
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
      {error && (
        <p className="error" style={{ marginTop: 12 }}>
          {error}
        </p>
      )}
    </Modal>
  );
}

function expiryStatus(end?: string): 'expired' | 'soon' | 'active' | null {
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
  if (status === 'expired') return <span className="badge expired">Expired</span>;
  if (status === 'soon') return <span className="badge pending">Expires soon</span>;
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

function LastUsedCell({ activity }: { activity?: AppCredentialSignInActivity }) {
  if (!activity) return <span className="muted">—</span>;
  const last = activity.signInActivity?.lastSignInDateTime;
  if (!last) return <span className="muted">Never</span>;
  return <span title={last}>{formatDateTime(last)}</span>;
}

function ResourceCell({ activity }: { activity?: AppCredentialSignInActivity }) {
  const name = activity?.signInActivity?.resourceDisplayName;
  return name ? <span>{name}</span> : <span className="muted">—</span>;
}

// Tags are stored on the directory object, so every write is a PATCH of the
// whole collection. Each caller gets its own instance (and its own error slot).
function useTagWriter(app: Application, onAppChange: (next: Application) => void) {
  const token = useGraphToken();
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const save = async (tags: string[]) => {
    setSaving(true);
    setError(null);
    try {
      await updateApplication(token, app.id, { tags });
      onAppChange({ ...app, tags });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      throw e;
    } finally {
      setSaving(false);
    }
  };

  return { saving, error, save };
}

function FavoriteToggle({
  app,
  onAppChange,
}: {
  app: Application;
  onAppChange: (next: Application) => void;
}) {
  const { saving, error, save } = useTagWriter(app, onAppChange);
  const favorite = isFavorite(app.tags);
  return (
    <>
      <FavoriteStar
        favorite={favorite}
        disabled={saving}
        onToggle={() => {
          void save(withFavorite(app.tags, !favorite)).catch(() => {});
        }}
      />
      {error && (
        <span className="error" style={{ fontSize: 12 }} title={error}>
          tag save failed
        </span>
      )}
    </>
  );
}

function TagsRow({
  app,
  onAppChange,
}: {
  app: Application;
  onAppChange: (next: Application) => void;
}) {
  const { saving, save } = useTagWriter(app, onAppChange);
  return <TagEditor tags={app.tags} onSave={save} disabled={saving} />;
}

// Every redirect URI on the app, tagged with the platform it is registered
// under. Read-only — editing lives in the Authentication tab.
function RedirectUriSummary({ app }: { app: Application }) {
  const entries: Array<{ platform: string; badge: string; uri: string }> = [
    ...(app.web?.redirectUris ?? []).map((uri) => ({
      platform: 'Web',
      badge: 'badge',
      uri,
    })),
    ...(app.spa?.redirectUris ?? []).map((uri) => ({
      platform: 'SPA',
      badge: 'badge delegated',
      uri,
    })),
    ...(app.publicClient?.redirectUris ?? []).map((uri) => ({
      platform: 'Public',
      badge: 'badge app',
      uri,
    })),
  ];

  if (entries.length === 0) return <span className="muted">—</span>;

  return (
    <div className="stack" style={{ gap: 4 }}>
      {entries.map((e) => (
        <div
          key={`${e.platform}:${e.uri}`}
          className="row"
          style={{ gap: 6, alignItems: 'flex-start' }}
        >
          <span className={e.badge} title={`Registered as ${e.platform}`}>
            {e.platform}
          </span>
          <span className="mono" style={{ fontSize: 12, minWidth: 0 }}>
            {e.uri}
          </span>
          <CopyButton value={e.uri} label="redirect URI" />
        </div>
      ))}
    </div>
  );
}

// Compact read-only summary of which OAuth 2.0 grant types the app is wired
// for — shown inline in the Details card. For toggling + full redirect URI
// editing the user switches to the dedicated Authentication tab.
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
    (app.passwordCredentials?.length ?? 0) > 0 || (app.keyCredentials?.length ?? 0) > 0;
  if (hasCred) {
    flows.push({
      label: 'Client credentials',
      title: 'App-only: uses a client secret, certificate, or FIC to obtain tokens without a user',
    });
  }

  if (flows.length === 0) {
    return (
      <span className="muted">
        No flows configured — add a redirect URI or credential in the Authentication or Credentials
        tab.
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
  const accessTokenIssuance = !!app.web?.implicitGrantSettings?.enableAccessTokenIssuance;
  const fallbackPublic = !!app.isFallbackPublicClient;

  const patch = async (body: UpdateApplicationPatch, local: Partial<Application>, key: string) => {
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

  const saveRedirectUris = async (platform: 'web' | 'spa' | 'publicClient', uris: string[]) => {
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
          initial={editPlatform === 'web' ? webUris : editPlatform === 'spa' ? spaUris : publicUris}
          onClose={() => setEditPlatform(null)}
          onSave={(uris) => saveRedirectUris(editPlatform, uris)}
        />
      )}
    </div>
  );
}

function FlowBadge({ active, label, hint }: { active: boolean; label: string; hint?: string }) {
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
        One URI per line. Must be an HTTPS URL (or <code>http://localhost</code>), except for
        public-client URIs which also allow custom schemes.
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

type ResourceAccessEntry = RequiredResourceAccess['resourceAccess'][number];

// Directory roles of the signed-in user, loaded once per page. `undefined`
// while loading, `null` when the lookup failed (we then let Graph decide).
function useMyActiveRoles() {
  const token = useGraphToken();
  const [roles, setRoles] = useState<Set<string> | null | undefined>(undefined);
  useEffect(() => {
    let cancelled = false;
    listMyActiveRoleTemplateIds(token)
      .then((r) => !cancelled && setRoles(r))
      .catch(() => !cancelled && setRoles(null));
    return () => {
      cancelled = true;
    };
  }, [token]);
  return roles;
}

// Why the signed-in user can't grant tenant-wide consent for `pending`, or
// null when they can. Global Admin / Privileged Role Admin can consent to
// anything; Cloud Application Admin / Application Admin to everything except
// Microsoft Graph (and Azure AD Graph) application permissions.
function consentBlocker(
  roles: Set<string> | null | undefined,
  pending: Array<{ resourceAppId: string; ra: ResourceAccessEntry }>,
): string | null {
  if (roles === undefined) return 'Checking your directory roles…';
  if (roles === null) return null;
  if (roles.has(ROLE_GLOBAL_ADMIN) || roles.has(ROLE_PRIVILEGED_ROLE_ADMIN)) return null;
  if (!roles.has(ROLE_CLOUD_APP_ADMIN) && !roles.has(ROLE_APP_ADMIN))
    return (
      'Requires an active Global Administrator, Privileged Role Administrator, ' +
      'Cloud Application Administrator or Application Administrator role. ' +
      'Activate one (e.g. in PIM) and reload the page.'
    );
  const needsPrivileged = pending.some(
    (p) => p.ra.type === 'Role' && PRIVILEGED_CONSENT_RESOURCES.has(p.resourceAppId),
  );
  if (needsPrivileged)
    return (
      'Microsoft Graph application permissions need an active Global ' +
      'Administrator or Privileged Role Administrator role.'
    );
  return null;
}

// Coloured glyph that tells the two permission kinds apart at a glance:
// a person for delegated (acts on behalf of a signed-in user), a server
// for application (app-only, no user present). Colours match .badge.app /
// .badge.delegated so the cue is consistent across the UI.
function PermissionKindIcon({ kind }: { kind: 'delegated' | 'application' }) {
  const label = kind === 'application' ? 'Application permission' : 'Delegated permission';
  return (
    <span className={`perm-kind-icon ${kind}`} title={label} aria-label={label} role="img">
      {kind === 'application' ? (
        <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden>
          <path d="M1.75 1A1.75 1.75 0 0 0 0 2.75v3.5C0 7.216.784 8 1.75 8h12.5A1.75 1.75 0 0 0 16 6.25v-3.5A1.75 1.75 0 0 0 14.25 1H1.75Zm0 1.5h12.5a.25.25 0 0 1 .25.25v3.5a.25.25 0 0 1-.25.25H1.75a.25.25 0 0 1-.25-.25v-3.5a.25.25 0 0 1 .25-.25ZM3 4.5a.75.75 0 1 0 1.5 0A.75.75 0 0 0 3 4.5ZM0 10.75C0 9.784.784 9 1.75 9h12.5c.966 0 1.75.784 1.75 1.75v3.5A1.75 1.75 0 0 1 14.25 16H1.75A1.75 1.75 0 0 1 0 14.25v-3.5Zm1.75-.25a.25.25 0 0 0-.25.25v3.5c0 .138.112.25.25.25h12.5a.25.25 0 0 0 .25-.25v-3.5a.25.25 0 0 0-.25-.25H1.75ZM3 12.5a.75.75 0 1 0 1.5 0 .75.75 0 0 0-1.5 0Z" />
        </svg>
      ) : (
        <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden>
          <path d="M10.561 8.073a6.005 6.005 0 0 1 3.432 5.142.75.75 0 1 1-1.498.07 4.5 4.5 0 0 0-8.99 0 .75.75 0 0 1-1.498-.07 6.004 6.004 0 0 1 3.431-5.142 3.999 3.999 0 1 1 5.123 0ZM10.5 5a2.5 2.5 0 1 0-5 0 2.5 2.5 0 0 0 5 0Z" />
        </svg>
      )}
    </span>
  );
}

function RequiredResourceAccessCard({
  app,
  sp,
  onSpChange,
  resourceSps,
  onChange,
}: {
  app: Application;
  sp: ServicePrincipal | null;
  onSpChange: (next: ServicePrincipal) => void;
  resourceSps: Record<string, ServicePrincipal | null>;
  onChange: (next: Application) => void;
}) {
  const token = useGraphToken();
  const [showAdd, setShowAdd] = useState(false);
  const [showConsent, setShowConsent] = useState(false);
  const [working, setWorking] = useState<string | null>(null);
  const [saveError, setSaveError] = useState<string | null>(null);
  const { confirm, confirmDialog } = useConfirm();
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set());
  // Tenant-wide grants held by this app's service principal. null = unknown
  // (still loading, or the lookup failed).
  const [grants, setGrants] = useState<OAuth2PermissionGrant[] | null>(null);
  const [assignments, setAssignments] = useState<AppRoleAssignment[] | null>(null);
  const roles = useMyActiveRoles();
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

  const loadGrantState = useCallback(
    async (clientSpId: string) => {
      const [g, a] = await Promise.all([
        listOAuth2GrantsForClient(token, clientSpId),
        listAppRoleAssignments(token, clientSpId),
      ]);
      setGrants(g);
      setAssignments(a);
      return { grants: g, assignments: a };
    },
    [token],
  );

  useEffect(() => {
    setGrants(null);
    setAssignments(null);
    if (!sp) {
      // No service principal in this tenant yet — nothing can be granted.
      setGrants([]);
      setAssignments([]);
      return;
    }
    loadGrantState(sp.id).catch(() => {
      /* status stays unknown */
    });
  }, [sp, loadGrantState]);

  const isGranted = (resourceAppId: string, ra: ResourceAccessEntry): boolean | undefined => {
    const resSp = resourceSps[resourceAppId];
    if (!resSp || grants === null || assignments === null) return undefined;
    if (ra.type === 'Role')
      return assignments.some((a) => a.appRoleId === ra.id && a.resourceId === resSp.id);
    const value = resSp.oauth2PermissionScopes?.find((s) => s.id === ra.id)?.value;
    if (!value) return undefined;
    return grants.some(
      (g) =>
        g.consentType === 'AllPrincipals' &&
        g.resourceId === resSp.id &&
        (g.scope ?? '').split(/\s+/).includes(value),
    );
  };

  const pending = groups.flatMap((g) =>
    g.resourceAccess
      .filter((ra) => isGranted(g.resourceAppId, ra) !== true)
      .map((ra) => ({ resourceAppId: g.resourceAppId, ra })),
  );
  const blocker = consentBlocker(roles, pending);

  const grantConsent = async () => {
    let clientSp = sp;
    if (!clientSp) {
      clientSp = await createServicePrincipalFromAppId(token, app.appId);
      onSpChange(clientSp);
    }
    // Re-read so we merge into the latest grants rather than a stale copy.
    const fresh = await loadGrantState(clientSp.id);
    for (const g of groups) {
      const resSp = resourceSps[g.resourceAppId];
      if (!resSp) continue;

      const scopes = g.resourceAccess
        .filter((ra) => ra.type === 'Scope')
        .map((ra) => resSp.oauth2PermissionScopes?.find((s) => s.id === ra.id)?.value)
        .filter((v): v is string => Boolean(v));
      if (scopes.length) {
        const existing = fresh.grants.find(
          (x) => x.consentType === 'AllPrincipals' && x.resourceId === resSp.id,
        );
        if (existing) {
          const current = (existing.scope ?? '').split(/\s+/).filter(Boolean);
          const merged = Array.from(new Set([...current, ...scopes]));
          if (merged.length !== current.length)
            await updateOAuth2Grant(token, existing.id, {
              scope: merged.join(' '),
            });
        } else {
          await createOAuth2Grant(token, {
            clientId: clientSp.id,
            consentType: 'AllPrincipals',
            principalId: null,
            resourceId: resSp.id,
            scope: scopes.join(' '),
          });
        }
      }

      for (const ra of g.resourceAccess) {
        if (ra.type !== 'Role') continue;
        const has = fresh.assignments.some(
          (a) => a.appRoleId === ra.id && a.resourceId === resSp.id,
        );
        if (!has)
          await createAppRoleAssignment(token, clientSp.id, {
            appRoleId: ra.id,
            principalId: clientSp.id,
            resourceId: resSp.id,
          });
      }
    }
    await loadGrantState(clientSp.id);
  };

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

  const removePermission = async (resourceAppId: string, id: string, label: string) => {
    const ok = await confirm({
      title: 'Remove permission?',
      message: `${label} will be removed from the manifest. Any consent already granted on the enterprise app is not revoked.`,
    });
    if (!ok) return;
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

  const removeResource = async (resourceAppId: string, name: string) => {
    const ok = await confirm({
      title: 'Remove resource?',
      message: `All permissions for ${name} will be removed from the manifest. Any consent already granted on the enterprise app is not revoked.`,
      confirmLabel: 'Remove resource',
    });
    if (!ok) return;
    const next = groups.filter((g) => g.resourceAppId !== resourceAppId);
    void commit(next, `res:${resourceAppId}`).catch(() => {});
  };

  const renderPermissionTable = (
    g: RequiredResourceAccess,
    resSp: ServicePrincipal | null | undefined,
    type: 'Scope' | 'Role',
  ) => {
    const entries = g.resourceAccess.filter((ra) => ra.type === type);
    if (entries.length === 0) return null;
    const resolving = resSp === undefined;
    const permKind = type === 'Role' ? 'application' : 'delegated';
    return (
      <div className={`perm-section ${permKind}`}>
        <div className="perm-section-header">
          <PermissionKindIcon kind={permKind} />
          <span className="perm-section-title">
            {type === 'Role' ? 'Application permissions' : 'Delegated permissions'}
          </span>
          <span className="perm-section-meta">
            {type === 'Role' ? 'roles' : 'scopes'} · {entries.length}
          </span>
        </div>
        <table className="table">
          <thead>
            <tr>
              <th>Permission</th>
              <th>Description</th>
              <th style={{ width: 120 }}>Admin consent</th>
              <th style={{ width: 100 }}></th>
            </tr>
          </thead>
          <tbody>
            {entries.map((ra) => {
              const role =
                type === 'Role' ? resSp?.appRoles?.find((x) => x.id === ra.id) : undefined;
              const scope =
                type === 'Scope'
                  ? resSp?.oauth2PermissionScopes?.find((x) => x.id === ra.id)
                  : undefined;
              const permValue = role?.value ?? scope?.value;
              const permTitle = role?.displayName ?? scope?.adminConsentDisplayName;
              const permDesc = role?.description ?? scope?.adminConsentDescription;
              const key = `perm:${g.resourceAppId}:${ra.id}`;
              const granted = isGranted(g.resourceAppId, ra);
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
                        <span className="mono" style={{ fontWeight: 600 }}>
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
                  <td className="muted" style={{ maxWidth: 380, fontSize: 12 }}>
                    {permDesc ?? ''}
                  </td>
                  <td>
                    {granted === undefined ? (
                      <span className="muted">—</span>
                    ) : granted ? (
                      <span className="badge granted">Granted</span>
                    ) : (
                      <span className="badge pending">Not granted</span>
                    )}
                  </td>
                  <td style={{ textAlign: 'right' }}>
                    <RemoveButton
                      label="Remove permission"
                      busy={working === key}
                      onClick={() =>
                        void removePermission(
                          g.resourceAppId,
                          ra.id,
                          permValue ?? permTitle ?? ra.id,
                        )
                      }
                    />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  };

  const consentDisabledReason =
    blocker ??
    (groups.length === 0
      ? 'No permissions to consent to'
      : pending.length === 0
        ? 'All permissions already have admin consent'
        : null);

  return (
    <div className="card">
      <div className="row" style={{ justifyContent: 'space-between', marginBottom: 12 }}>
        <div>
          <h3 style={{ margin: 0 }}>API permissions (manifest)</h3>
          <div className="muted" style={{ fontSize: 12, marginTop: 4 }}>
            Permissions declared in the application manifest. To take effect, they must be
            consented/granted on the enterprise app.
          </div>
        </div>
        <div className="row" style={{ gap: 8 }}>
          <button
            disabled={consentDisabledReason !== null}
            title={
              consentDisabledReason ??
              (roles === null
                ? "Couldn't verify your directory roles — Graph will reject the request if you lack them."
                : `Grant tenant-wide consent for ${pending.length} permission${pending.length === 1 ? '' : 's'}`)
            }
            onClick={() => setShowConsent(true)}
          >
            Admin consent
          </button>
          <button className="primary" onClick={() => setShowAdd(true)}>
            + Add
          </button>
        </div>
      </div>
      {saveError && (
        <p className="error" style={{ marginTop: 0 }}>
          {saveError}
        </p>
      )}
      {blocker && roles !== undefined && groups.length > 0 && pending.length > 0 && (
        <p className="muted" style={{ fontSize: 13, marginTop: 0 }}>
          Admin consent unavailable: {blocker}
        </p>
      )}

      {groups.length === 0 ? (
        <div className="muted">No required resource access declared.</div>
      ) : (
        <div>
          {groups.map((g) => {
            const resSp = resourceSps[g.resourceAppId];
            const resolving = resSp === undefined;
            const resName = resSp?.displayName ?? (resolving ? '…' : g.resourceAppId);
            const isCollapsed = collapsed.has(g.resourceAppId);
            const scopeCount = g.resourceAccess.filter((ra) => ra.type === 'Scope').length;
            const roleCount = g.resourceAccess.length - scopeCount;
            return (
              <div key={g.resourceAppId} className={`group-box${isCollapsed ? ' collapsed' : ''}`}>
                <div className="group-header" onClick={() => toggleCollapsed(g.resourceAppId)}>
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
                    <div className="group-heading">
                      <div className="group-title">{resName}</div>
                      <div className="group-subtitle">
                        {g.resourceAppId}
                        <CopyButton value={g.resourceAppId} label="resource appId" />
                        {' · '}
                        {scopeCount} scope{scopeCount === 1 ? '' : 's'}
                        {' · '}
                        {roleCount} role{roleCount === 1 ? '' : 's'}
                      </div>
                    </div>
                  </div>
                  <RemoveButton
                    label="Remove resource"
                    busy={working === `res:${g.resourceAppId}`}
                    onClick={(e) => {
                      e.stopPropagation();
                      void removeResource(g.resourceAppId, resName);
                    }}
                  />
                </div>
                {!isCollapsed && (
                  <div className="group-body">
                    {resSp === null && (
                      <div className="muted" style={{ fontSize: 12, padding: '8px 14px 0' }}>
                        This API has no service principal in this tenant, so its permissions can't
                        be consented to here.
                      </div>
                    )}
                    {renderPermissionTable(g, resSp, 'Scope')}
                    {renderPermissionTable(g, resSp, 'Role')}
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
      {showConsent && (
        <AdminConsentModal
          app={app}
          hasSp={Boolean(sp)}
          pending={pending}
          resourceSps={resourceSps}
          onClose={() => setShowConsent(false)}
          onConfirm={grantConsent}
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
      {confirmDialog}
    </div>
  );
}

function AdminConsentModal({
  app,
  hasSp,
  pending,
  resourceSps,
  onClose,
  onConfirm,
}: {
  app: Application;
  hasSp: boolean;
  pending: Array<{ resourceAppId: string; ra: ResourceAccessEntry }>;
  resourceSps: Record<string, ServicePrincipal | null>;
  onClose: () => void;
  onConfirm: () => Promise<void>;
}) {
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const byResource = new Map<string, ResourceAccessEntry[]>();
  for (const p of pending) {
    byResource.set(p.resourceAppId, [...(byResource.get(p.resourceAppId) ?? []), p.ra]);
  }
  const unavailable = Array.from(byResource.keys()).filter((id) => !resourceSps[id]);

  const submit = async () => {
    setBusy(true);
    setError(null);
    try {
      await onConfirm();
      onClose();
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setBusy(false);
    }
  };

  const label = (resSp: ServicePrincipal | null | undefined, ra: ResourceAccessEntry) =>
    (ra.type === 'Role'
      ? resSp?.appRoles?.find((x) => x.id === ra.id)?.value
      : resSp?.oauth2PermissionScopes?.find((x) => x.id === ra.id)?.value) ?? ra.id;

  return (
    <Modal
      title="Grant admin consent"
      onClose={() => !busy && onClose()}
      footer={
        <>
          <button disabled={busy} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={busy} onClick={submit}>
            {busy ? 'Granting…' : 'Grant consent'}
          </button>
        </>
      }
    >
      <p style={{ marginTop: 0 }}>
        Grant tenant-wide consent to <strong>{app.displayName}</strong> for the following
        permissions on behalf of all users in this tenant:
      </p>
      {!hasSp && (
        <p className="muted" style={{ fontSize: 13 }}>
          The app has no enterprise app (service principal) in this tenant yet — one will be created
          first.
        </p>
      )}
      {Array.from(byResource.entries()).map(([resourceAppId, entries]) => {
        const resSp = resourceSps[resourceAppId];
        return (
          <div key={resourceAppId} style={{ marginBottom: 10 }}>
            <div style={{ fontWeight: 600 }}>{resSp?.displayName ?? resourceAppId}</div>
            <div className="row wrap" style={{ gap: 6, marginTop: 4 }}>
              {entries.map((ra) => (
                <span
                  key={ra.id}
                  className={`badge ${ra.type === 'Role' ? 'app' : 'delegated'}`}
                  title={ra.type === 'Role' ? 'Application' : 'Delegated'}
                >
                  {label(resSp, ra)}
                </span>
              ))}
            </div>
          </div>
        );
      })}
      {unavailable.length > 0 && (
        <p className="muted" style={{ fontSize: 13 }}>
          Skipped — no service principal in this tenant for:{' '}
          <span className="mono">{unavailable.join(', ')}</span>
        </p>
      )}
      {error && <p className="error">{error}</p>}
    </Modal>
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
  const { confirm, confirmDialog } = useConfirm();
  const secrets: PasswordCredential[] = app.passwordCredentials ?? [];
  const certs: KeyCredential[] = app.keyCredentials ?? [];
  const [selectedSecret, setSelectedSecret] = useState<PasswordCredential | null>(null);
  const [selectedCert, setSelectedCert] = useState<KeyCredential | null>(null);
  const [showAddSecret, setShowAddSecret] = useState(false);
  const [showAddCert, setShowAddCert] = useState(false);
  const [newSecret, setNewSecret] = useState<PasswordCredential | null>(null);
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set());
  const [bulkBusy, setBulkBusy] = useState<'secrets' | 'certs' | null>(null);
  const [bulkError, setBulkError] = useState<string | null>(null);

  const expiredSecrets = secrets.filter((s) => expiryStatus(s.endDateTime) === 'expired');
  const expiredCerts = certs.filter((c) => expiryStatus(c.endDateTime) === 'expired');
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
    const ok = await confirm({
      title: `Remove ${n} expired client secret${n === 1 ? '' : 's'}?`,
      message:
        'Expired secrets no longer work, but removing them is permanent. Secrets that are still valid are not affected.',
      confirmLabel: 'Remove expired',
    });
    if (!ok) return;
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
    const ok = await confirm({
      title: `Remove ${n} expired certificate${n === 1 ? '' : 's'}?`,
      message:
        'Expired certificates no longer work, but removing them is permanent. Certificates that are still valid are not affected.',
      confirmLabel: 'Remove expired',
    });
    if (!ok) return;
    setBulkBusy('certs');
    setBulkError(null);
    try {
      const kept = certs.filter((c) => expiryStatus(c.endDateTime) !== 'expired');
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
          Last-used info unavailable — grant <span className="mono">AuditLog.Read.All</span> to
          enable.
        </p>
      )}
      {bulkError && (
        <p className="error" style={{ marginTop: 0 }}>
          {bulkError}
        </p>
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
              disabled={expiredSecrets.length === 0 || bulkBusy === 'secrets'}
              title={
                expiredSecrets.length === 0
                  ? 'No expired client secrets'
                  : `Remove ${expiredSecrets.length} expired`
              }
            >
              {bulkBusy === 'secrets'
                ? 'Removing…'
                : `Remove expired${expiredSecrets.length > 0 ? ` (${expiredSecrets.length})` : ''}`}
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
                : `Remove expired${expiredCerts.length > 0 ? ` (${expiredCerts.length})` : ''}`}
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
                <tr key={c.keyId} onClick={() => setSelectedCert(c)} style={{ cursor: 'pointer' }}>
                  <td>{c.displayName || <span className="muted">—</span>}</td>
                  <td className="mono" title={thumb || c.keyId} style={{ fontSize: 12 }}>
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
          appId={app.appId}
          secret={selectedSecret}
          activity={activityByKeyId.get(selectedSecret.keyId)}
          activityError={activityError}
          onClose={() => setSelectedSecret(null)}
          onDelete={() => handleRemoveSecret(selectedSecret.keyId)}
        />
      )}
      {selectedCert && (
        <CertificateAuditModal
          appId={app.appId}
          cert={selectedCert}
          activity={activityByKeyId.get(selectedCert.keyId)}
          activityError={activityError}
          onClose={() => setSelectedCert(null)}
          onDelete={() => handleRemoveCert(selectedCert.keyId)}
        />
      )}
      {showAddSecret && (
        <AddClientSecretModal onClose={() => setShowAddSecret(false)} onSubmit={handleAddSecret} />
      )}
      {showAddCert && (
        <AddCertificateModal onClose={() => setShowAddCert(false)} onSubmit={handleAddCert} />
      )}
      {newSecret && <NewSecretValueModal secret={newSecret} onClose={() => setNewSecret(null)} />}
      {confirmDialog}
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
        {action && <div onClick={(e) => e.stopPropagation()}>{action}</div>}
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

function AuditRow({ label, value, mono }: { label: string; value?: ReactNode; mono?: boolean }) {
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
  appId,
  secret,
  activity,
  activityError,
  onClose,
  onDelete,
}: {
  appId: string;
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
          Deleting this secret invalidates it immediately. Anything still using it will fail to
          authenticate.
        </p>
      )}
      {deleteError && <p className="error">{deleteError}</p>}
      <SignInEventSection activity={activity} activityError={activityError} />
      <CredentialSignInHistorySection appId={appId} keyId={secret.keyId} />
    </Modal>
  );
}

function CertificateAuditModal({
  appId,
  cert,
  activity,
  activityError,
  onClose,
  onDelete,
}: {
  appId: string;
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
          Deleting this certificate invalidates it immediately. Anything using it will fail to
          authenticate.
        </p>
      )}
      {deleteError && <p className="error">{deleteError}</p>}
      <SignInEventSection activity={activity} activityError={activityError} />
      <CredentialSignInHistorySection appId={appId} keyId={cert.keyId} />
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
        if (!cancelled) setEventError(e instanceof Error ? e.message : String(e));
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
          value={si?.lastSignInDateTime ? formatDateTime(si.lastSignInDateTime) : 'Never'}
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
            <AuditRow label="Resource ID" value={event?.resourceId ?? si?.resourceId} mono />
            <AuditRow label="Request ID" value={requestId} mono />
            <AuditRow
              label="Status"
              value={
                !event ? undefined : event.status?.errorCode ? (
                  <>
                    <span className="badge expired">Failed</span>
                    <div className="muted" style={{ fontSize: 11 }}>
                      {event.status.failureReason ?? `code ${event.status.errorCode}`}
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
            <AuditRow label="Target resource" value={si?.resourceDisplayName} />
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

function CredentialSignInHistorySection({ appId, keyId }: { appId: string; keyId: string }) {
  const token = useGraphToken();
  const [state, setState] = useState<
    | { kind: 'loading' }
    | { kind: 'loaded'; rows: SignIn[] }
    | { kind: 'forbidden' }
    | { kind: 'error'; message: string }
  >({ kind: 'loading' });

  useEffect(() => {
    let cancelled = false;
    setState({ kind: 'loading' });
    listAppOnlySignInsForCredential(token, appId, keyId)
      .then((rows) => {
        if (cancelled) return;
        rows.sort((a, b) => (b.createdDateTime ?? '').localeCompare(a.createdDateTime ?? ''));
        setState({ kind: 'loaded', rows });
      })
      .catch((e: unknown) => {
        if (cancelled) return;
        if (e instanceof GraphError && (e.status === 403 || e.status === 404)) {
          setState({ kind: 'forbidden' });
          return;
        }
        setState({
          kind: 'error',
          message: e instanceof Error ? e.message : String(e),
        });
      });
    return () => {
      cancelled = true;
    };
  }, [token, appId, keyId]);

  return (
    <>
      <h4 style={{ marginTop: 20, marginBottom: 4 }}>Sign-ins · last 30 days</h4>
      <div className="muted" style={{ fontSize: 12, marginBottom: 8 }}>
        App-only sign-ins that used this credential. Microsoft Entra retains 30 days of sign-in
        activity.
      </div>
      {state.kind === 'loading' && <span className="spinner" />}
      {state.kind === 'forbidden' && (
        <div className="muted" style={{ fontSize: 13 }}>
          Requires <span className="mono">AuditLog.Read.All</span>.
        </div>
      )}
      {state.kind === 'error' && (
        <p className="error" style={{ fontSize: 13 }}>
          {state.message}
        </p>
      )}
      {state.kind === 'loaded' && state.rows.length === 0 && (
        <div className="muted" style={{ fontSize: 13 }}>
          No sign-ins recorded in the last 30 days.
        </div>
      )}
      {state.kind === 'loaded' && state.rows.length > 0 && (
        <table className="table">
          <thead>
            <tr>
              <th>When</th>
              <th>Actor</th>
              <th>From IP</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            {state.rows.map((s) => {
              const ok = !s.status?.errorCode;
              const actor =
                s.servicePrincipalName ??
                s.appDisplayName ??
                s.userDisplayName ??
                s.userPrincipalName;
              return (
                <tr key={s.id}>
                  <td>{formatDateTime(s.createdDateTime)}</td>
                  <td>
                    <div>{actor ?? <span className="muted">—</span>}</div>
                    {s.servicePrincipalId && (
                      <div className="muted mono" style={{ fontSize: 11 }}>
                        {s.servicePrincipalId}
                      </div>
                    )}
                  </td>
                  <td className="mono">{s.ipAddress ?? <span className="muted">—</span>}</td>
                  <td>
                    {ok ? (
                      <span className="badge granted">Success</span>
                    ) : (
                      <>
                        <span className="badge expired">Failed</span>
                        {(s.status?.failureReason || s.status?.errorCode) && (
                          <div className="muted" style={{ fontSize: 11 }}>
                            {s.status?.failureReason ?? `code ${s.status?.errorCode}`}
                          </div>
                        )}
                      </>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}
    </>
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
        <select value={months} onChange={(e) => setMonths(Number(e.target.value))}>
          {SECRET_DURATIONS.map((d) => (
            <option key={d.months} value={d.months}>
              {d.label}
            </option>
          ))}
        </select>
      </label>
      <p className="muted" style={{ fontSize: 12 }}>
        Entra caps client secret lifetime at 24 months. Microsoft recommends certificates for
        production workloads.
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
        This is the only time the secret value will be shown. Copy it now and store it somewhere
        safe.
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
