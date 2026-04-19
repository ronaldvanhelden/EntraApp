import { useEffect, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import {
  deleteApplication,
  getApplication,
  updateApplication,
} from '../graph/applications';
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
  ServicePrincipal,
} from '../graph/types';
import { Modal } from '../components/Modal';

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

  useEffect(() => {
    setApp(null);
    setSp(null);
    setError(null);
    setEditing(false);
    setFic(null);
    setActivityByKeyId(new Map());
    setActivityError(null);
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
        <div>
          <h1>{app.displayName}</h1>
          <div className="subtitle mono">{app.appId}</div>
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
            <div className="mono">{app.id}</div>
            <div className="k">Application (client) ID</div>
            <div className="mono">{app.appId}</div>
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
          </div>
        </div>
      )}

      <CredentialsCard
        app={app}
        fic={fic}
        activityByKeyId={activityByKeyId}
        activityError={activityError}
      />

      <div className="card">
        <h3>Required resource access (manifest)</h3>
        {!app.requiredResourceAccess?.length ? (
          <div className="muted">No required resource access declared.</div>
        ) : (
          <table className="table">
            <thead>
              <tr>
                <th>Resource AppId</th>
                <th>Permissions declared</th>
              </tr>
            </thead>
            <tbody>
              {app.requiredResourceAccess.map((r) => (
                <tr key={r.resourceAppId}>
                  <td className="mono">{r.resourceAppId}</td>
                  <td>{r.resourceAccess.length}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
        <p className="muted" style={{ marginTop: 12, marginBottom: 0 }}>
          To grant permissions, manage the enterprise app (service principal)
          that represents this application in your tenant.
        </p>
      </div>

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

function CredentialsCard({
  app,
  fic,
  activityByKeyId,
  activityError,
}: {
  app: Application;
  fic: FederatedIdentityCredential[] | null;
  activityByKeyId: Map<string, AppCredentialSignInActivity>;
  activityError: string | null;
}) {
  const secrets: PasswordCredential[] = app.passwordCredentials ?? [];
  const certs: KeyCredential[] = app.keyCredentials ?? [];

  return (
    <div className="card">
      <h3>Client secrets &amp; certificates</h3>
      {activityError && (
        <p className="muted" style={{ fontSize: 13, marginTop: 0 }}>
          Last-used info unavailable — grant{' '}
          <span className="mono">AuditLog.Read.All</span> to enable.
        </p>
      )}

      <h4 style={{ marginTop: 8, marginBottom: 8 }}>
        Client secrets ({secrets.length})
      </h4>
      {secrets.length === 0 ? (
        <div className="muted">No client secrets configured.</div>
      ) : (
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
                <tr key={s.keyId}>
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
      )}

      <h4 style={{ marginTop: 20, marginBottom: 8 }}>
        Certificates ({certs.length})
      </h4>
      {certs.length === 0 ? (
        <div className="muted">No certificates configured.</div>
      ) : (
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
                <tr key={c.keyId}>
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
      )}

      <h4 style={{ marginTop: 20, marginBottom: 8 }}>
        Federated credentials ({fic?.length ?? 0})
      </h4>
      {fic === null ? (
        <div className="muted">
          <span className="spinner" /> Loading…
        </div>
      ) : fic.length === 0 ? (
        <div className="muted">No federated credentials configured.</div>
      ) : (
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
    </div>
  );
}
