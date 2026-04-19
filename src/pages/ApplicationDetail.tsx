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
import type { Application, ServicePrincipal } from '../graph/types';
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

  useEffect(() => {
    setApp(null);
    setSp(null);
    setError(null);
    setEditing(false);
    getApplication(token, id)
      .then(async (a) => {
        setApp(a);
        try {
          const maybe = await getServicePrincipalByAppId(token, a.appId);
          setSp(maybe ?? null);
        } catch {
          /* no SP yet — fine */
        }
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
