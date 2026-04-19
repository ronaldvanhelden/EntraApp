import { useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { createApplication, listApplications } from '../graph/applications';
import { createServicePrincipalFromAppId } from '../graph/servicePrincipals';
import type { Application } from '../graph/types';
import { Modal } from '../components/Modal';

const AUDIENCES = [
  'AzureADMyOrg',
  'AzureADMultipleOrgs',
  'AzureADandPersonalMicrosoftAccount',
  'PersonalMicrosoftAccount',
];

export function Applications() {
  const token = useGraphToken();
  const nav = useNavigate();
  const [apps, setApps] = useState<Application[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState('');
  const [showCreate, setShowCreate] = useState(false);

  useEffect(() => {
    setApps(null);
    setError(null);
    listApplications(token)
      .then(setApps)
      .catch((e) => setError(e.message));
  }, [token]);

  const filtered = useMemo(() => {
    if (!apps) return [];
    const q = search.trim().toLowerCase();
    if (!q) return apps;
    return apps.filter(
      (a) =>
        a.displayName?.toLowerCase().includes(q) ||
        a.appId?.toLowerCase().includes(q),
    );
  }, [apps, search]);

  return (
    <>
      <div className="page-header">
        <div>
          <h1>App registrations</h1>
          <div className="subtitle">
            Identity objects representing applications in your tenant
          </div>
        </div>
        <button className="primary" onClick={() => setShowCreate(true)}>
          + New app registration
        </button>
      </div>

      <div className="toolbar">
        <input
          className="search"
          placeholder="Filter by name or appId…"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
        />
        {apps && (
          <span className="muted">
            {filtered.length} of {apps.length}
          </span>
        )}
      </div>

      {error && <div className="card error">{error}</div>}
      {!apps && !error && <span className="spinner" />}

      {apps && (
        <div className="card" style={{ padding: 0 }}>
          <table className="table">
            <thead>
              <tr>
                <th>Display name</th>
                <th>Application (client) ID</th>
                <th>Sign-in audience</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((a) => (
                <tr key={a.id} onClick={() => nav(`/applications/${a.id}`)}>
                  <td>{a.displayName}</td>
                  <td className="mono">{a.appId}</td>
                  <td className="muted">{a.signInAudience}</td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr>
                  <td colSpan={3}>
                    <div className="empty">No applications match your filter.</div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {showCreate && (
        <CreateAppModal
          onClose={() => setShowCreate(false)}
          onCreated={(appId) => nav(`/applications/${appId}`)}
        />
      )}
    </>
  );
}

function CreateAppModal({
  onClose,
  onCreated,
}: {
  onClose: () => void;
  onCreated: (objectId: string) => void;
}) {
  const token = useGraphToken();
  const [displayName, setDisplayName] = useState('');
  const [audience, setAudience] = useState(AUDIENCES[0]);
  const [createSp, setCreateSp] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const submit = async () => {
    if (!displayName.trim()) return;
    setSubmitting(true);
    setError(null);
    try {
      const app = await createApplication(token, {
        displayName: displayName.trim(),
        signInAudience: audience,
      });
      if (createSp) {
        try {
          await createServicePrincipalFromAppId(token, app.appId);
        } catch (e: unknown) {
          // Non-fatal: app registration was created; surface but continue.
          setError(
            `App registered, but creating the enterprise app failed: ${
              e instanceof Error ? e.message : String(e)
            }`,
          );
        }
      }
      onCreated(app.id);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <Modal
      title="New app registration"
      onClose={() => !submitting && onClose()}
      footer={
        <>
          <button disabled={submitting} onClick={onClose}>
            Cancel
          </button>
          <button
            className="primary"
            disabled={submitting || !displayName.trim()}
            onClick={submit}
          >
            {submitting ? 'Creating…' : 'Create'}
          </button>
        </>
      }
    >
      <label className="field">
        <span>Display name</span>
        <input
          autoFocus
          value={displayName}
          onChange={(e) => setDisplayName(e.target.value)}
          placeholder="My new app"
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

      <label className="row" style={{ gap: 6 }}>
        <input
          type="checkbox"
          style={{ width: 'auto' }}
          checked={createSp}
          onChange={(e) => setCreateSp(e.target.checked)}
        />
        Also create the enterprise app (service principal) in this tenant
      </label>

      {error && <p className="error" style={{ marginTop: 12 }}>{error}</p>}
    </Modal>
  );
}
