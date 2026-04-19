import { useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { listApplications } from '../graph/applications';
import type { Application } from '../graph/types';

export function Applications() {
  const token = useGraphToken();
  const nav = useNavigate();
  const [apps, setApps] = useState<Application[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState('');

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
    </>
  );
}
