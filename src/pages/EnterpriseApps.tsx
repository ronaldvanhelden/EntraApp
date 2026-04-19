import { useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { listServicePrincipals } from '../graph/servicePrincipals';
import type { ServicePrincipal } from '../graph/types';

type Filter = 'all' | 'enterprise' | 'managed';

function classify(sp: ServicePrincipal): Filter {
  if (sp.tags?.includes('WindowsAzureActiveDirectoryIntegratedApp')) {
    return 'enterprise';
  }
  if (sp.servicePrincipalType === 'ManagedIdentity') return 'managed';
  return 'enterprise';
}

export function EnterpriseApps() {
  const token = useGraphToken();
  const nav = useNavigate();
  const [sps, setSps] = useState<ServicePrincipal[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState('');
  const [filter, setFilter] = useState<Filter>('all');

  useEffect(() => {
    setSps(null);
    setError(null);
    listServicePrincipals(token)
      .then((list) =>
        setSps(
          [...list].sort((a, b) =>
            (a.displayName ?? '').localeCompare(b.displayName ?? ''),
          ),
        ),
      )
      .catch((e) => setError(e.message));
  }, [token]);

  const filtered = useMemo(() => {
    if (!sps) return [];
    const q = search.trim().toLowerCase();
    return sps.filter((sp) => {
      if (filter !== 'all' && classify(sp) !== filter) return false;
      if (!q) return true;
      return (
        sp.displayName?.toLowerCase().includes(q) ||
        sp.appId?.toLowerCase().includes(q)
      );
    });
  }, [sps, search, filter]);

  return (
    <>
      <div className="page-header">
        <div>
          <h1>Enterprise apps</h1>
          <div className="subtitle">
            Service principals (tenant-local application instances)
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
        <div className="tabs" style={{ margin: 0, borderBottom: 'none' }}>
          <button
            className={filter === 'all' ? 'active' : ''}
            onClick={() => setFilter('all')}
          >
            All
          </button>
          <button
            className={filter === 'enterprise' ? 'active' : ''}
            onClick={() => setFilter('enterprise')}
          >
            Enterprise
          </button>
          <button
            className={filter === 'managed' ? 'active' : ''}
            onClick={() => setFilter('managed')}
          >
            Managed identities
          </button>
        </div>
        <div className="grow" />
        {sps && (
          <span className="muted">
            {filtered.length} of {sps.length}
          </span>
        )}
      </div>

      {error && <div className="card error">{error}</div>}
      {!sps && !error && <span className="spinner" />}

      {sps && (
        <div className="card" style={{ padding: 0 }}>
          <table className="table">
            <thead>
              <tr>
                <th>Display name</th>
                <th>Application ID</th>
                <th>Type</th>
                <th>Enabled</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((sp) => (
                <tr
                  key={sp.id}
                  onClick={() => nav(`/enterprise-apps/${sp.id}`)}
                >
                  <td>{sp.displayName}</td>
                  <td className="mono">{sp.appId}</td>
                  <td className="muted">
                    {sp.servicePrincipalType ?? 'Application'}
                  </td>
                  <td>
                    {sp.accountEnabled ? (
                      <span className="badge granted">Enabled</span>
                    ) : (
                      <span className="badge">Disabled</span>
                    )}
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr>
                  <td colSpan={4}>
                    <div className="empty">No enterprise apps match.</div>
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
