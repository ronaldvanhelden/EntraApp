import { useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { useCurrentTenantId } from '../auth/useCurrentTenantId';
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
  const currentTenantId = useCurrentTenantId();
  const nav = useNavigate();
  const [sps, setSps] = useState<ServicePrincipal[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState('');
  const [debouncedSearch, setDebouncedSearch] = useState('');
  const [filter, setFilter] = useState<Filter>('all');

  useEffect(() => {
    const handle = setTimeout(() => setDebouncedSearch(search.trim()), 250);
    return () => clearTimeout(handle);
  }, [search]);

  useEffect(() => {
    let cancelled = false;
    setSps(null);
    setError(null);
    listServicePrincipals(token, debouncedSearch || undefined)
      .then((list) => {
        if (cancelled) return;
        setSps(
          [...list].sort((a, b) =>
            (a.displayName ?? '').localeCompare(b.displayName ?? ''),
          ),
        );
      })
      .catch((e) => {
        if (!cancelled) setError(e.message);
      });
    return () => {
      cancelled = true;
    };
  }, [token, debouncedSearch]);

  const filtered = useMemo(() => {
    if (!sps) return [];
    if (filter === 'all') return sps;
    return sps.filter((sp) => classify(sp) === filter);
  }, [sps, filter]);

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
          placeholder="Search by name or appId…"
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
            {filter === 'all'
              ? `${sps.length} ${debouncedSearch ? 'matches' : 'total'}`
              : `${filtered.length} of ${sps.length}`}
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
                <th>Home tenant</th>
                <th>Enabled</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((sp) => {
                const external =
                  sp.appOwnerOrganizationId &&
                  currentTenantId &&
                  sp.appOwnerOrganizationId.toLowerCase() !==
                    currentTenantId.toLowerCase();
                return (
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
                      {external ? (
                        <span className="badge">External</span>
                      ) : sp.appOwnerOrganizationId ? (
                        <span className="badge granted">This tenant</span>
                      ) : (
                        <span className="muted">—</span>
                      )}
                    </td>
                    <td>
                      {sp.accountEnabled ? (
                        <span className="badge granted">Enabled</span>
                      ) : (
                        <span className="badge">Disabled</span>
                      )}
                    </td>
                  </tr>
                );
              })}
              {filtered.length === 0 && (
                <tr>
                  <td colSpan={5}>
                    <div className="empty">
                      {debouncedSearch
                        ? 'No enterprise apps match your search.'
                        : 'No enterprise apps found.'}
                    </div>
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
