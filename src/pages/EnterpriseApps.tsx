import { useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { useCurrentTenantId } from '../auth/useCurrentTenantId';
import { listServicePrincipals } from '../graph/servicePrincipals';
import type { ServicePrincipal } from '../graph/types';
import { AppIcon } from '../components/AppIcon';

type Filter = 'all' | 'enterprise' | 'managed' | 'legacy';
type TenantFilter = 'all' | 'this' | 'external';

function classify(sp: ServicePrincipal): Exclude<Filter, 'all'> {
  if (sp.servicePrincipalType === 'ManagedIdentity') return 'managed';
  if (sp.servicePrincipalType === 'Legacy') return 'legacy';
  return 'enterprise';
}

function isExternal(
  sp: ServicePrincipal,
  currentTenantId: string | undefined,
): boolean {
  return Boolean(
    sp.appOwnerOrganizationId &&
      currentTenantId &&
      sp.appOwnerOrganizationId.toLowerCase() !== currentTenantId.toLowerCase(),
  );
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
  const [tenantFilter, setTenantFilter] = useState<TenantFilter>('all');
  // When 'external' is chosen the user can further narrow to a single owning
  // tenant. Empty string means "any external tenant".
  const [externalTenantId, setExternalTenantId] = useState('');

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
    return sps.filter((sp) => {
      if (filter !== 'all' && classify(sp) !== filter) return false;
      if (tenantFilter === 'this' && isExternal(sp, currentTenantId))
        return false;
      if (tenantFilter === 'external') {
        if (!isExternal(sp, currentTenantId)) return false;
        if (
          externalTenantId &&
          sp.appOwnerOrganizationId?.toLowerCase() !==
            externalTenantId.toLowerCase()
        )
          return false;
      }
      return true;
    });
  }, [sps, filter, tenantFilter, externalTenantId, currentTenantId]);

  // Distinct external owning tenants seen in the current list — drives the
  // secondary dropdown when tenantFilter === 'external'.
  const externalTenants = useMemo(() => {
    if (!sps) return [] as string[];
    const set = new Set<string>();
    for (const sp of sps) {
      if (isExternal(sp, currentTenantId) && sp.appOwnerOrganizationId) {
        set.add(sp.appOwnerOrganizationId);
      }
    }
    return [...set].sort();
  }, [sps, currentTenantId]);

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
          <button
            className={filter === 'legacy' ? 'active' : ''}
            onClick={() => setFilter('legacy')}
            title="Legacy service principals — created before the app registration model, no backing application"
          >
            Legacy
          </button>
        </div>
        <div className="tabs" style={{ margin: 0, borderBottom: 'none' }}>
          <button
            className={tenantFilter === 'all' ? 'active' : ''}
            onClick={() => setTenantFilter('all')}
            title="Show apps from any home tenant"
          >
            Any tenant
          </button>
          <button
            className={tenantFilter === 'this' ? 'active' : ''}
            onClick={() => setTenantFilter('this')}
            title="Only apps whose home tenant is this tenant"
          >
            This tenant
          </button>
          <button
            className={tenantFilter === 'external' ? 'active' : ''}
            onClick={() => setTenantFilter('external')}
            title="Only apps from external tenants (multi-tenant apps, B2B)"
          >
            External
          </button>
        </div>
        {tenantFilter === 'external' && externalTenants.length > 0 && (
          <select
            value={externalTenantId}
            onChange={(e) => setExternalTenantId(e.target.value)}
            style={{ maxWidth: 280 }}
            title="Narrow to a specific external tenant"
          >
            <option value="">
              Any external tenant ({externalTenants.length})
            </option>
            {externalTenants.map((tid) => (
              <option key={tid} value={tid}>
                {tid}
              </option>
            ))}
          </select>
        )}
        <div className="grow" />
        {sps && (
          <span className="muted">
            {filter === 'all' && tenantFilter === 'all'
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
                    <td>
                      <div className="row" style={{ gap: 10, alignItems: 'center' }}>
                        <AppIcon
                          id={sp.appId || sp.id}
                          logoUrl={sp.info?.logoUrl}
                          title={sp.displayName}
                        />
                        <span>{sp.displayName}</span>
                      </div>
                    </td>
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
