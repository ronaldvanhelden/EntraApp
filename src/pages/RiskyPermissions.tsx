import { Fragment, useEffect, useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { useCurrentTenantId } from '../auth/useCurrentTenantId';
import {
  scanAppPermissions,
  type AppPermissionScan,
  type HeldPermission,
  type PermissionHolder,
} from '../graph/appPermissionScan';
import { listLastSignInByAppId } from '../graph/overview';
import { MS_GRAPH, isMicrosoftTenant } from '../lib/dangerousPermissions';
import { AppIcon } from '../components/AppIcon';
import { relativeTime } from '../components/ActivityTab';
import { SortTh, useSort } from '../components/SortableTable';

// Which permissions count: only the ones that lead to Global Administrator,
// those plus tenant-wide data and write access, or every app-only permission
// on the scanned APIs.
type Level = 'critical' | 'risky' | 'all';

const LEVELS: { id: Level; label: string; title: string }[] = [
  {
    id: 'critical',
    label: 'Critical',
    title: 'Permissions that lead to Global Administrator or to signing in as any user',
  },
  {
    id: 'risky',
    label: 'Critical and high',
    title: 'Also tenant-wide access to mail, files and chats, and write access to the directory',
  },
  { id: 'all', label: 'All app-only', title: 'Every application permission on the scanned APIs' },
];

const tierRank = (p: HeldPermission) => (p.risk?.tier === 'critical' ? 0 : p.risk ? 1 : 2);

function shown(p: HeldPermission, level: Level): boolean {
  if (level === 'all') return true;
  if (level === 'risky') return p.risk !== null;
  return p.risk?.tier === 'critical';
}

interface Row {
  holder: PermissionHolder;
  permissions: HeldPermission[];
  critical: number;
  high: number;
  lastSignIn: number | undefined;
}

export function RiskyPermissions() {
  const token = useGraphToken();
  const currentTenantId = useCurrentTenantId();
  const [scan, setScan] = useState<AppPermissionScan | null>(null);
  const [error, setError] = useState<string | null>(null);
  // null while loading or when the report isn't available.
  const [lastSignIns, setLastSignIns] = useState<Map<string, number> | null>(null);
  const [level, setLevel] = useState<Level>('risky');
  const [search, setSearch] = useState('');
  const [hideMicrosoft, setHideMicrosoft] = useState(false);
  const [open, setOpen] = useState<Set<string>>(new Set());

  useEffect(() => {
    let cancelled = false;
    setScan(null);
    setError(null);
    setLastSignIns(null);
    scanAppPermissions(token)
      .then((s) => !cancelled && setScan(s))
      .catch((e: unknown) => !cancelled && setError(e instanceof Error ? e.message : String(e)));
    listLastSignInByAppId(token)
      .then((m) => !cancelled && setLastSignIns(m))
      .catch(() => {
        // The sign-in report is a beta extra; the column stays empty.
      });
    return () => {
      cancelled = true;
    };
  }, [token]);

  const rows = useMemo(() => {
    if (!scan) return [] as Row[];
    const q = search.trim().toLowerCase();
    const out: Row[] = [];
    for (const holder of scan.holders) {
      if (hideMicrosoft && isMicrosoftTenant(holder.sp?.appOwnerOrganizationId)) continue;
      const permissions = holder.permissions
        .filter((p) => shown(p, level))
        .sort((a, b) => tierRank(a) - tierRank(b) || a.value.localeCompare(b.value));
      if (!permissions.length) continue;
      if (
        q &&
        !holder.displayName.toLowerCase().includes(q) &&
        !(holder.sp?.appId ?? '').toLowerCase().includes(q) &&
        !holder.id.toLowerCase().includes(q) &&
        !permissions.some((p) => p.value.toLowerCase().includes(q))
      )
        continue;
      out.push({
        holder,
        permissions,
        critical: permissions.filter((p) => p.risk?.tier === 'critical').length,
        high: permissions.filter((p) => p.risk?.tier === 'high').length,
        lastSignIn: holder.sp?.appId ? lastSignIns?.get(holder.sp.appId) : undefined,
      });
    }
    // Worst first; a column header overrides this.
    return out.sort(
      (a, b) =>
        b.critical - a.critical ||
        b.high - a.high ||
        a.holder.displayName.localeCompare(b.holder.displayName),
    );
  }, [scan, level, search, hideMicrosoft, lastSignIns]);

  const { sorted, sort } = useSort(rows, {
    name: (r) => r.holder.displayName,
    type: (r) => r.holder.sp?.servicePrincipalType ?? 'Application',
    tenant: (r) => homeTenant(r.holder, currentTenantId),
    enabled: (r) => (r.holder.sp ? Boolean(r.holder.sp.accountEnabled) : null),
    permissions: (r) => r.critical * 1000 + r.high + r.permissions.length / 1000,
    'last-sign-in': (r) => r.lastSignIn,
  });

  const totals = useMemo(() => {
    let critical = 0;
    let high = 0;
    for (const h of scan?.holders ?? []) {
      if (h.permissions.some((p) => p.risk?.tier === 'critical')) critical++;
      else if (h.permissions.some((p) => p.risk)) high++;
    }
    return { critical, high };
  }, [scan]);

  const toggle = (id: string) =>
    setOpen((prev) => {
      const next = new Set(prev);
      if (!next.delete(id)) next.add(id);
      return next;
    });

  const failed = scan?.resources.filter((r) => r.state === 'failed') ?? [];
  const scanned = scan?.resources.filter((r) => r.state === 'scanned') ?? [];

  return (
    <>
      <div className="page-header">
        <div>
          <h1>Risky permissions</h1>
          <div className="subtitle">
            Service principals holding dangerous application (app-only) permissions
          </div>
        </div>
      </div>

      {error && <div className="card error">{error}</div>}
      {!scan && !error && <span className="spinner" />}

      {scan && (
        <>
          <div className="stat-grid">
            <div className="stat-tile risk-tile critical">
              <div className="stat-value">{totals.critical}</div>
              <div className="stat-label">
                hold a critical permission: a route to Global Administrator
              </div>
            </div>
            <div className="stat-tile risk-tile high">
              <div className="stat-value">{totals.high}</div>
              <div className="stat-label">
                hold a high one: tenant-wide data, or write access to the directory
              </div>
            </div>
            <div className="stat-tile risk-tile">
              <div className="stat-value">{scan.holders.length}</div>
              <div className="stat-label">
                hold any app-only permission on {scanned.map((r) => r.name).join(', ') || 'no API'}
              </div>
            </div>
          </div>

          {failed.map((r) => (
            <div key={r.appId} className="card error">
              {r.name} could not be scanned, so its permissions are missing below: {r.error}
            </div>
          ))}

          <div className="toolbar">
            <input
              className="search"
              type="search"
              placeholder="Filter by app or permission…"
              aria-label="Filter by app or permission"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
            <div className="tabs" style={{ margin: 0, borderBottom: 'none' }}>
              {LEVELS.map((l) => (
                <button
                  key={l.id}
                  className={level === l.id ? 'active' : ''}
                  title={l.title}
                  onClick={() => setLevel(l.id)}
                >
                  {l.label}
                </button>
              ))}
            </div>
            <label className="row" style={{ gap: 8, alignItems: 'center', whiteSpace: 'nowrap' }}>
              <input
                type="checkbox"
                style={{ width: 'auto', flexShrink: 0, margin: 0 }}
                checked={hideMicrosoft}
                onChange={(e) => setHideMicrosoft(e.target.checked)}
              />
              Hide Microsoft’s own apps
            </label>
            <div className="grow" />
            <span className="muted">
              {rows.length} {rows.length === 1 ? 'service principal' : 'service principals'}
            </span>
          </div>

          <div className="card" style={{ padding: 0 }}>
            <table className="table">
              <thead>
                <tr>
                  <SortTh id="name" sort={sort}>
                    Service principal
                  </SortTh>
                  <SortTh id="type" sort={sort}>
                    Type
                  </SortTh>
                  <SortTh id="tenant" sort={sort}>
                    Home tenant
                  </SortTh>
                  <SortTh id="enabled" sort={sort}>
                    Enabled
                  </SortTh>
                  <SortTh id="permissions" sort={sort}>
                    Permissions
                  </SortTh>
                  <SortTh id="last-sign-in" sort={sort}>
                    Last sign-in
                  </SortTh>
                </tr>
              </thead>
              <tbody>
                {sorted.map((row) => {
                  const { holder } = row;
                  const sp = holder.sp;
                  const isOpen = open.has(holder.id);
                  const tenant = homeTenant(holder, currentTenantId);
                  return (
                    <Fragment key={holder.id}>
                      <tr
                        className={isOpen ? 'selected' : undefined}
                        onClick={() => toggle(holder.id)}
                        title={
                          isOpen
                            ? 'Hide what these permissions allow'
                            : 'Show what these permissions allow'
                        }
                      >
                        <td>
                          <div className="row" style={{ gap: 10, alignItems: 'center' }}>
                            <AppIcon
                              id={sp?.appId || holder.id}
                              logoUrl={sp?.info?.logoUrl}
                              title={holder.displayName}
                            />
                            <Link
                              to={`/enterprise-apps/${holder.id}`}
                              onClick={(e) => e.stopPropagation()}
                              title="Open the enterprise app to review or revoke"
                            >
                              {holder.displayName}
                            </Link>
                          </div>
                        </td>
                        <td className="muted">{sp?.servicePrincipalType ?? '—'}</td>
                        <td>
                          {tenant === 'This tenant' ? (
                            <span className="badge granted">This tenant</span>
                          ) : tenant ? (
                            <span className="badge">{tenant}</span>
                          ) : (
                            <span className="muted">—</span>
                          )}
                        </td>
                        <td>
                          {!sp ? (
                            <span className="muted">—</span>
                          ) : sp.accountEnabled ? (
                            <span className="badge granted">Enabled</span>
                          ) : (
                            <span className="badge">Disabled</span>
                          )}
                        </td>
                        <td>
                          <div className="perm-chips">
                            {row.permissions.map((p) => (
                              <PermissionChip key={p.assignmentId} permission={p} />
                            ))}
                          </div>
                        </td>
                        <td className="muted" style={{ whiteSpace: 'nowrap' }}>
                          {row.lastSignIn ? (
                            <span title={new Date(row.lastSignIn).toLocaleString()}>
                              {relativeTime(new Date(row.lastSignIn).toISOString())}
                            </span>
                          ) : lastSignIns && sp?.appId ? (
                            <span title="No sign-in in the service principal sign-in activity report">
                              None recorded
                            </span>
                          ) : (
                            '—'
                          )}
                        </td>
                      </tr>
                      {isOpen && (
                        <tr className="static">
                          <td colSpan={6} className="perm-detail">
                            <PermissionDetails permissions={row.permissions} />
                          </td>
                        </tr>
                      )}
                    </Fragment>
                  );
                })}
                {rows.length === 0 && (
                  <tr className="static">
                    <td colSpan={6}>
                      <div className="empty">
                        {search.trim()
                          ? 'Nothing matches the filter.'
                          : level === 'all'
                            ? 'No service principal holds an app-only permission on the scanned APIs.'
                            : 'No service principal holds a permission at this level.'}
                      </div>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          <p className="muted" style={{ fontSize: 12 }}>
            Scanned: {scanned.map((r) => r.name).join(', ') || 'nothing'}. App-only permissions on
            other APIs, delegated permissions and directory roles held by service principals are not
            part of this list. Critical and high are this app’s own rating of a fixed list of
            permissions; open a row to see what each one allows.
          </p>
        </>
      )}
    </>
  );
}

function homeTenant(holder: PermissionHolder, currentTenantId: string | undefined): string | null {
  const owner = holder.sp?.appOwnerOrganizationId;
  if (!owner) return null;
  if (currentTenantId && owner.toLowerCase() === currentTenantId.toLowerCase())
    return 'This tenant';
  return isMicrosoftTenant(owner) ? 'Microsoft' : 'External';
}

const chipTone = (p: HeldPermission) => p.risk?.tier ?? 'unrated';

function PermissionChip({ permission: p }: { permission: HeldPermission }) {
  return (
    <span
      className={`perm-chip ${chipTone(p)}`}
      title={`${p.resourceName}${p.risk ? ` · ${p.risk.tier}: ${p.risk.why}` : ''}`}
    >
      {p.resourceAppId !== MS_GRAPH && <span className="perm-chip-api">{p.resourceName}</span>}
      {p.value}
    </span>
  );
}

function PermissionDetails({ permissions }: { permissions: HeldPermission[] }) {
  return (
    <table className="perm-detail-table">
      <tbody>
        {permissions.map((p) => (
          <tr key={p.assignmentId}>
            <td>
              <span className={`priv-badge priv-${p.risk?.tier ?? 'unknown'}`}>
                {p.risk?.tier ?? 'Not rated'}
              </span>
            </td>
            <td>
              <span className="mono">{p.value}</span>
              <div className="muted">{p.resourceName}</div>
            </td>
            <td>
              {p.risk?.why ?? p.displayName ?? '—'}
              {p.grantedDateTime && (
                <div className="muted">
                  Granted {new Date(p.grantedDateTime).toLocaleDateString()}
                </div>
              )}
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}
