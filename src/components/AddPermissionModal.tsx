import { useEffect, useMemo, useState } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import {
  getServicePrincipal,
  getServicePrincipalByAppId,
  listServicePrincipals,
} from '../graph/servicePrincipals';
import type { ServicePrincipal } from '../graph/types';
import { Modal } from './Modal';

interface SubmitPayload {
  resource: ServicePrincipal;
  appRoleIds: string[];
  delegatedScopes: string[];
  consentType: 'AllPrincipals' | 'Principal';
  principalId?: string;
}

interface Props {
  onClose: () => void;
  onSubmit: (payload: SubmitPayload) => Promise<void>;
}

// Known well-known API app IDs for quick picking.
const COMMON_APIS: Array<{ appId: string; name: string }> = [
  { appId: '00000003-0000-0000-c000-000000000000', name: 'Microsoft Graph' },
  {
    appId: '00000002-0000-0ff1-ce00-000000000000',
    name: 'Office 365 Exchange Online',
  },
  {
    appId: '00000003-0000-0ff1-ce00-000000000000',
    name: 'Office 365 SharePoint Online',
  },
  {
    appId: 'c5393580-f805-4401-95e8-94b7a6ef2fc2',
    name: 'Office 365 Management APIs',
  },
  {
    appId: '797f4846-ba00-4fd7-ba43-dac1f8f63013',
    name: 'Azure Service Management',
  },
];

export function AddPermissionModal({ onClose, onSubmit }: Props) {
  const token = useGraphToken();

  const [resource, setResource] = useState<ServicePrincipal | null>(null);
  const [resourceLoading, setResourceLoading] = useState(false);
  const [search, setSearch] = useState('');
  const [results, setResults] = useState<ServicePrincipal[]>([]);
  const [error, setError] = useState<string | null>(null);

  const [tab, setTab] = useState<'delegated' | 'app'>('delegated');
  const [selectedApp, setSelectedApp] = useState<Set<string>>(new Set());
  const [selectedDel, setSelectedDel] = useState<Set<string>>(new Set());
  const [consentType, setConsentType] =
    useState<'AllPrincipals' | 'Principal'>('AllPrincipals');
  const [principalId, setPrincipalId] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [filter, setFilter] = useState('');

  // Live search of service principals while typing
  useEffect(() => {
    if (resource) return;
    if (search.trim().length < 2) {
      setResults([]);
      return;
    }
    const handle = setTimeout(() => {
      listServicePrincipals(token, search.trim())
        .then((r) => setResults(r.slice(0, 50)))
        .catch((e) => setError(e.message));
    }, 250);
    return () => clearTimeout(handle);
  }, [search, token, resource]);

  const selectResourceByAppId = async (appId: string, name: string) => {
    setError(null);
    setResourceLoading(true);
    try {
      const sp = await getServicePrincipalByAppId(token, appId);
      if (!sp) {
        setError(
          `Service principal for "${name}" not found. Add it via the Microsoft Entra portal first.`,
        );
        return;
      }
      const full = await getServicePrincipal(token, sp.id);
      setResource(full);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setResourceLoading(false);
    }
  };

  const selectResource = async (sp: ServicePrincipal) => {
    setError(null);
    setResourceLoading(true);
    try {
      const full = await getServicePrincipal(token, sp.id);
      setResource(full);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setResourceLoading(false);
    }
  };

  const delegatedScopes = useMemo(
    () => resource?.oauth2PermissionScopes?.filter((s) => s.isEnabled) ?? [],
    [resource],
  );
  const appRoles = useMemo(
    () =>
      resource?.appRoles?.filter(
        (r) => r.isEnabled && r.allowedMemberTypes?.includes('Application'),
      ) ?? [],
    [resource],
  );

  const filterFn = (text: string) =>
    !filter || text.toLowerCase().includes(filter.toLowerCase());

  const submit = async () => {
    if (!resource) return;
    setSubmitting(true);
    setError(null);
    try {
      await onSubmit({
        resource,
        appRoleIds: [...selectedApp],
        delegatedScopes: [...selectedDel].map(
          (id) =>
            delegatedScopes.find((s) => s.id === id)?.value ?? '',
        ).filter(Boolean),
        consentType,
        principalId:
          consentType === 'Principal' ? principalId.trim() : undefined,
      });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setSubmitting(false);
    }
  };

  const hasSelection = selectedApp.size > 0 || selectedDel.size > 0;

  return (
    <Modal
      title="Add API permission"
      onClose={onClose}
      footer={
        <>
          <button onClick={onClose}>Cancel</button>
          <button
            className="primary"
            disabled={!resource || !hasSelection || submitting}
            onClick={submit}
          >
            {submitting ? 'Granting…' : 'Grant selected'}
          </button>
        </>
      }
    >
      {!resource ? (
        <>
          <h3 style={{ marginTop: 0 }}>Select an API</h3>
          <div className="row wrap" style={{ marginBottom: 16 }}>
            {COMMON_APIS.map((api) => (
              <button
                key={api.appId}
                disabled={resourceLoading}
                onClick={() => selectResourceByAppId(api.appId, api.name)}
              >
                {api.name}
              </button>
            ))}
          </div>

          <label className="field">
            <span>Or search all APIs (by display name)</span>
            <input
              autoFocus
              placeholder="Type at least 2 characters…"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </label>

          {resourceLoading && (
            <div className="row" style={{ gap: 8 }}>
              <span className="spinner" />
              <span className="muted">Loading API…</span>
            </div>
          )}

          {results.length > 0 && (
            <div className="card" style={{ padding: 0, marginTop: 8 }}>
              <table className="table">
                <tbody>
                  {results.map((sp) => (
                    <tr key={sp.id} onClick={() => selectResource(sp)}>
                      <td>{sp.displayName}</td>
                      <td className="mono muted">{sp.appId}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {error && <p className="error">{error}</p>}
        </>
      ) : (
        <>
          <div className="row" style={{ justifyContent: 'space-between' }}>
            <div>
              <div style={{ fontSize: 15, fontWeight: 600 }}>
                {resource.displayName}
              </div>
              <div className="mono muted" style={{ fontSize: 12 }}>
                {resource.appId}
              </div>
            </div>
            <button
              className="ghost"
              onClick={() => {
                setResource(null);
                setSelectedApp(new Set());
                setSelectedDel(new Set());
                setSearch('');
              }}
            >
              ← Change API
            </button>
          </div>

          <div className="tabs" style={{ marginTop: 16 }}>
            <button
              className={tab === 'delegated' ? 'active' : ''}
              onClick={() => setTab('delegated')}
            >
              Delegated ({delegatedScopes.length})
            </button>
            <button
              className={tab === 'app' ? 'active' : ''}
              onClick={() => setTab('app')}
            >
              Application ({appRoles.length})
            </button>
          </div>

          <input
            placeholder="Filter permissions…"
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
            style={{ marginBottom: 12 }}
          />

          {tab === 'delegated' && (
            <>
              {delegatedScopes.length === 0 ? (
                <div className="empty">
                  This API exposes no delegated scopes.
                </div>
              ) : (
                <PermissionChecklist
                  items={delegatedScopes
                    .filter((s) =>
                      filterFn(`${s.value} ${s.adminConsentDisplayName}`),
                    )
                    .map((s) => ({
                      id: s.id,
                      title: s.value,
                      subtitle: s.adminConsentDisplayName,
                      description: s.adminConsentDescription,
                      badge: s.type === 'Admin' ? 'Admin' : 'User',
                    }))}
                  selected={selectedDel}
                  onToggle={(id) =>
                    setSelectedDel((prev) => toggleSet(prev, id))
                  }
                />
              )}

              <div
                className="card"
                style={{ padding: 12, marginTop: 16, marginBottom: 0 }}
              >
                <div className="row">
                  <label className="row" style={{ gap: 6 }}>
                    <input
                      type="radio"
                      style={{ width: 'auto' }}
                      checked={consentType === 'AllPrincipals'}
                      onChange={() => setConsentType('AllPrincipals')}
                    />
                    Consent for all users (admin)
                  </label>
                  <label className="row" style={{ gap: 6 }}>
                    <input
                      type="radio"
                      style={{ width: 'auto' }}
                      checked={consentType === 'Principal'}
                      onChange={() => setConsentType('Principal')}
                    />
                    Consent for single user
                  </label>
                </div>
                {consentType === 'Principal' && (
                  <input
                    style={{ marginTop: 8 }}
                    placeholder="User object ID"
                    value={principalId}
                    onChange={(e) => setPrincipalId(e.target.value)}
                  />
                )}
              </div>
            </>
          )}

          {tab === 'app' && (
            <>
              {appRoles.length === 0 ? (
                <div className="empty">
                  This API exposes no application roles.
                </div>
              ) : (
                <PermissionChecklist
                  items={appRoles
                    .filter((r) => filterFn(`${r.value} ${r.displayName}`))
                    .map((r) => ({
                      id: r.id,
                      title: r.value,
                      subtitle: r.displayName,
                      description: r.description,
                    }))}
                  selected={selectedApp}
                  onToggle={(id) =>
                    setSelectedApp((prev) => toggleSet(prev, id))
                  }
                />
              )}
            </>
          )}

          {error && <p className="error">{error}</p>}
        </>
      )}
    </Modal>
  );
}

function toggleSet(prev: Set<string>, id: string): Set<string> {
  const next = new Set(prev);
  if (next.has(id)) next.delete(id);
  else next.add(id);
  return next;
}

interface ChecklistItem {
  id: string;
  title: string;
  subtitle: string;
  description: string;
  badge?: string;
}

function PermissionChecklist({
  items,
  selected,
  onToggle,
}: {
  items: ChecklistItem[];
  selected: Set<string>;
  onToggle: (id: string) => void;
}) {
  return (
    <div className="stack" style={{ maxHeight: 360, overflowY: 'auto' }}>
      {items.map((it) => {
        const isSelected = selected.has(it.id);
        return (
          <label
            key={it.id}
            className="row"
            style={{
              alignItems: 'flex-start',
              gap: 10,
              padding: 8,
              border: '1px solid var(--border-subtle)',
              borderRadius: 6,
              background: isSelected ? 'var(--bg-hover)' : 'transparent',
              cursor: 'pointer',
            }}
          >
            <input
              type="checkbox"
              style={{ width: 'auto', marginTop: 3 }}
              checked={isSelected}
              onChange={() => onToggle(it.id)}
            />
            <div className="grow">
              <div className="row" style={{ gap: 8 }}>
                <span className="mono" style={{ fontWeight: 600 }}>
                  {it.title}
                </span>
                {it.badge && <span className="badge">{it.badge}</span>}
              </div>
              <div style={{ fontSize: 13 }}>{it.subtitle}</div>
              {it.description && (
                <div
                  className="muted"
                  style={{ fontSize: 12, marginTop: 2 }}
                >
                  {it.description}
                </div>
              )}
            </div>
          </label>
        );
      })}
      {items.length === 0 && (
        <div className="empty">No permissions match filter.</div>
      )}
    </div>
  );
}
