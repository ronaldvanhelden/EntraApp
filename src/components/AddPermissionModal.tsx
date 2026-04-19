import { useEffect, useMemo, useState } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import {
  getServicePrincipal,
  getServicePrincipalByAppId,
  listServicePrincipals,
} from '../graph/servicePrincipals';
import type { ServicePrincipal } from '../graph/types';
import type { PrincipalRef } from '../graph/directoryObjects';
import { Modal } from './Modal';
import { PrincipalPicker } from './PrincipalPicker';
import { ScopeLookupPanel } from './ScopeLookupPanel';
import { useScopeCatalog } from './PrivilegeBadge';
import { MS_GRAPH_APP_ID, type EndpointMatch } from '../graph/scopeCatalog';

export interface AddPermissionSubmitPayload {
  resource: ServicePrincipal;
  appRoleIds: string[];
  delegatedScopes: string[];
  consentType: 'AllPrincipals' | 'Principal';
  /** User/group object id when consentType === 'Principal' (delegated). */
  principalId?: string;
  /** Principal SP that the application-permission role assignment is on behalf of.
   *  Defaults to the client SP (undefined means "use clientSp"). */
  applicationPrincipalId?: string;
  applicationPrincipalName?: string;
}

interface Props {
  clientSp: ServicePrincipal;
  onClose: () => void;
  onSubmit: (payload: AddPermissionSubmitPayload) => Promise<void>;
}

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

export function AddPermissionModal({ clientSp, onClose, onSubmit }: Props) {
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
  const [delegatedPrincipal, setDelegatedPrincipal] =
    useState<PrincipalRef | null>(null);
  const [applicationPrincipal, setApplicationPrincipal] =
    useState<PrincipalRef | null>({
      id: clientSp.id,
      displayName: clientSp.displayName,
      kind: 'sp',
      subtitle: clientSp.appId,
    });
  const [submitting, setSubmitting] = useState(false);
  const [filter, setFilter] = useState('');
  const [mode, setMode] = useState<'api' | 'url'>('api');
  const catalog = useScopeCatalog();
  const [pendingPick, setPendingPick] = useState<
    | { scope: string; kind: 'delegated' | 'application' }
    | null
  >(null);

  useEffect(() => {
    if (!resource || !pendingPick) return;
    if (pendingPick.kind === 'delegated') {
      const id = resource.oauth2PermissionScopes?.find(
        (s) => s.value === pendingPick.scope,
      )?.id;
      if (id) {
        setSelectedDel((prev) => {
          const next = new Set(prev);
          next.add(id);
          return next;
        });
        setTab('delegated');
      }
    } else {
      const id = resource.appRoles?.find(
        (r) => r.value === pendingPick.scope,
      )?.id;
      if (id) {
        setSelectedApp((prev) => {
          const next = new Set(prev);
          next.add(id);
          return next;
        });
        setTab('app');
      }
    }
    setPendingPick(null);
  }, [resource, pendingPick]);

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
    if (consentType === 'Principal' && selectedDel.size > 0 && !delegatedPrincipal) {
      setError('Pick a user or group for single-principal consent.');
      return;
    }
    if (selectedApp.size > 0 && !applicationPrincipal) {
      setError('Pick a service principal that the application permission is assigned to.');
      return;
    }
    setSubmitting(true);
    setError(null);
    try {
      await onSubmit({
        resource,
        appRoleIds: [...selectedApp],
        delegatedScopes: [...selectedDel]
          .map((id) => delegatedScopes.find((s) => s.id === id)?.value ?? '')
          .filter(Boolean),
        consentType,
        principalId:
          consentType === 'Principal' ? delegatedPrincipal?.id : undefined,
        applicationPrincipalId: applicationPrincipal?.id,
        applicationPrincipalName: applicationPrincipal?.displayName,
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
          <div className="tabs" style={{ marginBottom: 16 }}>
            <button
              className={mode === 'api' ? 'active' : ''}
              onClick={() => setMode('api')}
            >
              By API
            </button>
            <button
              className={mode === 'url' ? 'active' : ''}
              onClick={() => setMode('url')}
            >
              By Graph URL
            </button>
          </div>
          {mode === 'api' ? (
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
              <h3 style={{ marginTop: 0 }}>Find scopes by Graph URL</h3>
              <p className="muted" style={{ fontSize: 12, marginTop: 0 }}>
                Only Microsoft Graph endpoints are in the public catalog.
                Picking a result loads the Microsoft Graph service principal
                and pre-selects the scope.
              </p>
              <ScopeLookupPanel
                catalog={catalog}
                onPick={async (match: EndpointMatch, kind) => {
                  setPendingPick({ scope: match.scope, kind });
                  await selectResourceByAppId(MS_GRAPH_APP_ID, 'Microsoft Graph');
                }}
              />
              {error && <p className="error">{error}</p>}
            </>
          )}
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
                    Consent for single user or group
                  </label>
                </div>
                {consentType === 'Principal' && (
                  <div style={{ marginTop: 10 }}>
                    <div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>
                      Principal (user or group)
                    </div>
                    <PrincipalPicker
                      kinds={['user', 'group']}
                      selected={delegatedPrincipal}
                      onChange={setDelegatedPrincipal}
                    />
                  </div>
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
              <div
                className="card"
                style={{ padding: 12, marginTop: 16, marginBottom: 0 }}
              >
                <div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>
                  Assign role to (principal service principal)
                </div>
                <PrincipalPicker
                  kinds={['sp']}
                  selected={applicationPrincipal}
                  onChange={setApplicationPrincipal}
                />
                <div className="muted" style={{ fontSize: 11, marginTop: 6 }}>
                  Defaults to this enterprise app. Change to grant the role to a
                  different service principal.
                </div>
              </div>
            </>
          )}

          {error && <p className="error">{error}</p>}
        </>
      )}
    </Modal>
  );
}

export function toggleSet(prev: Set<string>, id: string): Set<string> {
  const next = new Set(prev);
  if (next.has(id)) next.delete(id);
  else next.add(id);
  return next;
}

export interface ChecklistItem {
  id: string;
  title: string;
  subtitle: string;
  description: string;
  badge?: string;
}

export function PermissionChecklist({
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
