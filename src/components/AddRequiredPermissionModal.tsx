import { useEffect, useMemo, useState } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import {
  getServicePrincipal,
  getServicePrincipalByAppId,
  listServicePrincipals,
} from '../graph/servicePrincipals';
import type { RequiredResourceAccess, ServicePrincipal } from '../graph/types';
import { Modal } from './Modal';
import {
  PermissionChecklist,
  toggleSet,
} from './AddPermissionModal';
import { ScopeLookupPanel } from './ScopeLookupPanel';
import { useScopeCatalog } from './PrivilegeBadge';
import { MS_GRAPH_APP_ID, type EndpointMatch } from '../graph/scopeCatalog';

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

interface Props {
  /** Existing entries — used so we can pre-select already-declared permissions
   *  and merge (rather than replace) when the user picks the same API. */
  existing: RequiredResourceAccess[];
  onClose: () => void;
  onSubmit: (entry: RequiredResourceAccess) => Promise<void>;
}

export function AddRequiredPermissionModal({
  existing,
  onClose,
  onSubmit,
}: Props) {
  const token = useGraphToken();

  const [resource, setResource] = useState<ServicePrincipal | null>(null);
  const [resourceLoading, setResourceLoading] = useState(false);
  const [search, setSearch] = useState('');
  const [results, setResults] = useState<ServicePrincipal[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [tab, setTab] = useState<'delegated' | 'app'>('delegated');
  const [selectedApp, setSelectedApp] = useState<Set<string>>(new Set());
  const [selectedDel, setSelectedDel] = useState<Set<string>>(new Set());
  const [filter, setFilter] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [mode, setMode] = useState<'api' | 'url'>('api');
  const catalog = useScopeCatalog();
  // When the "By URL" flow picks a scope we remember it and apply the
  // selection as soon as the resource SP finishes loading (it exposes the
  // scope→id map we need).
  const [pendingPick, setPendingPick] = useState<
    | { scope: string; kind: 'delegated' | 'application' }
    | null
  >(null);

  const prepareSelectionForScope = (
    match: EndpointMatch,
    kind: 'delegated' | 'application',
  ) => {
    setPendingPick({ scope: match.scope, kind });
  };

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

  const pickResource = async (sp: ServicePrincipal) => {
    setError(null);
    setResourceLoading(true);
    try {
      const full = await getServicePrincipal(token, sp.id);
      setResource(full);
      seedFromExisting(full);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setResourceLoading(false);
    }
  };

  const pickResourceByAppId = async (appId: string, name: string) => {
    setError(null);
    setResourceLoading(true);
    try {
      const sp = await getServicePrincipalByAppId(token, appId);
      if (!sp) {
        setError(
          `Service principal for "${name}" not found in this tenant. Add it first.`,
        );
        return;
      }
      const full = await getServicePrincipal(token, sp.id);
      setResource(full);
      seedFromExisting(full);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setResourceLoading(false);
    }
  };

  const seedFromExisting = (sp: ServicePrincipal) => {
    const entry = existing.find((e) => e.resourceAppId === sp.appId);
    if (!entry) return;
    setSelectedApp(
      new Set(
        entry.resourceAccess.filter((a) => a.type === 'Role').map((a) => a.id),
      ),
    );
    setSelectedDel(
      new Set(
        entry.resourceAccess.filter((a) => a.type === 'Scope').map((a) => a.id),
      ),
    );
  };

  // Apply a "By URL" pick once the resource SP has loaded: resolve scope
  // value → id on the SP's scope/role arrays, then tick the checklist.
  useEffect(() => {
    if (!resource || !pendingPick) return;
    if (pendingPick.kind === 'delegated') {
      const id = resource.oauth2PermissionScopes?.find(
        (s) => s.value === pendingPick.scope,
      )?.id;
      if (id) {
        setSelectedDel((prev) => new Set(prev).add(id));
        setTab('delegated');
      }
    } else {
      const id = resource.appRoles?.find(
        (r) => r.value === pendingPick.scope,
      )?.id;
      if (id) {
        setSelectedApp((prev) => new Set(prev).add(id));
        setTab('app');
      }
    }
    setPendingPick(null);
  }, [resource, pendingPick]);

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
        resourceAppId: resource.appId,
        resourceAccess: [
          ...[...selectedApp].map((id) => ({ id, type: 'Role' as const })),
          ...[...selectedDel].map((id) => ({ id, type: 'Scope' as const })),
        ],
      });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setSubmitting(false);
    }
  };

  const hasSelection = selectedApp.size + selectedDel.size > 0;

  return (
    <Modal
      title="Add required permission"
      onClose={onClose}
      footer={
        <>
          <button onClick={onClose}>Cancel</button>
          <button
            className="primary"
            disabled={!resource || !hasSelection || submitting}
            onClick={submit}
          >
            {submitting ? 'Saving…' : 'Save to manifest'}
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
                    onClick={() => pickResourceByAppId(api.appId, api.name)}
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
                        <tr key={sp.id} onClick={() => pickResource(sp)}>
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
                onPick={async (match, kind) => {
                  await pickResourceByAppId(
                    MS_GRAPH_APP_ID,
                    'Microsoft Graph',
                  );
                  prepareSelectionForScope(match, kind);
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

          {tab === 'delegated' ? (
            delegatedScopes.length === 0 ? (
              <div className="empty">This API exposes no delegated scopes.</div>
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
            )
          ) : appRoles.length === 0 ? (
            <div className="empty">This API exposes no application roles.</div>
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
              onToggle={(id) => setSelectedApp((prev) => toggleSet(prev, id))}
            />
          )}

          <p className="muted" style={{ fontSize: 12, marginTop: 12 }}>
            Saving updates the application's manifest only. Permissions still
            need to be consented/granted on the enterprise app (service
            principal) before they take effect.
          </p>

          {error && <p className="error">{error}</p>}
        </>
      )}
    </Modal>
  );
}
