import { useState } from 'react';
import { useMsal } from '@azure/msal-react';
import { clearAuthConfig } from '../auth/config';
import { useAuthConfig } from '../auth/context';
import {
  buildDefaultUrl,
  buildTenantUrl,
  exitTenantOverride,
  isValidTenantId,
  loadTenantShortcuts,
  saveTenantShortcuts,
  type TenantShortcut,
} from '../auth/tenantParam';
import { CopyButton } from '../components/CopyButton';

export function Settings() {
  const { config, savedConfig, tenantOverride, setConfig } = useAuthConfig();
  const { instance } = useMsal();
  const [clientId, setClientId] = useState(savedConfig.clientId);
  const [tenantId, setTenantId] = useState(savedConfig.tenantId);
  const [saved, setSaved] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const save = async () => {
    setError(null);
    const trimmedClient = clientId.trim();
    if (!/^[0-9a-f-]{36}$/i.test(trimmedClient)) {
      setError('Client ID must be a valid GUID.');
      return;
    }
    const trimmedTenant = tenantId.trim() || 'common';
    if (!isValidTenantId(trimmedTenant)) {
      setError(
        'Tenant must be a GUID, a domain name, or common/organizations.',
      );
      return;
    }
    setConfig({ clientId: trimmedClient, tenantId: trimmedTenant });
    setSaved(true);
    // Sign out because MSAL will rebuild with new authority. Return to the URL
    // without ?tenant= so the saved default actually takes effect.
    try {
      await instance.logoutPopup({
        mainWindowRedirectUri: buildDefaultUrl('/settings'),
      });
    } catch {
      /* ignore */
    }
  };

  const reset = () => {
    clearAuthConfig();
    // Reload without ?tenant= so nothing survives the reset.
    exitTenantOverride('/settings');
  };

  return (
    <>
      <div className="page-header">
        <div>
          <h1>Settings</h1>
          <div className="subtitle">Authentication configuration</div>
        </div>
      </div>

      <div className="card" style={{ maxWidth: 640 }}>
        <h3>Active tenant</h3>
        <div className="kv">
          <div className="k">Signing in against</div>
          <div className="mono">
            {config.tenantId}{' '}
            {tenantOverride ? (
              <span className="badge delegated">from URL</span>
            ) : (
              <span className="badge">saved default</span>
            )}
          </div>
          <div className="k">Saved default</div>
          <div className="mono">{savedConfig.tenantId}</div>
        </div>
        {tenantOverride && (
          <>
            <p className="muted" style={{ fontSize: 13 }}>
              This tab is pinned to{' '}
              <span className="mono">{tenantOverride}</span> by the{' '}
              <span className="mono">?tenant=</span> parameter in the URL, so
              sign-in and every Graph call target that directory.
            </p>
            <div className="row" style={{ justifyContent: 'flex-end' }}>
              <button onClick={() => exitTenantOverride('/settings')}>
                Use default tenant instead
              </button>
            </div>
          </>
        )}
      </div>

      <TenantShortcuts currentTenantId={config.tenantId} />

      <div className="card" style={{ maxWidth: 640 }}>
        <h3>App registration used for sign-in</h3>
        <p className="muted" style={{ marginTop: 0 }}>
          Changing these values triggers a fresh sign-in
          {tenantOverride ? ' and releases the tenant pinned by the URL' : ''}.
        </p>

        <label className="field">
          <span>Client ID</span>
          <input
            value={clientId}
            onChange={(e) => setClientId(e.target.value)}
          />
        </label>

        <label className="field">
          <span>Default tenant (GUID, domain, or common/organizations)</span>
          <input
            value={tenantId}
            onChange={(e) => setTenantId(e.target.value)}
          />
        </label>

        {error && <p className="error">{error}</p>}
        {saved && <p style={{ color: 'var(--success)' }}>Saved.</p>}

        <div className="row" style={{ justifyContent: 'space-between' }}>
          <button className="danger" onClick={reset}>
            Reset configuration
          </button>
          <button className="primary" onClick={save}>
            Save
          </button>
        </div>
      </div>
    </>
  );
}

// Bookmarkable per-tenant links. Each one is this page's own URL plus
// ?tenant=<tenant>, so it can go on the bookmark bar, the Windows start menu,
// or a phone home screen.
function TenantShortcuts({ currentTenantId }: { currentTenantId: string }) {
  const [shortcuts, setShortcuts] = useState<TenantShortcut[]>(() =>
    loadTenantShortcuts(),
  );
  const [label, setLabel] = useState('');
  const [tenant, setTenant] = useState('');
  const [error, setError] = useState<string | null>(null);

  const persist = (next: TenantShortcut[]) => {
    saveTenantShortcuts(next);
    setShortcuts(next);
  };

  const add = () => {
    setError(null);
    const t = tenant.trim();
    if (!isValidTenantId(t)) {
      setError('Tenant must be a GUID or a domain name (e.g. debble.com).');
      return;
    }
    if (shortcuts.some((s) => s.tenantId.toLowerCase() === t.toLowerCase())) {
      setError('A shortcut for that tenant already exists.');
      return;
    }
    persist([...shortcuts, { label: label.trim() || t, tenantId: t }]);
    setLabel('');
    setTenant('');
  };

  const remove = (tenantId: string) => {
    persist(shortcuts.filter((s) => s.tenantId !== tenantId));
  };

  return (
    <div className="card" style={{ maxWidth: 640 }}>
      <h3>Tenant shortcuts</h3>
      <p className="muted" style={{ marginTop: 0, fontSize: 13 }}>
        Each shortcut opens EntraApp pinned to one tenant. Open it in a new tab,
        or bookmark it to jump straight into that directory — handy when your
        home tenant differs from the one you manage.
      </p>

      {shortcuts.length === 0 ? (
        <div className="muted" style={{ fontSize: 13 }}>
          No shortcuts yet.
        </div>
      ) : (
        <div className="stack">
          {shortcuts.map((s) => {
            const url = buildTenantUrl(s.tenantId);
            const active =
              s.tenantId.toLowerCase() === currentTenantId.toLowerCase();
            return (
              <div className="shortcut-row" key={s.tenantId}>
                <div className="shortcut-main">
                  <div className="shortcut-label">
                    <a href={url} target="_blank" rel="noreferrer">
                      {s.label}
                    </a>
                    {active && <span className="badge granted">active</span>}
                  </div>
                  <div className="mono shortcut-url" title={url}>
                    {url}
                  </div>
                </div>
                <CopyButton value={url} label="shortcut URL" />
                <button className="ghost" onClick={() => remove(s.tenantId)}>
                  Remove
                </button>
              </div>
            );
          })}
        </div>
      )}

      <div
        className="row wrap"
        style={{ marginTop: 12, alignItems: 'flex-end' }}
      >
        <label className="field grow" style={{ marginBottom: 0 }}>
          <span>Name</span>
          <input
            value={label}
            onChange={(e) => setLabel(e.target.value)}
            placeholder="Debble"
          />
        </label>
        <label className="field grow" style={{ marginBottom: 0 }}>
          <span>Tenant (GUID or domain)</span>
          <input
            value={tenant}
            onChange={(e) => setTenant(e.target.value)}
            placeholder="debble.com"
          />
        </label>
        <button onClick={add}>Add</button>
      </div>
      {error && <p className="error">{error}</p>}
    </div>
  );
}
