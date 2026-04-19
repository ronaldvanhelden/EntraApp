import { useState } from 'react';
import { computeRedirectUri, saveAuthConfig } from '../auth/config';
import { useAuthConfig } from '../auth/context';

export function SetupScreen() {
  const { config, setConfig } = useAuthConfig();
  const [clientId, setClientId] = useState(config.clientId);
  const [tenantId, setTenantId] = useState(config.tenantId || 'common');
  const [error, setError] = useState<string | null>(null);

  const save = () => {
    const trimmed = clientId.trim();
    if (!/^[0-9a-f-]{36}$/i.test(trimmed)) {
      setError('Client ID must be a valid GUID.');
      return;
    }
    const next = { clientId: trimmed, tenantId: tenantId.trim() || 'common' };
    saveAuthConfig(next);
    setConfig(next);
  };

  return (
    <div className="center">
      <div className="card" style={{ maxWidth: 560, width: '100%' }}>
        <h3>Configure EntraApp</h3>
        <p className="muted" style={{ marginTop: 0 }}>
          Enter the client (application) ID and tenant ID of an app registration
          this SPA should sign in with. The app must have <code>Single-page
          application</code> redirect URI configured for the host page.
        </p>

        <label className="field">
          <span>Client ID</span>
          <input
            value={clientId}
            onChange={(e) => setClientId(e.target.value)}
            placeholder="00000000-0000-0000-0000-000000000000"
            autoFocus
          />
        </label>

        <label className="field">
          <span>Tenant ID (or "common" / "organizations")</span>
          <input
            value={tenantId}
            onChange={(e) => setTenantId(e.target.value)}
            placeholder="common"
          />
        </label>

        {error && <p className="error">{error}</p>}

        <div className="row" style={{ justifyContent: 'flex-end' }}>
          <button className="primary" onClick={save}>
            Save & continue
          </button>
        </div>

        <div
          className="card"
          style={{ marginTop: 16, padding: 12, background: 'var(--bg)' }}
        >
          <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 4 }}>
            Before you sign in
          </div>
          <div className="muted" style={{ fontSize: 12 }}>
            Register this exact URL as a <strong>Single-page application</strong>{' '}
            redirect URI on the app registration (Authentication → Platform
            configurations → Single-page application → Add URI):
          </div>
          <div
            className="mono"
            style={{
              marginTop: 6,
              padding: 6,
              background: 'var(--bg-elevated)',
              border: '1px solid var(--border)',
              borderRadius: 4,
              wordBreak: 'break-all',
            }}
          >
            {computeRedirectUri()}
          </div>
        </div>
      </div>
    </div>
  );
}
