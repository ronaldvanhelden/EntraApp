import { useState } from 'react';
import { useMsal } from '@azure/msal-react';
import { saveAuthConfig, clearAuthConfig } from '../auth/config';
import { useAuthConfig } from '../auth/context';

export function Settings() {
  const { config, setConfig } = useAuthConfig();
  const { instance } = useMsal();
  const [clientId, setClientId] = useState(config.clientId);
  const [tenantId, setTenantId] = useState(config.tenantId);
  const [saved, setSaved] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const save = async () => {
    setError(null);
    const trimmed = clientId.trim();
    if (!/^[0-9a-f-]{36}$/i.test(trimmed)) {
      setError('Client ID must be a valid GUID.');
      return;
    }
    const next = { clientId: trimmed, tenantId: tenantId.trim() || 'common' };
    saveAuthConfig(next);
    setConfig(next);
    setSaved(true);
    // Sign out because MSAL will rebuild with new authority
    try {
      await instance.logoutPopup({ mainWindowRedirectUri: '/' });
    } catch {
      /* ignore */
    }
  };

  const reset = () => {
    clearAuthConfig();
    window.location.reload();
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
        <h3>App registration used for sign-in</h3>
        <p className="muted" style={{ marginTop: 0 }}>
          Changing these values triggers a fresh sign-in.
        </p>

        <label className="field">
          <span>Client ID</span>
          <input
            value={clientId}
            onChange={(e) => setClientId(e.target.value)}
          />
        </label>

        <label className="field">
          <span>Tenant ID</span>
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
