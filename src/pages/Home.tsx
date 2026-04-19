import { useEffect, useState } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import { getMe, type MeResponse } from '../graph/me';

export function Home() {
  const token = useGraphToken();
  const [me, setMe] = useState<MeResponse | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    getMe(token)
      .then(setMe)
      .catch((e) => setError(e.message));
  }, [token]);

  return (
    <>
      <div className="page-header">
        <div>
          <h1>Overview</h1>
          <div className="subtitle">
            Manage Entra ID app registrations and enterprise apps
          </div>
        </div>
      </div>

      <div className="card">
        <h3>Signed in as</h3>
        {error && <div className="error">{error}</div>}
        {!me && !error && <span className="spinner" />}
        {me && (
          <div className="kv">
            <div className="k">Display name</div>
            <div>{me.displayName}</div>
            <div className="k">UPN</div>
            <div className="mono">{me.userPrincipalName}</div>
            <div className="k">Object ID</div>
            <div className="mono">{me.id}</div>
          </div>
        )}
      </div>

      <div className="card">
        <h3>What you can do</h3>
        <ul className="muted" style={{ margin: 0, paddingLeft: 18 }}>
          <li>
            Browse <strong>app registrations</strong> and view metadata.
          </li>
          <li>
            Browse <strong>enterprise apps</strong> (service principals).
          </li>
          <li>
            Add / remove / grant <strong>API permissions</strong> on enterprise
            apps — both application (app-only) and delegated.
          </li>
        </ul>
      </div>
    </>
  );
}
