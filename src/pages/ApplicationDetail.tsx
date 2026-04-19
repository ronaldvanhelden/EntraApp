import { useEffect, useState } from 'react';
import { Link, useParams } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { getApplication } from '../graph/applications';
import { getServicePrincipalByAppId } from '../graph/servicePrincipals';
import type { Application, ServicePrincipal } from '../graph/types';

export function ApplicationDetail() {
  const { id = '' } = useParams();
  const token = useGraphToken();
  const [app, setApp] = useState<Application | null>(null);
  const [sp, setSp] = useState<ServicePrincipal | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setApp(null);
    setSp(null);
    setError(null);
    getApplication(token, id)
      .then(async (a) => {
        setApp(a);
        try {
          const maybe = await getServicePrincipalByAppId(token, a.appId);
          setSp(maybe ?? null);
        } catch {
          /* no SP yet — fine */
        }
      })
      .catch((e) => setError(e.message));
  }, [token, id]);

  if (error) return <div className="card error">{error}</div>;
  if (!app)
    return (
      <div className="center">
        <span className="spinner" />
      </div>
    );

  return (
    <>
      <div className="page-header">
        <div>
          <h1>{app.displayName}</h1>
          <div className="subtitle mono">{app.appId}</div>
        </div>
        {sp && (
          <Link
            to={`/enterprise-apps/${sp.id}`}
            className="primary"
            style={{ textDecoration: 'none' }}
          >
            <button className="primary">Manage enterprise app →</button>
          </Link>
        )}
      </div>

      <div className="card">
        <h3>Details</h3>
        <div className="kv">
          <div className="k">Object ID</div>
          <div className="mono">{app.id}</div>
          <div className="k">Application (client) ID</div>
          <div className="mono">{app.appId}</div>
          <div className="k">Sign-in audience</div>
          <div>{app.signInAudience}</div>
          <div className="k">Publisher domain</div>
          <div>{app.publisherDomain ?? '—'}</div>
          <div className="k">Created</div>
          <div>
            {app.createdDateTime
              ? new Date(app.createdDateTime).toLocaleString()
              : '—'}
          </div>
        </div>
      </div>

      <div className="card">
        <h3>Required resource access (manifest)</h3>
        {!app.requiredResourceAccess?.length ? (
          <div className="muted">No required resource access declared.</div>
        ) : (
          <table className="table">
            <thead>
              <tr>
                <th>Resource AppId</th>
                <th>Permissions declared</th>
              </tr>
            </thead>
            <tbody>
              {app.requiredResourceAccess.map((r) => (
                <tr key={r.resourceAppId}>
                  <td className="mono">{r.resourceAppId}</td>
                  <td>{r.resourceAccess.length}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
        <p className="muted" style={{ marginTop: 12, marginBottom: 0 }}>
          To add or grant permissions, manage the enterprise app (service
          principal) that represents this application in your tenant.
        </p>
      </div>
    </>
  );
}
