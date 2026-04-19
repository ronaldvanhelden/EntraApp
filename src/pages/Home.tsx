import { useEffect, useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { getMe, type MeResponse } from '../graph/me';
import {
  listRecentDirectoryAudits,
  type DirectoryAuditEvent,
} from '../graph/directoryAudits';

// Activities the Overview card surfaces. Names must match Microsoft Graph's
// activityDisplayName values exactly — they're case-sensitive strings.
const AUDIT_ACTIVITIES = [
  // App registration lifecycle
  'Add application',
  'Update application',
  'Delete application',
  'Hard Delete application',
  // Service principal lifecycle
  'Add service principal',
  'Update service principal',
  'Delete service principal',
  'Hard Delete service principal',
  // Credential rotation
  'Update application – Certificates and secrets management',
  // Consent & permission grants
  'Consent to application',
  'Add delegated permission grant',
  'Add app role assignment to service principal',
  'Add app role assignment grant to user',
];

export function Home() {
  const token = useGraphToken();
  const [me, setMe] = useState<MeResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [audits, setAudits] = useState<DirectoryAuditEvent[] | 'loading'>(
    'loading',
  );

  useEffect(() => {
    getMe(token)
      .then(setMe)
      .catch((e) => setError(e.message));
  }, [token]);

  useEffect(() => {
    setAudits('loading');
    listRecentDirectoryAudits(token, AUDIT_ACTIVITIES, 25)
      .then(setAudits)
      .catch(() => setAudits([]));
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

      <RecentActivityCard audits={audits} />

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

function RecentActivityCard({
  audits,
}: {
  audits: DirectoryAuditEvent[] | 'loading';
}) {
  return (
    <div className="card">
      <h3>Recent activity (last 30 days)</h3>
      <p className="muted" style={{ marginTop: 0, fontSize: 13 }}>
        App registration and enterprise app changes from the directory audit
        log. Requires <span className="mono">AuditLog.Read.All</span>.
      </p>
      {audits === 'loading' ? (
        <span className="spinner" />
      ) : audits.length === 0 ? (
        <div className="muted" style={{ fontSize: 13 }}>
          No recent directory changes, or audit logs not accessible.
        </div>
      ) : (
        <table className="table">
          <thead>
            <tr>
              <th>When</th>
              <th>Activity</th>
              <th>Target</th>
              <th>By</th>
              <th>Result</th>
            </tr>
          </thead>
          <tbody>
            {audits.map((e) => (
              <AuditRow key={e.id} event={e} />
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

function AuditRow({ event }: { event: DirectoryAuditEvent }) {
  const when = event.activityDateTime
    ? new Date(event.activityDateTime).toLocaleString()
    : '—';
  const who = useMemo(() => {
    const user = event.initiatedBy?.user;
    if (user?.displayName || user?.userPrincipalName) {
      return (
        <>
          <div>{user.displayName ?? user.userPrincipalName}</div>
          {user.userPrincipalName && user.displayName && (
            <div className="muted" style={{ fontSize: 12 }}>
              {user.userPrincipalName}
            </div>
          )}
        </>
      );
    }
    const app = event.initiatedBy?.app;
    if (app?.displayName || app?.appId) {
      return (
        <div>
          {app.displayName ?? app.appId}
          <span className="muted" style={{ fontSize: 12 }}> (app)</span>
        </div>
      );
    }
    return <span className="muted">—</span>;
  }, [event]);

  const target = useMemo(() => {
    const primary = event.targetResources?.[0];
    if (!primary) return <span className="muted">—</span>;
    const label = primary.displayName ?? primary.id ?? '—';
    const isDelete = /delete/i.test(event.activityDisplayName ?? '');
    const linkTo = !isDelete && primary.id ? hrefFor(primary) : null;
    const body = (
      <>
        <div>{label}</div>
        {primary.type && (
          <div className="muted" style={{ fontSize: 12 }}>
            {primary.type}
          </div>
        )}
      </>
    );
    return linkTo ? <Link to={linkTo}>{body}</Link> : <div>{body}</div>;
  }, [event]);

  return (
    <tr>
      <td style={{ whiteSpace: 'nowrap' }}>{when}</td>
      <td>{event.activityDisplayName ?? '—'}</td>
      <td>{target}</td>
      <td>{who}</td>
      <td>
        <ResultBadge result={event.result} />
      </td>
    </tr>
  );
}

function hrefFor(t: { id?: string; type?: string }): string | null {
  if (!t.id) return null;
  if (t.type === 'Application') return `/applications/${t.id}`;
  if (t.type === 'ServicePrincipal') return `/enterprise-apps/${t.id}`;
  return null;
}

function ResultBadge({ result }: { result?: string }) {
  if (!result) return <span className="muted">—</span>;
  const r = result.toLowerCase();
  if (r === 'success') return <span className="badge granted">success</span>;
  if (r === 'failure' || r === 'timeout')
    return <span className="badge expired">{result}</span>;
  return <span className="badge">{result}</span>;
}
