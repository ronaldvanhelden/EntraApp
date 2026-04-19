import { useEffect, useState } from 'react';
import { useParams } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { getServicePrincipal } from '../graph/servicePrincipals';
import type { ServicePrincipal } from '../graph/types';
import { PermissionsManager } from '../components/PermissionsManager';

type Tab = 'overview' | 'permissions';

export function EnterpriseAppDetail() {
  const { id = '' } = useParams();
  const token = useGraphToken();
  const [sp, setSp] = useState<ServicePrincipal | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [tab, setTab] = useState<Tab>('permissions');

  useEffect(() => {
    setSp(null);
    setError(null);
    getServicePrincipal(token, id)
      .then(setSp)
      .catch((e) => setError(e.message));
  }, [token, id]);

  if (error) return <div className="card error">{error}</div>;
  if (!sp)
    return (
      <div className="center">
        <span className="spinner" />
      </div>
    );

  return (
    <>
      <div className="page-header">
        <div>
          <h1>{sp.displayName}</h1>
          <div className="subtitle mono">{sp.appId}</div>
        </div>
      </div>

      <div className="tabs">
        <button
          className={tab === 'permissions' ? 'active' : ''}
          onClick={() => setTab('permissions')}
        >
          API permissions
        </button>
        <button
          className={tab === 'overview' ? 'active' : ''}
          onClick={() => setTab('overview')}
        >
          Overview
        </button>
      </div>

      {tab === 'overview' && (
        <div className="card">
          <h3>Details</h3>
          <div className="kv">
            <div className="k">Object ID</div>
            <div className="mono">{sp.id}</div>
            <div className="k">App ID</div>
            <div className="mono">{sp.appId}</div>
            <div className="k">Type</div>
            <div>{sp.servicePrincipalType ?? '—'}</div>
            <div className="k">Publisher</div>
            <div>{sp.publisherName ?? '—'}</div>
            <div className="k">Enabled</div>
            <div>{sp.accountEnabled ? 'Yes' : 'No'}</div>
            <div className="k">Tags</div>
            <div className="muted">{sp.tags?.join(', ') || '—'}</div>
          </div>
        </div>
      )}

      {tab === 'permissions' && <PermissionsManager clientSp={sp} />}
    </>
  );
}
