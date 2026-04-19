import { useEffect, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import {
  deleteServicePrincipal,
  getServicePrincipal,
  updateServicePrincipal,
} from '../graph/servicePrincipals';
import type { ServicePrincipal } from '../graph/types';
import { PermissionsManager } from '../components/PermissionsManager';
import { AssignedUsersManager } from '../components/AssignedUsersManager';
import { Modal } from '../components/Modal';

type Tab = 'overview' | 'permissions' | 'assignments';

export function EnterpriseAppDetail() {
  const { id = '' } = useParams();
  const token = useGraphToken();
  const nav = useNavigate();
  const [sp, setSp] = useState<ServicePrincipal | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [tab, setTab] = useState<Tab>('permissions');
  const [togglingEnabled, setTogglingEnabled] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [deleting, setDeleting] = useState(false);

  useEffect(() => {
    setSp(null);
    setError(null);
    getServicePrincipal(token, id)
      .then(setSp)
      .catch((e) => setError(e.message));
  }, [token, id]);

  const toggleEnabled = async () => {
    if (!sp) return;
    setTogglingEnabled(true);
    setError(null);
    try {
      const next = !sp.accountEnabled;
      await updateServicePrincipal(token, sp.id, { accountEnabled: next });
      setSp({ ...sp, accountEnabled: next });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setTogglingEnabled(false);
    }
  };

  const doDelete = async () => {
    if (!sp) return;
    setDeleting(true);
    setError(null);
    try {
      await deleteServicePrincipal(token, sp.id);
      nav('/enterprise-apps');
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setDeleting(false);
    }
  };

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
          className={tab === 'assignments' ? 'active' : ''}
          onClick={() => setTab('assignments')}
        >
          Users and groups
        </button>
        <button
          className={tab === 'overview' ? 'active' : ''}
          onClick={() => setTab('overview')}
        >
          Overview
        </button>
      </div>

      {tab === 'overview' && (
        <>
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
              <div className="row">
                {sp.accountEnabled ? (
                  <span className="badge granted">Enabled</span>
                ) : (
                  <span className="badge">Disabled</span>
                )}
                <button
                  onClick={toggleEnabled}
                  disabled={togglingEnabled}
                  style={{ marginLeft: 8 }}
                >
                  {togglingEnabled
                    ? '…'
                    : sp.accountEnabled
                      ? 'Disable'
                      : 'Enable'}
                </button>
              </div>
              <div className="k">Tags</div>
              <div className="muted">{sp.tags?.join(', ') || '—'}</div>
            </div>
          </div>

          <div className="card">
            <h3>Danger zone</h3>
            <div className="row" style={{ justifyContent: 'space-between' }}>
              <div className="muted" style={{ fontSize: 13 }}>
                Deletes the service principal. Users lose access immediately;
                the underlying app registration (if any) is not affected.
              </div>
              <button
                className="danger"
                onClick={() => setConfirmDelete(true)}
              >
                Delete enterprise app
              </button>
            </div>
          </div>
        </>
      )}

      {tab === 'permissions' && <PermissionsManager clientSp={sp} />}
      {tab === 'assignments' && <AssignedUsersManager sp={sp} />}

      {confirmDelete && (
        <Modal
          title="Delete enterprise app"
          onClose={() => !deleting && setConfirmDelete(false)}
          footer={
            <>
              <button
                disabled={deleting}
                onClick={() => setConfirmDelete(false)}
              >
                Cancel
              </button>
              <button className="danger" disabled={deleting} onClick={doDelete}>
                {deleting ? 'Deleting…' : 'Delete'}
              </button>
            </>
          }
        >
          <p>
            Delete <strong>{sp.displayName}</strong>? This removes the service
            principal and all its role assignments and grants. This cannot be
            undone.
          </p>
          <p className="mono muted" style={{ fontSize: 12 }}>
            {sp.id}
          </p>
        </Modal>
      )}
    </>
  );
}
