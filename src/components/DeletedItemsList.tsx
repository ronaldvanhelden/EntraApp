import { useEffect, useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import {
  daysUntilPurge,
  DELETED_RETENTION_DAYS,
  listDeletedApplications,
  listDeletedServicePrincipals,
  permanentlyDeleteItem,
  restoreDeletedItem,
} from '../graph/deletedItems';
import type { Application, ServicePrincipal } from '../graph/types';
import { AppIcon } from './AppIcon';
import { useConfirm } from './ConfirmDialog';

type Item = Application | ServicePrincipal;
type Kind = 'application' | 'servicePrincipal';

interface Props {
  kind: Kind;
  // Free-text filter from the page's search box (name or appId).
  search: string;
}

const COPY: Record<Kind, { noun: string; route: string; other: string }> = {
  application: {
    noun: 'app registration',
    route: '/applications',
    other:
      'Its enterprise app was deleted together with it and has to be restored separately under Enterprise apps → Deleted.',
  },
  servicePrincipal: {
    noun: 'enterprise app',
    route: '/enterprise-apps',
    other: 'Users and groups assignments and permission grants come back with it.',
  },
};

// Recycle-bin view for app registrations and enterprise apps: what was
// soft-deleted, when it will be purged, and the two things you can do with it.
export function DeletedItemsList({ kind, search }: Props) {
  const token = useGraphToken();
  const { confirm, confirmDialog } = useConfirm();
  const [items, setItems] = useState<Item[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busyId, setBusyId] = useState<string | null>(null);
  const [notice, setNotice] = useState<{ id: string; name: string } | null>(null);
  const copy = COPY[kind];

  useEffect(() => {
    let cancelled = false;
    setItems(null);
    setError(null);
    setNotice(null);
    (kind === 'application' ? listDeletedApplications(token) : listDeletedServicePrincipals(token))
      .then((list) => {
        if (cancelled) return;
        // Most recently deleted first — that's what people are looking for.
        setItems(
          [...list].sort((a, b) =>
            (b.deletedDateTime ?? '').localeCompare(a.deletedDateTime ?? ''),
          ),
        );
      })
      .catch((e: unknown) => {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e));
      });
    return () => {
      cancelled = true;
    };
  }, [token, kind]);

  const visible = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!items || !q) return items ?? [];
    return items.filter(
      (i) =>
        (i.displayName ?? '').toLowerCase().includes(q) ||
        (i.appId ?? '').toLowerCase().includes(q),
    );
  }, [items, search]);

  const drop = (id: string) => setItems((prev) => (prev ? prev.filter((i) => i.id !== id) : prev));

  const restore = async (item: Item) => {
    setBusyId(item.id);
    setError(null);
    setNotice(null);
    try {
      await restoreDeletedItem(token, item.id);
      drop(item.id);
      setNotice({ id: item.id, name: item.displayName });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusyId(null);
    }
  };

  const purge = async (item: Item) => {
    const ok = await confirm({
      title: `Permanently delete ${item.displayName}?`,
      message: (
        <>
          <p style={{ margin: '0 0 8px', fontSize: 14 }}>
            This removes the {copy.noun} from the recycle bin immediately instead of after{' '}
            {DELETED_RETENTION_DAYS} days. <strong>It can no longer be restored</strong>
            {kind === 'application' && ', and its secrets and certificates are gone for good'}.
          </p>
          <p className="mono muted" style={{ margin: 0, fontSize: 12 }}>
            {item.appId}
          </p>
        </>
      ),
      confirmLabel: 'Delete permanently',
    });
    if (!ok) return;
    setBusyId(item.id);
    setError(null);
    setNotice(null);
    try {
      await permanentlyDeleteItem(token, item.id);
      drop(item.id);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusyId(null);
    }
  };

  return (
    <>
      <div className="muted" style={{ fontSize: 12, marginBottom: 12 }}>
        Deleted {copy.noun}s stay in the recycle bin for {DELETED_RETENTION_DAYS} days and can be
        restored until then. {copy.other}
      </div>

      {error && <div className="card error">{error}</div>}
      {notice && (
        <div className="card" style={{ borderColor: 'var(--success)' }}>
          Restored <strong>{notice.name}</strong>.{' '}
          <Link to={`${copy.route}/${notice.id}`}>Open it →</Link>
        </div>
      )}
      {!items && !error && <span className="spinner" />}

      {items && (
        <div className="card" style={{ padding: 0 }}>
          <table className="table">
            <thead>
              <tr>
                <th>Display name</th>
                <th>Application ID</th>
                <th>Deleted</th>
                <th>Purged in</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {visible.map((i) => {
                const days = daysUntilPurge(i.deletedDateTime);
                const busy = busyId === i.id;
                return (
                  <tr key={i.id} style={{ cursor: 'default' }}>
                    <td>
                      <div className="row" style={{ gap: 10, alignItems: 'center' }}>
                        <AppIcon
                          id={i.appId || i.id}
                          logoUrl={i.info?.logoUrl}
                          title={i.displayName}
                        />
                        <span>{i.displayName}</span>
                      </div>
                    </td>
                    <td className="mono">{i.appId}</td>
                    <td>
                      {i.deletedDateTime ? (
                        new Date(i.deletedDateTime).toLocaleString()
                      ) : (
                        <span className="muted">—</span>
                      )}
                    </td>
                    <td>
                      {days === undefined ? (
                        <span className="muted">—</span>
                      ) : (
                        <span
                          className={`badge ${days <= 3 ? 'expired' : days <= 7 ? 'pending' : ''}`}
                        >
                          {days} {days === 1 ? 'day' : 'days'}
                        </span>
                      )}
                    </td>
                    <td className="row-actions">
                      <div className="row" style={{ gap: 6, justifyContent: 'flex-end' }}>
                        <button className="primary" disabled={busy} onClick={() => void restore(i)}>
                          {busy ? '…' : 'Restore'}
                        </button>
                        <button className="danger" disabled={busy} onClick={() => void purge(i)}>
                          Delete permanently
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
              {visible.length === 0 && (
                <tr>
                  <td colSpan={5}>
                    <div className="empty">
                      {search.trim()
                        ? `No deleted ${copy.noun}s match your search.`
                        : `No deleted ${copy.noun}s in the recycle bin.`}
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}
      {confirmDialog}
    </>
  );
}
