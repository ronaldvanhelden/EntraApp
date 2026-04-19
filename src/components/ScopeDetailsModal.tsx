import { useMemo } from 'react';
import { Modal } from './Modal';
import {
  privilegeLabel,
  type ScopeKind,
  type ScopeMeta,
} from '../graph/scopeCatalog';

interface Props {
  scope: string;
  resourceName?: string;
  kind: ScopeKind;
  catalog: Record<string, ScopeMeta> | null;
  onClose: () => void;
  // Description from the resource SP (oauth2PermissionScope or appRole). We
  // prefer this when present because it reflects what the API actually
  // exposes in this tenant; fall back to the public catalog otherwise.
  fallbackTitle?: string;
  fallbackDescription?: string;
}

export function ScopeDetailsModal({
  scope,
  resourceName,
  kind,
  catalog,
  onClose,
  fallbackTitle,
  fallbackDescription,
}: Props) {
  const meta = catalog?.[scope.toLowerCase()] ?? null;
  const scheme = kind === 'delegated' ? meta?.delegated : meta?.application;
  const level = scheme?.privilegeLevel;
  const { label: privLabel, tone } = privilegeLabel(level);

  const paths = useMemo(() => {
    if (!meta?.pathSets) return [] as { methods: string[]; paths: string[] }[];
    return meta.pathSets
      .filter((p) => p.kinds.includes(kind))
      .map((p) => ({ methods: p.methods, paths: p.paths }));
  }, [meta, kind]);

  return (
    <Modal title={scope} onClose={onClose}>
      <div className="kv">
        <div className="k">Resource</div>
        <div>{resourceName ?? 'Microsoft Graph'}</div>
        <div className="k">Type</div>
        <div>
          {kind === 'delegated' ? 'Delegated' : 'Application (app-only)'}
        </div>
        <div className="k">Privilege level</div>
        <div>
          {level === undefined ? (
            <span className="muted">Not in the public catalog</span>
          ) : (
            <span className={`priv-badge priv-${tone}`}>{privLabel}</span>
          )}
        </div>
        {scheme && (
          <>
            <div className="k">Admin consent</div>
            <div>{scheme.requiresAdminConsent ? 'Required' : 'Not required'}</div>
          </>
        )}
        {(fallbackTitle || scheme?.adminDescription) && (
          <>
            <div className="k">Description</div>
            <div>
              {fallbackTitle && (
                <div style={{ fontWeight: 500, marginBottom: 4 }}>
                  {fallbackTitle}
                </div>
              )}
              <div className="muted" style={{ fontSize: 13 }}>
                {scheme?.adminDescription ?? fallbackDescription ?? ''}
              </div>
              {scheme?.userDescription &&
                scheme.userDescription !== scheme.adminDescription && (
                  <div
                    className="muted"
                    style={{ fontSize: 12, marginTop: 6, fontStyle: 'italic' }}
                  >
                    As shown to end users: {scheme.userDescription}
                  </div>
                )}
            </div>
          </>
        )}
      </div>

      <h4 style={{ marginTop: 20, marginBottom: 8 }}>
        API paths unlocked ({paths.reduce((n, p) => n + p.paths.length, 0)})
      </h4>
      {!catalog ? (
        <div className="muted">
          <span className="spinner" /> Loading Microsoft Graph permissions
          catalog…
        </div>
      ) : paths.length === 0 ? (
        <div className="muted" style={{ fontSize: 13 }}>
          No path information in the public catalog for this scope.
          {meta === null &&
            ' (Scope may belong to an API other than Microsoft Graph.)'}
        </div>
      ) : (
        <table className="table">
          <thead>
            <tr>
              <th style={{ width: 140 }}>Methods</th>
              <th>Path</th>
            </tr>
          </thead>
          <tbody>
            {paths.flatMap((group) =>
              group.paths.map((p) => (
                <tr key={`${group.methods.join(',')}:${p}`}>
                  <td>
                    <div className="row" style={{ gap: 4, flexWrap: 'wrap' }}>
                      {group.methods.map((m) => (
                        <span key={m} className={`method-pill m-${m}`}>
                          {m}
                        </span>
                      ))}
                    </div>
                  </td>
                  <td className="mono" style={{ fontSize: 12 }}>
                    {p}
                  </td>
                </tr>
              )),
            )}
          </tbody>
        </table>
      )}
    </Modal>
  );
}
