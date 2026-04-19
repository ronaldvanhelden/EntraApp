import { useCallback, useEffect, useMemo, useState } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import type {
  AppRoleAssignment,
  OAuth2PermissionGrant,
  ServicePrincipal,
} from '../graph/types';
import {
  createAppRoleAssignment,
  createOAuth2Grant,
  deleteAppRoleAssignment,
  deleteOAuth2Grant,
  getServicePrincipal,
  listAppRoleAssignments,
  listOAuth2GrantsForClient,
  updateOAuth2Grant,
} from '../graph/servicePrincipals';
import { AddPermissionModal } from './AddPermissionModal';

interface Props {
  clientSp: ServicePrincipal;
}

interface RowBase {
  key: string;
  resourceId: string;
  resourceName: string;
  name: string;
  description: string;
}
interface AppRow extends RowBase {
  kind: 'app';
  assignmentId: string;
}
interface DelegatedRow extends RowBase {
  kind: 'delegated';
  grantId: string;
  scope: string;
  consentType: 'AllPrincipals' | 'Principal';
}
type Row = AppRow | DelegatedRow;

export function PermissionsManager({ clientSp }: Props) {
  const token = useGraphToken();
  const [assignments, setAssignments] = useState<AppRoleAssignment[] | null>(
    null,
  );
  const [grants, setGrants] = useState<OAuth2PermissionGrant[] | null>(null);
  const [resources, setResources] = useState<Record<string, ServicePrincipal>>(
    {},
  );
  const [error, setError] = useState<string | null>(null);
  const [working, setWorking] = useState<string | null>(null);
  const [showAdd, setShowAdd] = useState(false);

  const reload = useCallback(async () => {
    setError(null);
    setAssignments(null);
    setGrants(null);
    try {
      const [a, g] = await Promise.all([
        listAppRoleAssignments(token, clientSp.id),
        listOAuth2GrantsForClient(token, clientSp.id),
      ]);
      setAssignments(a);
      setGrants(g);

      const ids = new Set<string>();
      a.forEach((x) => ids.add(x.resourceId));
      g.forEach((x) => ids.add(x.resourceId));
      const missing = [...ids].filter((id) => !resources[id]);
      if (missing.length) {
        const fetched = await Promise.all(
          missing.map((id) => getServicePrincipal(token, id).catch(() => null)),
        );
        const next: Record<string, ServicePrincipal> = { ...resources };
        fetched.forEach((sp) => {
          if (sp) next[sp.id] = sp;
        });
        setResources(next);
      }
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [token, clientSp.id]);

  useEffect(() => {
    reload();
  }, [reload]);

  const rows: Row[] = useMemo(() => {
    const out: Row[] = [];
    if (assignments) {
      for (const a of assignments) {
        const res = resources[a.resourceId];
        const role = res?.appRoles?.find((r) => r.id === a.appRoleId);
        out.push({
          kind: 'app',
          key: `app:${a.id}`,
          assignmentId: a.id,
          resourceId: a.resourceId,
          resourceName: res?.displayName ?? a.resourceDisplayName ?? a.resourceId,
          name: role?.value ?? a.appRoleId,
          description: role?.description ?? '',
        });
      }
    }
    if (grants) {
      for (const g of grants) {
        const res = resources[g.resourceId];
        const scopes = (g.scope ?? '')
          .split(/\s+/)
          .map((s) => s.trim())
          .filter(Boolean);
        for (const scope of scopes) {
          const def = res?.oauth2PermissionScopes?.find(
            (s) => s.value === scope,
          );
          out.push({
            kind: 'delegated',
            key: `del:${g.id}:${scope}`,
            grantId: g.id,
            scope,
            consentType: g.consentType,
            resourceId: g.resourceId,
            resourceName: res?.displayName ?? g.resourceId,
            name: scope,
            description: def?.adminConsentDescription ?? '',
          });
        }
      }
    }
    return out.sort(
      (a, b) =>
        a.resourceName.localeCompare(b.resourceName) ||
        a.name.localeCompare(b.name),
    );
  }, [assignments, grants, resources]);

  const revoke = async (row: Row) => {
    setError(null);
    setWorking(row.key);
    try {
      if (row.kind === 'app') {
        await deleteAppRoleAssignment(token, clientSp.id, row.assignmentId);
      } else {
        const grant = grants?.find((g) => g.id === row.grantId);
        if (!grant) throw new Error('Grant not found');
        const remaining = (grant.scope ?? '')
          .split(/\s+/)
          .filter((s) => s && s !== row.scope)
          .join(' ');
        if (remaining) {
          await updateOAuth2Grant(token, grant.id, { scope: remaining });
        } else {
          await deleteOAuth2Grant(token, grant.id);
        }
      }
      await reload();
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setWorking(null);
    }
  };

  const addPermissions = async (payload: {
    resource: ServicePrincipal;
    appRoleIds: string[];
    delegatedScopes: string[];
    consentType: 'AllPrincipals' | 'Principal';
    principalId?: string;
  }) => {
    setError(null);
    try {
      // App-only: one appRoleAssignment per selected role
      for (const roleId of payload.appRoleIds) {
        await createAppRoleAssignment(token, clientSp.id, {
          appRoleId: roleId,
          principalId: clientSp.id,
          resourceId: payload.resource.id,
        });
      }

      // Delegated: merge into existing grant for (clientId,resourceId,consentType)
      if (payload.delegatedScopes.length) {
        const existing = grants?.find(
          (g) =>
            g.resourceId === payload.resource.id &&
            g.consentType === payload.consentType &&
            (payload.consentType === 'AllPrincipals' ||
              g.principalId === payload.principalId),
        );
        if (existing) {
          const merged = Array.from(
            new Set(
              [
                ...(existing.scope ?? '').split(/\s+/).filter(Boolean),
                ...payload.delegatedScopes,
              ].filter(Boolean),
            ),
          ).join(' ');
          await updateOAuth2Grant(token, existing.id, { scope: merged });
        } else {
          await createOAuth2Grant(token, {
            clientId: clientSp.id,
            consentType: payload.consentType,
            principalId:
              payload.consentType === 'Principal'
                ? payload.principalId ?? null
                : null,
            resourceId: payload.resource.id,
            scope: payload.delegatedScopes.join(' '),
          });
        }
      }

      // Keep resource in cache
      setResources((r) => ({ ...r, [payload.resource.id]: payload.resource }));
      setShowAdd(false);
      await reload();
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      throw e;
    }
  };

  const loading = assignments === null || grants === null;

  return (
    <>
      <div className="card">
        <div className="row" style={{ justifyContent: 'space-between' }}>
          <div>
            <h3 style={{ margin: 0 }}>Configured API permissions</h3>
            <div className="muted" style={{ fontSize: 12, marginTop: 4 }}>
              Granted permissions for this enterprise app. Delegated grants use{' '}
              <code>AllPrincipals</code> (admin consent) by default.
            </div>
          </div>
          <button className="primary" onClick={() => setShowAdd(true)}>
            + Add a permission
          </button>
        </div>
      </div>

      {error && <div className="card error">{error}</div>}

      {loading ? (
        <div className="center" style={{ height: 120 }}>
          <span className="spinner" />
        </div>
      ) : rows.length === 0 ? (
        <div className="card empty">
          No API permissions configured. Click <em>Add a permission</em> to
          grant one.
        </div>
      ) : (
        <div className="card" style={{ padding: 0 }}>
          <table className="table">
            <thead>
              <tr>
                <th>API</th>
                <th>Permission</th>
                <th>Type</th>
                <th>Description</th>
                <th style={{ width: 100 }}></th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.key} style={{ cursor: 'default' }}>
                  <td>{r.resourceName}</td>
                  <td className="mono">{r.name}</td>
                  <td>
                    {r.kind === 'app' ? (
                      <span className="badge app">Application</span>
                    ) : (
                      <span className="badge delegated">
                        Delegated
                        {r.consentType === 'Principal' ? ' (user)' : ''}
                      </span>
                    )}
                  </td>
                  <td
                    className="muted"
                    style={{ maxWidth: 420, fontSize: 12 }}
                  >
                    {r.description}
                  </td>
                  <td style={{ textAlign: 'right' }}>
                    <button
                      className="danger"
                      disabled={working === r.key}
                      onClick={() => revoke(r)}
                    >
                      {working === r.key ? '…' : 'Revoke'}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {showAdd && (
        <AddPermissionModal
          onClose={() => setShowAdd(false)}
          onSubmit={addPermissions}
        />
      )}
    </>
  );
}
