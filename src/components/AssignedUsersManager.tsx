import { useCallback, useEffect, useMemo, useState } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import {
  createAppRoleAssignedTo,
  deleteAppRoleAssignedTo,
  listAppRoleAssignedTo,
} from '../graph/servicePrincipals';
import type { AppRole, AppRoleAssignment, ServicePrincipal } from '../graph/types';
import {
  resolveDirectoryObjects,
  type PrincipalRef,
} from '../graph/directoryObjects';
import { Modal } from './Modal';
import { PrincipalPicker } from './PrincipalPicker';

interface Props {
  sp: ServicePrincipal;
}

// Microsoft's "Default Access" role id used when an SP exposes no custom roles
// or when assigning a user without picking a specific app role.
const DEFAULT_ACCESS_ROLE_ID = '00000000-0000-0000-0000-000000000000';

export function AssignedUsersManager({ sp }: Props) {
  const token = useGraphToken();
  const [assignments, setAssignments] = useState<AppRoleAssignment[] | null>(
    null,
  );
  const [principals, setPrincipals] = useState<Record<string, PrincipalRef>>({});
  const [error, setError] = useState<string | null>(null);
  const [working, setWorking] = useState<string | null>(null);
  const [showAdd, setShowAdd] = useState(false);

  const reload = useCallback(async () => {
    setError(null);
    setAssignments(null);
    try {
      const list = await listAppRoleAssignedTo(token, sp.id);
      setAssignments(list);
      const missing = list
        .map((a) => a.principalId)
        .filter((id) => id && !principals[id]);
      if (missing.length) {
        const resolved = await resolveDirectoryObjects(token, missing);
        setPrincipals((prev) => ({ ...prev, ...resolved }));
      }
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [token, sp.id]);

  useEffect(() => {
    reload();
  }, [reload]);

  const roleById = useMemo(() => {
    const m = new Map<string, AppRole>();
    (sp.appRoles ?? []).forEach((r) => m.set(r.id, r));
    return m;
  }, [sp.appRoles]);

  const assignableRoles = useMemo(
    () =>
      (sp.appRoles ?? []).filter(
        (r) =>
          r.isEnabled &&
          (r.allowedMemberTypes?.includes('User') ||
            r.allowedMemberTypes?.includes('Application')),
      ),
    [sp.appRoles],
  );

  const revoke = async (a: AppRoleAssignment) => {
    setError(null);
    setWorking(a.id);
    try {
      await deleteAppRoleAssignedTo(token, sp.id, a.id);
      await reload();
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setWorking(null);
    }
  };

  const add = async (p: PrincipalRef, roleId: string) => {
    await createAppRoleAssignedTo(token, sp.id, {
      appRoleId: roleId,
      principalId: p.id,
      resourceId: sp.id,
    });
    setPrincipals((prev) => ({ ...prev, [p.id]: p }));
    setShowAdd(false);
    await reload();
  };

  const loading = assignments === null;

  return (
    <>
      <div className="card">
        <div className="row" style={{ justifyContent: 'space-between' }}>
          <div>
            <h3 style={{ margin: 0 }}>Users and groups</h3>
            <div className="muted" style={{ fontSize: 12, marginTop: 4 }}>
              Principals (users, groups, service principals) assigned to this
              enterprise app&apos;s roles.
            </div>
          </div>
          <button className="primary" onClick={() => setShowAdd(true)}>
            + Add assignment
          </button>
        </div>
      </div>

      {error && <div className="card error">{error}</div>}

      {loading ? (
        <div className="center" style={{ height: 120 }}>
          <span className="spinner" />
        </div>
      ) : assignments.length === 0 ? (
        <div className="card empty">
          No principals assigned yet. Click <em>Add assignment</em> to grant
          access.
        </div>
      ) : (
        <div className="card" style={{ padding: 0 }}>
          <table className="table">
            <thead>
              <tr>
                <th>Principal</th>
                <th>Type</th>
                <th>Identifier</th>
                <th>Role</th>
                <th>Assigned</th>
                <th style={{ width: 100 }}></th>
              </tr>
            </thead>
            <tbody>
              {assignments.map((a) => {
                const p = principals[a.principalId];
                const role = roleById.get(a.appRoleId);
                const roleName =
                  a.appRoleId === DEFAULT_ACCESS_ROLE_ID
                    ? 'Default access'
                    : role?.displayName ?? role?.value ?? a.appRoleId;
                return (
                  <tr key={a.id} style={{ cursor: 'default' }}>
                    <td>
                      {p?.displayName ??
                        a.principalDisplayName ??
                        a.principalId}
                    </td>
                    <td>
                      <span className="badge">
                        {a.principalType ?? p?.kind ?? '—'}
                      </span>
                    </td>
                    <td className="mono muted" style={{ fontSize: 12 }}>
                      {p?.subtitle ?? a.principalId}
                    </td>
                    <td>
                      <div>{roleName}</div>
                      {role?.value && role.id !== DEFAULT_ACCESS_ROLE_ID && (
                        <div className="mono muted" style={{ fontSize: 11 }}>
                          {role.value}
                        </div>
                      )}
                    </td>
                    <td className="muted" style={{ fontSize: 12 }}>
                      {a.createdDateTime
                        ? new Date(a.createdDateTime).toLocaleDateString()
                        : '—'}
                    </td>
                    <td style={{ textAlign: 'right' }}>
                      <button
                        className="danger"
                        disabled={working === a.id}
                        onClick={() => revoke(a)}
                      >
                        {working === a.id ? '…' : 'Remove'}
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {showAdd && (
        <AddAssignmentModal
          sp={sp}
          assignableRoles={assignableRoles}
          onClose={() => setShowAdd(false)}
          onSubmit={add}
        />
      )}
    </>
  );
}

function AddAssignmentModal({
  sp,
  assignableRoles,
  onClose,
  onSubmit,
}: {
  sp: ServicePrincipal;
  assignableRoles: AppRole[];
  onClose: () => void;
  onSubmit: (p: PrincipalRef, roleId: string) => Promise<void>;
}) {
  const [principal, setPrincipal] = useState<PrincipalRef | null>(null);
  const [roleId, setRoleId] = useState<string>(
    assignableRoles[0]?.id ?? DEFAULT_ACCESS_ROLE_ID,
  );
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const submit = async () => {
    if (!principal) return;
    setSubmitting(true);
    setError(null);
    try {
      await onSubmit(principal, roleId);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <Modal
      title={`Assign principal to ${sp.displayName}`}
      onClose={onClose}
      footer={
        <>
          <button onClick={onClose}>Cancel</button>
          <button
            className="primary"
            disabled={!principal || submitting}
            onClick={submit}
          >
            {submitting ? 'Assigning…' : 'Assign'}
          </button>
        </>
      }
    >
      <label className="field">
        <span>Principal (user, group, or service principal)</span>
        <PrincipalPicker
          kinds={['user', 'group', 'sp']}
          selected={principal}
          onChange={setPrincipal}
          autoFocus
        />
      </label>

      <label className="field">
        <span>Role</span>
        <select value={roleId} onChange={(e) => setRoleId(e.target.value)}>
          <option value={DEFAULT_ACCESS_ROLE_ID}>
            Default access (no specific role)
          </option>
          {assignableRoles.map((r) => (
            <option key={r.id} value={r.id}>
              {r.displayName} ({r.value})
            </option>
          ))}
        </select>
      </label>

      {assignableRoles.length === 0 && (
        <div className="muted" style={{ fontSize: 12, marginTop: -4 }}>
          This app exposes no assignable roles. Only &quot;Default access&quot;
          is available.
        </div>
      )}

      {error && <p className="error">{error}</p>}
    </Modal>
  );
}
