import { useEffect, useMemo, useState } from 'react';
import { Link, useSearchParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import { PIM_ACTIVATE_SCOPES, PIM_ASSIGN_SCOPES, ROLE_READ_SCOPES } from '../auth/config';
import { useGraphToken } from '../auth/useGraphToken';
import type { GraphTokenFn } from '../graph/client';
import {
  deactivateRole,
  listMyRoles,
  listRoleSchedules,
  pimError,
  removeRole,
  type RoleSchedule,
} from '../graph/pim';
import { ConsentGate } from '../components/ConsentGate';
import { useConfirm } from '../components/ConfirmDialog';
import { ActivateRole, AssignRole } from '../components/PimDialogs';
import { RemoveButton } from '../components/RemoveButton';
import { SortTh, useSort } from '../components/SortableTable';

// Privileged Identity Management for Entra roles. "My roles" is where the
// signed-in admin turns an eligible role on; "All assignments" shows who is
// eligible for or holds which role, and is where roles are handed out.

const dash = <span className="muted">—</span>;
const day = (iso?: string | null) => (iso ? new Date(iso).toLocaleDateString() : null);
const time = (iso?: string | null) => (iso ? new Date(iso).toLocaleString() : null);
const scope = (s: string) => (s === '/' ? 'Whole directory' : <span className="mono">{s}</span>);
// PIM applies a change a moment after accepting it; reload once it has.
const SETTLE_MS = 2000;

const YEAR_MS = 365 * 864e5;
const overAYear = (iso?: string | null) =>
  Boolean(iso) && Date.now() - new Date(iso as string).getTime() > YEAR_MS;

// How one of my roles reached me: assigned to me, or to a group I am in.
function through(role: RoleSchedule) {
  if (role.via === undefined) return <span className="muted">Directly</span>;
  if (role.via === null)
    return (
      <span className="muted" title="Through a group that couldn't be identified">
        A group
      </span>
    );
  return <Link to={`/groups/${role.via.id}`}>{role.via.name}</Link>;
}

// When an eligibility was last used. Marked when that is over a year ago —
// or was never seen while the eligibility itself is older than a year — as
// a hint that it may no longer be needed.
function LastActivated({ row }: { row: RoleSchedule }) {
  if (row.kind !== 'eligible') return dash;
  if (row.principalKind && row.principalKind !== 'user')
    return (
      <span className="muted" title="A group's eligibility is activated by its members">
        By members
      </span>
    );
  if (row.lastActivated === undefined)
    return (
      <span className="muted" title="The activation history couldn't be read">
        Unknown
      </span>
    );
  if (row.lastActivated === null)
    return overAYear(row.startDateTime) ? (
      <span
        className="badge expired"
        title="Eligible for over a year without an activation in PIM's history"
      >
        Not in over a year
      </span>
    ) : (
      <span className="muted" title="No activation in PIM's history">
        Not yet
      </span>
    );
  return overAYear(row.lastActivated) ? (
    <span className="badge expired" title="Last activated over a year ago">
      {day(row.lastActivated)}
    </span>
  ) : (
    <>{day(row.lastActivated)}</>
  );
}

const PRINCIPAL_PATH = { user: '/users', group: '/groups', sp: '/enterprise-apps' } as const;

export function Pim() {
  const [params, setParams] = useSearchParams();
  const tab = params.get('tab') === 'all' ? 'all' : 'mine';
  const [assigning, setAssigning] = useState(false);
  // Bumped after a change so the lists load again.
  const [version, setVersion] = useState(0);
  const refresh = () => window.setTimeout(() => setVersion((v) => v + 1), SETTLE_MS);

  return (
    <>
      <div className="page-header">
        <div>
          <h1>Privileged roles</h1>
          <div className="subtitle">
            Activate your own admin roles and manage who is eligible for which role (PIM).
          </div>
        </div>
        <button className="primary" onClick={() => setAssigning(true)}>
          + Assign role
        </button>
      </div>
      <div className="card related-card">
        <div className="tabs">
          <button className={tab === 'mine' ? 'active' : ''} onClick={() => setParams({})}>
            My roles
          </button>
          <button
            className={tab === 'all' ? 'active' : ''}
            onClick={() => setParams({ tab: 'all' })}
          >
            All assignments
          </button>
        </div>
        <ConsentGate scopes={ROLE_READ_SCOPES} what="roles and their assignments">
          {(token) =>
            tab === 'mine' ? (
              <MyRoles token={token} version={version} onChanged={refresh} />
            ) : (
              <AllAssignments token={token} version={version} onChanged={refresh} />
            )
          }
        </ConsentGate>
      </div>
      {assigning && (
        <AssignRole
          onClose={() => setAssigning(false)}
          onDone={() => {
            setAssigning(false);
            setParams({ tab: 'all' });
            refresh();
          }}
        />
      )}
    </>
  );
}

function useRoles(load: () => Promise<RoleSchedule[]>, deps: unknown[]) {
  const [rows, setRows] = useState<RoleSchedule[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  useEffect(() => {
    let cancelled = false;
    setError(null);
    load()
      .then((r) => !cancelled && setRows(r))
      .catch((e: unknown) => !cancelled && setError(pimError(e)));
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);
  return { rows, error, setError };
}

interface PanelProps {
  token: GraphTokenFn;
  version: number;
  onChanged: () => void;
}

function MyRoles({ token, version, onChanged }: PanelProps) {
  const { instance, accounts } = useMsal();
  const userId = (instance.getActiveAccount() ?? accounts[0])?.localAccountId ?? '';
  const activateToken = useGraphToken(PIM_ACTIVATE_SCOPES);
  const { confirm, confirmDialog } = useConfirm();
  const { rows, error, setError } = useRoles(() => listMyRoles(token), [token, version]);
  const [activating, setActivating] = useState<RoleSchedule | null>(null);
  const [busy, setBusy] = useState<string | null>(null);

  const same = (a: RoleSchedule, b: RoleSchedule) =>
    a.roleDefinitionId === b.roleDefinitionId && a.directoryScopeId === b.directoryScopeId;
  const active = (rows ?? []).filter((r) => r.kind === 'active');
  const eligible = (rows ?? []).filter((r) => r.kind === 'eligible');
  // Roles held outright, without an eligibility behind them.
  const assigned = active.filter((a) => a.assignmentType !== 'Activated');

  const deactivate = async (role: RoleSchedule) => {
    const ok = await confirm({
      title: `Deactivate ${role.roleName}?`,
      message: 'You stay eligible and can activate the role again when you need it.',
      confirmLabel: 'Deactivate',
    });
    if (!ok) return;
    setBusy(role.key);
    setError(null);
    try {
      await deactivateRole(activateToken, {
        principalId: userId,
        roleDefinitionId: role.roleDefinitionId,
        directoryScopeId: role.directoryScopeId,
      });
      onChanged();
    } catch (e: unknown) {
      setError(pimError(e));
    } finally {
      setBusy(null);
    }
  };

  if (!rows && !error)
    return (
      <div className="related-note">
        <span className="spinner" />
      </div>
    );
  return (
    <>
      {error && <div className="error related-note">{error}</div>}
      {rows && eligible.length === 0 && assigned.length === 0 && (
        <div className="muted related-note">
          You have no admin roles in this directory, eligible or active.
        </div>
      )}
      {rows && (eligible.length > 0 || assigned.length > 0) && (
        <table className="table">
          <thead>
            <tr>
              <th>Role</th>
              <th>Through</th>
              <th>Scope</th>
              <th>Status</th>
              <th>Eligible until</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {eligible.map((role) => {
              const on = active.find((a) => a.assignmentType === 'Activated' && same(a, role));
              return (
                <tr key={role.key} className="static">
                  <td title={role.roleDescription}>{role.roleName}</td>
                  <td>{through(role)}</td>
                  <td>{scope(role.directoryScopeId)}</td>
                  <td>
                    {on ? (
                      <span className="badge granted">
                        Active{on.endDateTime ? ` until ${time(on.endDateTime)}` : ''}
                      </span>
                    ) : (
                      <span className="badge pending">Eligible</span>
                    )}
                  </td>
                  <td>{day(role.endDateTime) ?? <span className="muted">No end date</span>}</td>
                  <td className="row-actions">
                    {on ? (
                      <button
                        className="danger"
                        disabled={busy === role.key}
                        onClick={() => void deactivate(role)}
                      >
                        {busy === role.key ? 'Deactivating…' : 'Deactivate'}
                      </button>
                    ) : (
                      <button className="primary" onClick={() => setActivating(role)}>
                        Activate
                      </button>
                    )}
                  </td>
                </tr>
              );
            })}
            {assigned.map((role) => (
              <tr key={role.key} className="static">
                <td title={role.roleDescription}>{role.roleName}</td>
                <td>{through(role)}</td>
                <td>{scope(role.directoryScopeId)}</td>
                <td>
                  <span className="badge granted">Always active</span>
                </td>
                <td>{day(role.endDateTime) ?? dash}</td>
                <td></td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      {activating && (
        <ActivateRole
          role={activating}
          userId={userId}
          onClose={() => setActivating(null)}
          onDone={onChanged}
        />
      )}
      {confirmDialog}
    </>
  );
}

function AllAssignments({ token, version, onChanged }: PanelProps) {
  const assignToken = useGraphToken(PIM_ASSIGN_SCOPES);
  const { confirm, confirmDialog } = useConfirm();
  const { rows, error, setError } = useRoles(() => listRoleSchedules(token), [token, version]);
  const [search, setSearch] = useState('');
  const [kind, setKind] = useState<'all' | RoleSchedule['kind']>('all');
  const [busy, setBusy] = useState<string | null>(null);

  const shown = useMemo(() => {
    const q = search.trim().toLowerCase();
    return (rows ?? []).filter(
      (r) =>
        (kind === 'all' || r.kind === kind) &&
        (!q || r.roleName.toLowerCase().includes(q) || r.principalName.toLowerCase().includes(q)),
    );
  }, [rows, search, kind]);
  const { sorted, sort } = useSort(shown, {
    who: (r) => r.principalName,
    through: (r) => r.via?.name,
    role: (r) => r.roleName,
    type: (r) =>
      r.kind === 'eligible'
        ? 'Eligible'
        : r.assignmentType === 'Activated'
          ? 'Activated'
          : 'Active',
    scope: (r) => r.directoryScopeId,
    ends: (r) => r.endDateTime,
    activated: (r) => (r.kind === 'eligible' ? r.lastActivated : null),
  });

  const remove = async (row: RoleSchedule) => {
    const ok = await confirm({
      title: `Remove ${row.roleName} from ${row.principalName}?`,
      message:
        row.kind === 'eligible'
          ? `${row.principalName} can no longer activate this role. An activation that is running now ends as well.`
          : `${row.principalName} loses this role straight away.`,
      confirmLabel: 'Remove role',
    });
    if (!ok) return;
    setBusy(row.key);
    setError(null);
    try {
      await removeRole(assignToken, row);
      onChanged();
    } catch (e: unknown) {
      setError(pimError(e));
    } finally {
      setBusy(null);
    }
  };

  return (
    <>
      <div className="row related-note" style={{ gap: 8 }}>
        <input
          className="search"
          style={{ maxWidth: 320 }}
          placeholder="Filter by role or name…"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
        />
        <select
          style={{ width: 'auto' }}
          value={kind}
          onChange={(e) => setKind(e.target.value as typeof kind)}
        >
          <option value="all">Eligible and active</option>
          <option value="eligible">Eligible only</option>
          <option value="active">Active only</option>
        </select>
        <div className="grow" />
        {rows && <span className="muted">{shown.length} shown</span>}
      </div>
      {error && <div className="error related-note">{error}</div>}
      {!rows && !error && (
        <div className="related-note">
          <span className="spinner" />
        </div>
      )}
      {rows && shown.length === 0 && (
        <div className="muted related-note">No role assignments match.</div>
      )}
      {shown.length > 0 && (
        <table className="table">
          <thead>
            <tr>
              <SortTh id="who" sort={sort}>
                Who
              </SortTh>
              <SortTh id="through" sort={sort}>
                Through
              </SortTh>
              <SortTh id="role" sort={sort}>
                Role
              </SortTh>
              <SortTh id="type" sort={sort}>
                Type
              </SortTh>
              <SortTh id="scope" sort={sort}>
                Scope
              </SortTh>
              <SortTh id="ends" sort={sort}>
                Ends
              </SortTh>
              <SortTh id="activated" sort={sort}>
                Last activated
              </SortTh>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {sorted.map((r) => {
              const activated = r.assignmentType === 'Activated';
              const path = r.principalKind && PRINCIPAL_PATH[r.principalKind];
              return (
                <tr key={r.key} className="static">
                  <td>
                    {path ? (
                      <Link to={`${path}/${r.principalId}`}>{r.principalName}</Link>
                    ) : (
                      r.principalName
                    )}
                    {r.principalKind && r.principalKind !== 'user' && (
                      <span className="muted">
                        {' · '}
                        {r.principalKind === 'group' ? 'group' : 'app'}
                      </span>
                    )}
                  </td>
                  <td>{through(r)}</td>
                  <td title={r.roleDescription}>{r.roleName}</td>
                  <td>
                    {r.kind === 'eligible' ? (
                      <span className="badge pending">Eligible</span>
                    ) : (
                      <span
                        className="badge granted"
                        title={activated ? 'Activated from an eligibility' : 'Assigned by an admin'}
                      >
                        {activated ? 'Activated' : 'Active'}
                      </span>
                    )}
                  </td>
                  <td>{scope(r.directoryScopeId)}</td>
                  <td>
                    {(activated ? time(r.endDateTime) : day(r.endDateTime)) ?? (
                      <span className="muted">No end date</span>
                    )}
                  </td>
                  <td>
                    <LastActivated row={r} />
                  </td>
                  <td className="row-actions">
                    {!activated && !r.via && (
                      <RemoveButton
                        label="Remove this role"
                        busy={busy === r.key}
                        onClick={() => void remove(r)}
                      />
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}
      {confirmDialog}
    </>
  );
}
