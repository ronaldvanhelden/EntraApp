import { useEffect, useMemo, useState } from 'react';
import { useMsal } from '@azure/msal-react';
import { PIM_ACTIVATE_SCOPES, PIM_ASSIGN_SCOPES, ROLE_READ_SCOPES } from '../auth/config';
import { useGraphToken } from '../auth/useGraphToken';
import type { GraphTokenFn } from '../graph/client';
import type { PrincipalRef } from '../graph/directoryObjects';
import {
  activateRole,
  assignRole,
  getActivationRules,
  listRoleDefinitions,
  pimError,
  stepUpFor,
  type ActivationRules,
  type AssignmentEnd,
  type RoleDefinition,
  type RoleRequestResult,
  type RoleSchedule,
  type StepUp,
} from '../graph/pim';
import { Modal } from './Modal';
import { PrincipalPicker } from './PrincipalPicker';
import { Field } from './PrincipalEditors';

const scopeLabel = (scope: string) => (scope === '/' ? 'the whole directory' : scope);

// Turn one of your own eligible roles on for a few hours.
export function ActivateRole({
  role,
  userId,
  onClose,
  onDone,
}: {
  role: RoleSchedule;
  // The signed-in user: the one the role is activated for.
  userId: string;
  onClose: () => void;
  // Called once the request went in (activated or waiting for approval).
  onDone: () => void;
}) {
  const { instance, accounts } = useMsal();
  const token = useGraphToken(PIM_ACTIVATE_SCOPES);
  const readToken = useGraphToken(ROLE_READ_SCOPES);
  const [rules, setRules] = useState<ActivationRules | null>(null);
  const [hours, setHours] = useState(1);
  const [justification, setJustification] = useState('');
  const [ticket, setTicket] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [stepUp, setStepUp] = useState<StepUp | null>(null);
  const [result, setResult] = useState<RoleRequestResult | null>(null);

  useEffect(() => {
    let cancelled = false;
    void getActivationRules(readToken, role.roleDefinitionId).then((r) => {
      if (cancelled) return;
      setRules(r);
      setHours(Math.min(r.maxHours, 8));
    });
    return () => {
      cancelled = true;
    };
  }, [readToken, role.roleDefinitionId]);

  const max = rules?.maxHours ?? 8;
  const valid =
    hours > 0 &&
    hours <= max &&
    (!rules?.justification || justification.trim().length > 0) &&
    (!rules?.ticket || ticket.trim().length > 0);

  const send = async (getToken: GraphTokenFn) => {
    setBusy(true);
    setError(null);
    try {
      const res = await activateRole(
        getToken,
        {
          principalId: userId,
          roleDefinitionId: role.roleDefinitionId,
          directoryScopeId: role.directoryScopeId,
        },
        { hours, justification: justification.trim(), ticketNumber: ticket.trim() || undefined },
      );
      setStepUp(null);
      setResult(res);
      onDone();
    } catch (e: unknown) {
      const needed = stepUpFor(e);
      if (needed) setStepUp(needed);
      else setError(pimError(e));
    } finally {
      setBusy(false);
    }
  };

  // The role asks for a stronger sign-in. Runs from its own click so the
  // browser lets the sign-in window open.
  const verify = async () => {
    const account = instance.getActiveAccount() ?? accounts[0];
    if (!account || !stepUp) return;
    setBusy(true);
    setError(null);
    try {
      const fresh = await instance.acquireTokenPopup({
        account,
        scopes: PIM_ACTIVATE_SCOPES,
        ...(stepUp.claims
          ? { claims: stepUp.claims }
          : { extraQueryParameters: { amr_values: 'mfa' } }),
      });
      await send(async () => fresh.accessToken);
    } catch (e: unknown) {
      setError(pimError(e));
      setBusy(false);
    }
  };

  if (result) {
    const waiting = /pending/i.test(result.status ?? '') && !/provision/i.test(result.status ?? '');
    return (
      <Modal
        title={waiting ? 'Waiting for approval' : 'Role activated'}
        onClose={onClose}
        footer={
          <button className="primary" onClick={onClose}>
            Done
          </button>
        }
      >
        <p style={{ marginTop: 0 }}>
          {waiting ? (
            <>
              Your request for <strong>{role.roleName}</strong> was sent to the approvers. The role
              becomes active once one of them approves it.
            </>
          ) : (
            <>
              <strong>{role.roleName}</strong> is active for {hours} hour{hours === 1 ? '' : 's'}.
            </>
          )}
        </p>
        <p className="muted" style={{ marginBottom: 0, fontSize: 13 }}>
          It can take a few minutes before the role works everywhere. If something still says you
          lack permission after that, sign out and back in.
        </p>
      </Modal>
    );
  }

  return (
    <Modal
      title={`Activate ${role.roleName}`}
      onClose={() => !busy && onClose()}
      footer={
        <>
          <button disabled={busy} onClick={onClose}>
            Cancel
          </button>
          {stepUp ? (
            <button className="primary" disabled={busy} onClick={() => void verify()}>
              {busy ? 'Verifying…' : 'Verify and activate'}
            </button>
          ) : (
            <button
              className="primary"
              disabled={busy || !valid || !rules}
              onClick={() => void send(token)}
            >
              {busy ? 'Activating…' : 'Activate'}
            </button>
          )}
        </>
      }
    >
      <p className="muted" style={{ marginTop: 0, fontSize: 13 }}>
        For {scopeLabel(role.directoryScopeId)}.{' '}
        {rules?.approval && 'This role needs approval before it becomes active. '}
        {rules?.mfa && 'You may be asked to verify with MFA.'}
      </p>
      {!rules ? (
        <span className="spinner" />
      ) : (
        <>
          <Field label={`Duration in hours (at most ${max})`}>
            <input
              type="number"
              min={0.5}
              max={max}
              step={0.5}
              value={hours}
              onChange={(e) => setHours(Number(e.target.value))}
            />
          </Field>
          <Field label={`Reason${rules.justification ? '' : ' (optional)'}`}>
            <textarea
              rows={2}
              value={justification}
              autoFocus
              onChange={(e) => setJustification(e.target.value)}
              placeholder="Why you need the role now"
            />
          </Field>
          {rules.ticket && (
            <Field label="Ticket number">
              <input value={ticket} onChange={(e) => setTicket(e.target.value)} />
            </Field>
          )}
        </>
      )}
      {stepUp && (
        <p className="error" style={{ marginBottom: 0 }}>
          This role asks for an extra verification of your sign-in. Choose “Verify and activate” to
          do that now.
        </p>
      )}
      {error && (
        <p className="error" style={{ marginBottom: 0 }}>
          {error}
        </p>
      )}
    </Modal>
  );
}

const DURATIONS: [label: string, months: number | null][] = [
  ['Permanent', null],
  ['1 month', 1],
  ['3 months', 3],
  ['6 months', 6],
  ['1 year', 12],
];

// Make someone eligible for a role, or give them the role outright.
export function AssignRole({
  principal,
  onClose,
  onDone,
}: {
  // Fixed principal (the user page); otherwise one is picked here.
  principal?: PrincipalRef;
  onClose: () => void;
  onDone: () => void;
}) {
  const token = useGraphToken(PIM_ASSIGN_SCOPES);
  const readToken = useGraphToken(ROLE_READ_SCOPES);
  const [roles, setRoles] = useState<RoleDefinition[] | null>(null);
  const [filter, setFilter] = useState('');
  const [roleId, setRoleId] = useState('');
  const [picked, setPicked] = useState<PrincipalRef | null>(principal ?? null);
  const [kind, setKind] = useState<RoleSchedule['kind']>('eligible');
  const [months, setMonths] = useState<number | null>(12);
  const [justification, setJustification] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    listRoleDefinitions(readToken)
      .then((r) => !cancelled && setRoles(r))
      .catch((e: unknown) => !cancelled && setError(pimError(e)));
    return () => {
      cancelled = true;
    };
  }, [readToken]);

  const shown = useMemo(() => {
    const q = filter.trim().toLowerCase();
    return (roles ?? []).filter((r) => !q || (r.displayName ?? '').toLowerCase().includes(q));
  }, [roles, filter]);
  const role = roles?.find((r) => r.id === roleId);
  const valid = Boolean(picked && role && justification.trim());

  const submit = async () => {
    if (!picked || !role) return;
    let end: AssignmentEnd = { permanent: true };
    if (months !== null) {
      const until = new Date();
      until.setMonth(until.getMonth() + months);
      end = { permanent: false, endDateTime: until.toISOString() };
    }
    setBusy(true);
    setError(null);
    try {
      await assignRole(
        token,
        kind,
        { principalId: picked.id, roleDefinitionId: role.id, directoryScopeId: '/' },
        end,
        justification.trim(),
      );
      onDone();
    } catch (e: unknown) {
      setError(pimError(e));
      setBusy(false);
    }
  };

  return (
    <Modal
      title={principal ? `Assign a role to ${principal.displayName}` : 'Assign a role'}
      onClose={() => !busy && onClose()}
      footer={
        <>
          <button disabled={busy} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={busy || !valid} onClick={() => void submit()}>
            {busy ? 'Assigning…' : 'Assign'}
          </button>
        </>
      }
    >
      {!principal && (
        <div style={{ marginBottom: 12 }}>
          <PrincipalPicker
            kinds={['user', 'group', 'sp']}
            selected={picked}
            onChange={setPicked}
            placeholder="Search who gets the role (min 2 chars)…"
            autoFocus
          />
          <p className="muted" style={{ fontSize: 12, margin: '4px 0 0' }}>
            A group only works when it was created as role-assignable.
          </p>
        </div>
      )}
      <Field label="Role">
        <input
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
          placeholder="Filter roles…"
          autoFocus={Boolean(principal)}
        />
      </Field>
      {!roles && !error ? (
        <span className="spinner" />
      ) : (
        <select
          size={7}
          value={roleId}
          onChange={(e) => setRoleId(e.target.value)}
          style={{ width: '100%', marginBottom: 4 }}
        >
          {shown.map((r) => (
            <option key={r.id} value={r.id} title={r.description}>
              {r.displayName ?? r.id}
              {r.isBuiltIn === false ? ' (custom)' : ''}
            </option>
          ))}
        </select>
      )}
      <p className="muted" style={{ fontSize: 12, marginTop: 0, minHeight: 32 }}>
        {role?.description}
      </p>
      <div className="form-grid">
        <Field label="Type">
          <select value={kind} onChange={(e) => setKind(e.target.value as RoleSchedule['kind'])}>
            <option value="eligible">Eligible — must activate it to use it</option>
            <option value="active">Active — has it all the time</option>
          </select>
        </Field>
        <Field label="Ends">
          <select
            value={months ?? ''}
            onChange={(e) => setMonths(e.target.value ? Number(e.target.value) : null)}
          >
            {DURATIONS.map(([label, m]) => (
              <option key={label} value={m ?? ''}>
                {label}
              </option>
            ))}
          </select>
        </Field>
      </div>
      <Field label="Reason">
        <textarea
          rows={2}
          value={justification}
          onChange={(e) => setJustification(e.target.value)}
          placeholder="Recorded in the PIM audit history"
        />
      </Field>
      <p className="muted" style={{ fontSize: 12, margin: 0 }}>
        Assigned for the whole directory. The role's PIM settings decide whether a permanent
        assignment is allowed; if not, pick an end date.
      </p>
      {error && (
        <p className="error" style={{ marginBottom: 0 }}>
          {error}
        </p>
      )}
    </Modal>
  );
}
