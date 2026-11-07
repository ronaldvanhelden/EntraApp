import { useState } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import { updateApplication } from '../graph/applications';
import type { AppRole, Application, OAuth2PermissionScope } from '../graph/types';
import { CopyButton } from './CopyButton';
import { useConfirm } from './ConfirmDialog';
import { Modal } from './Modal';
import { RemoveButton } from './RemoveButton';

// Manages what this app registration *exposes* to other clients — the
// "Expose an API" and "App roles" blades in the Entra portal:
//
//   • delegated permissions (scopes) under application.api.oauth2PermissionScopes
//   • application roles under application.appRoles
//
// Graph replaces each collection wholesale on PATCH, and refuses to drop an
// entry that is still enabled, so delete is a two-step disable-then-remove.

interface Props {
  app: Application;
  onAppChange: (next: Application) => void;
}

export function ExposedApiManager({ app, onAppChange }: Props) {
  const token = useGraphToken();
  const { confirm, confirmDialog } = useConfirm();
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [scopeModal, setScopeModal] = useState<'closed' | 'new' | OAuth2PermissionScope>('closed');
  const [roleModal, setRoleModal] = useState<'closed' | 'new' | AppRole>('closed');

  const scopes = app.api?.oauth2PermissionScopes ?? [];
  const roles = app.appRoles ?? [];
  // Roles Entra synthesises can't be edited here — only ones defined on the app.
  const editableRole = (r: AppRole) => !r.origin || r.origin === 'Application';

  const saveScopes = async (next: OAuth2PermissionScope[]) => {
    await updateApplication(token, app.id, { api: { oauth2PermissionScopes: next } });
    onAppChange({ ...app, api: { ...app.api, oauth2PermissionScopes: next } });
  };
  const saveRoles = async (next: AppRole[]) => {
    await updateApplication(token, app.id, { appRoles: next });
    onAppChange({ ...app, appRoles: next });
  };

  const run = async (key: string, fn: () => Promise<void>) => {
    setBusy(key);
    setError(null);
    try {
      await fn();
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      throw e;
    } finally {
      setBusy(null);
    }
  };

  const upsertScope = (s: OAuth2PermissionScope) =>
    run(`scope:${s.id}`, () =>
      saveScopes(
        scopes.some((x) => x.id === s.id)
          ? scopes.map((x) => (x.id === s.id ? s : x))
          : [...scopes, s],
      ),
    );
  const toggleScope = (s: OAuth2PermissionScope) =>
    upsertScope({ ...s, isEnabled: !s.isEnabled }).catch(() => {});
  const deleteScope = async (s: OAuth2PermissionScope) => {
    const ok = await confirm({
      title: `Delete scope ${s.value}?`,
      message: (
        <>
          <p style={{ margin: '0 0 8px', fontSize: 14 }}>
            Clients that request <span className="mono">{s.value}</span> will stop getting tokens
            with it, and any consent already granted for it is dropped.
            {s.isEnabled && ' The scope is disabled first, then removed.'}
          </p>
        </>
      ),
      confirmLabel: 'Delete scope',
    });
    if (!ok) return;
    await run(`scope:${s.id}`, async () => {
      if (s.isEnabled) {
        await saveScopes(scopes.map((x) => (x.id === s.id ? { ...x, isEnabled: false } : x)));
      }
      await saveScopes(scopes.filter((x) => x.id !== s.id));
    }).catch(() => {});
  };

  const upsertRole = (r: AppRole) =>
    run(`role:${r.id}`, () =>
      saveRoles(
        roles.some((x) => x.id === r.id)
          ? roles.map((x) => (x.id === r.id ? r : x))
          : [...roles, r],
      ),
    );
  const toggleRole = (r: AppRole) => upsertRole({ ...r, isEnabled: !r.isEnabled }).catch(() => {});
  const deleteRole = async (r: AppRole) => {
    const ok = await confirm({
      title: `Delete app role ${r.value || r.displayName}?`,
      message: (
        <p style={{ margin: 0, fontSize: 14 }}>
          Users, groups and apps assigned this role lose it; tokens stop carrying{' '}
          <span className="mono">{r.value}</span>.
          {r.isEnabled && ' The role is disabled first, then removed.'}
        </p>
      ),
      confirmLabel: 'Delete role',
    });
    if (!ok) return;
    await run(`role:${r.id}`, async () => {
      if (r.isEnabled) {
        await saveRoles(roles.map((x) => (x.id === r.id ? { ...x, isEnabled: false } : x)));
      }
      await saveRoles(roles.filter((x) => x.id !== r.id));
    }).catch(() => {});
  };

  const hasUri = (app.identifierUris?.length ?? 0) > 0;

  return (
    <>
      {error && <div className="card error">{error}</div>}

      <div className="card">
        <div className="row" style={{ justifyContent: 'space-between' }}>
          <div>
            <h3 style={{ margin: 0 }}>Delegated permissions (scopes)</h3>
            <div className="muted" style={{ fontSize: 12, marginTop: 4 }}>
              Permissions a client can request on behalf of a signed-in user, e.g.{' '}
              <span className="mono">api://…/Orders.Read</span>.
            </div>
          </div>
          <button
            className="primary"
            style={{ whiteSpace: 'nowrap' }}
            onClick={() => setScopeModal('new')}
            disabled={!!busy}
          >
            + Add scope
          </button>
        </div>
        {!hasUri && (
          <p className="muted" style={{ fontSize: 12, marginBottom: 0 }}>
            This app has no Application ID URI yet, so clients can't request its scopes. Set one
            under Overview → Edit (e.g. <span className="mono">api://{app.appId}</span>).
          </p>
        )}
        {scopes.length === 0 ? (
          <div className="muted" style={{ marginTop: 12 }}>
            No scopes defined.
          </div>
        ) : (
          <table className="table" style={{ marginTop: 12 }}>
            <thead>
              <tr>
                <th>Scope</th>
                <th>Shown to admins as</th>
                <th>Who can consent</th>
                <th>Status</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {scopes.map((s) => (
                <tr key={s.id} onClick={() => setScopeModal(s)}>
                  <td>
                    <span className="mono">{s.value}</span>
                    <div className="mono muted" style={{ fontSize: 11 }}>
                      {s.id}
                      <CopyButton value={s.id} label="scope id" />
                    </div>
                  </td>
                  <td>
                    {s.adminConsentDisplayName}
                    <div className="muted" style={{ fontSize: 11 }}>
                      {s.adminConsentDescription}
                    </div>
                  </td>
                  <td>
                    {s.type === 'Admin' ? (
                      <span className="badge pending">Admins only</span>
                    ) : (
                      <span className="badge delegated">Users</span>
                    )}
                  </td>
                  <td>
                    {s.isEnabled ? (
                      <span className="badge granted">Enabled</span>
                    ) : (
                      <span className="badge">Disabled</span>
                    )}
                  </td>
                  <td className="row-actions" onClick={(e) => e.stopPropagation()}>
                    <div className="row" style={{ gap: 6, justifyContent: 'flex-end' }}>
                      <button
                        disabled={busy === `scope:${s.id}`}
                        onClick={() => void toggleScope(s)}
                      >
                        {s.isEnabled ? 'Disable' : 'Enable'}
                      </button>
                      <RemoveButton
                        label="Delete scope"
                        busy={busy === `scope:${s.id}`}
                        onClick={() => void deleteScope(s)}
                      />
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      <div className="card">
        <div className="row" style={{ justifyContent: 'space-between' }}>
          <div>
            <h3 style={{ margin: 0 }}>App roles</h3>
            <div className="muted" style={{ fontSize: 12, marginTop: 4 }}>
              Roles assigned to users, groups or other apps; they arrive in the{' '}
              <span className="mono">roles</span> claim. Roles allowed for applications are what
              other apps request as application permissions.
            </div>
          </div>
          <button
            className="primary"
            style={{ whiteSpace: 'nowrap' }}
            onClick={() => setRoleModal('new')}
            disabled={!!busy}
          >
            + Add role
          </button>
        </div>
        {roles.length === 0 ? (
          <div className="muted" style={{ marginTop: 12 }}>
            No app roles defined.
          </div>
        ) : (
          <table className="table" style={{ marginTop: 12 }}>
            <thead>
              <tr>
                <th>Role</th>
                <th>Display name</th>
                <th>Allowed members</th>
                <th>Status</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {roles.map((r) => (
                <tr
                  key={r.id}
                  onClick={() => editableRole(r) && setRoleModal(r)}
                  style={editableRole(r) ? undefined : { cursor: 'default' }}
                >
                  <td>
                    <span className="mono">{r.value || <span className="muted">—</span>}</span>
                    <div className="mono muted" style={{ fontSize: 11 }}>
                      {r.id}
                      <CopyButton value={r.id} label="role id" />
                    </div>
                  </td>
                  <td>
                    {r.displayName}
                    <div className="muted" style={{ fontSize: 11 }}>
                      {r.description}
                    </div>
                  </td>
                  <td>
                    <div className="row wrap" style={{ gap: 4 }}>
                      {r.allowedMemberTypes.includes('User') && (
                        <span className="badge delegated">Users/groups</span>
                      )}
                      {r.allowedMemberTypes.includes('Application') && (
                        <span className="badge app">Applications</span>
                      )}
                    </div>
                  </td>
                  <td>
                    {r.isEnabled ? (
                      <span className="badge granted">Enabled</span>
                    ) : (
                      <span className="badge">Disabled</span>
                    )}
                    {!editableRole(r) && (
                      <div className="muted" style={{ fontSize: 11 }}>
                        built-in ({r.origin})
                      </div>
                    )}
                  </td>
                  <td className="row-actions" onClick={(e) => e.stopPropagation()}>
                    {editableRole(r) && (
                      <div className="row" style={{ gap: 6, justifyContent: 'flex-end' }}>
                        <button
                          disabled={busy === `role:${r.id}`}
                          onClick={() => void toggleRole(r)}
                        >
                          {r.isEnabled ? 'Disable' : 'Enable'}
                        </button>
                        <RemoveButton
                          label="Delete role"
                          busy={busy === `role:${r.id}`}
                          onClick={() => void deleteRole(r)}
                        />
                      </div>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {scopeModal !== 'closed' && (
        <ScopeModal
          existing={scopeModal === 'new' ? undefined : scopeModal}
          taken={scopes.map((s) => s.value.toLowerCase())}
          onClose={() => setScopeModal('closed')}
          onSubmit={async (s) => {
            await upsertScope(s);
            setScopeModal('closed');
          }}
        />
      )}
      {roleModal !== 'closed' && (
        <RoleModal
          existing={roleModal === 'new' ? undefined : roleModal}
          taken={roles.map((r) => r.value.toLowerCase()).filter(Boolean)}
          onClose={() => setRoleModal('closed')}
          onSubmit={async (r) => {
            await upsertRole(r);
            setRoleModal('closed');
          }}
        />
      )}
      {confirmDialog}
    </>
  );
}

// A permission's value becomes part of the scope string clients send; Entra
// rejects whitespace and a handful of characters.
const VALUE_RE = /^[^\s:]{1,120}$/;

function ScopeModal({
  existing,
  taken,
  onClose,
  onSubmit,
}: {
  existing?: OAuth2PermissionScope;
  taken: string[];
  onClose: () => void;
  onSubmit: (s: OAuth2PermissionScope) => Promise<void>;
}) {
  const [value, setValue] = useState(existing?.value ?? '');
  const [type, setType] = useState<'User' | 'Admin'>(existing?.type ?? 'User');
  const [adminName, setAdminName] = useState(existing?.adminConsentDisplayName ?? '');
  const [adminDesc, setAdminDesc] = useState(existing?.adminConsentDescription ?? '');
  const [userName, setUserName] = useState(existing?.userConsentDisplayName ?? '');
  const [userDesc, setUserDesc] = useState(existing?.userConsentDescription ?? '');
  const [enabled, setEnabled] = useState(existing?.isEnabled ?? true);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const v = value.trim();
  const duplicate =
    v.toLowerCase() !== existing?.value.toLowerCase() && taken.includes(v.toLowerCase());
  const valid = VALUE_RE.test(v) && !duplicate && adminName.trim() && adminDesc.trim();

  const submit = async () => {
    if (!valid) return;
    setSubmitting(true);
    setError(null);
    try {
      await onSubmit({
        id: existing?.id ?? crypto.randomUUID(),
        value: v,
        type,
        adminConsentDisplayName: adminName.trim(),
        adminConsentDescription: adminDesc.trim(),
        userConsentDisplayName: userName.trim() || undefined,
        userConsentDescription: userDesc.trim() || undefined,
        isEnabled: enabled,
      });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setSubmitting(false);
    }
  };

  return (
    <Modal
      title={existing ? `Edit scope ${existing.value}` : 'New scope'}
      onClose={() => !submitting && onClose()}
      footer={
        <>
          <button disabled={submitting} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={submitting || !valid} onClick={submit}>
            {submitting ? 'Saving…' : existing ? 'Save' : 'Add scope'}
          </button>
        </>
      }
    >
      <label className="field">
        <span>Scope name (value)</span>
        <input
          autoFocus
          className="mono"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          placeholder="Orders.Read"
        />
        {duplicate && (
          <span className="error" style={{ marginTop: 4 }}>
            A scope with this name already exists.
          </span>
        )}
        {v && !VALUE_RE.test(v) && (
          <span className="error" style={{ marginTop: 4 }}>
            No spaces or colons; max 120 characters.
          </span>
        )}
      </label>
      <label className="field">
        <span>Who can consent</span>
        <select value={type} onChange={(e) => setType(e.target.value as 'User' | 'Admin')}>
          <option value="User">Admins and users</option>
          <option value="Admin">Admins only</option>
        </select>
      </label>
      <label className="field">
        <span>Admin consent display name</span>
        <input
          value={adminName}
          onChange={(e) => setAdminName(e.target.value)}
          placeholder="Read orders"
        />
      </label>
      <label className="field">
        <span>Admin consent description</span>
        <textarea
          rows={2}
          value={adminDesc}
          onChange={(e) => setAdminDesc(e.target.value)}
          placeholder="Allows the app to read orders on behalf of the signed-in user."
        />
      </label>
      <label className="field">
        <span>User consent display name (optional)</span>
        <input value={userName} onChange={(e) => setUserName(e.target.value)} />
      </label>
      <label className="field">
        <span>User consent description (optional)</span>
        <textarea rows={2} value={userDesc} onChange={(e) => setUserDesc(e.target.value)} />
      </label>
      <label className="row" style={{ gap: 8 }}>
        <input
          type="checkbox"
          style={{ width: 'auto' }}
          checked={enabled}
          onChange={(e) => setEnabled(e.target.checked)}
        />
        <span>Enabled</span>
      </label>
      {error && <p className="error">{error}</p>}
    </Modal>
  );
}

function RoleModal({
  existing,
  taken,
  onClose,
  onSubmit,
}: {
  existing?: AppRole;
  taken: string[];
  onClose: () => void;
  onSubmit: (r: AppRole) => Promise<void>;
}) {
  const [value, setValue] = useState(existing?.value ?? '');
  const [displayName, setDisplayName] = useState(existing?.displayName ?? '');
  const [description, setDescription] = useState(existing?.description ?? '');
  const [forUsers, setForUsers] = useState(existing?.allowedMemberTypes.includes('User') ?? true);
  const [forApps, setForApps] = useState(
    existing?.allowedMemberTypes.includes('Application') ?? false,
  );
  const [enabled, setEnabled] = useState(existing?.isEnabled ?? true);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const v = value.trim();
  const duplicate =
    !!v && v.toLowerCase() !== existing?.value.toLowerCase() && taken.includes(v.toLowerCase());
  const valid =
    VALUE_RE.test(v) &&
    !duplicate &&
    displayName.trim() &&
    description.trim() &&
    (forUsers || forApps);

  const submit = async () => {
    if (!valid) return;
    setSubmitting(true);
    setError(null);
    try {
      await onSubmit({
        id: existing?.id ?? crypto.randomUUID(),
        value: v,
        displayName: displayName.trim(),
        description: description.trim(),
        allowedMemberTypes: [...(forUsers ? ['User'] : []), ...(forApps ? ['Application'] : [])],
        isEnabled: enabled,
        origin: existing?.origin,
      });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setSubmitting(false);
    }
  };

  return (
    <Modal
      title={existing ? `Edit role ${existing.value || existing.displayName}` : 'New app role'}
      onClose={() => !submitting && onClose()}
      footer={
        <>
          <button disabled={submitting} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={submitting || !valid} onClick={submit}>
            {submitting ? 'Saving…' : existing ? 'Save' : 'Add role'}
          </button>
        </>
      }
    >
      <label className="field">
        <span>Role value (goes in the roles claim)</span>
        <input
          autoFocus
          className="mono"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          placeholder="Orders.Admin"
        />
        {duplicate && (
          <span className="error" style={{ marginTop: 4 }}>
            A role with this value already exists.
          </span>
        )}
        {v && !VALUE_RE.test(v) && (
          <span className="error" style={{ marginTop: 4 }}>
            No spaces or colons; max 120 characters.
          </span>
        )}
      </label>
      <label className="field">
        <span>Display name</span>
        <input
          value={displayName}
          onChange={(e) => setDisplayName(e.target.value)}
          placeholder="Order administrators"
        />
      </label>
      <label className="field">
        <span>Description</span>
        <textarea
          rows={2}
          value={description}
          onChange={(e) => setDescription(e.target.value)}
          placeholder="Can create, update and cancel orders."
        />
      </label>
      <div className="field">
        <span
          style={{ display: 'block', fontSize: 12, color: 'var(--text-muted)', marginBottom: 4 }}
        >
          Allowed member types
        </span>
        <label className="row" style={{ gap: 8 }}>
          <input
            type="checkbox"
            style={{ width: 'auto' }}
            checked={forUsers}
            onChange={(e) => setForUsers(e.target.checked)}
          />
          <span>Users and groups</span>
        </label>
        <label className="row" style={{ gap: 8 }}>
          <input
            type="checkbox"
            style={{ width: 'auto' }}
            checked={forApps}
            onChange={(e) => setForApps(e.target.checked)}
          />
          <span>Applications (app-only / application permission)</span>
        </label>
        {!forUsers && !forApps && (
          <span className="error" style={{ fontSize: 12 }}>
            Pick at least one.
          </span>
        )}
      </div>
      <label className="row" style={{ gap: 8 }}>
        <input
          type="checkbox"
          style={{ width: 'auto' }}
          checked={enabled}
          onChange={(e) => setEnabled(e.target.checked)}
        />
        <span>Enabled</span>
      </label>
      {error && <p className="error">{error}</p>}
    </Modal>
  );
}
