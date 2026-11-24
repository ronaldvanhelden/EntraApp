import { useEffect, useState, type ReactNode } from 'react';
import { useMsal } from '@azure/msal-react';
import { useGraphToken } from '../auth/useGraphToken';
import { GROUP_WRITE_SCOPES, USER_PASSWORD_SCOPES, USER_WRITE_SCOPES } from '../auth/config';
import {
  createGroup,
  createUser,
  generatePassword,
  listVerifiedDomains,
  setUserPassword,
  type NewGroupKind,
} from '../graph/principals';
import type { QuickGroup, QuickUser } from '../graph/quickSearch';
import { CopyButton } from './CopyButton';
import { Modal } from './Modal';

// "New user" / "New group" dialogs for the Users and Groups pages. onCreated
// hands back the new object so the list can show it at once — Graph's search
// index lags behind writes. Editing happens on the detail pages.

const message = (e: unknown) => (e instanceof Error ? e.message : String(e));

// Letters, digits and . _ - only: what Graph accepts in a mail nickname.
const toNickname = (s: string) =>
  s
    .normalize('NFKD')
    .replace(/[^\w.\- ]/g, '')
    .trim()
    .replace(/\s+/g, '.')
    .toLowerCase();

export function Field({ label, children }: { label: ReactNode; children: ReactNode }) {
  return (
    <label className="field">
      <span>{label}</span>
      {children}
    </label>
  );
}

export function Check({
  checked,
  onChange,
  children,
}: {
  checked: boolean;
  onChange: (v: boolean) => void;
  children: ReactNode;
}) {
  return (
    <label className="row" style={{ gap: 6 }}>
      <input
        type="checkbox"
        style={{ width: 'auto' }}
        checked={checked}
        onChange={(e) => onChange(e.target.checked)}
      />
      {children}
    </label>
  );
}

// ---------------------------------------------------------------------- Users

interface NewUserProps {
  onClose: () => void;
  onCreated: (user: QuickUser) => void;
}

export function NewUser({ onClose, onCreated }: NewUserProps) {
  const readToken = useGraphToken();
  const writeToken = useGraphToken(USER_WRITE_SCOPES);
  const { instance, accounts } = useMsal();
  const [displayName, setDisplayName] = useState('');
  const [givenName, setGivenName] = useState('');
  const [surname, setSurname] = useState('');
  const [alias, setAlias] = useState('');
  const [aliasTouched, setAliasTouched] = useState(false);
  const [domains, setDomains] = useState<string[]>([]);
  const [domain, setDomain] = useState('');
  const [enabled, setEnabled] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [created, setCreated] = useState<{ upn: string; password: string } | null>(null);

  useEffect(() => {
    let cancelled = false;
    // If the tenant's domains can't be read, the signed-in user's own domain
    // is a safe guess.
    const own = (instance.getActiveAccount() ?? accounts[0])?.username.split('@')[1];
    listVerifiedDomains(readToken)
      .catch(() => [] as string[])
      .then((list) => {
        if (cancelled) return;
        const all = list.length ? list : own ? [own] : [];
        setDomains(all);
        setDomain((d) => d || all[0] || '');
      });
    return () => {
      cancelled = true;
    };
  }, [readToken, instance, accounts]);

  const effectiveAlias = aliasTouched ? alias : toNickname(displayName);
  const valid = displayName.trim() && effectiveAlias && domain;

  const submit = async () => {
    if (!valid) return;
    setSubmitting(true);
    setError(null);
    const password = generatePassword();
    const upn = `${effectiveAlias}@${domain}`;
    try {
      const user = await createUser(writeToken, {
        displayName: displayName.trim(),
        userPrincipalName: upn,
        mailNickname: effectiveAlias,
        ...(givenName.trim() ? { givenName: givenName.trim() } : {}),
        ...(surname.trim() ? { surname: surname.trim() } : {}),
        accountEnabled: enabled,
        password,
      });
      onCreated({ userType: 'Member', accountEnabled: enabled, ...user });
      setCreated({ upn: user.userPrincipalName ?? upn, password });
    } catch (e: unknown) {
      setError(message(e));
    } finally {
      setSubmitting(false);
    }
  };

  if (created) {
    return (
      <Modal
        title="User created"
        onClose={onClose}
        footer={
          <button className="primary" onClick={onClose}>
            Done
          </button>
        }
      >
        <p style={{ marginTop: 0 }}>
          Pass these on to the user. The password is shown only now and must be changed at first
          sign-in.
        </p>
        <Field label="Sign-in name">
          <div className="row" style={{ gap: 8 }}>
            <span className="mono">{created.upn}</span>
            <CopyButton value={created.upn} label="sign-in name" />
          </div>
        </Field>
        <Field label="Temporary password">
          <div className="row" style={{ gap: 8 }}>
            <span className="mono">{created.password}</span>
            <CopyButton value={created.password} label="temporary password" />
          </div>
        </Field>
      </Modal>
    );
  }

  return (
    <Modal
      title="New user"
      onClose={() => !submitting && onClose()}
      footer={
        <>
          <button disabled={submitting} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={submitting || !valid} onClick={submit}>
            {submitting ? 'Creating…' : 'Create'}
          </button>
        </>
      }
    >
      <Field label="Display name">
        <input
          autoFocus
          value={displayName}
          onChange={(e) => setDisplayName(e.target.value)}
          placeholder="Ada Lovelace"
        />
      </Field>
      <div className="row" style={{ gap: 12 }}>
        <div style={{ flex: 1 }}>
          <Field label="First name">
            <input value={givenName} onChange={(e) => setGivenName(e.target.value)} />
          </Field>
        </div>
        <div style={{ flex: 1 }}>
          <Field label="Last name">
            <input value={surname} onChange={(e) => setSurname(e.target.value)} />
          </Field>
        </div>
      </div>
      <Field label="Sign-in name (user principal name)">
        <div className="row" style={{ gap: 6 }}>
          <input
            style={{ flex: 1 }}
            value={effectiveAlias}
            onChange={(e) => {
              setAliasTouched(true);
              setAlias(e.target.value.replace(/[^\w.\-]/g, ''));
            }}
            placeholder="ada.lovelace"
            spellCheck={false}
          />
          <span className="muted">@</span>
          <select style={{ flex: 1 }} value={domain} onChange={(e) => setDomain(e.target.value)}>
            {domains.map((d) => (
              <option key={d} value={d}>
                {d}
              </option>
            ))}
          </select>
        </div>
      </Field>
      <Check checked={enabled} onChange={setEnabled}>
        Account enabled (can sign in)
      </Check>
      <p className="muted" style={{ fontSize: 12, marginBottom: 0 }}>
        A temporary password is generated and shown after the user is created.
      </p>
      {error && (
        <p className="error" style={{ marginTop: 12 }}>
          {error}
        </p>
      )}
    </Modal>
  );
}

// ------------------------------------------------------------ Reset password

interface ResetPasswordProps {
  user: Pick<QuickUser, 'id' | 'displayName' | 'userPrincipalName'> & {
    onPremisesSyncEnabled?: boolean | null;
  };
  onClose: () => void;
  // Called once the password has been replaced.
  onDone: () => void;
}

// Replace a user's password: either a generated temporary one (a reset) or
// one typed in here. The new password is shown once, to pass on.
export function ResetPassword({ user, onClose, onDone }: ResetPasswordProps) {
  const token = useGraphToken(USER_PASSWORD_SCOPES);
  const [mode, setMode] = useState<'generate' | 'custom'>('generate');
  const [custom, setCustom] = useState('');
  const [forceChange, setForceChange] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [done, setDone] = useState<string | null>(null);
  const name = user.displayName ?? user.userPrincipalName ?? user.id;
  // Entra's own rules (complexity, banned words) are checked by Graph; this
  // only stops an obviously too short one.
  const valid = mode === 'generate' || custom.length >= 8;

  const submit = async () => {
    if (!valid) return;
    const password = mode === 'generate' ? generatePassword() : custom;
    setSubmitting(true);
    setError(null);
    try {
      await setUserPassword(token, user.id, password, forceChange);
      setDone(password);
      onDone();
    } catch (e: unknown) {
      setError(message(e));
    } finally {
      setSubmitting(false);
    }
  };

  if (done) {
    return (
      <Modal
        title="Password changed"
        onClose={onClose}
        footer={
          <button className="primary" onClick={onClose}>
            Done
          </button>
        }
      >
        <p style={{ marginTop: 0 }}>
          Pass this on to {name}. It is shown only now
          {forceChange ? ' and must be changed at the next sign-in.' : '.'}
        </p>
        {user.userPrincipalName && (
          <Field label="Sign-in name">
            <div className="row" style={{ gap: 8 }}>
              <span className="mono">{user.userPrincipalName}</span>
              <CopyButton value={user.userPrincipalName} label="sign-in name" />
            </div>
          </Field>
        )}
        <Field label={forceChange ? 'Temporary password' : 'Password'}>
          <div className="row" style={{ gap: 8 }}>
            <span className="mono">{done}</span>
            <CopyButton value={done} label="password" />
          </div>
        </Field>
      </Modal>
    );
  }

  const radio = (value: 'generate' | 'custom', label: string) => (
    <label className="row" style={{ gap: 6, marginBottom: 6 }}>
      <input
        type="radio"
        name="password-mode"
        style={{ width: 'auto' }}
        checked={mode === value}
        onChange={() => setMode(value)}
      />
      {label}
    </label>
  );

  return (
    <Modal
      title={`Reset password — ${name}`}
      onClose={() => !submitting && onClose()}
      footer={
        <>
          <button disabled={submitting} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={submitting || !valid} onClick={submit}>
            {submitting ? 'Saving…' : mode === 'generate' ? 'Reset password' : 'Set password'}
          </button>
        </>
      }
    >
      <p style={{ marginTop: 0, fontSize: 14 }}>
        The current password stops working straight away. Sessions that are already signed in aren't
        ended by this.
      </p>
      {user.onPremisesSyncEnabled && (
        <p className="muted" style={{ fontSize: 12 }}>
          This user is synced from on-premises Active Directory: the change only goes through if
          password writeback is set up.
        </p>
      )}
      {radio('generate', 'Generate a temporary password')}
      {radio('custom', 'Set a password myself')}
      {mode === 'custom' && (
        <Field label="New password (at least 8 characters)">
          <input
            autoFocus
            className="mono"
            value={custom}
            onChange={(e) => setCustom(e.target.value)}
            autoComplete="off"
            spellCheck={false}
          />
        </Field>
      )}
      <div style={{ marginTop: 10 }}>
        <Check checked={forceChange} onChange={setForceChange}>
          Require a new password at the next sign-in
        </Check>
      </div>
      {error && (
        <p className="error" style={{ marginTop: 12 }}>
          {error}
        </p>
      )}
    </Modal>
  );
}

// --------------------------------------------------------------------- Groups

interface NewGroupProps {
  onClose: () => void;
  onCreated: (group: QuickGroup) => void;
}

export function NewGroup({ onClose, onCreated }: NewGroupProps) {
  const token = useGraphToken(GROUP_WRITE_SCOPES);
  const [displayName, setDisplayName] = useState('');
  const [description, setDescription] = useState('');
  const [kind, setKind] = useState<NewGroupKind>('security');
  const [nickname, setNickname] = useState('');
  const [nicknameTouched, setNicknameTouched] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const name = displayName.trim();
  const effectiveNickname = nicknameTouched ? nickname : toNickname(displayName);
  const valid = name.length > 0 && effectiveNickname.length > 0;

  const submit = async () => {
    if (!valid) return;
    setSubmitting(true);
    setError(null);
    try {
      const group = await createGroup(token, {
        displayName: name,
        description: description.trim() || undefined,
        mailNickname: effectiveNickname,
        kind,
      });
      onCreated(group);
    } catch (e: unknown) {
      setError(message(e));
      setSubmitting(false);
    }
  };

  return (
    <Modal
      title="New group"
      onClose={() => !submitting && onClose()}
      footer={
        <>
          <button disabled={submitting} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={submitting || !valid} onClick={submit}>
            {submitting ? 'Creating…' : 'Create'}
          </button>
        </>
      }
    >
      <Field label="Display name">
        <input
          autoFocus
          value={displayName}
          onChange={(e) => setDisplayName(e.target.value)}
          placeholder="Finance team"
        />
      </Field>
      <Field label="Description">
        <textarea rows={3} value={description} onChange={(e) => setDescription(e.target.value)} />
      </Field>
      <Field label="Type">
        <select value={kind} onChange={(e) => setKind(e.target.value as NewGroupKind)}>
          <option value="security">Security — grants access to apps and resources</option>
          <option value="m365">Microsoft 365 — shared mailbox, calendar and files</option>
        </select>
      </Field>
      <Field label={kind === 'm365' ? 'Mail alias' : 'Mail nickname'}>
        <input
          value={effectiveNickname}
          onChange={(e) => {
            setNicknameTouched(true);
            setNickname(e.target.value.replace(/[^\w.\-]/g, ''));
          }}
          spellCheck={false}
        />
      </Field>
      <p className="muted" style={{ fontSize: 12, marginBottom: 0 }}>
        You become the owner. The group opens after it is created.
      </p>
      {error && (
        <p className="error" style={{ marginTop: 12 }}>
          {error}
        </p>
      )}
    </Modal>
  );
}
