import { Fragment, useEffect, useMemo, useState, type ReactNode } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import {
  extractKeyIdFromSignIn,
  listAppOnlySignInsForApp,
  listUserSignInsForApp,
} from '../graph/signIns';
import type {
  KeyCredential,
  PasswordCredential,
  ServicePrincipal,
  SignIn,
} from '../graph/types';
import { getApplicationByAppId } from '../graph/applications';
import { Modal } from './Modal';

interface Props {
  sp: ServicePrincipal;
}

type Filter = 'all' | 'delegated' | 'app';

export function SignInAuditTab({ sp }: Props) {
  const token = useGraphToken();
  const [delegated, setDelegated] = useState<SignIn[] | null>(null);
  const [appOnly, setAppOnly] = useState<SignIn[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [filter, setFilter] = useState<Filter>('all');
  // Keyed by lowercased keyId AND lowercased thumbprint (both reference forms
  // that sign-in logs use).
  const [credentialHint, setCredentialHint] = useState<
    Record<string, { keyId: string; label: string }>
  >({});
  const [selected, setSelected] = useState<SignIn | null>(null);

  useEffect(() => {
    setDelegated(null);
    setAppOnly(null);
    setError(null);

    listUserSignInsForApp(token, sp.appId)
      .then(setDelegated)
      .catch((e: unknown) => {
        setDelegated([]);
        setError(e instanceof Error ? e.message : String(e));
      });
    listAppOnlySignInsForApp(token, sp.appId)
      .then(setAppOnly)
      .catch((e: unknown) => {
        setAppOnly([]);
        setError(e instanceof Error ? e.message : String(e));
      });

    // Best-effort: fetch credentials of the underlying app registration so we
    // can label keyIds the sign-in log references. Not all tenants have the
    // app object visible to the caller — swallow failure.
    getApplicationByAppId(token, sp.appId)
      .then((app) => {
        if (!app) return;
        const map: Record<string, { keyId: string; label: string }> = {};
        (app.passwordCredentials ?? []).forEach((c: PasswordCredential) => {
          map[c.keyId.toLowerCase()] = {
            keyId: c.keyId,
            label: c.displayName ? `secret · ${c.displayName}` : 'secret',
          };
        });
        (app.keyCredentials ?? []).forEach((c: KeyCredential) => {
          const label = c.displayName
            ? `certificate · ${c.displayName}`
            : 'certificate';
          map[c.keyId.toLowerCase()] = { keyId: c.keyId, label };
          const thumb = thumbprintFromB64(c.customKeyIdentifier);
          if (thumb) {
            map[thumb.toLowerCase()] = { keyId: c.keyId, label };
          }
        });
        setCredentialHint(map);
      })
      .catch(() => {});
  }, [token, sp.appId]);

  const rows: SignIn[] = useMemo(() => {
    const merged = [...(delegated ?? []), ...(appOnly ?? [])];
    merged.sort((a, b) =>
      (b.createdDateTime ?? '').localeCompare(a.createdDateTime ?? ''),
    );
    if (filter === 'delegated')
      return merged.filter((s) => !isAppOnly(s));
    if (filter === 'app') return merged.filter(isAppOnly);
    return merged;
  }, [delegated, appOnly, filter]);

  const loading = delegated === null || appOnly === null;
  const totalDelegated = delegated?.length ?? 0;
  const totalApp = appOnly?.length ?? 0;

  return (
    <>
      <div className="card">
        <div className="row" style={{ justifyContent: 'space-between' }}>
          <div>
            <h3 style={{ margin: 0 }}>Sign-in audit</h3>
            <div className="muted" style={{ fontSize: 12, marginTop: 4 }}>
              Most recent sign-ins for this app. Microsoft Entra retains 30
              days of sign-in activity. Requires{' '}
              <span className="mono">AuditLog.Read.All</span>.
            </div>
          </div>
          <div className="tabs">
            <button
              className={filter === 'all' ? 'active' : ''}
              onClick={() => setFilter('all')}
            >
              All ({totalDelegated + totalApp})
            </button>
            <button
              className={filter === 'delegated' ? 'active' : ''}
              onClick={() => setFilter('delegated')}
            >
              Delegated ({totalDelegated})
            </button>
            <button
              className={filter === 'app' ? 'active' : ''}
              onClick={() => setFilter('app')}
            >
              App-only ({totalApp})
            </button>
          </div>
        </div>
      </div>

      {error && <div className="card error">{error}</div>}

      {loading ? (
        <div className="center" style={{ height: 120 }}>
          <span className="spinner" />
        </div>
      ) : rows.length === 0 ? (
        <div className="card empty">
          No sign-ins recorded in the last 30 days.
        </div>
      ) : (
        <div className="card" style={{ padding: 0 }}>
          <table className="table">
            <thead>
              <tr>
                <th>When</th>
                <th>Flow</th>
                <th>Actor</th>
                <th>Resource</th>
                <th>Credential</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((s) => (
                <SignInRow
                  key={s.id}
                  signIn={s}
                  credentialHint={credentialHint}
                  onClick={() => setSelected(s)}
                />
              ))}
            </tbody>
          </table>
        </div>
      )}
      {selected && (
        <SignInDetailModal
          signIn={selected}
          credentialHint={credentialHint}
          onClose={() => setSelected(null)}
        />
      )}
    </>
  );
}

function isAppOnly(s: SignIn): boolean {
  return (s.signInEventTypes ?? []).some(
    (t) => t === 'servicePrincipal' || t === 'managedIdentity',
  );
}

function SignInRow({
  signIn,
  credentialHint,
  onClick,
}: {
  signIn: SignIn;
  credentialHint: Record<string, { keyId: string; label: string }>;
  onClick: () => void;
}) {
  const appOnly = isAppOnly(signIn);
  const actorName = appOnly
    ? signIn.servicePrincipalName ?? signIn.appDisplayName
    : signIn.userDisplayName ?? signIn.userPrincipalName;
  const actorSubtitle = appOnly
    ? signIn.servicePrincipalId
    : signIn.userPrincipalName;
  const status = signIn.status;
  const ok = !status?.errorCode;

  return (
    <tr onClick={onClick} style={{ cursor: 'pointer' }}>
      <td>
        {signIn.createdDateTime
          ? new Date(signIn.createdDateTime).toLocaleString()
          : '—'}
        {signIn.ipAddress && (
          <div className="mono muted" style={{ fontSize: 11 }}>
            {signIn.ipAddress}
          </div>
        )}
      </td>
      <td>
        {appOnly ? (
          <span className="badge app">App-only</span>
        ) : (
          <span className="badge delegated">
            Delegated
            {signIn.isInteractive === false ? ' (non-int.)' : ''}
          </span>
        )}
        {signIn.authenticationProtocol && (
          <div className="muted" style={{ fontSize: 11, marginTop: 2 }}>
            {signIn.authenticationProtocol}
          </div>
        )}
      </td>
      <td>
        <div>{actorName ?? <span className="muted">—</span>}</div>
        {actorSubtitle && (
          <div className="muted mono" style={{ fontSize: 11 }}>
            {actorSubtitle}
          </div>
        )}
      </td>
      <td>
        <div>{signIn.resourceDisplayName ?? <span className="muted">—</span>}</div>
        {signIn.clientAppUsed && (
          <div className="muted" style={{ fontSize: 11 }}>
            via {signIn.clientAppUsed}
          </div>
        )}
      </td>
      <td>
        <CredentialCell signIn={signIn} credentialHint={credentialHint} />
      </td>
      <td>
        {ok ? (
          <span className="badge granted">Success</span>
        ) : (
          <>
            <span className="badge expired">Failed</span>
            <div className="muted" style={{ fontSize: 11 }}>
              {status?.failureReason ?? `code ${status?.errorCode}`}
            </div>
          </>
        )}
      </td>
    </tr>
  );
}

function CredentialCell({
  signIn,
  credentialHint,
}: {
  signIn: SignIn;
  credentialHint: Record<string, { keyId: string; label: string }>;
}) {
  const appOnly = isAppOnly(signIn);
  if (!appOnly) {
    // Delegated flows — credential is user-held (password/MFA), not the app's.
    const methods = signIn.authenticationMethodsUsed ?? [];
    return methods.length ? (
      <span className="muted" style={{ fontSize: 12 }}>
        {methods.join(', ')}
      </span>
    ) : (
      <span className="muted">User auth</span>
    );
  }

  const type = signIn.clientCredentialType;
  const label = credentialTypeLabel(type);
  const keyId = extractKeyIdFromSignIn(signIn);
  const hint = keyId ? credentialHint[keyId.toLowerCase()] : undefined;

  return (
    <>
      <div>{label ?? <span className="muted">—</span>}</div>
      {type === 'federatedIdentityCredential' && signIn.federatedCredentialId && (
        <div className="mono muted" style={{ fontSize: 11 }}>
          fic {signIn.federatedCredentialId}
        </div>
      )}
      {keyId ? (
        <div
          className="mono muted"
          title={keyId}
          style={{ fontSize: 11 }}
        >
          {hint ? `${hint.label} · ` : ''}
          {shorten(keyId)}
        </div>
      ) : (
        // Older events or some tenants omit the dedicated credential fields.
        // Click the row to see the raw event — authenticationProcessingDetails
        // sometimes still has enough info to identify the key.
        type &&
        type !== 'none' && (
          <div
            className="muted"
            style={{ fontSize: 11 }}
            title="Graph didn't return a credential id for this event — click the row for raw details"
          >
            (key not recorded)
          </div>
        )
      )}
    </>
  );
}

function credentialTypeLabel(type?: string): string {
  switch (type) {
    case 'none':
      return 'Public client (none)';
    case 'clientSecret':
      return 'Client secret';
    case 'clientAssertion':
      return 'Certificate (client assertion)';
    case 'federatedIdentityCredential':
      return 'Federated identity credential';
    case 'managedIdentity':
      return 'Managed identity';
    case undefined:
      return '—';
    default:
      return type;
  }
}

function shorten(s: string): string {
  return s.length > 13 ? `${s.slice(0, 8)}…${s.slice(-4)}` : s;
}

function thumbprintFromB64(b64?: string | null): string | undefined {
  if (!b64) return undefined;
  try {
    const bin = atob(b64);
    let hex = '';
    for (let i = 0; i < bin.length; i++) {
      hex += bin.charCodeAt(i).toString(16).padStart(2, '0');
    }
    return hex.toUpperCase();
  } catch {
    return undefined;
  }
}

function SignInDetailModal({
  signIn,
  credentialHint,
  onClose,
}: {
  signIn: SignIn;
  credentialHint: Record<string, { keyId: string; label: string }>;
  onClose: () => void;
}) {
  const appOnly = isAppOnly(signIn);
  const actorName = appOnly
    ? signIn.servicePrincipalName ?? signIn.appDisplayName
    : signIn.userDisplayName ?? signIn.userPrincipalName;
  const actorSub = appOnly
    ? signIn.servicePrincipalId
    : signIn.userPrincipalName;
  const credKey = extractKeyIdFromSignIn(signIn);
  const credHint = credKey ? credentialHint[credKey.toLowerCase()] : undefined;

  return (
    <Modal title="Sign-in details" onClose={onClose}>
      <div className="kv">
        <Row label="When">
          {signIn.createdDateTime
            ? new Date(signIn.createdDateTime).toLocaleString()
            : '—'}
        </Row>
        <Row label="Flow">
          {appOnly ? 'App-only (service principal)' : 'Delegated (user)'}
          {signIn.isInteractive === true && ' · interactive'}
          {signIn.isInteractive === false && ' · non-interactive'}
        </Row>
        <Row label="Auth protocol">{signIn.authenticationProtocol ?? '—'}</Row>
        <Row label="Actor">
          <div>{actorName ?? '—'}</div>
          {actorSub && (
            <div className="muted mono" style={{ fontSize: 11 }}>
              {actorSub}
            </div>
          )}
        </Row>
        <Row label="From IP" mono>
          {signIn.ipAddress ?? '—'}
        </Row>
        <Row label="Client">{signIn.clientAppUsed ?? '—'}</Row>
        <Row label="Resource">
          <div>{signIn.resourceDisplayName ?? '—'}</div>
          {signIn.resourceId && (
            <div className="mono muted" style={{ fontSize: 11 }}>
              {signIn.resourceId}
            </div>
          )}
        </Row>
        <Row label="Credential type">
          {appOnly
            ? credentialTypeLabel(signIn.clientCredentialType)
            : (signIn.authenticationMethodsUsed ?? []).join(', ') || 'User auth'}
        </Row>
        {appOnly && (
          <>
            <Row label="Credential">
              {credHint ? (
                <>
                  <div>{credHint.label}</div>
                  <div className="mono muted" style={{ fontSize: 11 }}>
                    keyId {credHint.keyId}
                  </div>
                </>
              ) : credKey ? (
                <div className="mono">{credKey}</div>
              ) : signIn.federatedCredentialId ? (
                <div className="mono">
                  fic {signIn.federatedCredentialId}
                </div>
              ) : (
                <span className="muted">—</span>
              )}
            </Row>
            {signIn.servicePrincipalCredentialKeyId && (
              <Row label="Credential keyId" mono>
                {signIn.servicePrincipalCredentialKeyId}
              </Row>
            )}
            {signIn.servicePrincipalCredentialThumbprint && (
              <Row label="Certificate thumbprint" mono>
                {signIn.servicePrincipalCredentialThumbprint}
              </Row>
            )}
          </>
        )}
        <Row label="Request ID" mono>
          {signIn.id}
        </Row>
        <Row label="Status">
          {signIn.status?.errorCode ? (
            <>
              <span className="badge expired">Failed</span>
              <div className="muted" style={{ fontSize: 12, marginTop: 4 }}>
                {signIn.status.failureReason ??
                  `code ${signIn.status.errorCode}`}
              </div>
              {signIn.status.additionalDetails && (
                <div className="muted" style={{ fontSize: 12 }}>
                  {signIn.status.additionalDetails}
                </div>
              )}
            </>
          ) : (
            <span className="badge granted">Success</span>
          )}
        </Row>
      </div>

      {(signIn.authenticationProcessingDetails?.length ?? 0) > 0 && (
        <>
          <h4 style={{ marginTop: 20, marginBottom: 8 }}>
            Authentication processing details
          </h4>
          <div className="kv">
            {signIn.authenticationProcessingDetails!.map((d, i) => (
              <Fragment key={i}>
                <div className="k">{d.key ?? '—'}</div>
                <div className="mono" style={{ wordBreak: 'break-all' }}>
                  {d.value ?? '—'}
                </div>
              </Fragment>
            ))}
          </div>
        </>
      )}

      <details style={{ marginTop: 20 }}>
        <summary className="muted" style={{ cursor: 'pointer', fontSize: 12 }}>
          Raw event ({Object.keys(signIn).length} fields)
        </summary>
        <pre
          className="mono"
          style={{
            fontSize: 11,
            marginTop: 8,
            padding: 12,
            background: 'var(--bg)',
            border: '1px solid var(--border-subtle)',
            borderRadius: 4,
            overflow: 'auto',
            maxHeight: 320,
            whiteSpace: 'pre-wrap',
            wordBreak: 'break-all',
          }}
        >
          {JSON.stringify(signIn, null, 2)}
        </pre>
      </details>
    </Modal>
  );
}

function Row({
  label,
  mono,
  children,
}: {
  label: string;
  mono?: boolean;
  children: ReactNode;
}) {
  return (
    <>
      <div className="k">{label}</div>
      <div className={mono ? 'mono' : undefined}>{children}</div>
    </>
  );
}
