import { useEffect, useMemo, useState } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import {
  getSignIn,
  isAppOnlySignIn,
  listSignInsUsingCredential,
  type CredentialRef,
  type CredentialSignIns,
} from '../graph/signIns';
import type {
  AppCredentialSignInActivity,
  FederatedIdentityCredential,
  SignIn,
} from '../graph/types';
import { formatLocation, relativeTime, SignInDetailModal } from './ActivityTab';

// Everything known about how one credential has been used: the Entra
// "last used" report, plus every sign-in in the last 30 days that named the
// credential, broken down by resource, origin, user and failure. Rendered
// inside the secret / certificate / federated-credential dialogs.

interface Props {
  appId: string;
  label: string;
  ref_: CredentialRef;
  fic?: FederatedIdentityCredential;
  // appCredentialSignInActivities entry — secrets and certificates only.
  report?: AppCredentialSignInActivity;
  reportError?: string | null;
}

type Load<T> = { kind: 'loading' } | { kind: 'ok'; data: T } | { kind: 'error'; message: string };

interface Bucket {
  key: string;
  title: string;
  subtitle?: string;
  count: number;
  failed: number;
  last?: string;
  sample: SignIn;
}

function bucketize(
  rows: SignIn[],
  keyOf: (s: SignIn) => { key: string; title: string; subtitle?: string } | undefined,
): Bucket[] {
  const map = new Map<string, Bucket>();
  for (const s of rows) {
    const k = keyOf(s);
    if (!k) continue;
    const b = map.get(k.key) ?? { ...k, count: 0, failed: 0, sample: s };
    b.count++;
    if (s.status?.errorCode) b.failed++;
    if (s.createdDateTime && (!b.last || s.createdDateTime > b.last)) {
      b.last = s.createdDateTime;
      b.sample = s;
    }
    map.set(k.key, b);
  }
  return [...map.values()].sort(
    (a, b) => b.count - a.count || (b.last ?? '').localeCompare(a.last ?? ''),
  );
}

const PROTOCOL_LABEL: Record<string, string> = {
  clientCredentials: 'client credentials',
  authorizationCodeWithPkce: 'auth code (PKCE)',
  authorizationCodeWithoutPkce: 'auth code',
  refreshTokenGrant: 'refresh token',
  onBehalfOf: 'on-behalf-of',
  oAuth2: 'OAuth 2.0',
  ropc: 'ROPC (password)',
  deviceCode: 'device code',
  saml20: 'SAML 2.0',
  wsFederation: 'WS-Fed',
};
const protocol = (s: SignIn) =>
  s.authenticationProtocol
    ? (PROTOCOL_LABEL[s.authenticationProtocol] ?? s.authenticationProtocol)
    : undefined;

export function CredentialUsage({ appId, label, ref_, fic, report, reportError }: Props) {
  const token = useGraphToken();
  const [log, setLog] = useState<Load<CredentialSignIns>>({ kind: 'loading' });
  const [reportEvent, setReportEvent] = useState<Load<SignIn | null>>({ kind: 'loading' });
  const [selected, setSelected] = useState<SignIn | null>(null);
  const [showSameType, setShowSameType] = useState(false);
  const [showAll, setShowAll] = useState(false);

  const refKey = `${ref_.kind}:${ref_.keyId ?? ''}:${ref_.thumbprint ?? ''}:${ref_.ficId ?? ''}`;
  useEffect(() => {
    let cancelled = false;
    setLog({ kind: 'loading' });
    setSelected(null);
    setShowAll(false);
    listSignInsUsingCredential(token, appId, ref_)
      .then((r) => {
        if (!cancelled) setLog({ kind: 'ok', data: r });
      })
      .catch((e: unknown) => {
        if (!cancelled)
          setLog({ kind: 'error', message: e instanceof Error ? e.message : String(e) });
      });
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [token, appId, refKey]);

  // The report's last-sign-in points at one event by request id; fetch it so
  // "last used" can show who/where even when it's older than the 30-day log
  // window (Entra keeps the event only as long as the log does, so this can
  // legitimately 404).
  const requestId = report?.signInActivity?.requestId;
  useEffect(() => {
    let cancelled = false;
    if (!requestId) {
      setReportEvent({ kind: 'ok', data: null });
      return;
    }
    setReportEvent({ kind: 'loading' });
    getSignIn(token, requestId)
      .then((s) => {
        if (!cancelled) setReportEvent({ kind: 'ok', data: s });
      })
      .catch((e: unknown) => {
        if (!cancelled)
          setReportEvent({ kind: 'error', message: e instanceof Error ? e.message : String(e) });
      });
    return () => {
      cancelled = true;
    };
  }, [token, requestId]);

  const rows = log.kind === 'ok' ? log.data.attributed : [];
  const sameType = log.kind === 'ok' ? log.data.sameType : [];
  const stats = useMemo(() => {
    const failed = rows.filter((s) => !!s.status?.errorCode).length;
    const appOnly = rows.filter(isAppOnlySignIn).length;
    return {
      total: rows.length,
      failed,
      appOnly,
      delegated: rows.length - appOnly,
      last: rows[0]?.createdDateTime,
      first: rows[rows.length - 1]?.createdDateTime,
      ips: new Set(rows.map((s) => s.ipAddress).filter(Boolean)).size,
      resources: new Set(rows.map((s) => s.resourceId ?? s.resourceDisplayName).filter(Boolean))
        .size,
      users: new Set(rows.map((s) => s.userPrincipalName).filter(Boolean)).size,
    };
  }, [rows]);

  const resources = useMemo(
    () =>
      bucketize(rows, (s) =>
        s.resourceDisplayName || s.resourceId
          ? {
              key: s.resourceId ?? s.resourceDisplayName!,
              title: s.resourceDisplayName ?? s.resourceId!,
              subtitle: s.resourceDisplayName ? s.resourceId : undefined,
            }
          : undefined,
      ),
    [rows],
  );
  const origins = useMemo(
    () =>
      bucketize(rows, (s) =>
        s.ipAddress
          ? {
              key: s.ipAddress,
              title: s.ipAddress,
              subtitle:
                [
                  formatLocation(s),
                  s.autonomousSystemNumber ? `AS${s.autonomousSystemNumber}` : undefined,
                ]
                  .filter(Boolean)
                  .join(' · ') || undefined,
            }
          : undefined,
      ),
    [rows],
  );
  const users = useMemo(
    () =>
      bucketize(rows, (s) =>
        s.userPrincipalName || s.userDisplayName
          ? {
              key: s.userId ?? s.userPrincipalName ?? s.userDisplayName!,
              title: s.userDisplayName ?? s.userPrincipalName!,
              subtitle: s.userDisplayName ? s.userPrincipalName : undefined,
            }
          : undefined,
      ),
    [rows],
  );
  const sameTypeUsers = useMemo(
    () =>
      bucketize(sameType, (s) =>
        s.userPrincipalName || s.userDisplayName
          ? {
              key: s.userId ?? s.userPrincipalName ?? s.userDisplayName!,
              title: s.userDisplayName ?? s.userPrincipalName!,
              subtitle: s.userDisplayName ? s.userPrincipalName : undefined,
            }
          : undefined,
      ),
    [sameType],
  );
  const failures = useMemo(
    () =>
      bucketize(
        rows.filter((s) => !!s.status?.errorCode),
        (s) => ({
          key: String(s.status!.errorCode),
          title: s.status!.failureReason ?? `Error ${s.status!.errorCode}`,
          subtitle: `code ${s.status!.errorCode}`,
        }),
      ),
    [rows],
  );
  const protocols = useMemo(
    () =>
      bucketize(rows, (s) =>
        protocol(s) ? { key: protocol(s)!, title: protocol(s)! } : undefined,
      ),
    [rows],
  );

  const forbidden = log.kind === 'ok' && log.data.errors.some((e) => e.forbidden);
  const credentialHint = { label, keyId: ref_.keyId, ficId: ref_.ficId };
  const reportLast = report?.signInActivity?.lastSignInDateTime;

  return (
    <>
      <h4 style={{ marginTop: 20, marginBottom: 8 }}>Usage</h4>

      {/* Headline numbers */}
      {log.kind === 'loading' ? (
        <div className="row" style={{ gap: 8 }}>
          <span className="spinner" />
          <span className="muted">Reading sign-in log…</span>
        </div>
      ) : log.kind === 'error' ? (
        <p className="error" style={{ fontSize: 13 }}>
          {log.message}
        </p>
      ) : (
        <div className="stat-row">
          <Stat label="Sign-ins · 30 d" value={stats.total} />
          <Stat label="App-only" value={stats.appOnly} />
          <Stat label="Delegated" value={stats.delegated} />
          <Stat label="Failed" value={stats.failed} tone={stats.failed ? 'bad' : undefined} />
          <Stat label="IPs" value={stats.ips} />
          <Stat label="Resources" value={stats.resources} />
          {stats.users > 0 && <Stat label="Users" value={stats.users} />}
          <Stat
            label="Last seen"
            value={stats.last ? relativeTime(stats.last) : 'none'}
            title={stats.last ? new Date(stats.last).toLocaleString() : undefined}
          />
        </div>
      )}
      {forbidden && (
        <p className="muted" style={{ fontSize: 12 }}>
          Some sign-in queries were refused — reading the log needs{' '}
          <span className="mono">AuditLog.Read.All</span> and an Entra ID P1 licence.
        </p>
      )}
      {log.kind === 'ok' &&
        log.data.errors
          .filter((e) => !e.forbidden)
          .map((e) => (
            <p key={e.flow} className="error" style={{ fontSize: 12 }}>
              {e.flow}: {e.message}
            </p>
          ))}

      {/* Entra report (secrets/certs) — reaches beyond the 30-day log window */}
      {ref_.kind !== 'federated' && (
        <Section
          title="Last recorded use"
          hint="From Entra's credential report. Reaches back further than the sign-in log but lags by hours."
        >
          {reportError ? (
            <div className="muted" style={{ fontSize: 13 }}>
              Report unavailable: {reportError}
            </div>
          ) : !report || !reportLast ? (
            <div className="muted" style={{ fontSize: 13 }}>
              Never used, as far as the report knows.
            </div>
          ) : (
            <div className="kv">
              <KV label="When">
                <span title={new Date(reportLast).toLocaleString()}>
                  {relativeTime(reportLast)}
                </span>{' '}
                <span className="muted">({new Date(reportLast).toLocaleString()})</span>
              </KV>
              <KV label="Resource">
                {report.signInActivity?.resourceDisplayName ?? '—'}
                {report.signInActivity?.resourceId && (
                  <div className="mono muted" style={{ fontSize: 11 }}>
                    {report.signInActivity.resourceId}
                  </div>
                )}
              </KV>
              <KV label="Key">
                {report.keyType}
                {report.keyUsage ? ` · ${report.keyUsage}` : ''}
              </KV>
              {reportEvent.kind === 'ok' && reportEvent.data && (
                <KV label="That sign-in">
                  <EventSummary s={reportEvent.data} />{' '}
                  <button className="link" onClick={() => setSelected(reportEvent.data)}>
                    <span className="mono" style={{ fontSize: 12 }}>
                      details →
                    </span>
                  </button>
                </KV>
              )}
              {reportEvent.kind === 'error' && (
                <KV label="That sign-in">
                  <span className="muted" style={{ fontSize: 12 }}>
                    Event {requestId} is no longer in the log.
                  </span>
                </KV>
              )}
            </div>
          )}
        </Section>
      )}

      {/* Federated identity: what gets exchanged for this credential */}
      {fic && (
        <Section title="Federated identity">
          <div className="kv">
            <KV label="Issuer" mono>
              {fic.issuer}
            </KV>
            <KV label="Subject" mono>
              {fic.subject}
            </KV>
            <KV label="Audiences" mono>
              {fic.audiences.join(', ')}
            </KV>
            {fic.description && <KV label="Description">{fic.description}</KV>}
          </div>
          <p className="muted" style={{ fontSize: 12, margin: '8px 0 0' }}>
            Every sign-in below is a token from this issuer, for this subject, exchanged for an
            Entra token. Entra's credential report doesn't cover federated credentials, so the
            30-day log is the only usage record.
          </p>
        </Section>
      )}

      {log.kind === 'ok' && rows.length === 0 && (
        <Section title="Sign-ins · last 30 days">
          <div className="muted" style={{ fontSize: 13 }}>
            No sign-in in the last 30 days named this credential.
          </div>
        </Section>
      )}

      {rows.length > 0 && (
        <>
          <Section title="What it accessed" hint="Resources this credential obtained tokens for.">
            <BucketTable buckets={resources} head="Resource" onPick={setSelected} />
          </Section>

          <Section
            title="Where from"
            hint="Source IPs. For app-only flows this is the workload holding the credential — a server, pipeline agent or developer machine."
          >
            <BucketTable buckets={origins} head="IP address" onPick={setSelected} mono />
          </Section>

          {users.length > 0 && (
            <Section
              title="Users signed in through it"
              hint="Delegated flows where this exact key was recorded alongside the user."
            >
              <BucketTable buckets={users} head="User" onPick={setSelected} />
            </Section>
          )}

          {protocols.length > 1 && (
            <Section title="How">
              <BucketTable buckets={protocols} head="Grant / protocol" onPick={setSelected} />
            </Section>
          )}

          {failures.length > 0 && (
            <Section title="Failures">
              <BucketTable buckets={failures} head="Reason" onPick={setSelected} />
            </Section>
          )}

          <Section title={`All sign-ins · ${rows.length}${rows.length >= 100 ? '+' : ''}`}>
            <table className="table">
              <thead>
                <tr>
                  <th>When</th>
                  <th>Flow</th>
                  <th>Who / what</th>
                  <th>Resource</th>
                  <th>From</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {(showAll ? rows : rows.slice(0, 15)).map((s) => (
                  <EventRow key={s.id} s={s} onClick={() => setSelected(s)} />
                ))}
              </tbody>
            </table>
            {rows.length > 15 && !showAll && (
              <button className="ghost" style={{ marginTop: 6 }} onClick={() => setShowAll(true)}>
                Show all {rows.length}
              </button>
            )}
          </Section>
        </>
      )}

      {sameType.length > 0 && (
        <Section
          title={`Delegated sign-ins that used a ${ref_.kind === 'secret' ? 'client secret' : ref_.kind === 'certificate' ? 'certificate' : 'federated credential'} · ${sameType.length}`}
          hint="Entra recorded the credential type for these user sign-ins but not which key. They belong to this credential only if the app has no other of the same type."
        >
          {!showSameType ? (
            <button className="ghost" onClick={() => setShowSameType(true)}>
              Show {sameType.length} sign-ins by {sameTypeUsers.length}{' '}
              {sameTypeUsers.length === 1 ? 'user' : 'users'}
            </button>
          ) : (
            <BucketTable buckets={sameTypeUsers} head="User" onPick={setSelected} />
          )}
        </Section>
      )}

      {selected && (
        <SignInDetailModal
          signIn={selected}
          appId={appId}
          credential={credentialHint}
          onClose={() => setSelected(null)}
        />
      )}
    </>
  );
}

function Stat({
  label,
  value,
  tone,
  title,
}: {
  label: string;
  value: number | string;
  tone?: 'bad';
  title?: string;
}) {
  return (
    <div className={`stat${tone ? ` ${tone}` : ''}`} title={title}>
      <div className="stat-value">{value}</div>
      <div className="stat-label">{label}</div>
    </div>
  );
}

function Section({
  title,
  hint,
  children,
}: {
  title: string;
  hint?: string;
  children: React.ReactNode;
}) {
  return (
    <div style={{ marginTop: 16 }}>
      <h4 style={{ margin: '0 0 2px' }}>{title}</h4>
      {hint && (
        <div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>
          {hint}
        </div>
      )}
      {children}
    </div>
  );
}

function KV({
  label,
  mono,
  children,
}: {
  label: string;
  mono?: boolean;
  children: React.ReactNode;
}) {
  return (
    <>
      <div className="k">{label}</div>
      <div
        className={mono ? 'mono' : undefined}
        style={mono ? { overflowWrap: 'anywhere' } : undefined}
      >
        {children}
      </div>
    </>
  );
}

function EventSummary({ s }: { s: SignIn }) {
  const appOnly = isAppOnlySignIn(s);
  const who = appOnly
    ? (s.servicePrincipalName ?? s.appDisplayName ?? 'the app itself')
    : (s.userDisplayName ?? s.userPrincipalName ?? 'a user');
  const bits = [
    appOnly ? 'app-only' : 'delegated',
    protocol(s),
    s.ipAddress ? `from ${s.ipAddress}` : undefined,
    formatLocation(s),
    s.status?.errorCode ? `failed: ${s.status.failureReason ?? s.status.errorCode}` : 'success',
  ].filter(Boolean);
  return (
    <>
      <div>{who}</div>
      <div className="muted" style={{ fontSize: 11 }}>
        {bits.join(' · ')}
      </div>
    </>
  );
}

function BucketTable({
  buckets,
  head,
  onPick,
  mono,
}: {
  buckets: Bucket[];
  head: string;
  onPick: (s: SignIn) => void;
  mono?: boolean;
}) {
  return (
    <table className="table">
      <thead>
        <tr>
          <th>{head}</th>
          <th>Sign-ins</th>
          <th>Last</th>
        </tr>
      </thead>
      <tbody>
        {buckets.map((b) => (
          <tr key={b.key} onClick={() => onPick(b.sample)} title="Open the most recent sign-in">
            <td>
              <div className={mono ? 'mono' : undefined}>{b.title}</div>
              {b.subtitle && (
                <div className="mono muted" style={{ fontSize: 11 }}>
                  {b.subtitle}
                </div>
              )}
            </td>
            <td>
              {b.count}
              {b.failed > 0 && (
                <span className="badge expired" style={{ marginLeft: 6 }}>
                  {b.failed} failed
                </span>
              )}
            </td>
            <td title={b.last ? new Date(b.last).toLocaleString() : undefined}>
              {b.last ? relativeTime(b.last) : '—'}
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function EventRow({ s, onClick }: { s: SignIn; onClick: () => void }) {
  const appOnly = isAppOnlySignIn(s);
  const ok = !s.status?.errorCode;
  return (
    <tr onClick={onClick} style={{ cursor: 'pointer' }}>
      <td title={s.createdDateTime}>
        {s.createdDateTime ? new Date(s.createdDateTime).toLocaleString() : '—'}
      </td>
      <td>
        <span className={`badge ${appOnly ? 'app' : 'delegated'}`}>
          {appOnly ? 'App-only' : s.isInteractive === false ? 'Non-interactive' : 'Delegated'}
        </span>
        {protocol(s) && (
          <div className="muted" style={{ fontSize: 11, marginTop: 2 }}>
            {protocol(s)}
          </div>
        )}
      </td>
      <td>
        {appOnly ? (
          <>
            <div>{s.servicePrincipalName ?? s.appDisplayName ?? '—'}</div>
            {s.managedServiceIdentity?.msiType && (
              <div className="muted" style={{ fontSize: 11 }}>
                managed identity · {s.managedServiceIdentity.msiType}
              </div>
            )}
          </>
        ) : (
          <>
            <div>{s.userDisplayName ?? s.userPrincipalName ?? '—'}</div>
            {s.userDisplayName && s.userPrincipalName && (
              <div className="mono muted" style={{ fontSize: 11 }}>
                {s.userPrincipalName}
              </div>
            )}
          </>
        )}
      </td>
      <td>{s.resourceDisplayName ?? <span className="muted">—</span>}</td>
      <td>
        <div className="mono" style={{ fontSize: 12 }}>
          {s.ipAddress ?? '—'}
        </div>
        {formatLocation(s) && (
          <div className="muted" style={{ fontSize: 11 }}>
            {formatLocation(s)}
          </div>
        )}
      </td>
      <td>
        {ok ? (
          <span className="badge granted">Success</span>
        ) : (
          <>
            <span className="badge expired">Failed</span>
            <div className="muted" style={{ fontSize: 11 }}>
              {s.status?.failureReason ?? `code ${s.status?.errorCode}`}
            </div>
          </>
        )}
      </td>
    </tr>
  );
}
