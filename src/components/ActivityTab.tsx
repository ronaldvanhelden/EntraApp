import { Fragment, useEffect, useMemo, useState, type ReactNode } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import { GraphError } from '../graph/client';
import { getApplicationByAppId } from '../graph/applications';
import {
  buildCredentialActivityMap,
  listAppCredentialSignInActivities,
  listFederatedIdentityCredentials,
} from '../graph/credentials';
import { getServicePrincipalSignInActivity } from '../graph/servicePrincipals';
import {
  extractKeyIdFromSignIn,
  isAppOnlySignIn,
  listSignInsForApp,
  listSignInsForResource,
  SIGN_IN_EVENT_TYPES,
  type SignInEventType,
} from '../graph/signIns';
import type {
  AppCredentialSignInActivity,
  Application,
  FederatedIdentityCredential,
  ServicePrincipal,
  ServicePrincipalSignInActivity,
  SignIn,
} from '../graph/types';
import { Modal } from './Modal';

// Activity tab shared by app registrations and enterprise apps. Pulls every
// usage signal Entra exposes for one appId so an admin can judge whether the
// object is still in use before deleting it:
//
//   • the raw sign-in log, per event class, with the app as *client*
//     (interactive user, non-interactive user, app-only, managed identity);
//   • the same four classes with the app as *resource* — someone obtained a
//     token *for* this app's API, so deleting it breaks a caller;
//   • the per-credential "last used" report (secrets / certificates);
//   • the per-service-principal "last sign-in" report.
//
// The sign-in log keeps 30 days. The two reports summarise further back but
// lag the log by hours and don't cover federated identity credentials, so
// federated usage is derived from the log only.

interface Props {
  appId: string;
  // Pass the app registration / its federated credentials when the caller
  // already has them (ApplicationDetail). Omit (undefined) to let the tab look
  // them up best-effort; null means "known to be unreadable".
  app?: Application | null;
  fics?: FederatedIdentityCredential[] | null;
  // Service principal when known — its own credentials (SAML signing certs,
  // SP-held secrets) are matched against the log too.
  sp?: ServicePrincipal | null;
}

const PAGE = 50;
const MAX_PAGES = 2;
const CAP = PAGE * MAX_PAGES;

type Role = 'client' | 'resource';
interface Signal {
  role: Role;
  type: SignInEventType;
}
const SIGNALS: Signal[] = (['client', 'resource'] as Role[]).flatMap((role) =>
  SIGN_IN_EVENT_TYPES.map((type) => ({ role, type })),
);
const signalKey = (s: Signal) => `${s.role}:${s.type}`;

type Load<T> =
  | { kind: 'loading' }
  | { kind: 'ok'; data: T }
  | { kind: 'unavailable'; reason: string; forbidden: boolean };

function unavailable(e: unknown): Load<never> {
  if (e instanceof GraphError) {
    return {
      kind: 'unavailable',
      forbidden: e.status === 403,
      reason: e.status === 403 ? 'Requires AuditLog.Read.All (and Entra ID P1)' : e.message,
    };
  }
  return {
    kind: 'unavailable',
    forbidden: false,
    reason: e instanceof Error ? e.message : String(e),
  };
}

type EventKind = 'interactive' | 'nonInteractive' | 'appOnly' | 'managedIdentity';
type Filter = 'all' | EventKind | 'resource';

function eventKind(s: SignIn): EventKind {
  const types = s.signInEventTypes ?? [];
  if (types.includes('managedIdentity')) return 'managedIdentity';
  if (types.includes('servicePrincipal')) return 'appOnly';
  if (types.includes('nonInteractiveUser') || s.isInteractive === false) return 'nonInteractive';
  return 'interactive';
}

const KIND_LABEL: Record<EventKind, string> = {
  interactive: 'Interactive',
  nonInteractive: 'Non-interactive',
  appOnly: 'App-only',
  managedIdentity: 'Managed identity',
};

const SIGNAL_LABEL: Record<SignInEventType, string> = {
  interactiveUser: 'Users, interactive',
  nonInteractiveUser: 'Users, non-interactive (token refresh, SSO)',
  servicePrincipal: 'App-only (client credentials)',
  managedIdentity: 'Managed identity',
};

interface KnownCredential {
  key: string;
  kind: 'secret' | 'certificate' | 'federated';
  label: string;
  owner: 'application' | 'servicePrincipal';
  keyId?: string;
  thumbprint?: string;
  ficId?: string;
  expires?: string;
}

interface EventRow {
  s: SignIn;
  roles: Set<Role>;
}

export function ActivityTab({ appId, app, fics, sp }: Props) {
  const token = useGraphToken();
  const [signals, setSignals] = useState<Record<string, Load<SignIn[]>>>({});
  const [credReport, setCredReport] = useState<Load<Map<string, AppCredentialSignInActivity>>>({
    kind: 'loading',
  });
  const [spReport, setSpReport] = useState<Load<ServicePrincipalSignInActivity | undefined>>({
    kind: 'loading',
  });
  const [resolvedApp, setResolvedApp] = useState<Application | null | undefined>(app);
  const [resolvedFics, setResolvedFics] = useState<
    FederatedIdentityCredential[] | null | undefined
  >(fics);
  const [filter, setFilter] = useState<Filter>('all');
  const [selected, setSelected] = useState<SignIn | null>(null);

  useEffect(() => {
    let cancelled = false;
    setSignals(Object.fromEntries(SIGNALS.map((s) => [signalKey(s), { kind: 'loading' }])));
    setCredReport({ kind: 'loading' });
    setSpReport({ kind: 'loading' });
    setFilter('all');
    setSelected(null);

    for (const sig of SIGNALS) {
      const p =
        sig.role === 'client'
          ? listSignInsForApp(token, appId, sig.type, PAGE, MAX_PAGES)
          : listSignInsForResource(token, appId, sig.type, PAGE, MAX_PAGES);
      p.then((rows) => {
        if (!cancelled)
          setSignals((prev) => ({ ...prev, [signalKey(sig)]: { kind: 'ok', data: rows } }));
      }).catch((e: unknown) => {
        if (!cancelled) setSignals((prev) => ({ ...prev, [signalKey(sig)]: unavailable(e) }));
      });
    }
    listAppCredentialSignInActivities(token, appId)
      .then((rows) => {
        if (!cancelled) setCredReport({ kind: 'ok', data: buildCredentialActivityMap(rows) });
      })
      .catch((e: unknown) => {
        if (!cancelled) setCredReport(unavailable(e));
      });
    getServicePrincipalSignInActivity(token, appId)
      .then((r) => {
        if (!cancelled) setSpReport({ kind: 'ok', data: r });
      })
      .catch((e: unknown) => {
        if (!cancelled) setSpReport(unavailable(e));
      });
    return () => {
      cancelled = true;
    };
  }, [token, appId]);

  // Resolve the app registration (for credential names) when the caller
  // didn't hand it over. Foreign-tenant SPs have no readable application.
  useEffect(() => {
    if (app !== undefined) {
      setResolvedApp(app);
      return;
    }
    let cancelled = false;
    setResolvedApp(undefined);
    getApplicationByAppId(token, appId)
      .then((a) => {
        if (!cancelled) setResolvedApp(a ?? null);
      })
      .catch(() => {
        if (!cancelled) setResolvedApp(null);
      });
    return () => {
      cancelled = true;
    };
  }, [token, appId, app]);

  useEffect(() => {
    if (fics !== undefined) {
      setResolvedFics(fics);
      return;
    }
    if (!resolvedApp) {
      setResolvedFics(resolvedApp === null ? null : undefined);
      return;
    }
    let cancelled = false;
    listFederatedIdentityCredentials(token, resolvedApp.id)
      .then((r) => {
        if (!cancelled) setResolvedFics(r);
      })
      .catch(() => {
        if (!cancelled) setResolvedFics(null);
      });
    return () => {
      cancelled = true;
    };
  }, [token, fics, resolvedApp]);

  const credentials = useMemo<KnownCredential[]>(() => {
    const out: KnownCredential[] = [];
    const seen = new Set<string>();
    const add = (owner: KnownCredential['owner'], a?: Application | ServicePrincipal | null) => {
      for (const c of a?.passwordCredentials ?? []) {
        if (seen.has(c.keyId)) continue;
        seen.add(c.keyId);
        out.push({
          key: `secret:${c.keyId}`,
          kind: 'secret',
          owner,
          keyId: c.keyId,
          label: c.displayName || 'Client secret',
          expires: c.endDateTime,
        });
      }
      for (const c of a?.keyCredentials ?? []) {
        if (seen.has(c.keyId)) continue;
        seen.add(c.keyId);
        out.push({
          key: `cert:${c.keyId}`,
          kind: 'certificate',
          owner,
          keyId: c.keyId,
          thumbprint: thumbprintFromB64(c.customKeyIdentifier),
          label: c.displayName || 'Certificate',
          expires: c.endDateTime,
        });
      }
    };
    add('application', resolvedApp);
    if (sp && sp.appId === appId) add('servicePrincipal', sp);
    for (const f of resolvedFics ?? []) {
      out.push({
        key: `fic:${f.id}`,
        kind: 'federated',
        owner: 'application',
        ficId: f.id,
        label: f.name,
      });
    }
    return out;
  }, [resolvedApp, resolvedFics, sp, appId]);

  // Lookup tables from the reference forms the sign-in log uses.
  const credentialIndex = useMemo(() => {
    const byRef = new Map<string, KnownCredential>();
    const byFic = new Map<string, KnownCredential>();
    for (const c of credentials) {
      if (c.keyId) byRef.set(c.keyId.toLowerCase(), c);
      if (c.thumbprint) byRef.set(c.thumbprint.toLowerCase(), c);
      if (c.ficId) byFic.set(c.ficId.toLowerCase(), c);
    }
    return { byRef, byFic };
  }, [credentials]);

  const matchCredential = (s: SignIn): KnownCredential | undefined => {
    if (!isAppOnlySignIn(s)) return undefined;
    if (s.federatedCredentialId) {
      const f = credentialIndex.byFic.get(s.federatedCredentialId.toLowerCase());
      if (f) return f;
    }
    const ref = extractKeyIdFromSignIn(s);
    return ref ? credentialIndex.byRef.get(ref.toLowerCase()) : undefined;
  };

  const events = useMemo<EventRow[]>(() => {
    const map = new Map<string, EventRow>();
    for (const sig of SIGNALS) {
      const l = signals[signalKey(sig)];
      if (l?.kind !== 'ok') continue;
      for (const s of l.data) {
        const row = map.get(s.id) ?? { s, roles: new Set<Role>() };
        row.roles.add(sig.role);
        map.set(s.id, row);
      }
    }
    return [...map.values()].sort((a, b) =>
      (b.s.createdDateTime ?? '').localeCompare(a.s.createdDateTime ?? ''),
    );
  }, [signals]);

  const signalsDone = SIGNALS.every(
    (s) => signals[signalKey(s)] && signals[signalKey(s)].kind !== 'loading',
  );
  const okSignals = SIGNALS.filter((s) => signals[signalKey(s)]?.kind === 'ok');
  const unavailableSignals = SIGNALS.filter((s) => signals[signalKey(s)]?.kind === 'unavailable');

  // Per-credential usage derived from the log.
  const credentialUsage = useMemo(() => {
    const usage = new Map<string, { last?: string; count: number }>();
    const unknown = new Map<string, { ref: string; type?: string; last?: string; count: number }>();
    for (const { s, roles } of events) {
      if (!roles.has('client') || !isAppOnlySignIn(s)) continue;
      const c = matchCredential(s);
      const when = s.createdDateTime;
      if (c) {
        const u = usage.get(c.key) ?? { count: 0 };
        u.count++;
        if (when && (!u.last || when > u.last)) u.last = when;
        usage.set(c.key, u);
      } else {
        const ref =
          s.federatedCredentialId ??
          extractKeyIdFromSignIn(s) ??
          (s.clientCredentialType && s.clientCredentialType !== 'none'
            ? `(${s.clientCredentialType}, key not recorded)`
            : '(no credential — public client)');
        const u = unknown.get(ref) ?? { ref, type: s.clientCredentialType, count: 0 };
        u.count++;
        if (when && (!u.last || when > u.last)) u.last = when;
        unknown.set(ref, u);
      }
    }
    return { usage, unknown: [...unknown.values()] };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [events, credentialIndex]);

  const latestCredentialReport = useMemo(() => {
    if (credReport.kind !== 'ok') return undefined;
    let latest: string | undefined;
    for (const a of credReport.data.values()) {
      const t = a.signInActivity?.lastSignInDateTime;
      if (t && (!latest || t > latest)) latest = t;
    }
    return latest;
  }, [credReport]);

  const spReportLast =
    spReport.kind === 'ok' ? spReport.data?.lastSignInActivity?.lastSignInDateTime : undefined;
  const latestAny = [events[0]?.s.createdDateTime, spReportLast, latestCredentialReport]
    .filter((v): v is string => !!v)
    .sort()
    .pop();

  const verdict: 'loading' | 'unknown' | 'active' | 'quiet' = !signalsDone
    ? 'loading'
    : okSignals.length === 0 && credReport.kind !== 'ok' && spReport.kind !== 'ok'
      ? 'unknown'
      : latestAny
        ? 'active'
        : 'quiet';

  const filtered = useMemo(() => {
    if (filter === 'all') return events;
    if (filter === 'resource') return events.filter((e) => e.roles.has('resource'));
    return events.filter((e) => e.roles.has('client') && eventKind(e.s) === filter);
  }, [events, filter]);

  const count = (pred: (e: EventRow) => boolean) => events.filter(pred).length;
  const chip = (f: Filter, label: string, n: number) => (
    <button className={filter === f ? 'active' : ''} onClick={() => setFilter(f)}>
      {label} ({n})
    </button>
  );

  const audience = resolvedApp?.signInAudience;
  const multiTenant = !!audience && audience !== 'AzureADMyOrg';

  return (
    <>
      <div className="card">
        <h3>Is this app still in use?</h3>
        <div className="muted" style={{ fontSize: 12, marginTop: -6, marginBottom: 12 }}>
          Every usage signal Entra exposes for app ID <span className="mono">{appId}</span>. The
          sign-in log covers the last 30 days; the usage reports reach further back but lag by
          hours.
        </div>

        <VerdictBanner
          verdict={verdict}
          latest={latestAny}
          eventCount={events.length}
          unavailable={unavailableSignals.length}
        />

        <SignalTable
          title="As a client — this app obtained tokens"
          signals={signals}
          role="client"
        />
        <SignalTable
          title="As a resource — tokens were issued for this app's API"
          signals={signals}
          role="resource"
          hint="If anything shows here, other apps or users depend on this app. Deleting it breaks them."
        />

        <h4 style={{ margin: '16px 0 6px' }}>Entra usage report</h4>
        <SpReportView report={spReport} />

        <ul className="muted" style={{ fontSize: 12, margin: '14px 0 0', paddingLeft: 18 }}>
          <li>
            The sign-in log only keeps 30 days. "Nothing in 30 days" is not proof of "never used".
          </li>
          <li>
            Federated credential usage is only visible in the log — the reports don't cover it.
          </li>
          {multiTenant && (
            <li>
              This is a multi-tenant app (<span className="mono">{audience}</span>). Sign-ins in
              other tenants are not visible from here.
            </li>
          )}
          <li>
            References without a sign-in (known client apps, Conditional Access policies, app role
            assignments, code that holds the client ID) aren't detectable from the log.
          </li>
        </ul>
      </div>

      <div className="card">
        <h3>Credentials</h3>
        <div className="muted" style={{ fontSize: 12, marginTop: -6, marginBottom: 10 }}>
          "Last used" comes from the credential report (secrets and certificates only). "Last seen"
          and the count come from the app-only sign-ins in the last 30 days.
        </div>
        {credReport.kind === 'unavailable' && (
          <div className="muted" style={{ fontSize: 12, marginBottom: 8 }}>
            Credential report unavailable: {credReport.reason}
          </div>
        )}
        {resolvedApp === null && (
          <div className="muted" style={{ fontSize: 12, marginBottom: 8 }}>
            The app registration isn't readable from this tenant, so credentials seen in the log are
            listed by key id only.
          </div>
        )}
        <CredentialsTable
          credentials={credentials}
          usage={credentialUsage.usage}
          unknown={credentialUsage.unknown}
          report={credReport.kind === 'ok' ? credReport.data : undefined}
          loading={resolvedApp === undefined || resolvedFics === undefined || !signalsDone}
        />
      </div>

      <div className="card">
        <div className="row" style={{ justifyContent: 'space-between' }}>
          <h3 style={{ margin: 0 }}>Sign-ins · last 30 days</h3>
          <div className="tabs" style={{ marginBottom: 0 }}>
            {chip('all', 'All', events.length)}
            {chip(
              'interactive',
              'Interactive',
              count((e) => e.roles.has('client') && eventKind(e.s) === 'interactive'),
            )}
            {chip(
              'nonInteractive',
              'Non-interactive',
              count((e) => e.roles.has('client') && eventKind(e.s) === 'nonInteractive'),
            )}
            {chip(
              'appOnly',
              'App-only',
              count((e) => e.roles.has('client') && eventKind(e.s) === 'appOnly'),
            )}
            {chip(
              'managedIdentity',
              'Managed identity',
              count((e) => e.roles.has('client') && eventKind(e.s) === 'managedIdentity'),
            )}
            {chip(
              'resource',
              'As resource',
              count((e) => e.roles.has('resource')),
            )}
          </div>
        </div>
        {unavailableSignals.length > 0 && (
          <div className="muted" style={{ fontSize: 12, marginTop: 8 }}>
            {unavailableSignals.length === SIGNALS.length
              ? `Sign-in log unavailable: ${(signals[signalKey(unavailableSignals[0])] as { reason: string }).reason}`
              : `${unavailableSignals.length} of ${SIGNALS.length} queries failed — see the tables above.`}
          </div>
        )}
      </div>

      {!signalsDone ? (
        <div className="center" style={{ height: 120 }}>
          <span className="spinner" />
        </div>
      ) : filtered.length === 0 ? (
        <div className="card empty">
          {okSignals.length === 0
            ? 'Sign-in log could not be read.'
            : 'No sign-ins recorded in the last 30 days.'}
        </div>
      ) : (
        <div className="card" style={{ padding: 0 }}>
          <table className="table">
            <thead>
              <tr>
                <th>When</th>
                <th>Kind</th>
                <th>Actor</th>
                <th>{filter === 'resource' ? 'Client app' : 'Resource'}</th>
                <th>Credential</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((e) => (
                <SignInRow
                  key={e.s.id}
                  row={e}
                  appId={appId}
                  credential={matchCredential(e.s)}
                  onClick={() => setSelected(e.s)}
                />
              ))}
            </tbody>
          </table>
          {events.length >= CAP && (
            <div className="muted" style={{ fontSize: 12, padding: '8px 12px' }}>
              Showing the {CAP} most recent events per query; older events in the 30-day window are
              not listed.
            </div>
          )}
        </div>
      )}

      {selected && (
        <SignInDetailModal
          signIn={selected}
          appId={appId}
          credential={matchCredential(selected)}
          onClose={() => setSelected(null)}
        />
      )}
    </>
  );
}

function VerdictBanner({
  verdict,
  latest,
  eventCount,
  unavailable,
}: {
  verdict: 'loading' | 'unknown' | 'active' | 'quiet';
  latest?: string;
  eventCount: number;
  unavailable: number;
}) {
  if (verdict === 'loading') {
    return (
      <div className="verdict">
        <span className="spinner" />
        <div>Reading sign-in log and usage reports…</div>
      </div>
    );
  }
  if (verdict === 'unknown') {
    return (
      <div className="verdict unknown">
        <span className="badge">Unknown</span>
        <div>
          None of the usage signals could be read. Check that the signed-in account has{' '}
          <span className="mono">AuditLog.Read.All</span> and the tenant has an Entra ID P1 licence.
        </div>
      </div>
    );
  }
  if (verdict === 'active') {
    return (
      <div className="verdict active">
        <span className="badge pending">In use</span>
        <div>
          Last activity <strong title={latest}>{relativeTime(latest!)}</strong> (
          {formatDateTime(latest)})
          {eventCount > 0 && (
            <> · {eventCount >= CAP ? `${CAP}+` : eventCount} sign-ins in the last 30 days</>
          )}
          . Deleting this app would interrupt whoever is using it.
        </div>
      </div>
    );
  }
  return (
    <div className="verdict quiet">
      <span className="badge granted">No activity found</span>
      <div>
        No sign-ins in the last 30 days and the usage reports have no last-sign-in for this app.
        {unavailable > 0 && (
          <>
            {' '}
            <strong>
              {unavailable} of {SIGNALS.length} log queries failed
            </strong>
            , so this is a partial view.
          </>
        )}{' '}
        See the caveats below before deleting.
      </div>
    </div>
  );
}

function SignalTable({
  title,
  signals,
  role,
  hint,
}: {
  title: string;
  signals: Record<string, Load<SignIn[]>>;
  role: Role;
  hint?: string;
}) {
  return (
    <>
      <h4 style={{ margin: '16px 0 2px' }}>{title}</h4>
      {hint && (
        <div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>
          {hint}
        </div>
      )}
      <table className="table">
        <thead>
          <tr>
            <th>Flow</th>
            <th>Last sign-in</th>
            <th>Sign-ins (30 d)</th>
          </tr>
        </thead>
        <tbody>
          {SIGN_IN_EVENT_TYPES.map((type) => {
            const l = signals[signalKey({ role, type })];
            const rows = l?.kind === 'ok' ? l.data : [];
            const last = rows
              .map((s) => s.createdDateTime ?? '')
              .sort()
              .pop();
            return (
              <tr key={type} style={{ cursor: 'default' }}>
                <td>{SIGNAL_LABEL[type]}</td>
                <td>
                  {!l || l.kind === 'loading' ? (
                    <span className="spinner" />
                  ) : l.kind === 'unavailable' ? (
                    <span className="muted" title={l.reason}>
                      unavailable
                    </span>
                  ) : last ? (
                    <span title={formatDateTime(last)}>{relativeTime(last)}</span>
                  ) : (
                    <span className="muted">none</span>
                  )}
                </td>
                <td>
                  {l?.kind === 'ok' ? (
                    rows.length >= CAP ? (
                      `${CAP}+`
                    ) : rows.length === 0 ? (
                      <span className="muted">0</span>
                    ) : (
                      rows.length
                    )
                  ) : (
                    <span className="muted">—</span>
                  )}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </>
  );
}

function SpReportView({ report }: { report: Load<ServicePrincipalSignInActivity | undefined> }) {
  if (report.kind === 'loading') return <span className="spinner" />;
  if (report.kind === 'unavailable')
    return (
      <div className="muted" style={{ fontSize: 12 }}>
        Unavailable: {report.reason}
      </div>
    );
  const r = report.data;
  if (!r)
    return (
      <div className="muted" style={{ fontSize: 12 }}>
        No entry for this app — the report has never recorded a sign-in for it (or hasn't caught
        up).
      </div>
    );
  const rows: [string, string | undefined][] = [
    ['Any flow', r.lastSignInActivity?.lastSignInDateTime],
    ['Delegated, as client', r.delegatedClientSignInActivity?.lastSignInDateTime],
    ['Delegated, as resource', r.delegatedResourceSignInActivity?.lastSignInDateTime],
    ['App-only, as client', r.applicationAuthenticationClientSignInActivity?.lastSignInDateTime],
    [
      'App-only, as resource',
      r.applicationAuthenticationResourceSignInActivity?.lastSignInDateTime,
    ],
  ];
  return (
    <div className="kv">
      {rows.map(([label, when]) => (
        <Fragment key={label}>
          <div className="k">{label}</div>
          <div>
            {when ? (
              <>
                {relativeTime(when)} <span className="muted">({formatDateTime(when)})</span>
              </>
            ) : (
              <span className="muted">—</span>
            )}
          </div>
        </Fragment>
      ))}
    </div>
  );
}

function CredentialsTable({
  credentials,
  usage,
  unknown,
  report,
  loading,
}: {
  credentials: KnownCredential[];
  usage: Map<string, { last?: string; count: number }>;
  unknown: { ref: string; type?: string; last?: string; count: number }[];
  report?: Map<string, AppCredentialSignInActivity>;
  loading: boolean;
}) {
  if (loading && credentials.length === 0) return <span className="spinner" />;
  if (credentials.length === 0 && unknown.length === 0)
    return (
      <div className="muted" style={{ fontSize: 13 }}>
        No credentials, and no app-only sign-ins in the last 30 days.
      </div>
    );
  const kindLabel = { secret: 'Secret', certificate: 'Certificate', federated: 'Federated' };
  return (
    <table className="table">
      <thead>
        <tr>
          <th>Credential</th>
          <th>Kind</th>
          <th>Last used (report)</th>
          <th>Last seen (log)</th>
          <th>Sign-ins (30 d)</th>
        </tr>
      </thead>
      <tbody>
        {credentials.map((c) => {
          const u = usage.get(c.key);
          const rep = c.keyId ? report?.get(c.keyId) : undefined;
          const repLast = rep?.signInActivity?.lastSignInDateTime;
          const expired = !!c.expires && c.expires < new Date().toISOString();
          return (
            <tr key={c.key} style={{ cursor: 'default' }}>
              <td>
                {c.label}
                <div className="mono muted" style={{ fontSize: 11 }}>
                  {c.keyId ?? c.ficId}
                  {c.owner === 'servicePrincipal' && ' · on enterprise app'}
                  {expired && (
                    <span className="badge expired" style={{ marginLeft: 6 }}>
                      Expired
                    </span>
                  )}
                </div>
              </td>
              <td>{kindLabel[c.kind]}</td>
              <td>
                {c.kind === 'federated' ? (
                  <span
                    className="muted"
                    title="The credential report doesn't cover federated credentials"
                  >
                    n/a
                  </span>
                ) : !report ? (
                  <span className="muted">—</span>
                ) : repLast ? (
                  <span title={formatDateTime(repLast)}>
                    {relativeTime(repLast)}
                    {rep?.signInActivity?.resourceDisplayName && (
                      <div className="muted" style={{ fontSize: 11 }}>
                        → {rep.signInActivity.resourceDisplayName}
                      </div>
                    )}
                  </span>
                ) : (
                  <span className="muted">never</span>
                )}
              </td>
              <td>
                {u?.last ? (
                  <span title={formatDateTime(u.last)}>{relativeTime(u.last)}</span>
                ) : (
                  <span className="muted">none</span>
                )}
              </td>
              <td>{u ? u.count >= CAP ? `${CAP}+` : u.count : <span className="muted">0</span>}</td>
            </tr>
          );
        })}
        {unknown.map((u) => (
          <tr key={`unknown:${u.ref}`} style={{ cursor: 'default' }}>
            <td>
              <span className="muted">Not among the current credentials</span>
              <div className="mono muted" style={{ fontSize: 11 }}>
                {u.ref}
              </div>
            </td>
            <td>{credentialTypeLabel(u.type)}</td>
            <td>
              <span className="muted">—</span>
            </td>
            <td>
              {u.last ? <span title={formatDateTime(u.last)}>{relativeTime(u.last)}</span> : '—'}
            </td>
            <td>{u.count >= CAP ? `${CAP}+` : u.count}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function SignInRow({
  row,
  appId,
  credential,
  onClick,
}: {
  row: EventRow;
  appId: string;
  credential?: KnownCredential;
  onClick: () => void;
}) {
  const { s, roles } = row;
  const kind = eventKind(s);
  const appOnly = kind === 'appOnly' || kind === 'managedIdentity';
  const asResourceOnly = roles.has('resource') && !roles.has('client');
  const actorName = appOnly
    ? (s.servicePrincipalName ?? s.appDisplayName)
    : (s.userDisplayName ?? s.userPrincipalName);
  const actorSubtitle = appOnly ? s.servicePrincipalId : s.userPrincipalName;
  const ok = !s.status?.errorCode;

  return (
    <tr onClick={onClick} style={{ cursor: 'pointer' }}>
      <td>
        {formatDateTime(s.createdDateTime)}
        {s.ipAddress && (
          <div className="mono muted" style={{ fontSize: 11 }}>
            {s.ipAddress}
          </div>
        )}
      </td>
      <td>
        <span className={`badge ${appOnly ? 'app' : 'delegated'}`}>{KIND_LABEL[kind]}</span>
        {asResourceOnly && (
          <div className="muted" style={{ fontSize: 11, marginTop: 2 }}>
            as resource
          </div>
        )}
        {s.authenticationProtocol && (
          <div className="muted" style={{ fontSize: 11, marginTop: 2 }}>
            {s.authenticationProtocol}
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
        {asResourceOnly ? (
          <>
            <div>{s.appDisplayName ?? <span className="muted">—</span>}</div>
            {s.appId && s.appId !== appId && (
              <div className="muted mono" style={{ fontSize: 11 }}>
                {s.appId}
              </div>
            )}
          </>
        ) : (
          <>
            <div>{s.resourceDisplayName ?? <span className="muted">—</span>}</div>
            {s.clientAppUsed && (
              <div className="muted" style={{ fontSize: 11 }}>
                via {s.clientAppUsed}
              </div>
            )}
          </>
        )}
      </td>
      <td>
        <CredentialCell signIn={s} credential={credential} />
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

function CredentialCell({ signIn, credential }: { signIn: SignIn; credential?: KnownCredential }) {
  if (!isAppOnlySignIn(signIn)) {
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
  const ref = signIn.federatedCredentialId ?? extractKeyIdFromSignIn(signIn);
  return (
    <>
      <div>{credential ? credential.label : credentialTypeLabel(type)}</div>
      {credential && (
        <div className="muted" style={{ fontSize: 11 }}>
          {credentialTypeLabel(type)}
        </div>
      )}
      {ref ? (
        <div className="mono muted" title={ref} style={{ fontSize: 11 }}>
          {shorten(ref)}
        </div>
      ) : (
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

function formatDateTime(iso?: string): string {
  return iso ? new Date(iso).toLocaleString() : '—';
}

export function relativeTime(iso: string): string {
  const diff = Date.now() - new Date(iso).getTime();
  if (!Number.isFinite(diff)) return iso;
  const m = Math.round(diff / 60000);
  if (m < 1) return 'just now';
  if (m < 60) return `${m} min ago`;
  const h = Math.round(m / 60);
  if (h < 48) return `${h} h ago`;
  const d = Math.round(h / 24);
  if (d < 60) return `${d} days ago`;
  const mo = Math.round(d / 30);
  if (mo < 24) return `${mo} months ago`;
  return `${Math.round(mo / 12)} years ago`;
}

function SignInDetailModal({
  signIn,
  appId,
  credential,
  onClose,
}: {
  signIn: SignIn;
  appId: string;
  credential?: KnownCredential;
  onClose: () => void;
}) {
  const appOnly = isAppOnlySignIn(signIn);
  const actorName = appOnly
    ? (signIn.servicePrincipalName ?? signIn.appDisplayName)
    : (signIn.userDisplayName ?? signIn.userPrincipalName);
  const actorSub = appOnly ? signIn.servicePrincipalId : signIn.userPrincipalName;
  const credKey = signIn.federatedCredentialId ?? extractKeyIdFromSignIn(signIn);
  const asResource = !!signIn.appId && signIn.appId !== appId;

  return (
    <Modal title="Sign-in details" onClose={onClose}>
      <div className="kv">
        <Row label="When">{formatDateTime(signIn.createdDateTime)}</Row>
        <Row label="Flow">
          {KIND_LABEL[eventKind(signIn)]}
          {asResource ? ' · this app was the resource' : ' · this app was the client'}
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
        <Row label="Client app">
          <div>{signIn.appDisplayName ?? '—'}</div>
          {signIn.appId && (
            <div className="mono muted" style={{ fontSize: 11 }}>
              {signIn.appId}
            </div>
          )}
        </Row>
        <Row label="From IP" mono>
          {signIn.ipAddress ?? '—'}
        </Row>
        <Row label="Client type">{signIn.clientAppUsed ?? '—'}</Row>
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
              {credential ? (
                <>
                  <div>{credential.label}</div>
                  <div className="mono muted" style={{ fontSize: 11 }}>
                    {credential.keyId ?? credential.ficId}
                  </div>
                </>
              ) : credKey ? (
                <div className="mono">{credKey}</div>
              ) : (
                <span className="muted">—</span>
              )}
            </Row>
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
                {signIn.status.failureReason ?? `code ${signIn.status.errorCode}`}
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
          <h4 style={{ marginTop: 20, marginBottom: 8 }}>Authentication processing details</h4>
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

function Row({ label, mono, children }: { label: string; mono?: boolean; children: ReactNode }) {
  return (
    <>
      <div className="k">{label}</div>
      <div className={mono ? 'mono' : undefined}>{children}</div>
    </>
  );
}
