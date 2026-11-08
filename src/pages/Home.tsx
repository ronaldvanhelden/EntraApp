import { useEffect, useMemo, useState, type ReactNode } from 'react';
import { Link } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { getMe, type MeResponse } from '../graph/me';
import { listRecentDirectoryAudits, type DirectoryAuditEvent } from '../graph/directoryAudits';
import { listApplications, listOwnedApplications } from '../graph/applications';
import { listDeletedApplications, daysUntilPurge } from '../graph/deletedItems';
import { listMyActiveRoleNames } from '../graph/directoryRoles';
import {
  countCollection,
  listFavoriteApplications,
  listLastSignInByAppId,
  listOwnerlessApplications,
} from '../graph/overview';
import type { Application } from '../graph/types';
import { AppIcon } from '../components/AppIcon';

// Activities the Overview card surfaces. Names must match Microsoft Graph's
// activityDisplayName values exactly — they're case-sensitive strings.
const AUDIT_ACTIVITIES = [
  // App registration lifecycle
  'Add application',
  'Update application',
  'Delete application',
  'Hard Delete application',
  // Service principal lifecycle
  'Add service principal',
  'Update service principal',
  'Delete service principal',
  'Hard Delete service principal',
  // Credential rotation
  'Update application – Certificates and secrets management',
  // Consent & permission grants
  'Consent to application',
  'Add delegated permission grant',
  'Add app role assignment to service principal',
  'Add app role assignment grant to user',
];

const DAY = 864e5;
const SOON_DAYS = 30;
const PURGE_WARN_DAYS = 7;
const UNUSED_DAYS = 30;

type Loadable<T> = T | 'loading' | 'error';
type Scope = 'mine' | 'tenant';
const SCOPE_KEY = 'entraapp.overview.scope';

function readScope(): Scope {
  try {
    return localStorage.getItem(SCOPE_KEY) === 'tenant' ? 'tenant' : 'mine';
  } catch {
    return 'mine';
  }
}

function useLoad<T>(load: (() => Promise<T>) | null, deps: unknown[]): Loadable<T> {
  const [state, setState] = useState<Loadable<T>>('loading');
  useEffect(() => {
    if (!load) return;
    let cancelled = false;
    setState('loading');
    load()
      .then((v) => !cancelled && setState(v))
      .catch(() => !cancelled && setState('error'));
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);
  return state;
}

export function Home() {
  const token = useGraphToken();
  const [scope, setScopeState] = useState<Scope>(readScope);
  const setScope = (s: Scope) => {
    setScopeState(s);
    try {
      localStorage.setItem(SCOPE_KEY, s);
    } catch {
      // Not persisted; the choice still applies for this visit.
    }
  };

  const me = useLoad<MeResponse>(() => getMe(token), [token]);
  const roles = useLoad(() => listMyActiveRoleNames(token), [token]);
  const owned = useLoad(() => listOwnedApplications(token), [token]);
  const favorites = useLoad(() => listFavoriteApplications(token), [token]);
  const deleted = useLoad(() => listDeletedApplications(token), [token]);
  const appCount = useLoad(() => countCollection(token, 'applications'), [token]);
  const spCount = useLoad(() => countCollection(token, 'servicePrincipals'), [token]);
  const lastSignIns = useLoad(() => listLastSignInByAppId(token), [token]);
  const audits = useLoad(() => listRecentDirectoryAudits(token, AUDIT_ACTIVITIES, 50), [token]);
  // Tenant-wide lists are only fetched once someone asks for them.
  const tenantApps = useLoad(scope === 'tenant' ? () => listApplications(token) : null, [
    token,
    scope,
  ]);
  const ownerless = useLoad(scope === 'tenant' ? () => listOwnerlessApplications(token) : null, [
    token,
    scope,
  ]);

  const scopedApps = scope === 'tenant' ? tenantApps : owned;

  return (
    <>
      <div className="page-header">
        <div>
          <h1>Overview</h1>
          <div className="subtitle">
            {me === 'loading' ? (
              <span className="spinner" />
            ) : me === 'error' ? (
              'Signed in'
            ) : (
              <>
                Signed in as <strong>{me.displayName}</strong>
              </>
            )}
            {Array.isArray(roles) && roles.length > 0 && (
              <>
                {' · '}
                {roles.join(', ')}
              </>
            )}
            {Array.isArray(roles) && roles.length === 0 && ' · no active admin roles'}
          </div>
        </div>
      </div>

      <div className="stat-grid">
        <StatTile label="App registrations" value={appCount} to="/applications" />
        <StatTile
          label="Owned by me"
          value={Array.isArray(owned) ? owned.length : owned}
          to="/applications?view=owned"
        />
        <StatTile label="Enterprise apps" value={spCount} to="/enterprise-apps" />
        <StatTile
          label="Deleted (recycle bin)"
          value={Array.isArray(deleted) ? deleted.length : deleted}
          to="/applications?view=deleted"
        />
      </div>

      <AttentionCard
        scope={scope}
        onScope={setScope}
        apps={scopedApps}
        deleted={deleted}
        ownerless={scope === 'tenant' ? ownerless : null}
        lastSignIns={lastSignIns}
      />

      <div className="overview-grid">
        <MyAppsCard favorites={favorites} owned={owned} />
        <ActivityCard
          audits={audits}
          meId={me !== 'loading' && me !== 'error' ? me.id : undefined}
          owned={owned}
        />
      </div>
    </>
  );
}

// ---------------------------------------------------------------- Stat tiles

function StatTile({ label, value, to }: { label: string; value: Loadable<number>; to: string }) {
  return (
    <Link to={to} className="stat-tile">
      <div className="stat-value">
        {value === 'loading' ? <span className="spinner" /> : value === 'error' ? '—' : value}
      </div>
      <div className="stat-label">{label}</div>
    </Link>
  );
}

// ----------------------------------------------------------- Needs attention

type Severity = 'danger' | 'warning' | 'info';

interface Issue {
  key: string;
  app: Pick<Application, 'id' | 'displayName' | 'info'>;
  text: string;
  sort: number;
  to: string;
  state?: unknown;
}

interface IssueGroup {
  key: string;
  severity: Severity;
  title: string;
  hint: string;
  issues: Issue[];
}

function credentialIssues(app: Application, now: number) {
  const out: { group: 'expired' | 'expiring' | 'cleanup'; issue: Issue }[] = [];
  const kinds: [string, { endDateTime?: string }[] | undefined][] = [
    ['Secret', app.passwordCredentials],
    ['Certificate', app.keyCredentials],
  ];
  for (const [kind, creds] of kinds) {
    const ends = (creds ?? [])
      .map((c) => (c.endDateTime ? new Date(c.endDateTime).getTime() : NaN))
      .filter(Number.isFinite);
    if (ends.length === 0) continue;
    const expired = ends.filter((e) => e < now);
    const valid = ends.filter((e) => e >= now);
    const healthy = valid.filter((e) => e - now >= SOON_DAYS * DAY);
    const base = {
      app,
      to: `/applications/${app.id}`,
      state: { tab: 'credentials' },
    };
    if (expired.length > 0 && valid.length === 0) {
      const last = Math.max(...expired);
      out.push({
        group: 'expired',
        issue: {
          ...base,
          key: `${app.id}-${kind}-expired`,
          text: `${kind} expired ${relative(last, now)}`,
          sort: last,
        },
      });
    } else if (valid.length > 0 && healthy.length === 0) {
      const last = Math.max(...valid);
      out.push({
        group: 'expiring',
        issue: {
          ...base,
          key: `${app.id}-${kind}-expiring`,
          text: `${kind} expires ${relative(last, now)}`,
          sort: last,
        },
      });
    }
    if (expired.length > 0 && valid.length > 0) {
      const n = expired.length;
      out.push({
        group: 'cleanup',
        issue: {
          ...base,
          key: `${app.id}-${kind}-cleanup`,
          text: `${n} expired ${kind.toLowerCase()}${n === 1 ? '' : 's'} can be removed`,
          sort: -n,
        },
      });
    }
  }
  return out;
}

function buildGroups(
  apps: Application[] | null,
  deleted: Application[] | null,
  ownerless: Application[] | null,
  lastSignIns: Map<string, number> | null,
): IssueGroup[] {
  const now = Date.now();
  const unused: Issue[] = [];
  if (lastSignIns) {
    for (const app of apps ?? []) {
      // Brand-new registrations haven't had the chance to be used yet.
      const created = app.createdDateTime ? new Date(app.createdDateTime).getTime() : NaN;
      if (!Number.isFinite(created) || now - created < UNUSED_DAYS * DAY) continue;
      const last = app.appId ? lastSignIns.get(app.appId) : undefined;
      if (last !== undefined && now - last < UNUSED_DAYS * DAY) continue;
      unused.push({
        key: `${app.id}-unused`,
        app,
        text: last === undefined ? 'No sign-in recorded' : `Last sign-in ${relative(last, now)}`,
        // Never-used first, then the longest idle.
        sort: last ?? 0,
        to: `/applications/${app.id}`,
        state: { tab: 'activity' },
      });
    }
  }
  const expired: Issue[] = [];
  const expiring: Issue[] = [];
  const cleanup: Issue[] = [];
  for (const app of apps ?? []) {
    for (const { group, issue } of credentialIssues(app, now)) {
      (group === 'expired' ? expired : group === 'expiring' ? expiring : cleanup).push(issue);
    }
  }
  const purge: Issue[] = (deleted ?? []).flatMap((d) => {
    const days = daysUntilPurge(d.deletedDateTime);
    if (days === undefined || days > PURGE_WARN_DAYS) return [];
    return [
      {
        key: `${d.id}-purge`,
        app: d,
        text:
          days === 0
            ? 'Permanently removed today'
            : `Permanently removed in ${days} day${days === 1 ? '' : 's'}`,
        sort: days,
        to: '/applications?view=deleted',
      },
    ];
  });
  const noOwner: Issue[] = (ownerless ?? []).map((a) => ({
    key: `${a.id}-owner`,
    app: a,
    text: 'No owners',
    sort: 0,
    to: `/applications/${a.id}`,
  }));

  return [
    {
      key: 'expired',
      severity: 'danger' as const,
      title: 'Expired credentials',
      hint: 'Every secret or certificate of this kind has expired — sign-ins using it fail.',
      issues: expired.sort((a, b) => b.sort - a.sort),
    },
    {
      key: 'expiring',
      severity: 'warning' as const,
      title: `Expiring within ${SOON_DAYS} days`,
      hint: 'No credential of this kind lasts beyond the window — rotate before it runs out.',
      issues: expiring.sort((a, b) => a.sort - b.sort),
    },
    {
      key: 'purge',
      severity: 'warning' as const,
      title: 'About to be purged',
      hint: `Deleted app registrations that leave the recycle bin within ${PURGE_WARN_DAYS} days.`,
      issues: purge.sort((a, b) => a.sort - b.sort),
    },
    {
      key: 'ownerless',
      severity: 'info' as const,
      title: 'Without an owner',
      hint: 'Nobody but admins can manage these app registrations.',
      issues: noOwner.sort((a, b) => name(a).localeCompare(name(b))),
    },
    {
      key: 'unused',
      severity: 'info' as const,
      title: `Unused for ${UNUSED_DAYS}+ days`,
      hint: 'No sign-ins as client or as API — a candidate to clean up or deactivate.',
      issues: unused.sort((a, b) => a.sort - b.sort || name(a).localeCompare(name(b))),
    },
    {
      key: 'cleanup',
      severity: 'info' as const,
      title: 'Expired credentials to clean up',
      hint: 'A valid replacement exists; the expired one is just clutter.',
      issues: cleanup.sort((a, b) => a.sort - b.sort || name(a).localeCompare(name(b))),
    },
  ].filter((g) => g.issues.length > 0);
}

const name = (i: Issue) => i.app.displayName ?? '';

function AttentionCard({
  scope,
  onScope,
  apps,
  deleted,
  ownerless,
  lastSignIns,
}: {
  scope: Scope;
  onScope: (s: Scope) => void;
  apps: Loadable<Application[]>;
  deleted: Loadable<Application[]>;
  ownerless: Loadable<Application[]> | null;
  lastSignIns: Loadable<Map<string, number>>;
}) {
  const groups = useMemo(
    () =>
      buildGroups(
        Array.isArray(apps) ? apps : null,
        Array.isArray(deleted) ? deleted : null,
        ownerless && Array.isArray(ownerless) ? ownerless : null,
        lastSignIns instanceof Map ? lastSignIns : null,
      ),
    [apps, deleted, ownerless, lastSignIns],
  );
  const loading =
    apps === 'loading' ||
    deleted === 'loading' ||
    ownerless === 'loading' ||
    lastSignIns === 'loading';
  const failed: string[] = [];
  if (apps === 'error') failed.push(scope === 'tenant' ? 'app registrations' : 'your apps');
  if (deleted === 'error') failed.push('the recycle bin');
  if (ownerless === 'error') failed.push('app owners');
  if (lastSignIns === 'error') failed.push('sign-in activity (needs AuditLog.Read.All)');

  return (
    <div className="card">
      <div className="card-title-row">
        <h3>Needs attention</h3>
        <div className="segmented" role="group" aria-label="Which apps to check">
          <button className={scope === 'mine' ? 'active' : ''} onClick={() => onScope('mine')}>
            My apps
          </button>
          <button className={scope === 'tenant' ? 'active' : ''} onClick={() => onScope('tenant')}>
            Whole tenant
          </button>
        </div>
      </div>
      {failed.length > 0 && (
        <div className="muted" style={{ fontSize: 13, marginBottom: 8 }}>
          Couldn't load {failed.join(' or ')}.
        </div>
      )}
      {groups.map((g) => (
        <AttentionGroup key={g.key} group={g} />
      ))}
      {loading && (
        <div className="muted" style={{ fontSize: 13 }}>
          <span className="spinner" /> Checking{' '}
          {scope === 'tenant' ? 'every app registration' : 'your apps'}…
        </div>
      )}
      {!loading && groups.length === 0 && failed.length === 0 && (
        <div className="all-clear">
          ✓ Nothing needs attention
          {scope === 'mine' ? ' in the apps you own.' : ' across the tenant.'}
        </div>
      )}
    </div>
  );
}

const PREVIEW = 5;

function AttentionGroup({ group }: { group: IssueGroup }) {
  const [open, setOpen] = useState(false);
  const shown = open ? group.issues : group.issues.slice(0, PREVIEW);
  return (
    <div className={`attention-group sev-${group.severity}`}>
      <div className="attention-head">
        <span className={`attention-count sev-${group.severity}`}>{group.issues.length}</span>
        <div>
          <div className="attention-title">{group.title}</div>
          <div className="muted" style={{ fontSize: 12 }}>
            {group.hint}
          </div>
        </div>
      </div>
      <ul className="item-list">
        {shown.map((i) => (
          <li key={i.key}>
            <AppLink app={i.app} to={i.to} state={i.state} />
            <span className="item-meta">{i.text}</span>
          </li>
        ))}
      </ul>
      {group.issues.length > PREVIEW && (
        <button className="ghost show-more" onClick={() => setOpen(!open)}>
          {open ? 'Show fewer' : `Show all ${group.issues.length}`}
        </button>
      )}
    </div>
  );
}

// ------------------------------------------------------------------ My apps

function MyAppsCard({
  favorites,
  owned,
}: {
  favorites: Loadable<Application[]>;
  owned: Loadable<Application[]>;
}) {
  const sortByName = (list: Application[]) =>
    [...list].sort((a, b) => (a.displayName ?? '').localeCompare(b.displayName ?? ''));
  return (
    <div className="card">
      <h3>Favorites</h3>
      <AppList
        state={favorites}
        render={(list) => sortByName(list).slice(0, 10)}
        empty="Star an app registration to pin it here."
        more={(n) => <Link to="/applications?tag=debble%3AFavorite">All {n} favorites →</Link>}
      />
      <h3 style={{ marginTop: 20 }}>Owned by me</h3>
      <AppList
        state={owned}
        render={(list) => sortByName(list).slice(0, 8)}
        empty="You don't own any app registrations."
        more={(n) => <Link to="/applications?view=owned">All {n} owned →</Link>}
      />
    </div>
  );
}

function AppList({
  state,
  render,
  empty,
  more,
}: {
  state: Loadable<Application[]>;
  render: (list: Application[]) => Application[];
  empty: string;
  more: (total: number) => ReactNode;
}) {
  if (state === 'loading') return <span className="spinner" />;
  if (state === 'error') return <div className="muted">Couldn't load.</div>;
  if (state.length === 0)
    return (
      <div className="muted" style={{ fontSize: 13 }}>
        {empty}
      </div>
    );
  const shown = render(state);
  return (
    <>
      <ul className="item-list">
        {shown.map((a) => (
          <li key={a.id}>
            <AppLink app={a} to={`/applications/${a.id}`} />
          </li>
        ))}
      </ul>
      {state.length > shown.length && <div className="list-more">{more(state.length)}</div>}
    </>
  );
}

function AppLink({
  app,
  to,
  state,
}: {
  app: Pick<Application, 'id' | 'displayName' | 'info'>;
  to: string;
  state?: unknown;
}) {
  return (
    <Link to={to} state={state} className="app-link">
      <AppIcon id={app.id} logoUrl={app.info?.logoUrl} size={20} />
      <span>{app.displayName ?? app.id}</span>
    </Link>
  );
}

// ---------------------------------------------------------- Recent activity

type ActivityFilter = 'all' | 'me' | 'myApps';

function ActivityCard({
  audits,
  meId,
  owned,
}: {
  audits: Loadable<DirectoryAuditEvent[]>;
  meId?: string;
  owned: Loadable<Application[]>;
}) {
  const [filter, setFilter] = useState<ActivityFilter>('all');
  const [limit, setLimit] = useState(10);

  const filtered = useMemo(() => {
    if (!Array.isArray(audits)) return [];
    if (filter === 'me') return audits.filter((e) => meId && e.initiatedBy?.user?.id === meId);
    if (filter === 'myApps') {
      if (!Array.isArray(owned)) return [];
      const ids = new Set(owned.map((a) => a.id));
      const names = new Set(owned.map((a) => a.displayName).filter(Boolean));
      // Service principal events carry the SP's id, not the application's, so
      // those are matched on display name.
      return audits.filter((e) =>
        e.targetResources?.some(
          (t) => (t.id && ids.has(t.id)) || (t.displayName && names.has(t.displayName)),
        ),
      );
    }
    return audits;
  }, [audits, filter, meId, owned]);

  const shown = filtered.slice(0, limit);
  const byDay = groupByDay(shown);

  return (
    <div className="card">
      <div className="card-title-row">
        <h3>Recent activity</h3>
        <div className="segmented" role="group" aria-label="Filter activity">
          {(
            [
              ['all', 'Everything'],
              ['me', 'By me'],
              ['myApps', 'My apps'],
            ] as const
          ).map(([k, label]) => (
            <button
              key={k}
              className={filter === k ? 'active' : ''}
              onClick={() => {
                setFilter(k);
                setLimit(10);
              }}
            >
              {label}
            </button>
          ))}
        </div>
      </div>
      {audits === 'loading' ? (
        <span className="spinner" />
      ) : audits === 'error' || audits.length === 0 ? (
        <div className="muted" style={{ fontSize: 13 }}>
          No recent changes, or the audit log isn't accessible (needs{' '}
          <span className="mono">AuditLog.Read.All</span>).
        </div>
      ) : filtered.length === 0 ? (
        <div className="muted" style={{ fontSize: 13 }}>
          Nothing matching in the last {audits.length} changes.
        </div>
      ) : (
        <>
          {byDay.map(([day, events]) => (
            <div key={day} className="activity-day">
              <div className="activity-day-label">{day}</div>
              <ul className="item-list">
                {events.map((e) => (
                  <ActivityItem key={e.id} event={e} />
                ))}
              </ul>
            </div>
          ))}
          {filtered.length > limit && (
            <button className="ghost show-more" onClick={() => setLimit(limit + 15)}>
              Show more
            </button>
          )}
        </>
      )}
    </div>
  );
}

function groupByDay(events: DirectoryAuditEvent[]): [string, DirectoryAuditEvent[]][] {
  const out = new Map<string, DirectoryAuditEvent[]>();
  const today = new Date().toDateString();
  const yesterday = new Date(Date.now() - DAY).toDateString();
  for (const e of events) {
    const d = e.activityDateTime ? new Date(e.activityDateTime) : null;
    const key = !d
      ? 'Unknown date'
      : d.toDateString() === today
        ? 'Today'
        : d.toDateString() === yesterday
          ? 'Yesterday'
          : d.toLocaleDateString(undefined, { weekday: 'short', day: 'numeric', month: 'short' });
    const list = out.get(key);
    if (list) list.push(e);
    else out.set(key, [e]);
  }
  return [...out.entries()];
}

function ActivityItem({ event }: { event: DirectoryAuditEvent }) {
  const time = event.activityDateTime
    ? new Date(event.activityDateTime).toLocaleTimeString(undefined, {
        hour: '2-digit',
        minute: '2-digit',
      })
    : '';
  const target = event.targetResources?.[0];
  const isDelete = /delete/i.test(event.activityDisplayName ?? '');
  const href = !isDelete && target ? hrefFor(target) : null;
  const targetLabel = target?.displayName ?? target?.id;
  const user = event.initiatedBy?.user;
  const app = event.initiatedBy?.app;
  const who = user?.displayName ?? user?.userPrincipalName ?? app?.displayName ?? app?.appId;
  const failed = event.result && event.result.toLowerCase() !== 'success';

  return (
    <li className="activity-item">
      <span className="activity-time">{time}</span>
      <div className="activity-body">
        <div>
          {event.activityDisplayName ?? 'Change'}
          {targetLabel && (
            <>
              {' · '}
              {href ? <Link to={href}>{targetLabel}</Link> : <strong>{targetLabel}</strong>}
            </>
          )}
          {failed && (
            <>
              {' '}
              <span className="badge expired">{event.result}</span>
            </>
          )}
        </div>
        {who && (
          <div className="muted" style={{ fontSize: 12 }}>
            by {who}
            {!user && app && ' (app)'}
          </div>
        )}
      </div>
    </li>
  );
}

function hrefFor(t: { id?: string; type?: string }): string | null {
  if (!t.id) return null;
  if (t.type === 'Application') return `/applications/${t.id}`;
  if (t.type === 'ServicePrincipal') return `/enterprise-apps/${t.id}`;
  return null;
}

// ------------------------------------------------------------------ helpers

const rtf = new Intl.RelativeTimeFormat(undefined, { numeric: 'auto' });

function relative(when: number, now: number): string {
  const days = Math.round((when - now) / DAY);
  if (Math.abs(days) >= 730) return rtf.format(Math.round(days / 365), 'year');
  if (Math.abs(days) >= 60) return rtf.format(Math.round(days / 30), 'month');
  return rtf.format(days, 'day');
}
