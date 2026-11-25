import { useEffect, useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import {
  listObjectAudits,
  type DirectoryAuditEvent,
  type DirectoryAuditModifiedProperty,
  type ObjectAudits,
} from '../graph/directoryAudits';
import { resolveDirectoryObjects } from '../graph/directoryObjects';
import { relativeTime } from './ActivityTab';

// Changes tab shared by app registrations and enterprise apps: what the
// directory audit log recorded about the object — when, what, by whom, and the
// old and new value of each changed property. The log keeps about 30 days
// (7 on free tenants).

export interface ChangeTarget {
  id: string;
  // Shown per row when there is more than one target, e.g. "App registration".
  label: string;
}

interface Change {
  name: string;
  from: string;
  to: string;
  // In-app page of the object the value names (an owner), if it has one.
  href?: string;
}

// Audit values are JSON-encoded: "\"x\"" for a string, "[\"a\",\"b\"]" for a
// list. Unwrap to what a person would read.
function readable(raw: string | null | undefined): string {
  if (raw == null || raw === '') return '';
  let v: unknown;
  try {
    v = JSON.parse(raw);
  } catch {
    return raw;
  }
  const one = (x: unknown) => (typeof x === 'string' ? x : JSON.stringify(x));
  if (Array.isArray(v)) return v.map(one).join(', ');
  return v == null ? '' : one(v);
}

// "Included Updated Properties" only repeats the names of the other entries.
const NOISE = new Set(['Included Updated Properties']);

const TARGET_PATH: Record<string, string> = {
  user: '/users',
  serviceprincipal: '/enterprise-apps',
  group: '/groups',
};

// Owner events don't change a property: they name the app and the owner, and
// the log pads them with the app's own id, name and app id as "empty → value".
// Say who was added or removed instead of repeating what the app is.
function ownerChanges(
  e: DirectoryAuditEvent,
  ids: Set<string>,
  names: Record<string, string>,
): Change[] | null {
  const activity = e.activityDisplayName ?? '';
  if (!/owner/i.test(activity)) return null;
  const owners = (e.targetResources ?? []).filter((t) => t.id && !ids.has(t.id));
  if (owners.length === 0) return null;
  const removed = /remove|delete/i.test(activity);
  return owners.map((t) => {
    // Service principals often come without a name; the log then only has
    // their app id, tucked away in a TargetId property.
    const fallback = (t.modifiedProperties ?? []).find((p) =>
      p.displayName?.startsWith('TargetId.'),
    );
    const name =
      t.displayName ||
      t.userPrincipalName ||
      names[t.id!] ||
      readable(fallback?.newValue ?? fallback?.oldValue) ||
      t.id!;
    const path = TARGET_PATH[(t.type ?? '').toLowerCase()];
    return {
      name: 'Owner',
      from: removed ? name : '',
      to: removed ? '' : name,
      href: path ? `${path}/${t.id}` : undefined,
    };
  });
}

function changesOf(
  e: DirectoryAuditEvent,
  ids: Set<string>,
  names: Record<string, string>,
): Change[] {
  const owner = ownerChanges(e, ids, names);
  if (owner) return owner;
  const targets = e.targetResources ?? [];
  // Prefer what changed on this object; some events only carry the change on
  // the other party (e.g. the user who was assigned).
  const own = targets.filter((t) => t.id && ids.has(t.id));
  const props = (ts: typeof targets): DirectoryAuditModifiedProperty[] =>
    ts.flatMap((t) => t.modifiedProperties ?? []);
  const list = props(own).length > 0 ? props(own) : props(targets);
  return list
    .filter((p) => p.displayName && !NOISE.has(p.displayName))
    .map((p) => ({ name: p.displayName!, from: readable(p.oldValue), to: readable(p.newValue) }))
    .filter((c) => c.from !== c.to);
}

function Actor({ event }: { event: DirectoryAuditEvent }) {
  const user = event.initiatedBy?.user;
  const app = event.initiatedBy?.app;
  if (user?.id || user?.userPrincipalName || user?.displayName) {
    const name = user.displayName || user.userPrincipalName || user.id;
    return user.id ? (
      <Link to={`/users/${user.id}`} title={user.userPrincipalName ?? undefined}>
        {name}
      </Link>
    ) : (
      <span>{name}</span>
    );
  }
  if (app?.displayName || app?.appId) {
    return (
      <span title="An application made this change">
        {app.displayName ?? app.appId} <span className="muted">(app)</span>
      </span>
    );
  }
  return <span className="muted">Unknown</span>;
}

const CLIP = 140;

function Value({ text, open, href }: { text: string; open: boolean; href?: string }) {
  if (!text) return <span className="muted">empty</span>;
  if (href) return <Link to={href}>{text}</Link>;
  return <code>{open || text.length <= CLIP ? text : `${text.slice(0, CLIP)}…`}</code>;
}

function Changes({ changes }: { changes: Change[] }) {
  const [open, setOpen] = useState(false);
  if (changes.length === 0) return <span className="muted">—</span>;
  const long = changes.some((c) => c.from.length > CLIP || c.to.length > CLIP);
  return (
    <div className="changes">
      {changes.map((c, i) => (
        <div key={i} className="change">
          <span className="change-name">{c.name}</span>
          <Value text={c.from} open={open} href={c.href} />
          <span className="muted" aria-label="changed to">
            →
          </span>
          <Value text={c.to} open={open} href={c.href} />
        </div>
      ))}
      {long && (
        <button className="link" onClick={() => setOpen((o) => !o)}>
          {open ? 'Show less' : 'Show full values'}
        </button>
      )}
    </div>
  );
}

export function ChangeHistory({ targets }: { targets: ChangeTarget[] }) {
  const token = useGraphToken();
  // undefined = loading, null = the audit log can't be read.
  const [audits, setAudits] = useState<ObjectAudits | null | undefined>(undefined);
  const [error, setError] = useState<string | null>(null);
  const [filter, setFilter] = useState('');
  // Off: only real changes, each kind once. Automated deployments re-apply the
  // same thing on every run, which otherwise buries what actually changed.
  const [everything, setEverything] = useState(false);
  const key = targets.map((t) => t.id).join(',');

  useEffect(() => {
    let cancelled = false;
    setAudits(undefined);
    setError(null);
    listObjectAudits(token, key.split(',').filter(Boolean))
      .then((a) => !cancelled && setAudits(a))
      .catch((e) => !cancelled && setError(e instanceof Error ? e.message : String(e)));
    return () => {
      cancelled = true;
    };
  }, [token, key]);

  // Names of owners the log only knows by id (service principals, mostly).
  const [names, setNames] = useState<Record<string, string>>({});
  useEffect(() => {
    const own = new Set(key.split(','));
    const unnamed = (audits?.events ?? [])
      .filter((e) => /owner/i.test(e.activityDisplayName ?? ''))
      .flatMap((e) => e.targetResources ?? [])
      .filter((t) => t.id && !own.has(t.id) && !t.displayName && !t.userPrincipalName)
      .map((t) => t.id!);
    if (unnamed.length === 0) return;
    let cancelled = false;
    void resolveDirectoryObjects(token, unnamed).then((found) => {
      if (cancelled) return;
      setNames(Object.fromEntries(Object.entries(found).map(([id, p]) => [id, p.displayName])));
    });
    return () => {
      cancelled = true;
    };
  }, [token, audits, key]);

  const rows = useMemo(() => {
    const ids = new Set(targets.map((t) => t.id));
    return (audits?.events ?? []).map((event) => {
      const changes = changesOf(event, ids, names);
      const on = targets.filter((t) => event.targetResources?.some((r) => r.id === t.id));
      const by = event.initiatedBy;
      const text = [
        event.activityDisplayName,
        event.category,
        by?.user?.displayName,
        by?.user?.userPrincipalName,
        by?.app?.displayName,
        ...changes.flatMap((c) => [c.name, c.from, c.to]),
      ]
        .join(' ')
        .toLowerCase();
      return { event, changes, on, text };
    });
    // `key` stands in for `targets`, which is a new array on every render.
  }, [audits, key, names]);

  if (error) return <div className="error">{error}</div>;
  if (audits === undefined)
    return (
      <div className="empty">
        <span className="spinner" />
      </div>
    );
  if (audits === null)
    return (
      <div className="empty">
        The audit log can't be read. It needs the AuditLog.Read.All permission and a role such as
        Reports Reader or Security Reader.
      </div>
    );

  const q = filter.trim().toLowerCase();
  const matching = q ? rows.filter((r) => r.text.includes(q)) : rows;
  // Newest first, as the log returns them. Unless every event is asked for,
  // events that recorded no change are dropped (failures stay) and identical
  // ones — same activity, by the same actor, with the same change — collapse
  // into the most recent, which then says how often and since when.
  type Row = (typeof rows)[number] & { repeats: number; first?: string };
  const shown: Row[] = [];
  if (everything) {
    shown.push(...matching.map((r) => ({ ...r, repeats: 1 })));
  } else {
    const seen = new Map<string, Row>();
    for (const r of matching) {
      const failed = Boolean(r.event.result) && r.event.result !== 'success';
      if (r.changes.length === 0 && !failed) continue;
      const by = r.event.initiatedBy;
      const signature = JSON.stringify([
        r.event.activityDisplayName,
        r.event.result,
        by?.user?.id ?? by?.user?.userPrincipalName ?? by?.app?.appId ?? by?.app?.displayName,
        r.on.map((t) => t.id),
        r.changes,
      ]);
      const earlier = seen.get(signature);
      if (earlier) {
        earlier.repeats += 1;
        earlier.first = r.event.activityDateTime;
      } else {
        const row: Row = { ...r, repeats: 1 };
        seen.set(signature, row);
        shown.push(row);
      }
    }
  }
  const folded = matching.length - shown.length;

  return (
    <>
      <div className="toolbar">
        <input
          className="search"
          type="search"
          placeholder="Filter by activity, person or property…"
          aria-label="Filter changes"
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
        />
        <label
          className="row"
          style={{ gap: 8, alignItems: 'center', margin: '0 12px', whiteSpace: 'nowrap' }}
        >
          <input
            type="checkbox"
            style={{ width: 'auto', flexShrink: 0, margin: 0 }}
            checked={everything}
            onChange={(e) => setEverything(e.target.checked)}
          />
          Show every event
        </label>
        <span className="muted">
          {shown.length.toLocaleString()} {everything ? 'events' : 'changes'}
          {folded > 0 && ` · ${folded.toLocaleString()} repeats and events without a change hidden`}
          {audits.more ? ' (most recent only)' : ''} · the audit log keeps about 30 days
        </span>
      </div>
      {shown.length === 0 ? (
        <div className="empty">
          {rows.length === 0
            ? 'Nothing was changed in the period the audit log covers.'
            : matching.length === 0
              ? 'No changes match the filter.'
              : 'Only events without a recorded change — tick “Show every event” to see them.'}
        </div>
      ) : (
        <div className="card">
          <table className="table">
            <thead>
              <tr>
                <th>When</th>
                <th>Activity</th>
                <th>Changed by</th>
                <th>Changes</th>
              </tr>
            </thead>
            <tbody>
              {shown.map(({ event, changes, on, repeats, first }) => (
                <tr key={event.id} className="static">
                  <td className="muted nowrap">
                    {event.activityDateTime ? (
                      <span title={new Date(event.activityDateTime).toLocaleString()}>
                        {relativeTime(event.activityDateTime)}
                      </span>
                    ) : (
                      '—'
                    )}
                    {repeats > 1 && (
                      <div
                        style={{ fontSize: 12 }}
                        title={first ? `First: ${new Date(first).toLocaleString()}` : undefined}
                      >
                        {repeats}× {first ? `since ${relativeTime(first)}` : ''}
                      </div>
                    )}
                  </td>
                  <td>
                    {event.activityDisplayName ?? '—'}
                    {event.result && event.result !== 'success' && (
                      <>
                        {' '}
                        <span className="badge expired">{event.result}</span>
                      </>
                    )}
                    {targets.length > 1 && on.length > 0 && (
                      <div className="muted" style={{ fontSize: 12 }}>
                        {on.map((t) => t.label).join(', ')}
                      </div>
                    )}
                  </td>
                  <td>
                    <Actor event={event} />
                  </td>
                  <td>
                    <Changes changes={changes} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </>
  );
}
