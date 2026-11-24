import { useCallback, useEffect, useMemo, useState, type ReactNode } from 'react';
import { Link, useNavigate, useSearchParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import { useGraphToken } from '../auth/useGraphToken';
import { useCurrentTenantId } from '../auth/useCurrentTenantId';
import { GROUP_WRITE_SCOPES, USER_WRITE_SCOPES } from '../auth/config';
import {
  countGroupMembers,
  countUserApps,
  listLastSignIns,
  listUserGroupIds,
  type UserApps,
  type UserGroups,
  quickSearchGroups,
  isGuest,
  quickSearchUsers,
  type GroupKindFilter,
  type QuickGroup,
  type QuickSearchResult,
  type QuickUser,
  type SearchDepth,
  type UserKindFilter,
} from '../graph/quickSearch';
import { deleteGroup, deleteUser, listVerifiedDomains } from '../graph/principals';
import { groupKind } from '../components/DirectorySearch';
import { entraGroupUrl, entraUserUrl } from '../lib/entraLinks';
import { ExternalLinkIcon } from '../components/ButtonIcons';
import { useConfirm, type ConfirmOptions } from '../components/ConfirmDialog';
import { GroupSiteLink, useGroupSites } from '../components/GroupSites';
import { Identicon } from '../components/Identicon';
import { NewGroup, NewUser } from '../components/PrincipalEditors';
import { relativeTime } from '../components/ActivityTab';
import { QueryInput } from '../components/QueryInput';
import { RemoveButton } from '../components/RemoveButton';
import type { Entity } from '../lib/searchQuery';
import { SortTh, useSort, type SortValue } from '../components/SortableTable';

// Users and Groups: create and delete here; a row opens the detail page
// (/users/:id, /groups/:id) where the object is edited. A tenant can hold far
// more of either than is sensible to load, so one page is fetched and the
// search runs on the server.

const USER_ENTITIES: Entity[] = ['user'];
const GROUP_ENTITIES: Entity[] = ['group'];

const LIST_PAGE: SearchDepth = { top: 200, pages: 1 };

type TokenFn = () => Promise<string>;

interface Column<T> {
  header: string;
  cell: (item: T) => ReactNode;
  // What the column sorts on.
  sort: (item: T) => SortValue;
  className?: string;
}

interface CreatorProps<T> {
  onClose: () => void;
  onCreated: (item: T) => void;
}

interface ListProps<T> {
  title: string;
  subtitle: string;
  // Singular and plural, lower case: "user" / "users".
  noun: string;
  nouns: string;
  placeholder: string;
  // What the search box completes field names for.
  entities: Entity[];
  load: (token: TokenFn, q: string, depth: SearchDepth) => Promise<QuickSearchResult<T>>;
  link: (item: T, tenantId?: string) => string;
  columns: Column<T>[];
  // Extra toolbar controls (e.g. tabs) shown next to the search box.
  filters?: ReactNode;
  // Whether a freshly created item belongs in the current view.
  showsNew?: (item: T) => boolean;
  writeScopes: string[];
  // In-app path of the detail page.
  detail: (item: T) => string;
  Creator: (props: CreatorProps<T>) => ReactNode;
  // Go to the new item's detail page once it exists (the "New user" dialog
  // stays open instead, to show the temporary password).
  openAfterCreate?: boolean;
  remove: (token: TokenFn, id: string) => Promise<void>;
  confirmRemove: (item: T) => ConfirmOptions;
  // A reason this row can't be deleted, if any.
  locked?: (item: T) => string | undefined;
}

function PrincipalList<T extends { id: string; displayName?: string }>({
  title,
  subtitle,
  noun,
  nouns,
  placeholder,
  entities,
  load,
  link,
  columns,
  filters,
  showsNew,
  writeScopes,
  detail,
  Creator,
  openAfterCreate,
  remove,
  confirmRemove,
  locked,
}: ListProps<T>) {
  const token = useGraphToken();
  const writeToken = useGraphToken(writeScopes);
  const tenantId = useCurrentTenantId();
  const nav = useNavigate();
  const { confirm, confirmDialog } = useConfirm();
  // ?q= seeds the search, so the Overview search can hand over to this list.
  const [params] = useSearchParams();
  const [search, setSearch] = useState(() => params.get('q') ?? '');
  const [debounced, setDebounced] = useState(() => search.trim());
  const [result, setResult] = useState<QuickSearchResult<T> | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [creating, setCreating] = useState(false);
  const [busyId, setBusyId] = useState<string | null>(null);
  const [actionError, setActionError] = useState<string | null>(null);

  useEffect(() => {
    const handle = setTimeout(() => setDebounced(search.trim()), 250);
    return () => clearTimeout(handle);
  }, [search]);

  useEffect(() => {
    let cancelled = false;
    setResult(null);
    setError(null);
    load(token, debounced, LIST_PAGE)
      .then((r) => !cancelled && setResult(r))
      .catch((e: unknown) => {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e));
      });
    return () => {
      cancelled = true;
    };
  }, [token, debounced, load]);

  // Graph's search index trails behind writes, so a reload right after a
  // change could miss it. The list is patched in place instead.
  const onCreated = (item: T) => {
    if (openAfterCreate) {
      nav(detail(item));
      return;
    }
    setResult((prev) =>
      !prev || (showsNew && !showsNew(item)) ? prev : { ...prev, items: [item, ...prev.items] },
    );
  };

  const onRemove = async (item: T) => {
    if (!(await confirm(confirmRemove(item)))) return;
    setBusyId(item.id);
    setActionError(null);
    try {
      await remove(writeToken, item.id);
      setResult((prev) =>
        prev ? { ...prev, items: prev.items.filter((i) => i.id !== item.id) } : prev,
      );
    } catch (e: unknown) {
      setActionError(
        `Couldn't delete ${item.displayName ?? item.id}: ${e instanceof Error ? e.message : String(e)}`,
      );
    } finally {
      setBusyId(null);
    }
  };

  const { sorted, sort } = useSort(result?.items ?? [], {
    name: (i) => i.displayName,
    ...Object.fromEntries(columns.map((c) => [sortKey(c.header), c.sort])),
  });

  return (
    <>
      <div className="page-header">
        <div>
          <h1>{title}</h1>
          <div className="subtitle">{subtitle}</div>
        </div>
        <button className="primary" onClick={() => setCreating(true)}>
          + New {noun}
        </button>
      </div>

      <div className="toolbar">
        <QueryInput
          className="search"
          entities={entities}
          placeholder={placeholder}
          value={search}
          onChange={setSearch}
        />
        {filters}
        <div className="grow" />
        {result && (
          <span className="muted">
            {result.more
              ? `First ${result.items.length} — search to narrow down`
              : `${result.items.length} ${debounced ? 'matches' : 'total'}`}
          </span>
        )}
      </div>

      {error && <div className="card error">{error}</div>}
      {actionError && <div className="card error">{actionError}</div>}
      {!result && !error && <span className="spinner" />}

      {result && (
        <div className="card" style={{ padding: 0 }}>
          <table className="table">
            <thead>
              <tr>
                <SortTh id="name" sort={sort}>
                  Display name
                </SortTh>
                {columns.map((c) => (
                  <SortTh key={c.header} id={sortKey(c.header)} sort={sort}>
                    {c.header}
                  </SortTh>
                ))}
                <th></th>
              </tr>
            </thead>
            <tbody>
              {sorted.map((item) => {
                const name = item.displayName ?? item.id;
                const lock = locked?.(item);
                return (
                  <tr key={item.id} onClick={() => nav(detail(item))}>
                    <td>
                      <div className="row" style={{ gap: 10, alignItems: 'center' }}>
                        <Identicon id={item.id} title={name} />
                        {/* A real link too, so it can be opened in a new tab. */}
                        <Link
                          to={detail(item)}
                          style={{ color: 'inherit' }}
                          onClick={(e) => e.stopPropagation()}
                        >
                          {name}
                        </Link>
                      </div>
                    </td>
                    {columns.map((c) => (
                      <td key={c.header} className={c.className}>
                        {c.cell(item)}
                      </td>
                    ))}
                    <td className="row-actions" onClick={(e) => e.stopPropagation()}>
                      <div className="row" style={{ gap: 6, justifyContent: 'flex-end' }}>
                        <a
                          href={link(item, tenantId)}
                          target="_blank"
                          rel="noreferrer"
                          className="muted"
                          style={{ display: 'inline-flex', padding: 4 }}
                          title="Open in the Entra admin center"
                          aria-label={`Open ${name} in the Entra admin center`}
                        >
                          <ExternalLinkIcon />
                        </a>
                        <RemoveButton
                          label={lock ?? `Delete ${noun}`}
                          disabled={Boolean(lock)}
                          busy={busyId === item.id}
                          onClick={() => void onRemove(item)}
                        />
                      </div>
                    </td>
                  </tr>
                );
              })}
              {result.items.length === 0 && (
                <tr>
                  <td colSpan={columns.length + 2}>
                    <div className="empty">
                      {debounced ? `No ${nouns} match your search.` : `No ${nouns} found.`}
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {creating && <Creator onClose={() => setCreating(false)} onCreated={onCreated} />}
      {confirmDialog}
    </>
  );
}

const dash = <span className="muted">—</span>;

// "User principal name" → "user-principal-name", for ?sort=.
const sortKey = (header: string) => header.toLowerCase().replace(/\s+/g, '-');
const time = (iso?: string | null) => (iso ? new Date(iso).getTime() || null : null);

// Date in the list, full timestamp on hover.
function created(value?: string | null): ReactNode {
  const d = value ? new Date(value) : null;
  if (!d || Number.isNaN(d.getTime())) return dash;
  return (
    <span title={d.toLocaleString()}>
      {d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' })}
    </span>
  );
}

// Group and app counts and the last sign-in arrive after the list (see
// Users), so they are columns of their own slotted in per render.
const userColumns = (late: Column<QuickUser>[]): Column<QuickUser>[] => [
  {
    header: 'User principal name',
    cell: (u) => u.userPrincipalName ?? dash,
    sort: (u) => u.userPrincipalName,
  },
  {
    header: 'Type',
    sort: (u) => (isGuest(u) ? 'Guest' : (u.userType ?? 'Member')),
    cell: (u) =>
      isGuest(u) ? (
        <span
          className="badge pending"
          title={u.userType === 'Guest' ? undefined : 'External account (#EXT#) set to Member'}
        >
          Guest
        </span>
      ) : (
        <span className="muted">{u.userType ?? 'Member'}</span>
      ),
  },
  {
    header: 'Enabled',
    sort: (u) => u.accountEnabled !== false,
    cell: (u) =>
      u.accountEnabled === false ? (
        <span className="badge">Disabled</span>
      ) : (
        <span className="badge granted">Enabled</span>
      ),
  },
  ...late,
  {
    header: 'Created',
    cell: (u) => created(u.createdDateTime),
    sort: (u) => time(u.createdDateTime),
    className: 'muted nowrap',
  },
  { header: 'Object ID', cell: (u) => u.id, sort: (u) => u.id, className: 'mono' },
];

// The member count arrives after the list (see Groups), so it's a column of
// its own slotted in per render.
const groupColumns = (
  members: Column<QuickGroup>,
  site: Column<QuickGroup> | null,
): Column<QuickGroup>[] => [
  { header: 'Type', cell: (g) => groupKind(g), sort: (g) => groupKind(g), className: 'muted' },
  members,
  { header: 'Mail', cell: (g) => g.mail ?? dash, sort: (g) => g.mail },
  ...(site ? [site] : []),
  {
    header: 'Created',
    cell: (g) => created(g.createdDateTime),
    sort: (g) => time(g.createdDateTime),
    className: 'muted nowrap',
  },
  { header: 'Object ID', cell: (g) => g.id, sort: (g) => g.id, className: 'mono' },
];

const plural = (n: number, one: string) => `${n.toLocaleString()} ${one}${n === 1 ? '' : 's'}`;

const GROUPS_COLUMN = (counts: Map<string, UserGroups | null>): Column<QuickUser> => ({
  header: 'Groups',
  sort: (u) => counts.get(u.id)?.directCount,
  cell: (u) => {
    const g = counts.get(u.id);
    if (g === undefined) return <span className="spinner" />;
    if (g === null) return dash;
    const nested = Math.max(g.all.length - g.direct.length, 0);
    return (
      <span
        title={
          `Direct member of ${plural(g.directCount, 'group')}` +
          (nested > 0 ? ` · in ${nested.toLocaleString()} more through nesting` : '')
        }
      >
        {g.directCount.toLocaleString()}
      </span>
    );
  },
});

const APPS_COLUMN = (counts: Map<string, UserApps | null>): Column<QuickUser> => ({
  header: 'Enterprise apps',
  sort: (u) => counts.get(u.id)?.total,
  cell: (u) => {
    const a = counts.get(u.id);
    if (a === undefined) return <span className="spinner" />;
    if (a === null) return dash;
    return (
      <span
        title={
          `${plural(a.direct, 'app')} assigned to the user · ${a.viaGroups.toLocaleString()} more through groups` +
          (a.nestedOnly > 0
            ? `\n${plural(a.nestedOnly, 'app')} assigned to a parent group only — not counted: Entra doesn't pass an app assignment on to nested groups`
            : '') +
          (a.more ? '\nAt least this many: an assignment list was too long to read in full' : '')
        }
      >
        {a.total.toLocaleString()}
        {a.more && '+'}
        {a.nestedOnly > 0 && <span className="muted"> (+{a.nestedOnly})</span>}
      </span>
    );
  },
});

type UserTab = 'all' | UserKindFilter;
const USER_TABS: { tab: UserTab; label: string; title: string }[] = [
  { tab: 'all', label: 'All', title: 'Every user in this tenant' },
  { tab: 'Member', label: 'Members', title: 'Users who belong to this tenant' },
  {
    tab: 'Guest',
    label: 'Guests',
    title:
      'Users invited from outside this tenant, including external (#EXT#) accounts set to Member',
  },
  {
    tab: 'shared',
    label: 'Shared mailboxes',
    title:
      'Mailbox accounts nobody signs in to: disabled members with a mail address and no ' +
      'licence. Room and equipment mailboxes show up here too.',
  },
];
const USER_NOUNS: Record<UserTab, string> = {
  all: 'users',
  Member: 'members',
  Guest: 'guests',
  shared: 'shared mailboxes',
};

export function Users() {
  const { instance, accounts } = useMsal();
  const meId = (instance.getActiveAccount() ?? accounts[0])?.localAccountId;
  // ?type=Member|Guest|shared keeps the tab in the URL, so it survives a reload.
  const [params, setParams] = useSearchParams();
  const typeParam = params.get('type');
  const tab: UserTab = USER_TABS.find((t) => t.tab === typeParam)?.tab ?? 'all';
  const setTab = (next: UserTab) =>
    setParams(
      (prev) => {
        const p = new URLSearchParams(prev);
        if (next === 'all') p.delete('type');
        else p.set('type', next);
        return p;
      },
      { replace: true },
    );
  // The tenant's onmicrosoft.com domains, needed to put "#EXT#" accounts
  // under Guests. If they can't be read, the tabs split on userType alone.
  const token = useGraphToken();
  const [extDomains, setExtDomains] = useState<string[]>([]);
  useEffect(() => {
    let cancelled = false;
    listVerifiedDomains(token)
      .then((all) => all.filter((d) => /\.onmicrosoft\.com$/i.test(d)))
      .catch(() => [] as string[])
      .then((list) => !cancelled && list.length > 0 && setExtDomains(list));
    return () => {
      cancelled = true;
    };
  }, [token]);
  // Last sign-ins are fetched after the rows are in. In the map: undefined =
  // still loading, null = never signed in. 'unavailable' = can't be read here.
  const [signIns, setSignIns] = useState<Map<string, string | null> | 'unavailable'>(
    () => new Map(),
  );
  // Group memberships and app assignments, also fetched after the rows are
  // in: undefined = still loading, null = couldn't be read.
  const [groupCounts, setGroupCounts] = useState<Map<string, UserGroups | null>>(() => new Map());
  const [appCounts, setAppCounts] = useState<Map<string, UserApps | null>>(() => new Map());
  const loadAccess = useCallback(async (t: TokenFn, ids: string[]) => {
    const fill =
      <V,>(found: Map<string, V>) =>
      (prev: Map<string, V | null>) => {
        const next = new Map(prev);
        for (const id of ids) next.set(id, found.get(id) ?? null);
        return next;
      };
    const groups = await listUserGroupIds(t, ids).catch(() => new Map<string, UserGroups>());
    setGroupCounts(fill(groups));
    const apps = await countUserApps(t, ids, groups).catch(() => new Map<string, UserApps>());
    setAppCounts(fill(apps));
  }, []);
  const load = useCallback(
    async (t: TokenFn, q: string, depth: SearchDepth) => {
      const result = await quickSearchUsers(
        t,
        q,
        depth,
        tab === 'all' ? undefined : tab,
        extDomains,
      );
      const ids = result.items.map((u) => u.id);
      if (ids.length > 0) void loadAccess(t, ids);
      if (ids.length > 0)
        void listLastSignIns(t, ids)
          .catch(() => null)
          .then((found) =>
            setSignIns((prev) => {
              if (!found) return 'unavailable';
              const next = new Map(prev === 'unavailable' ? [] : prev);
              // A user whose lookup failed is shown as unknown, not as loading.
              for (const id of ids) next.set(id, found.has(id) ? found.get(id)! : '');
              return next;
            }),
          );
      return result;
    },
    [tab, extDomains, loadAccess],
  );
  const columns = useMemo(
    () =>
      userColumns([
        GROUPS_COLUMN(groupCounts),
        APPS_COLUMN(appCounts),
        {
          header: 'Last sign-in',
          className: 'muted nowrap',
          sort: (u) => (signIns === 'unavailable' ? null : time(signIns.get(u.id))),
          cell: (u) => {
            if (signIns === 'unavailable')
              return (
                <span title="Sign-in activity needs AuditLog.Read.All and Entra ID P1 or P2">
                  —
                </span>
              );
            const at = signIns.get(u.id);
            if (at === undefined) return <span className="spinner" />;
            if (at === null) return 'Never';
            if (at === '') return dash;
            return <span title={new Date(at).toLocaleString()}>{relativeTime(at)}</span>;
          },
        },
      ]),
    [signIns, groupCounts, appCounts],
  );

  return (
    <PrincipalList
      title="Users"
      subtitle="Members, guests and shared mailboxes in this tenant"
      noun="user"
      nouns={USER_NOUNS[tab]}
      placeholder="Search by name, UPN, mail or object ID…"
      entities={USER_ENTITIES}
      load={load}
      link={(u, tenantId) => entraUserUrl(u.id, tenantId)}
      columns={columns}
      filters={
        <div className="tabs" style={{ margin: 0, borderBottom: 'none' }}>
          {USER_TABS.map((t) => (
            <button
              key={t.tab}
              className={tab === t.tab ? 'active' : ''}
              title={t.title}
              onClick={() => setTab(t.tab)}
            >
              {t.label}
            </button>
          ))}
        </div>
      }
      // New users are created as enabled members.
      showsNew={() => tab === 'all' || tab === 'Member'}
      writeScopes={USER_WRITE_SCOPES}
      detail={(u) => `/users/${u.id}`}
      Creator={NewUser}
      remove={deleteUser}
      locked={(u) => (u.id === meId ? "You can't delete your own account" : undefined)}
      confirmRemove={(u) => ({
        title: `Delete ${u.displayName ?? 'this user'}?`,
        message: (
          <p style={{ margin: 0, fontSize: 14 }}>
            <strong>{u.userPrincipalName ?? u.id}</strong> will no longer be able to sign in and
            loses access to every app. Deleted users can be restored from the Entra admin center for
            30 days.
          </p>
        ),
        confirmLabel: 'Delete user',
      })}
    />
  );
}

type GroupTab = 'all' | GroupKindFilter;
const GROUP_TABS: { tab: GroupTab; label: string; title: string }[] = [
  { tab: 'all', label: 'All', title: 'Every group in this tenant' },
  {
    tab: 'security',
    label: 'Security',
    title: 'Security groups, including mail-enabled ones — used to grant access',
  },
  {
    tab: 'unified',
    label: 'Microsoft 365',
    title: 'Microsoft 365 (unified) groups — shared mailbox, calendar and files',
  },
  {
    tab: 'distribution',
    label: 'Distribution',
    title: 'Distribution lists — mail only, not usable for granting access',
  },
];
const GROUP_NOUNS: Record<GroupTab, string> = {
  all: 'groups',
  security: 'security groups',
  unified: 'Microsoft 365 groups',
  distribution: 'distribution groups',
};

const isUnified = (g: QuickGroup) => Boolean(g.groupTypes?.includes('Unified'));

// Which tab a group belongs under; mirrors the server-side filters.
function groupTab(g: QuickGroup): GroupKindFilter | undefined {
  if (g.groupTypes?.includes('Unified')) return 'unified';
  if (g.securityEnabled) return 'security';
  return g.mailEnabled ? 'distribution' : undefined;
}

export function Groups() {
  // ?type=security|unified|distribution keeps the tab in the URL, so it survives a reload.
  const [params, setParams] = useSearchParams();
  const typeParam = params.get('type');
  const tab: GroupTab = GROUP_TABS.find((t) => t.tab === typeParam)?.tab ?? 'all';
  const setTab = (next: GroupTab) =>
    setParams(
      (prev) => {
        const p = new URLSearchParams(prev);
        if (next === 'all') p.delete('type');
        else p.set('type', next);
        return p;
      },
      { replace: true },
    );
  // Member counts are fetched after the rows are in, so the list doesn't
  // wait for them; undefined = still loading, null = couldn't be read.
  const [counts, setCounts] = useState<Map<string, number | null>>(() => new Map());
  const { sites, needsConsent, request: requestSites, grant: grantSites } = useGroupSites();
  const load = useCallback(
    async (token: TokenFn, q: string, depth: SearchDepth) => {
      const result = await quickSearchGroups(token, q, depth, tab === 'all' ? undefined : tab);
      const ids = result.items.map((g) => g.id);
      requestSites(result.items.filter(isUnified).map((g) => g.id));
      void countGroupMembers(token, ids)
        .catch(() => new Map<string, number>())
        .then((found) =>
          setCounts((prev) => {
            const next = new Map(prev);
            for (const id of ids) next.set(id, found.get(id) ?? null);
            return next;
          }),
        );
      return result;
    },
    [tab, requestSites],
  );
  // Only Microsoft 365 groups have a site, so the other tabs skip the column.
  const showSites = tab === 'all' || tab === 'unified';
  const columns = useMemo(
    () =>
      groupColumns(
        {
          header: 'Members',
          sort: (g) => counts.get(g.id),
          cell: (g) => {
            const n = counts.get(g.id);
            if (n === undefined) return <span className="spinner" />;
            return n === null ? dash : <span title="Direct members">{n.toLocaleString()}</span>;
          },
        },
        showSites
          ? {
              header: 'SharePoint site',
              sort: (g) => sites.get(g.id)?.url,
              cell: (g) =>
                isUnified(g) ? (
                  <GroupSiteLink
                    site={sites.get(g.id)}
                    needsConsent={needsConsent}
                    onGrant={grantSites}
                  />
                ) : (
                  dash
                ),
            }
          : null,
      ),
    [counts, showSites, sites, needsConsent, grantSites],
  );

  return (
    <PrincipalList
      title="Groups"
      subtitle="Security, Microsoft 365 and distribution groups"
      noun="group"
      nouns={GROUP_NOUNS[tab]}
      placeholder="Search by name, mail or object ID…"
      entities={GROUP_ENTITIES}
      load={load}
      link={(g, tenantId) => entraGroupUrl(g.id, tenantId)}
      columns={columns}
      filters={
        <div className="tabs" style={{ margin: 0, borderBottom: 'none' }}>
          {GROUP_TABS.map((t) => (
            <button
              key={t.tab}
              className={tab === t.tab ? 'active' : ''}
              title={t.title}
              onClick={() => setTab(t.tab)}
            >
              {t.label}
            </button>
          ))}
        </div>
      }
      // A group made or edited here only stays if it fits the open tab.
      showsNew={(g) => tab === 'all' || groupTab(g) === tab}
      writeScopes={GROUP_WRITE_SCOPES}
      detail={(g) => `/groups/${g.id}`}
      Creator={NewGroup}
      openAfterCreate
      remove={deleteGroup}
      confirmRemove={(g) => {
        const restorable = g.groupTypes?.includes('Unified');
        return {
          title: `Delete ${g.displayName ?? 'this group'}?`,
          message: (
            <p style={{ margin: 0, fontSize: 14 }}>
              Everyone who gets access to apps or resources through this group loses it.{' '}
              {restorable ? (
                'Microsoft 365 groups can be restored from the Entra admin center for 30 days.'
              ) : (
                <strong>This is permanent — only Microsoft 365 groups can be restored.</strong>
              )}
            </p>
          ),
          confirmLabel: 'Delete group',
        };
      }}
    />
  );
}
