import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { Link, useLocation, useNavigate, useSearchParams } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { useCurrentTenantId } from '../auth/useCurrentTenantId';
import {
  FULL_SEARCH,
  quickSearchApplications,
  quickSearchGroups,
  quickSearchServicePrincipals,
  isGuest,
  quickSearchUsers,
  type QuickSearchResult,
} from '../graph/quickSearch';
import { isFavorite } from '../lib/tags';
import { AppIcon } from '../components/AppIcon';
import {
  AppRegistrationIcon,
  EnterpriseAppIcon,
  ExternalLinkIcon,
  GroupIcon,
  UserIcon,
} from '../components/ButtonIcons';
import { Highlight, groupKind } from '../components/DirectorySearch';
import { Identicon } from '../components/Identicon';
import { QueryInput } from '../components/QueryInput';
import { SortTh, useSort } from '../components/SortableTable';
import type { Entity } from '../lib/searchQuery';

// Full results for a directory search (/search?q=…): everything the Overview
// search box finds, in one table that can be narrowed by type and by a few
// properties of the hits. Query and filters live in the URL, so coming back
// from an app returns to the same view.

const MIN_CHARS = 2;
const ALL_ENTITIES: Entity[] = ['app', 'sp', 'user', 'group'];

type Kind = 'app' | 'sp' | 'user' | 'group';

const KINDS: { kind: Kind; label: string; plural: string; icon: ReactNode }[] = [
  {
    kind: 'app',
    label: 'App registration',
    plural: 'App registrations',
    icon: <AppRegistrationIcon />,
  },
  { kind: 'sp', label: 'Enterprise app', plural: 'Enterprise apps', icon: <EnterpriseAppIcon /> },
  { kind: 'user', label: 'User', plural: 'Users', icon: <UserIcon /> },
  { kind: 'group', label: 'Group', plural: 'Groups', icon: <GroupIcon /> },
];

interface Row {
  key: string;
  kind: Kind;
  name: string;
  sub?: string;
  // The identifier people paste around: appId for apps, object id otherwise.
  id: string;
  to: string;
  external?: boolean;
  icon: ReactNode;
  // Properties the result can be narrowed by; also shown as badges.
  flags: string[];
}

const DANGER_FLAGS = new Set(['Disabled']);

interface Bucket {
  loading: boolean;
  error: boolean;
  rows: Row[];
  more: boolean;
}
const EMPTY: Bucket = { loading: false, error: false, rows: [], more: false };
const LOADING: Bucket = { ...EMPTY, loading: true };

export function Search() {
  const token = useGraphToken();
  const tenantId = useCurrentTenantId();
  const nav = useNavigate();
  const location = useLocation();
  const [params, setParams] = useSearchParams();
  const setParam = (key: string, value: string | null) =>
    setParams(
      (prev) => {
        const next = new URLSearchParams(prev);
        if (value) next.set(key, value);
        else next.delete(key);
        return next;
      },
      { replace: true },
    );

  const query = (params.get('q') ?? '').trim();
  const typeParam = params.get('type');
  const type = KINDS.some((k) => k.kind === typeParam) ? (typeParam as Kind) : null;
  const flag = params.get('show');

  const [text, setText] = useState(query);
  const lastPushed = useRef(query);
  useEffect(() => {
    const q = text.trim();
    if (q === query) return;
    const handle = setTimeout(() => {
      lastPushed.current = q;
      setParam('q', q || null);
    }, 300);
    return () => clearTimeout(handle);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [text, query]);
  // The URL changed from outside (browser back): show its search.
  useEffect(() => {
    if (query === lastPushed.current) return;
    lastPushed.current = query;
    setText(query);
  }, [query]);

  const [buckets, setBuckets] = useState<Record<Kind, Bucket>>({
    app: EMPTY,
    sp: EMPTY,
    user: EMPTY,
    group: EMPTY,
  });
  const tooShort = query.length < MIN_CHARS;

  useEffect(() => {
    if (tooShort) {
      setBuckets({ app: EMPTY, sp: EMPTY, user: EMPTY, group: EMPTY });
      return;
    }
    let cancelled = false;
    setBuckets({ app: LOADING, sp: LOADING, user: LOADING, group: LOADING });
    const run = <T,>(
      kind: Kind,
      load: Promise<QuickSearchResult<T>>,
      toRow: (item: T) => Omit<Row, 'kind'>,
    ) =>
      load
        .then((r) => ({
          ...EMPTY,
          rows: r.items.map((i) => ({ ...toRow(i), kind })),
          more: r.more,
        }))
        .catch(() => ({ ...EMPTY, error: true }))
        .then((bucket) => !cancelled && setBuckets((prev) => ({ ...prev, [kind]: bucket })));

    run('app', quickSearchApplications(token, query, FULL_SEARCH), (a) => ({
      key: `app-${a.id}`,
      name: a.displayName ?? a.id,
      sub: a.uniqueName ?? undefined,
      id: a.appId,
      to: `/applications/${a.id}`,
      icon: <AppIcon id={a.appId || a.id} logoUrl={a.info?.logoUrl} size={24} />,
      flags: isFavorite(a.tags) ? ['Favorite'] : [],
    }));
    run('sp', quickSearchServicePrincipals(token, query, FULL_SEARCH), (sp) => {
      const flags: string[] = [];
      if (sp.accountEnabled === false) flags.push('Disabled');
      if (sp.servicePrincipalType === 'ManagedIdentity') flags.push('Managed identity');
      if (sp.servicePrincipalType === 'Legacy') flags.push('Legacy');
      if (
        sp.appOwnerOrganizationId &&
        tenantId &&
        sp.appOwnerOrganizationId.toLowerCase() !== tenantId.toLowerCase()
      )
        flags.push('From another tenant');
      return {
        key: `sp-${sp.id}`,
        name: sp.displayName ?? sp.id,
        sub: sp.publisherName,
        id: sp.appId,
        to: `/enterprise-apps/${sp.id}`,
        icon: <AppIcon id={sp.appId || sp.id} logoUrl={sp.info?.logoUrl} size={24} />,
        flags,
      };
    });
    run('user', quickSearchUsers(token, query, FULL_SEARCH), (u) => {
      const flags: string[] = [];
      if (u.accountEnabled === false) flags.push('Disabled');
      if (isGuest(u)) flags.push('Guest');
      return {
        key: `user-${u.id}`,
        name: u.displayName ?? u.id,
        sub: u.userPrincipalName ?? u.mail,
        id: u.id,
        to: `/users/${u.id}`,
        icon: <Identicon id={u.id} size={24} />,
        flags,
      };
    });
    run('group', quickSearchGroups(token, query, FULL_SEARCH), (g) => ({
      key: `group-${g.id}`,
      name: g.displayName ?? g.id,
      sub: [groupKind(g), g.mail].filter(Boolean).join(' · '),
      id: g.id,
      to: `/groups/${g.id}`,
      icon: <Identicon id={g.id} size={24} />,
      flags: [],
    }));
    return () => {
      cancelled = true;
    };
  }, [token, query, tooShort, tenantId]);

  const all = useMemo(() => KINDS.flatMap((k) => buckets[k.kind].rows), [buckets]);
  const ofType = useMemo(() => (type ? all.filter((r) => r.kind === type) : all), [all, type]);
  // Properties present in the current type selection, most common first.
  const flagFacets = useMemo(() => {
    const counts = new Map<string, number>();
    for (const r of ofType) for (const f of r.flags) counts.set(f, (counts.get(f) ?? 0) + 1);
    return [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  }, [ofType]);
  const activeFlag = flag && flagFacets.some(([f]) => f === flag) ? flag : null;
  const visible = useMemo(
    () => (activeFlag ? ofType.filter((r) => r.flags.includes(activeFlag)) : ofType),
    [ofType, activeFlag],
  );

  // Without a sort: grouped by type, best match first within each.
  const { sorted, sort } = useSort(visible, {
    name: (r) => r.name,
    type: (r) => KINDS.findIndex((k) => k.kind === r.kind),
    id: (r) => r.id,
  });

  const loading = KINDS.some((k) => buckets[k.kind].loading);
  const failed = KINDS.filter((k) => buckets[k.kind].error).map((k) => k.plural.toLowerCase());
  const capped = KINDS.filter((k) => buckets[k.kind].more).map((k) => k.plural.toLowerCase());
  const listKey = location.pathname + location.search;
  const open = (r: Row) => {
    if (r.external) window.open(r.to, '_blank', 'noopener,noreferrer');
    else nav(r.to, { state: { fromList: listKey } });
  };

  return (
    <>
      <div className="page-header">
        <div>
          <h1>Search</h1>
          <div className="subtitle">
            App registrations, enterprise apps, users and groups in this tenant
          </div>
        </div>
      </div>

      <div className="toolbar">
        <QueryInput
          className="search search-wide"
          entities={ALL_ENTITIES}
          autoFocus
          ariaLabel="Search the directory"
          placeholder="Search apps, users and groups by name or ID…"
          value={text}
          onChange={setText}
        />
        {!tooShort && (
          <span className="muted">
            {loading ? (
              <span className="spinner" />
            ) : visible.length === all.length ? (
              `${all.length} result${all.length === 1 ? '' : 's'}`
            ) : (
              `${visible.length} of ${all.length}`
            )}
          </span>
        )}
      </div>

      {tooShort ? (
        <div className="card">
          <div className="empty">Type at least {MIN_CHARS} characters to search.</div>
        </div>
      ) : (
        <>
          <div className="row wrap filter-chips">
            <button
              className={type === null ? 'primary' : 'ghost'}
              onClick={() => setParam('type', null)}
            >
              All ({all.length})
            </button>
            {KINDS.map((k) => {
              const b = buckets[k.kind];
              return (
                <button
                  key={k.kind}
                  className={type === k.kind ? 'primary' : 'ghost'}
                  disabled={!b.loading && b.rows.length === 0 && type !== k.kind}
                  onClick={() => setParam('type', type === k.kind ? null : k.kind)}
                >
                  {k.icon}
                  {k.plural} ({b.loading ? '…' : b.rows.length}
                  {b.more && '+'})
                </button>
              );
            })}
            {flagFacets.length > 0 && <span className="filter-sep" aria-hidden="true" />}
            {flagFacets.map(([f, n]) => (
              <button
                key={f}
                className={activeFlag === f ? 'primary' : 'ghost'}
                aria-pressed={activeFlag === f}
                onClick={() => setParam('show', activeFlag === f ? null : f)}
              >
                {f === 'Favorite' && '★ '}
                {f} ({n})
              </button>
            ))}
          </div>

          {failed.length > 0 && (
            <div className="card error">Couldn't search {failed.join(', ')}.</div>
          )}

          <div className="card" style={{ padding: 0 }}>
            <table className="table">
              <thead>
                <tr>
                  <SortTh id="name" sort={sort}>
                    Name
                  </SortTh>
                  <SortTh id="type" sort={sort}>
                    Type
                  </SortTh>
                  <SortTh id="id" sort={sort}>
                    ID
                  </SortTh>
                  <th />
                </tr>
              </thead>
              <tbody>
                {sorted.map((r) => {
                  const name = <Highlight text={r.name} query={query} />;
                  return (
                    <tr key={r.key} onClick={() => open(r)}>
                      <td>
                        <div className="row" style={{ gap: 10 }}>
                          {r.icon}
                          <div style={{ minWidth: 0 }}>
                            <div>
                              {/* A real link too, so it can be opened in a new tab. */}
                              {r.external ? (
                                <a
                                  className="result-link"
                                  href={r.to}
                                  target="_blank"
                                  rel="noreferrer"
                                  onClick={(e) => e.stopPropagation()}
                                >
                                  {name}
                                </a>
                              ) : (
                                <Link
                                  className="result-link"
                                  to={r.to}
                                  state={{ fromList: listKey }}
                                  onClick={(e) => e.stopPropagation()}
                                >
                                  {name}
                                </Link>
                              )}
                              {r.flags
                                .filter((f) => f !== 'Favorite')
                                .map((f) => (
                                  <span
                                    key={f}
                                    className={`badge${DANGER_FLAGS.has(f) ? ' expired' : ''}`}
                                    style={{ marginLeft: 6 }}
                                  >
                                    {f}
                                  </span>
                                ))}
                              {r.flags.includes('Favorite') && (
                                <span className="dsearch-star" title="Favorite">
                                  {' '}
                                  ★
                                </span>
                              )}
                            </div>
                            {r.sub && (
                              <div className="muted" style={{ fontSize: 12 }}>
                                <Highlight text={r.sub} query={query} />
                              </div>
                            )}
                          </div>
                        </div>
                      </td>
                      <td className="muted" style={{ whiteSpace: 'nowrap' }}>
                        {KINDS.find((k) => k.kind === r.kind)?.label}
                      </td>
                      <td className="mono">{r.id}</td>
                      <td className="muted">
                        {r.external && (
                          <span title="Opens in the Entra admin center">
                            <ExternalLinkIcon />
                          </span>
                        )}
                      </td>
                    </tr>
                  );
                })}
                {visible.length === 0 && (
                  <tr>
                    <td colSpan={4}>
                      <div className="empty">
                        {loading ? (
                          <span className="spinner" />
                        ) : all.length === 0 ? (
                          `Nothing matches “${query}”.`
                        ) : (
                          'Nothing matches these filters.'
                        )}
                      </div>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {capped.length > 0 && (
            <div className="muted" style={{ fontSize: 12 }}>
              Showing the first matches only for {capped.join(', ')} — refine the search to see the
              rest.
            </div>
          )}
        </>
      )}
    </>
  );
}
