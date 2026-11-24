import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { useCurrentTenantId } from '../auth/useCurrentTenantId';
import {
  quickSearchApplications,
  quickSearchGroups,
  quickSearchServicePrincipals,
  isGuest,
  quickSearchUsers,
  type QuickGroup,
  type QuickSearchResult,
  type QuickUser,
} from '../graph/quickSearch';
import type { Application, ServicePrincipal } from '../graph/types';
import { isFavorite } from '../lib/tags';
import { AppIcon } from './AppIcon';
import { AppRegistrationIcon, EnterpriseAppIcon, GroupIcon, UserIcon } from './ButtonIcons';
import { Identicon } from './Identicon';

// Search box for the Overview: type a name or appId and jump straight to an
// app registration, enterprise app, user or group. Results drop
// down in four sections; arrow keys pick a row and Enter opens it. Enter
// without a pick opens the full, filterable results page (/search).

const MIN_CHARS = 2;
const SHOWN = 6;
// Users and groups are the side dish: fewer rows keep the apps in view.
const SHOWN_PRINCIPALS = 4;

interface Section<T> {
  loading: boolean;
  error: boolean;
  result: QuickSearchResult<T> | null;
}
const IDLE = { loading: false, error: false, result: null };

interface Option {
  key: string;
  to: string;
  // Opens another site (the Entra admin center) in a new tab.
  external?: boolean;
  node: ReactNode;
}

export function groupKind(g: QuickGroup): string {
  const dynamic = g.groupTypes?.includes('DynamicMembership') ? ' · dynamic' : '';
  if (g.groupTypes?.includes('Unified')) return `Microsoft 365${dynamic}`;
  if (g.securityEnabled) return `${g.mailEnabled ? 'Mail-enabled security' : 'Security'}${dynamic}`;
  return g.mailEnabled ? 'Distribution' : 'Group';
}

export function DirectorySearch() {
  const token = useGraphToken();
  const tenantId = useCurrentTenantId();
  const nav = useNavigate();
  const [text, setText] = useState('');
  const [query, setQuery] = useState('');
  const [open, setOpen] = useState(false);
  // -1: nothing picked yet — Enter then opens the full results page.
  const [active, setActive] = useState(-1);
  const [apps, setApps] = useState<Section<Application>>(IDLE);
  const [sps, setSps] = useState<Section<ServicePrincipal>>(IDLE);
  const [users, setUsers] = useState<Section<QuickUser>>(IDLE);
  const [groups, setGroups] = useState<Section<QuickGroup>>(IDLE);
  const rootRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const trimmed = text.trim();
  const tooShort = trimmed.length < MIN_CHARS;

  useEffect(() => {
    if (tooShort) {
      setQuery('');
      return;
    }
    const handle = setTimeout(() => setQuery(trimmed), 250);
    return () => clearTimeout(handle);
  }, [trimmed, tooShort]);

  useEffect(() => {
    if (!query) {
      setApps(IDLE);
      setSps(IDLE);
      setUsers(IDLE);
      setGroups(IDLE);
      return;
    }
    let cancelled = false;
    // Earlier results stay on screen until the new ones arrive.
    setApps((s) => ({ ...s, loading: true, error: false }));
    setSps((s) => ({ ...s, loading: true, error: false }));
    setUsers((s) => ({ ...s, loading: true, error: false }));
    setGroups((s) => ({ ...s, loading: true, error: false }));
    quickSearchApplications(token, query)
      .then((result) => !cancelled && setApps({ loading: false, error: false, result }))
      .catch(() => !cancelled && setApps({ loading: false, error: true, result: null }));
    quickSearchServicePrincipals(token, query)
      .then((result) => !cancelled && setSps({ loading: false, error: false, result }))
      .catch(() => !cancelled && setSps({ loading: false, error: true, result: null }));
    quickSearchUsers(token, query)
      .then((result) => !cancelled && setUsers({ loading: false, error: false, result }))
      .catch(() => !cancelled && setUsers({ loading: false, error: true, result: null }));
    quickSearchGroups(token, query)
      .then((result) => !cancelled && setGroups({ loading: false, error: false, result }))
      .catch(() => !cancelled && setGroups({ loading: false, error: true, result: null }));
    setActive(-1);
    return () => {
      cancelled = true;
    };
  }, [token, query]);

  // "/" focuses the search from anywhere on the page.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== '/' || e.metaKey || e.ctrlKey || e.altKey) return;
      const t = e.target as HTMLElement | null;
      if (t && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName))) return;
      e.preventDefault();
      inputRef.current?.focus();
      inputRef.current?.select();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, []);

  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener('mousedown', onDown);
    return () => document.removeEventListener('mousedown', onDown);
  }, [open]);

  const appOptions = useMemo<Option[]>(() => {
    const r = apps.result;
    if (!r) return [];
    const rows: Option[] = r.items.slice(0, SHOWN).map((a) => ({
      key: `app-${a.id}`,
      to: `/applications/${a.id}`,
      node: (
        <>
          <AppIcon id={a.appId || a.id} logoUrl={a.info?.logoUrl} size={24} />
          <span className="dsearch-text">
            <span className="dsearch-name">
              {isFavorite(a.tags) && <span className="dsearch-star">★ </span>}
              <Highlight text={a.displayName ?? a.id} query={query} />
            </span>
            {a.uniqueName && <span className="dsearch-sub mono">{a.uniqueName}</span>}
          </span>
          <span className="dsearch-id mono">{a.appId}</span>
        </>
      ),
    }));
    if (r.more || r.items.length > SHOWN) {
      rows.push({
        key: 'app-all',
        to: `/applications?q=${encodeURIComponent(query)}`,
        node: <span className="dsearch-all">All matching app registrations →</span>,
      });
    }
    return rows;
  }, [apps.result, query]);

  const spOptions = useMemo<Option[]>(() => {
    const r = sps.result;
    if (!r) return [];
    const rows: Option[] = r.items.slice(0, SHOWN).map((sp) => {
      const kind =
        sp.servicePrincipalType === 'ManagedIdentity'
          ? 'Managed identity'
          : sp.servicePrincipalType === 'Legacy'
            ? 'Legacy'
            : sp.publisherName;
      return {
        key: `sp-${sp.id}`,
        to: `/enterprise-apps/${sp.id}`,
        node: (
          <>
            <AppIcon id={sp.appId || sp.id} logoUrl={sp.info?.logoUrl} size={24} />
            <span className="dsearch-text">
              <span className="dsearch-name">
                <Highlight text={sp.displayName ?? sp.id} query={query} />
                {sp.accountEnabled === false && (
                  <>
                    {' '}
                    <span className="badge expired">Disabled</span>
                  </>
                )}
              </span>
              {kind && <span className="dsearch-sub">{kind}</span>}
            </span>
            <span className="dsearch-id mono">{sp.appId}</span>
          </>
        ),
      };
    });
    if (r.more || r.items.length > SHOWN) {
      rows.push({
        key: 'sp-all',
        to: `/enterprise-apps?q=${encodeURIComponent(query)}`,
        node: <span className="dsearch-all">All matching enterprise apps →</span>,
      });
    }
    return rows;
  }, [sps.result, query]);

  const userOptions = useMemo<Option[]>(
    () =>
      (users.result?.items ?? []).slice(0, SHOWN_PRINCIPALS).map((u) => ({
        key: `user-${u.id}`,
        to: `/users/${u.id}`,
        node: (
          <>
            <Identicon id={u.id} size={24} title={u.displayName ?? u.id} />
            <span className="dsearch-text">
              <span className="dsearch-name">
                <Highlight text={u.displayName ?? u.id} query={query} />
                {isGuest(u) && (
                  <>
                    {' '}
                    <span className="badge">Guest</span>
                  </>
                )}
                {u.accountEnabled === false && (
                  <>
                    {' '}
                    <span className="badge expired">Disabled</span>
                  </>
                )}
              </span>
              {(u.userPrincipalName || u.mail) && (
                <span className="dsearch-sub">
                  <Highlight text={u.userPrincipalName ?? u.mail ?? ''} query={query} />
                </span>
              )}
            </span>
          </>
        ),
      })),
    [users.result, query, tenantId],
  );

  const groupOptions = useMemo<Option[]>(
    () =>
      (groups.result?.items ?? []).slice(0, SHOWN_PRINCIPALS).map((g) => ({
        key: `group-${g.id}`,
        to: `/groups/${g.id}`,
        node: (
          <>
            <Identicon id={g.id} size={24} title={g.displayName ?? g.id} />
            <span className="dsearch-text">
              <span className="dsearch-name">
                <Highlight text={g.displayName ?? g.id} query={query} />
              </span>
              <span className="dsearch-sub">
                {groupKind(g)}
                {g.mail && ` · ${g.mail}`}
              </span>
            </span>
          </>
        ),
      })),
    [groups.result, query, tenantId],
  );

  const options = useMemo(
    () => [...appOptions, ...spOptions, ...userOptions, ...groupOptions],
    [appOptions, spOptions, userOptions, groupOptions],
  );
  const activeIndex = Math.min(active, options.length - 1);
  const showPanel = open && trimmed.length > 0;

  useEffect(() => {
    if (!showPanel || activeIndex < 0) return;
    document.getElementById(`dsearch-opt-${activeIndex}`)?.scrollIntoView({ block: 'nearest' });
  }, [activeIndex, showPanel]);

  const resultsUrl = `/search?q=${encodeURIComponent(trimmed)}`;

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      setOpen(true);
      if (options.length === 0) return;
      // From "nothing picked", down lands on the first row and up on the last.
      setActive(
        e.key === 'ArrowDown'
          ? (activeIndex + 1) % options.length
          : activeIndex <= 0
            ? options.length - 1
            : activeIndex - 1,
      );
    } else if (e.key === 'Enter') {
      const hit = showPanel && activeIndex >= 0 ? options[activeIndex] : undefined;
      if (hit) {
        e.preventDefault();
        if (hit.external) window.open(hit.to, '_blank', 'noopener,noreferrer');
        else nav(hit.to);
      } else if (!tooShort) {
        e.preventDefault();
        nav(resultsUrl);
      }
    } else if (e.key === 'Escape') {
      if (open && trimmed) setOpen(false);
      else setText('');
    }
  };

  const renderSection = (
    title: string,
    icon: ReactNode,
    section: Section<unknown>,
    rows: Option[],
    offset: number,
    noun: string,
    // Sections without a list page to hand over to say so instead.
    shown?: number,
  ) => (
    <div className="dsearch-section" role="group" aria-label={title}>
      <div className="dsearch-heading">
        {icon}
        {title}
        {section.loading && <span className="spinner" />}
      </div>
      {rows.map((o, i) => {
        const props = {
          id: `dsearch-opt-${offset + i}`,
          role: 'option',
          'aria-selected': offset + i === activeIndex,
          tabIndex: -1,
          className: `dsearch-option${offset + i === activeIndex ? ' active' : ''}`,
        };
        return o.external ? (
          <a key={o.key} {...props} href={o.to} target="_blank" rel="noreferrer">
            {o.node}
          </a>
        ) : (
          <Link key={o.key} {...props} to={o.to}>
            {o.node}
          </Link>
        );
      })}
      {shown !== undefined &&
        section.result &&
        (section.result.more || section.result.items.length > shown) && (
          <div className="dsearch-note dsearch-more">
            More {noun} match — keep typing to narrow it down.
          </div>
        )}
      {section.error && <div className="dsearch-note">Couldn't search {noun}.</div>}
      {!section.error && !section.loading && section.result && rows.length === 0 && (
        <div className="dsearch-note">
          No {noun} match “{query}”.
        </div>
      )}
    </div>
  );

  return (
    <div
      className="dsearch"
      ref={rootRef}
      onBlur={(e) => {
        if (!e.currentTarget.contains(e.relatedTarget)) setOpen(false);
      }}
    >
      <svg className="dsearch-icon" viewBox="0 0 24 24" aria-hidden="true">
        <circle cx="11" cy="11" r="7" />
        <path d="M20 20l-3.5-3.5" />
      </svg>
      <input
        ref={inputRef}
        type="search"
        role="combobox"
        aria-expanded={showPanel}
        aria-controls="dsearch-panel"
        aria-autocomplete="list"
        aria-activedescendant={
          showPanel && activeIndex >= 0 ? `dsearch-opt-${activeIndex}` : undefined
        }
        aria-label="Search apps, users and groups"
        placeholder="Search apps, users and groups by name or ID…"
        autoComplete="off"
        spellCheck={false}
        value={text}
        onChange={(e) => {
          setText(e.target.value);
          setOpen(true);
        }}
        onFocus={() => setOpen(true)}
        onKeyDown={onKeyDown}
      />
      {!text && <kbd className="dsearch-kbd">/</kbd>}
      {showPanel && (
        <div
          className="dsearch-panel"
          id="dsearch-panel"
          role="listbox"
          // Keeps focus in the input, so clicking a result doesn't close the
          // panel before the click lands.
          onMouseDown={(e) => e.preventDefault()}
        >
          {tooShort ? (
            <div className="dsearch-note">Keep typing — at least {MIN_CHARS} characters.</div>
          ) : (
            <>
              {renderSection(
                'App registrations',
                <AppRegistrationIcon />,
                apps,
                appOptions,
                0,
                'app registrations',
              )}
              {renderSection(
                'Enterprise apps',
                <EnterpriseAppIcon />,
                sps,
                spOptions,
                appOptions.length,
                'enterprise apps',
              )}
              {renderSection(
                'Users',
                <UserIcon />,
                users,
                userOptions,
                appOptions.length + spOptions.length,
                'users',
                SHOWN_PRINCIPALS,
              )}
              {renderSection(
                'Groups',
                <GroupIcon />,
                groups,
                groupOptions,
                appOptions.length + spOptions.length + userOptions.length,
                'groups',
                SHOWN_PRINCIPALS,
              )}
              <Link className="dsearch-option dsearch-results-link" tabIndex={-1} to={resultsUrl}>
                <span className="dsearch-all">See all results for “{trimmed}” →</span>
              </Link>
              <div className="dsearch-foot">
                <span>
                  <kbd>↑</kbd> <kbd>↓</kbd> to pick
                </span>
                <span>
                  <kbd>Enter</kbd> {activeIndex >= 0 ? 'to open' : 'for all results'}
                </span>
                <span>
                  <kbd>Esc</kbd> to close
                </span>
              </div>
            </>
          )}
        </div>
      )}
    </div>
  );
}

export function Highlight({ text, query }: { text: string; query: string }) {
  const i = query ? text.toLowerCase().indexOf(query.toLowerCase()) : -1;
  if (i < 0) return <>{text}</>;
  return (
    <>
      {text.slice(0, i)}
      <mark>{text.slice(i, i + query.length)}</mark>
      {text.slice(i + query.length)}
    </>
  );
}
