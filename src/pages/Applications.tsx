import { useEffect, useMemo, useRef, useState } from 'react';
import { useLocation, useNavigate, useSearchParams } from 'react-router-dom';
import { restoreScroll, saveScroll } from '../lib/scrollMemory';
import { useGraphToken } from '../auth/useGraphToken';
import {
  createApplication,
  listApplications,
  listDeactivatedApplications,
  listOwnedApplications,
  listUsedTags,
  updateApplication,
  type UsedTag,
} from '../graph/applications';
import { createServicePrincipalFromAppId } from '../graph/servicePrincipals';
import type { Application } from '../graph/types';
import { Modal } from '../components/Modal';
import { AppIcon } from '../components/AppIcon';
import { DeletedItemsList } from '../components/DeletedItemsList';
import { FavoriteStar, TagChip } from '../components/TagChip';
import {
  FAVORITE_TAG,
  TAG_PREFIX,
  addTag,
  hasTag,
  isFavorite,
  isOurTag,
  ourTags,
  removeTag,
  sortTags,
  tagLabel,
  toStoredTag,
  withFavorite,
} from '../lib/tags';

type View = 'active' | 'owned' | 'deactivated' | 'deleted';
const VIEWS: View[] = ['active', 'owned', 'deactivated', 'deleted'];

const AUDIENCES = [
  'AzureADMyOrg',
  'AzureADMultipleOrgs',
  'AzureADandPersonalMicrosoftAccount',
  'PersonalMicrosoftAccount',
];

export function Applications() {
  const token = useGraphToken();
  const nav = useNavigate();
  const location = useLocation();
  // Tab, search and tag filter live in the URL, so coming back from an app
  // registration (back button or browser back) returns to the same list.
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
  const [apps, setApps] = useState<Application[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState(() => params.get('q') ?? '');
  const debouncedSearch = (params.get('q') ?? '').trim();
  const [showCreate, setShowCreate] = useState(false);
  // 'owned' is the subset the signed-in user is an owner of; 'deactivated'
  // lists app registrations with isDisabled = true; 'deleted' is the directory
  // recycle bin, rendered by its own component.
  const viewParam = params.get('view');
  const view: View = (VIEWS as string[]).includes(viewParam ?? '') ? (viewParam as View) : 'active';
  // Switching tabs also clears the tag filter.
  const setView = (v: View) =>
    setParams(
      (prev) => {
        const next = new URLSearchParams(prev);
        if (v === 'active') next.delete('view');
        else next.set('view', v);
        next.delete('tag');
        return next;
      },
      { replace: true },
    );
  const listed = view !== 'deleted';
  const tagFilter = params.get('tag');
  const setTagFilter = (t: string | null) => setParam('tag', t);
  const [tagBusyId, setTagBusyId] = useState<string | null>(null);
  const [tagError, setTagError] = useState<string | null>(null);
  // Row ids ticked for bulk tagging. Cleared whenever the list reloads.
  const [selected, setSelected] = useState<Set<string>>(() => new Set());
  const [bulkBusy, setBulkBusy] = useState(false);
  const [usedTags, setUsedTags] = useState<UsedTag[] | null>(null);

  useEffect(() => {
    const q = search.trim();
    if (q === debouncedSearch) return;
    const handle = setTimeout(() => {
      lastPushedSearch.current = q;
      setParam('q', q || null);
    }, 250);
    return () => clearTimeout(handle);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [search, debouncedSearch]);

  // The URL changed from outside (sidebar link, browser back): show its search.
  const lastPushedSearch = useRef(debouncedSearch);
  useEffect(() => {
    if (debouncedSearch === lastPushedSearch.current) return;
    lastPushedSearch.current = debouncedSearch;
    setSearch(debouncedSearch);
  }, [debouncedSearch]);

  // Back on this list from an app registration: once rows are in, scroll to
  // where the user was.
  const listKey = location.pathname + location.search;
  useEffect(() => {
    if (apps) restoreScroll(listKey);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [apps]);
  const openApp = (id: string) => {
    saveScroll(listKey);
    nav(`/applications/${id}`, { state: { fromList: listKey } });
  };

  useEffect(() => {
    if (view === 'deleted') return;
    let cancelled = false;
    setApps(null);
    setError(null);
    setSelected(new Set());
    const q = debouncedSearch || undefined;
    (view === 'owned'
      ? listOwnedApplications(token, q)
      : view === 'deactivated'
        ? listDeactivatedApplications(token, q)
        : listApplications(token, q)
    )
      .then((list) => {
        if (cancelled) return;
        // Favorites first, then alphabetical. Sorting happens only on load, so
        // starring a row doesn't make it jump away under the cursor.
        setApps(
          [...list].sort((a, b) => {
            const favDelta = Number(isFavorite(b.tags)) - Number(isFavorite(a.tags));
            if (favDelta !== 0) return favDelta;
            return (a.displayName ?? '').localeCompare(b.displayName ?? '');
          }),
        );
      })
      .catch((e) => {
        if (!cancelled) setError(e.message);
      });
    return () => {
      cancelled = true;
    };
  }, [token, debouncedSearch, view]);

  // Facets over our own (debble:-prefixed) tags in the loaded set, most-used
  // first so the useful ones surface. Tags set outside EntraApp aren't offered
  // as filters.
  const tagFacets = useMemo(() => {
    const counts = new Map<string, { tag: string; count: number }>();
    for (const a of apps ?? []) {
      for (const t of ourTags(a.tags)) {
        const key = t.toLowerCase();
        const existing = counts.get(key);
        if (existing) existing.count += 1;
        else counts.set(key, { tag: t, count: 1 });
      }
    }
    return [...counts.values()].sort((a, b) => {
      if (a.tag.toLowerCase() === FAVORITE_TAG.toLowerCase()) return -1;
      if (b.tag.toLowerCase() === FAVORITE_TAG.toLowerCase()) return 1;
      return b.count - a.count || tagLabel(a.tag).localeCompare(tagLabel(b.tag));
    });
  }, [apps]);

  const visible = useMemo(() => {
    if (!apps) return [];
    if (!tagFilter) return apps;
    return apps.filter((a) => hasTag(a.tags, tagFilter));
  }, [apps, tagFilter]);

  // Only rows on screen count as selected, so a tag filter never lets a bulk
  // action reach apps the user can't see.
  const selectedApps = useMemo(
    () => visible.filter((a) => selected.has(a.id)),
    [visible, selected],
  );
  const allVisibleSelected = visible.length > 0 && selectedApps.length === visible.length;

  const toggleSelected = (id: string) =>
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

  const toggleAllVisible = () =>
    setSelected(allVisibleSelected ? new Set() : new Set(visible.map((a) => a.id)));

  // Tenant-wide tag suggestions for the bulk bar, fetched the first time
  // something is selected.
  useEffect(() => {
    if (selectedApps.length === 0 || usedTags) return;
    let cancelled = false;
    listUsedTags(token)
      .then((t) => {
        if (!cancelled) setUsedTags(t);
      })
      .catch(() => {
        if (!cancelled) setUsedTags([]);
      });
    return () => {
      cancelled = true;
    };
  }, [token, selectedApps.length, usedTags]);

  // Applies a tag change to every selected app. Graph PATCHes tags wholesale
  // per app, so each one is its own request; they run a few at a time and
  // each row updates as soon as its write succeeds. Failures are collected
  // and reported together rather than stopping the batch.
  const bulkApply = async (transform: (tags: string[] | undefined) => string[]) => {
    const targets = selectedApps
      .map((app) => ({ app, next: transform(app.tags) }))
      .filter(({ app, next }) => !sameTags(app.tags, next));
    if (targets.length === 0) return;
    setBulkBusy(true);
    setTagError(null);
    const failures: string[] = [];
    await runPool(targets, 4, async ({ app, next }) => {
      try {
        await updateApplication(token, app.id, { tags: next });
        setApps((prev) =>
          prev ? prev.map((a) => (a.id === app.id ? { ...a, tags: next } : a)) : prev,
        );
      } catch (e: unknown) {
        failures.push(`${app.displayName}: ${e instanceof Error ? e.message : String(e)}`);
      }
    });
    if (failures.length) {
      setTagError(
        `Tag update failed for ${failures.length} of ${targets.length} apps — ${failures.join('; ')}`,
      );
    }
    setUsedTags(null);
    setBulkBusy(false);
  };

  // Tags live on the app registration in Entra, so the star is a Graph write.
  // Applied optimistically and rolled back if the PATCH fails.
  const toggleFavorite = async (app: Application) => {
    const next = withFavorite(app.tags, !isFavorite(app.tags));
    const previous = app.tags;
    const apply = (tags: string[] | undefined) =>
      setApps((prev) => (prev ? prev.map((a) => (a.id === app.id ? { ...a, tags } : a)) : prev));

    apply(next);
    setTagBusyId(app.id);
    setTagError(null);
    try {
      await updateApplication(token, app.id, { tags: next });
    } catch (e: unknown) {
      apply(previous);
      setTagError(e instanceof Error ? e.message : String(e));
    } finally {
      setTagBusyId(null);
    }
  };

  return (
    <>
      <div className="page-header">
        <div>
          <h1>App registrations</h1>
          <div className="subtitle">Identity objects representing applications in your tenant</div>
        </div>
        <button className="primary" onClick={() => setShowCreate(true)}>
          + New app registration
        </button>
      </div>

      <div className="toolbar">
        <input
          className="search"
          placeholder="Search by name, unique name, appId, or tag…"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
        />
        <div className="tabs" style={{ margin: 0, borderBottom: 'none' }}>
          <button
            className={view === 'active' ? 'active' : ''}
            onClick={() => {
              setView('active');
            }}
          >
            Active
          </button>
          <button
            className={view === 'owned' ? 'active' : ''}
            onClick={() => {
              setView('owned');
            }}
            title="App registrations you are listed as an owner of"
          >
            Owned
          </button>
          <button
            className={view === 'deactivated' ? 'active' : ''}
            onClick={() => {
              setView('deactivated');
            }}
            title="App registrations that are deactivated: no new tokens, configuration kept"
          >
            Deactivated
          </button>
          <button
            className={view === 'deleted' ? 'active' : ''}
            onClick={() => setView('deleted')}
            title="Soft-deleted app registrations in the directory recycle bin"
          >
            Deleted
          </button>
        </div>
        {apps && listed && (
          <span className="muted">
            {tagFilter
              ? `${visible.length} of ${apps.length}`
              : `${apps.length} ${debouncedSearch ? 'matches' : 'total'}`}
          </span>
        )}
      </div>

      {view === 'deleted' && <DeletedItemsList kind="application" search={debouncedSearch} />}

      {listed && tagFacets.length > 0 && (
        <div className="row wrap" style={{ gap: 6, marginBottom: 12 }}>
          <button
            className={tagFilter === null ? 'primary' : 'ghost'}
            onClick={() => setTagFilter(null)}
          >
            All
          </button>
          {tagFacets.map((f) => (
            <button
              key={f.tag}
              title={f.tag}
              className={tagFilter?.toLowerCase() === f.tag.toLowerCase() ? 'primary' : 'ghost'}
              onClick={() =>
                setTagFilter(tagFilter?.toLowerCase() === f.tag.toLowerCase() ? null : f.tag)
              }
            >
              {f.tag.toLowerCase() === FAVORITE_TAG.toLowerCase() && '★ '}
              {tagLabel(f.tag)} ({f.count})
            </button>
          ))}
        </div>
      )}

      {listed && error && <div className="card error">{error}</div>}
      {listed && tagError && <div className="card error">{tagError}</div>}
      {listed && !apps && !error && <span className="spinner" />}

      {listed && apps && selectedApps.length > 0 && (
        <BulkTagBar
          selected={selectedApps}
          suggestions={usedTags ?? tagFacets}
          busy={bulkBusy}
          onAdd={(tag) => bulkApply((tags) => addTag(tags, tag))}
          onRemove={(tag) => bulkApply((tags) => removeTag(tags, tag))}
          onClear={() => setSelected(new Set())}
        />
      )}

      {listed && apps && (
        <div className="card" style={{ padding: 0 }}>
          <table className="table">
            <thead>
              <tr>
                <th className="select-cell">
                  <Checkbox
                    checked={allVisibleSelected}
                    indeterminate={selectedApps.length > 0 && !allVisibleSelected}
                    disabled={visible.length === 0 || bulkBusy}
                    label="Select all visible app registrations"
                    onChange={toggleAllVisible}
                  />
                </th>
                <th>Display name</th>
                <th>Application (client) ID</th>
                <th>Tags</th>
                <th>Credentials</th>
                <th>Sign-in audience</th>
              </tr>
            </thead>
            <tbody>
              {visible.map((a) => (
                <tr
                  key={a.id}
                  className={selected.has(a.id) ? 'selected' : undefined}
                  onClick={() => openApp(a.id)}
                >
                  <td className="select-cell" onClick={(e) => e.stopPropagation()}>
                    <Checkbox
                      checked={selected.has(a.id)}
                      disabled={bulkBusy}
                      label={`Select ${a.displayName}`}
                      onChange={() => toggleSelected(a.id)}
                    />
                  </td>
                  <td>
                    <div className="row" style={{ gap: 10, alignItems: 'center' }}>
                      <FavoriteStar
                        favorite={isFavorite(a.tags)}
                        disabled={tagBusyId === a.id}
                        onToggle={() => toggleFavorite(a)}
                      />
                      <AppIcon
                        id={a.appId || a.id}
                        logoUrl={a.info?.logoUrl}
                        title={a.displayName}
                      />
                      {hasExpiredCredential(a) && (
                        <span
                          title="One or more credentials have expired"
                          style={{ color: 'var(--danger, #d13438)' }}
                        >
                          ⚠
                        </span>
                      )}
                      <div style={{ minWidth: 0 }}>
                        <div>{a.displayName}</div>
                        {a.uniqueName && (
                          <div
                            className="mono muted"
                            style={{ fontSize: 11 }}
                            title={`Unique name: ${a.uniqueName}`}
                          >
                            {a.uniqueName}
                          </div>
                        )}
                      </div>
                    </div>
                  </td>
                  <td className="mono">{a.appId}</td>
                  <td>
                    {a.tags?.length ? (
                      <div className="row wrap" style={{ gap: 4 }}>
                        {sortTags(a.tags).map((t) => (
                          <TagChip
                            key={t}
                            tag={t}
                            title={`Filter by tag ${tagLabel(t)}`}
                            onClick={() => setTagFilter(t)}
                          />
                        ))}
                      </div>
                    ) : (
                      <span className="muted">—</span>
                    )}
                  </td>
                  <td>
                    <CredentialBadges app={a} />
                  </td>
                  <td className="muted">{a.signInAudience}</td>
                </tr>
              ))}
              {visible.length === 0 && (
                <tr>
                  <td colSpan={6}>
                    <div className="empty">
                      {view === 'deactivated' && !tagFilter && !debouncedSearch
                        ? 'No deactivated app registrations.'
                        : tagFilter
                          ? `No applications tagged “${tagLabel(tagFilter)}”.`
                          : debouncedSearch
                            ? 'No applications match your search.'
                            : view === 'owned'
                              ? 'You are not an owner of any app registrations.'
                              : 'No applications found.'}
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {showCreate && (
        <CreateAppModal
          onClose={() => setShowCreate(false)}
          onCreated={(appId) => nav(`/applications/${appId}`)}
        />
      )}
    </>
  );
}

function CreateAppModal({
  onClose,
  onCreated,
}: {
  onClose: () => void;
  onCreated: (objectId: string) => void;
}) {
  const token = useGraphToken();
  const [displayName, setDisplayName] = useState('');
  const [audience, setAudience] = useState(AUDIENCES[0]);
  const [createSp, setCreateSp] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const submit = async () => {
    if (!displayName.trim()) return;
    setSubmitting(true);
    setError(null);
    try {
      const app = await createApplication(token, {
        displayName: displayName.trim(),
        signInAudience: audience,
      });
      if (createSp) {
        try {
          await createServicePrincipalFromAppId(token, app.appId);
        } catch (e: unknown) {
          // Non-fatal: app registration was created; surface but continue.
          setError(
            `App registered, but creating the enterprise app failed: ${
              e instanceof Error ? e.message : String(e)
            }`,
          );
        }
      }
      onCreated(app.id);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <Modal
      title="New app registration"
      onClose={() => !submitting && onClose()}
      footer={
        <>
          <button disabled={submitting} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={submitting || !displayName.trim()} onClick={submit}>
            {submitting ? 'Creating…' : 'Create'}
          </button>
        </>
      }
    >
      <label className="field">
        <span>Display name</span>
        <input
          autoFocus
          value={displayName}
          onChange={(e) => setDisplayName(e.target.value)}
          placeholder="My new app"
        />
      </label>

      <label className="field">
        <span>Sign-in audience</span>
        <select value={audience} onChange={(e) => setAudience(e.target.value)}>
          {AUDIENCES.map((a) => (
            <option key={a} value={a}>
              {a}
            </option>
          ))}
        </select>
      </label>

      <label className="row" style={{ gap: 6 }}>
        <input
          type="checkbox"
          style={{ width: 'auto' }}
          checked={createSp}
          onChange={(e) => setCreateSp(e.target.checked)}
        />
        Also create the enterprise app (service principal) in this tenant
      </label>

      {error && (
        <p className="error" style={{ marginTop: 12 }}>
          {error}
        </p>
      )}
    </Modal>
  );
}

function sameTags(a: string[] | undefined, b: string[]): boolean {
  const x = a ?? [];
  return x.length === b.length && x.every((t, i) => t === b[i]);
}

async function runPool<T>(items: T[], size: number, worker: (item: T) => Promise<void>) {
  let i = 0;
  const lanes = Array.from({ length: Math.min(size, items.length) }, async () => {
    while (i < items.length) await worker(items[i++]);
  });
  await Promise.all(lanes);
}

function Checkbox({
  checked,
  indeterminate = false,
  disabled,
  label,
  onChange,
}: {
  checked: boolean;
  indeterminate?: boolean;
  disabled?: boolean;
  label: string;
  onChange: () => void;
}) {
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => {
    if (ref.current) ref.current.indeterminate = indeterminate;
  }, [indeterminate]);
  return (
    <input
      ref={ref}
      type="checkbox"
      aria-label={label}
      title={label}
      checked={checked}
      disabled={disabled}
      onChange={onChange}
    />
  );
}

// Shown above the table while rows are ticked: add a tag to all of them, or
// remove one of our tags that some of them carry.
function BulkTagBar({
  selected,
  suggestions,
  busy,
  onAdd,
  onRemove,
  onClear,
}: {
  selected: Application[];
  suggestions: Array<{ tag: string; count: number }>;
  busy: boolean;
  onAdd: (tag: string) => Promise<void>;
  onRemove: (tag: string) => Promise<void>;
  onClear: () => void;
}) {
  const [value, setValue] = useState('');

  // Our tags present on the selection, with how many selected apps carry each.
  const present = useMemo(() => {
    const counts = new Map<string, { tag: string; count: number }>();
    for (const a of selected) {
      for (const t of ourTags(a.tags)) {
        const key = t.toLowerCase();
        const hit = counts.get(key);
        if (hit) hit.count += 1;
        else counts.set(key, { tag: t, count: 1 });
      }
    }
    return [...counts.values()].sort((a, b) => tagLabel(a.tag).localeCompare(tagLabel(b.tag)));
  }, [selected]);

  const add = async () => {
    const stored = toStoredTag(value);
    if (!stored) return;
    // Reuse the casing already in use in the tenant, so "prod" joins "Prod".
    const existing = suggestions.find((s) => s.tag.toLowerCase() === stored.toLowerCase());
    await onAdd(existing?.tag ?? stored);
    setValue('');
  };

  const n = selected.length;
  return (
    <div className="card bulk-bar">
      <strong>{n} selected</strong>
      <span className="row" style={{ gap: 4 }}>
        <span className="mono muted" style={{ fontSize: 12 }}>
          {TAG_PREFIX}
        </span>
        <input
          list="bulk-tag-suggestions"
          value={value}
          disabled={busy}
          placeholder="Tag to add"
          style={{ width: 170 }}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') void add();
          }}
        />
        <datalist id="bulk-tag-suggestions">
          {suggestions
            .filter((s) => isOurTag(s.tag))
            .map((s) => (
              <option key={s.tag} value={tagLabel(s.tag)} />
            ))}
        </datalist>
        <button className="primary" disabled={busy || !toStoredTag(value)} onClick={add}>
          {busy ? 'Saving…' : `Add to ${n}`}
        </button>
      </span>
      {present.length > 0 && (
        <span className="row wrap" style={{ gap: 4 }}>
          <span className="muted" style={{ fontSize: 12 }}>
            Remove:
          </span>
          {present.map((p) => (
            <TagChip
              key={p.tag}
              tag={p.tag}
              disabled={busy}
              title={`Remove ${tagLabel(p.tag)} from the ${p.count} selected ${
                p.count === 1 ? 'app' : 'apps'
              } that have it`}
              onRemove={() => void onRemove(p.tag)}
            />
          ))}
        </span>
      )}
      <button className="ghost" disabled={busy} onClick={onClear} style={{ marginLeft: 'auto' }}>
        Clear selection
      </button>
    </div>
  );
}

type ExpiryState = 'none' | 'active' | 'soon' | 'expired';

function credentialGroupState(creds: Array<{ endDateTime?: string }> | undefined): ExpiryState {
  if (!creds || creds.length === 0) return 'none';
  const now = Date.now();
  const soonMs = 30 * 24 * 60 * 60 * 1000;
  let worst: 'active' | 'soon' = 'active';
  for (const c of creds) {
    const end = c.endDateTime ? new Date(c.endDateTime).getTime() : NaN;
    if (!Number.isFinite(end)) continue;
    if (end < now) return 'expired';
    if (end - now < soonMs) worst = 'soon';
  }
  return worst;
}

function hasExpiredCredential(app: Application): boolean {
  return (
    credentialGroupState(app.passwordCredentials) === 'expired' ||
    credentialGroupState(app.keyCredentials) === 'expired'
  );
}

function CredentialBadges({ app }: { app: Application }) {
  const secretState = credentialGroupState(app.passwordCredentials);
  const certState = credentialGroupState(app.keyCredentials);
  const ficCount = app.federatedIdentityCredentials?.length ?? 0;

  if (secretState === 'none' && certState === 'none' && ficCount === 0) {
    return <span className="muted">—</span>;
  }
  return (
    <div className="row" style={{ gap: 6, flexWrap: 'wrap' }}>
      {secretState !== 'none' && (
        <CredentialBadge
          label="Secret"
          count={app.passwordCredentials?.length ?? 0}
          state={secretState}
        />
      )}
      {certState !== 'none' && (
        <CredentialBadge label="Cert" count={app.keyCredentials?.length ?? 0} state={certState} />
      )}
      {ficCount > 0 && <CredentialBadge label="FIC" count={ficCount} state="active" />}
    </div>
  );
}

function CredentialBadge({
  label,
  count,
  state,
}: {
  label: string;
  count: number;
  state: ExpiryState;
}) {
  const cls =
    state === 'expired' ? 'badge expired' : state === 'soon' ? 'badge pending' : 'badge granted';
  const title =
    state === 'expired'
      ? `${label}: one or more expired`
      : state === 'soon'
        ? `${label}: one or more expire within 30 days`
        : `${label}: active`;
  return (
    <span className={cls} title={title}>
      {label} {count}
      {state === 'expired' && ' ⚠'}
    </span>
  );
}
