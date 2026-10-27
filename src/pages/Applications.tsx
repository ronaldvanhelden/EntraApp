import { useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useGraphToken } from '../auth/useGraphToken';
import { createApplication, listApplications, updateApplication } from '../graph/applications';
import { createServicePrincipalFromAppId } from '../graph/servicePrincipals';
import type { Application } from '../graph/types';
import { Modal } from '../components/Modal';
import { AppIcon } from '../components/AppIcon';
import { FavoriteStar, TagChip } from '../components/TagChip';
import {
  FAVORITE_TAG,
  hasTag,
  isFavorite,
  ourTags,
  sortTags,
  tagLabel,
  withFavorite,
} from '../lib/tags';

const AUDIENCES = [
  'AzureADMyOrg',
  'AzureADMultipleOrgs',
  'AzureADandPersonalMicrosoftAccount',
  'PersonalMicrosoftAccount',
];

export function Applications() {
  const token = useGraphToken();
  const nav = useNavigate();
  const [apps, setApps] = useState<Application[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState('');
  const [debouncedSearch, setDebouncedSearch] = useState('');
  const [showCreate, setShowCreate] = useState(false);
  const [tagFilter, setTagFilter] = useState<string | null>(null);
  const [tagBusyId, setTagBusyId] = useState<string | null>(null);
  const [tagError, setTagError] = useState<string | null>(null);

  useEffect(() => {
    const handle = setTimeout(() => setDebouncedSearch(search.trim()), 250);
    return () => clearTimeout(handle);
  }, [search]);

  useEffect(() => {
    let cancelled = false;
    setApps(null);
    setError(null);
    listApplications(token, debouncedSearch || undefined)
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
  }, [token, debouncedSearch]);

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
        {apps && (
          <span className="muted">
            {tagFilter
              ? `${visible.length} of ${apps.length}`
              : `${apps.length} ${debouncedSearch ? 'matches' : 'total'}`}
          </span>
        )}
      </div>

      {tagFacets.length > 0 && (
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

      {error && <div className="card error">{error}</div>}
      {tagError && <div className="card error">{tagError}</div>}
      {!apps && !error && <span className="spinner" />}

      {apps && (
        <div className="card" style={{ padding: 0 }}>
          <table className="table">
            <thead>
              <tr>
                <th>Display name</th>
                <th>Application (client) ID</th>
                <th>Tags</th>
                <th>Credentials</th>
                <th>Sign-in audience</th>
              </tr>
            </thead>
            <tbody>
              {visible.map((a) => (
                <tr key={a.id} onClick={() => nav(`/applications/${a.id}`)}>
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
                  <td colSpan={5}>
                    <div className="empty">
                      {tagFilter
                        ? `No applications tagged “${tagLabel(tagFilter)}”.`
                        : debouncedSearch
                          ? 'No applications match your search.'
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
