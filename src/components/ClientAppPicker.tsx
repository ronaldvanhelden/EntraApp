import { useEffect, useMemo, useState } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import { listServicePrincipals } from '../graph/servicePrincipals';
import type { ServicePrincipal } from '../graph/types';
import { FIRST_PARTY_APPS, getFirstPartyApp } from '../lib/firstPartyApps';
import { THIRD_PARTY_APPS, getThirdPartyApp } from '../lib/thirdPartyApps';

// Picks a client application by appId for "Authorized client applications".
// Offers both Microsoft first-party apps (from a static catalog — most have no
// service principal in the tenant until first sign-in) and service principals
// that exist in this tenant. A raw GUID can still be typed for clients that
// live in another tenant.

// Tenants Microsoft's own first-party apps are registered in.
const MICROSOFT_TENANTS = new Set([
  'f8cdef31-a31e-4b4a-93e4-5f571e91255a', // Microsoft Services
  '72f988bf-86f1-41af-91ab-2d7cd011db47', // Microsoft (corp)
  '33e01921-4d64-4f8c-a055-5bdaffd5e33d', // Microsoft (AME)
  'cdc5aeea-15c5-4db6-b079-fcadd2505dc2', // Microsoft (PME)
]);

export function isMicrosoftApp(
  appId: string,
  sp?: Pick<ServicePrincipal, 'appOwnerOrganizationId'> | null,
) {
  if (getFirstPartyApp(appId)) return true;
  const owner = sp?.appOwnerOrganizationId?.toLowerCase();
  return !!owner && MICROSOFT_TENANTS.has(owner);
}

export interface ClientAppChoice {
  appId: string;
  displayName: string | null;
  microsoft: boolean;
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

interface Props {
  selected: ClientAppChoice | null;
  onChange: (c: ClientAppChoice | null) => void;
  /** Lowercased appIds to mark as already added. */
  taken?: string[];
  disabled?: boolean;
  autoFocus?: boolean;
}

export function MicrosoftBadge() {
  return (
    <span className="badge app" title="Microsoft first-party application">
      Microsoft
    </span>
  );
}

export function ThirdPartyBadge({ note }: { note?: string }) {
  return (
    <span
      className="badge pending"
      title={`Known third-party application${note ? `: ${note}` : ''}`}
    >
      Third-party
    </span>
  );
}

// Badges for a client: Microsoft first-party or a known third-party app.
export function KnownAppBadges({ appId, microsoft }: { appId: string; microsoft: boolean }) {
  const tp = getThirdPartyApp(appId);
  return (
    <>
      {microsoft && !tp && <MicrosoftBadge />}
      {tp && <ThirdPartyBadge note={tp.note} />}
    </>
  );
}

// Investigation hint from the catalogs, if any.
export function knownAppNote(appId: string): string | undefined {
  return getFirstPartyApp(appId)?.note ?? getThirdPartyApp(appId)?.note;
}

export function ClientAppPicker({ selected, onChange, taken = [], disabled, autoFocus }: Props) {
  const token = useGraphToken();
  const [search, setSearch] = useState('');
  const [tenantResults, setTenantResults] = useState<ServicePrincipal[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const q = search.trim();
  const ql = q.toLowerCase();

  useEffect(() => {
    if (selected || q.length < 2) {
      setTenantResults([]);
      setError(null);
      return;
    }
    let cancelled = false;
    const handle = setTimeout(async () => {
      setLoading(true);
      setError(null);
      try {
        const sps = await listServicePrincipals(token, q);
        if (!cancelled) setTenantResults(sps.slice(0, 25));
      } catch (e: unknown) {
        if (!cancelled) {
          setError(e instanceof Error ? e.message : String(e));
          setTenantResults([]);
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    }, 250);
    return () => {
      cancelled = true;
      clearTimeout(handle);
    };
  }, [token, q, selected]);

  const tenantIds = useMemo(
    () => new Set(tenantResults.map((sp) => sp.appId.toLowerCase())),
    [tenantResults],
  );

  // Catalog matches not already shown as a tenant service principal.
  const firstParty = useMemo(() => {
    const list = ql
      ? FIRST_PARTY_APPS.filter(
          (a) => a.displayName.toLowerCase().includes(ql) || a.appId.startsWith(ql),
        )
      : FIRST_PARTY_APPS;
    return list.filter((a) => !tenantIds.has(a.appId));
  }, [ql, tenantIds]);
  const thirdParty = useMemo(() => {
    const list = ql
      ? THIRD_PARTY_APPS.filter(
          (a) => a.displayName.toLowerCase().includes(ql) || a.appId.startsWith(ql),
        )
      : THIRD_PARTY_APPS;
    return list.filter((a) => !tenantIds.has(a.appId));
  }, [ql, tenantIds]);

  if (selected) {
    return (
      <div
        className="row"
        style={{
          gap: 8,
          padding: 8,
          border: '1px solid var(--border-subtle)',
          borderRadius: 6,
          background: 'var(--bg)',
        }}
      >
        <div className="grow">
          <div className="row" style={{ gap: 6 }}>
            <strong>{selected.displayName ?? 'Unknown application'}</strong>
            <KnownAppBadges appId={selected.appId} microsoft={selected.microsoft} />
          </div>
          <div className="mono muted" style={{ fontSize: 12 }}>
            {selected.appId}
          </div>
          {knownAppNote(selected.appId) && (
            <div className="muted" style={{ fontSize: 12 }}>
              {knownAppNote(selected.appId)}
            </div>
          )}
        </div>
        {!disabled && (
          <button className="ghost" onClick={() => onChange(null)}>
            Change
          </button>
        )}
      </div>
    );
  }

  const isTaken = (appId: string) => taken.includes(appId.toLowerCase());
  const pick = (c: ClientAppChoice) => {
    if (!isTaken(c.appId)) onChange(c);
  };

  const row = (c: ClientAppChoice, key: string) => (
    <tr
      key={key}
      onClick={() => pick(c)}
      style={isTaken(c.appId) ? { opacity: 0.5, cursor: 'default' } : undefined}
    >
      <td>
        <div className="row" style={{ gap: 6 }}>
          <span>{c.displayName}</span>
          <KnownAppBadges appId={c.appId} microsoft={c.microsoft} />
          {isTaken(c.appId) && <span className="badge">Already added</span>}
        </div>
        <div className="mono muted" style={{ fontSize: 12 }}>
          {c.appId}
        </div>
        {knownAppNote(c.appId) && (
          <div className="muted" style={{ fontSize: 12 }}>
            {knownAppNote(c.appId)}
          </div>
        )}
      </td>
    </tr>
  );

  const groupHeader = (label: string) => (
    <tr>
      <th style={{ position: 'sticky', top: 0, background: 'var(--bg-elevated, var(--bg))' }}>
        {label}
      </th>
    </tr>
  );

  // A full GUID nobody lists — offer it as-is (client from another tenant).
  const rawGuid =
    GUID_RE.test(q) && !tenantIds.has(ql) && !getFirstPartyApp(q) && !getThirdPartyApp(q)
      ? ql
      : null;

  return (
    <div className="stack" style={{ gap: 6 }}>
      <input
        autoFocus={autoFocus}
        disabled={disabled}
        placeholder="Search by name or client ID — apps in this tenant, Microsoft and known third-party apps"
        value={search}
        onChange={(e) => setSearch(e.target.value)}
      />
      {error && <div className="error">{error}</div>}
      <div
        style={{
          maxHeight: 280,
          overflowY: 'auto',
          border: '1px solid var(--border-subtle)',
          borderRadius: 6,
        }}
      >
        <table className="table">
          <tbody>
            {rawGuid && (
              <>
                {groupHeader('Client ID')}
                {row(
                  {
                    appId: rawGuid,
                    displayName: 'Use this client ID',
                    microsoft: false,
                  },
                  `raw:${rawGuid}`,
                )}
              </>
            )}
            {q.length >= 2 && (
              <>
                {groupHeader('In this tenant')}
                {loading ? (
                  <tr>
                    <td>
                      <span className="spinner" /> <span className="muted">Searching…</span>
                    </td>
                  </tr>
                ) : tenantResults.length === 0 ? (
                  <tr>
                    <td className="muted">No service principals match.</td>
                  </tr>
                ) : (
                  tenantResults.map((sp) =>
                    row(
                      {
                        appId: sp.appId,
                        displayName: sp.displayName,
                        microsoft: isMicrosoftApp(sp.appId, sp),
                      },
                      `sp:${sp.id}`,
                    ),
                  )
                )}
              </>
            )}
            {groupHeader('Microsoft first-party applications')}
            {firstParty.length === 0 ? (
              <tr>
                <td className="muted">No Microsoft apps match.</td>
              </tr>
            ) : (
              firstParty.map((a) =>
                row(
                  { appId: a.appId, displayName: a.displayName, microsoft: true },
                  `fp:${a.appId}`,
                ),
              )
            )}
            {groupHeader('Known third-party applications')}
            {thirdParty.length === 0 ? (
              <tr>
                <td className="muted">No known third-party apps match.</td>
              </tr>
            ) : (
              thirdParty.map((a) =>
                row(
                  { appId: a.appId, displayName: a.displayName, microsoft: false },
                  `tp:${a.appId}`,
                ),
              )
            )}
          </tbody>
        </table>
      </div>
      <div className="muted" style={{ fontSize: 12 }}>
        Client in another tenant? Paste its application (client) ID.
      </div>
    </div>
  );
}
