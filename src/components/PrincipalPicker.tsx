import { useEffect, useMemo, useState } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import {
  searchGroups,
  searchUsers,
  type PrincipalKind,
  type PrincipalRef,
} from '../graph/directoryObjects';
import { listServicePrincipals } from '../graph/servicePrincipals';

interface Props {
  kinds: PrincipalKind[];
  selected: PrincipalRef | null;
  onChange: (p: PrincipalRef | null) => void;
  placeholder?: string;
  autoFocus?: boolean;
}

const KIND_LABEL: Record<PrincipalKind, string> = {
  user: 'User',
  group: 'Group',
  sp: 'Service principal',
};

export function PrincipalPicker({
  kinds,
  selected,
  onChange,
  placeholder,
  autoFocus,
}: Props) {
  const token = useGraphToken();
  const [kind, setKind] = useState<PrincipalKind>(kinds[0]);
  const [search, setSearch] = useState('');
  const [results, setResults] = useState<PrincipalRef[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!kinds.includes(kind)) setKind(kinds[0]);
  }, [kinds, kind]);

  useEffect(() => {
    if (selected) return;
    const q = search.trim();
    if (q.length < 2) {
      setResults([]);
      setError(null);
      return;
    }
    const handle = setTimeout(async () => {
      setLoading(true);
      setError(null);
      try {
        let refs: PrincipalRef[] = [];
        if (kind === 'user') {
          const users = await searchUsers(token, q);
          refs = users.map((u) => ({
            id: u.id,
            displayName: u.displayName,
            kind: 'user',
            subtitle: u.userPrincipalName || u.mail,
          }));
        } else if (kind === 'group') {
          const groups = await searchGroups(token, q);
          refs = groups.map((g) => ({
            id: g.id,
            displayName: g.displayName,
            kind: 'group',
            subtitle: g.mail || g.mailNickname,
          }));
        } else {
          const sps = await listServicePrincipals(token, q);
          refs = sps.slice(0, 25).map((sp) => ({
            id: sp.id,
            displayName: sp.displayName,
            kind: 'sp',
            subtitle: sp.appId,
          }));
        }
        refs.sort((a, b) => a.displayName.localeCompare(b.displayName));
        setResults(refs);
      } catch (e: unknown) {
        setError(e instanceof Error ? e.message : String(e));
        setResults([]);
      } finally {
        setLoading(false);
      }
    }, 250);
    return () => clearTimeout(handle);
  }, [search, token, kind, selected]);

  const tabs = useMemo(() => kinds, [kinds]);

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
        <span className="badge">{KIND_LABEL[selected.kind]}</span>
        <div className="grow">
          <div style={{ fontWeight: 600 }}>{selected.displayName}</div>
          {selected.subtitle && (
            <div className="mono muted" style={{ fontSize: 12 }}>
              {selected.subtitle}
            </div>
          )}
        </div>
        <button className="ghost" onClick={() => onChange(null)}>
          Clear
        </button>
      </div>
    );
  }

  return (
    <div className="stack" style={{ gap: 6 }}>
      {tabs.length > 1 && (
        <div className="row" style={{ gap: 4 }}>
          {tabs.map((k) => (
            <button
              key={k}
              className={kind === k ? 'primary' : ''}
              style={{ padding: '4px 10px', fontSize: 12 }}
              onClick={() => setKind(k)}
            >
              {KIND_LABEL[k]}
            </button>
          ))}
        </div>
      )}
      <input
        autoFocus={autoFocus}
        placeholder={placeholder ?? `Search ${KIND_LABEL[kind].toLowerCase()}s (min 2 chars)…`}
        value={search}
        onChange={(e) => setSearch(e.target.value)}
      />
      {loading && (
        <div className="row" style={{ gap: 8 }}>
          <span className="spinner" />
          <span className="muted">Searching…</span>
        </div>
      )}
      {error && <div className="error">{error}</div>}
      {!loading && results.length > 0 && (
        <div
          style={{
            maxHeight: 220,
            overflowY: 'auto',
            border: '1px solid var(--border-subtle)',
            borderRadius: 6,
          }}
        >
          <table className="table">
            <tbody>
              {results.map((r) => (
                <tr key={r.id} onClick={() => onChange(r)}>
                  <td>
                    <div>{r.displayName}</div>
                    {r.subtitle && (
                      <div className="mono muted" style={{ fontSize: 12 }}>
                        {r.subtitle}
                      </div>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {!loading && search.trim().length >= 2 && results.length === 0 && !error && (
        <div className="muted" style={{ fontSize: 12 }}>
          No matches.
        </div>
      )}
    </div>
  );
}
