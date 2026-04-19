import { useMemo, useState } from 'react';
import {
  findScopesForEndpoint,
  privilegeLabel,
  type EndpointMatch,
  type ScopeKind,
  type ScopeMeta,
} from '../graph/scopeCatalog';

const METHODS = ['GET', 'POST', 'PATCH', 'PUT', 'DELETE'] as const;

interface Props {
  catalog: Record<string, ScopeMeta> | null;
  // Optional: when a user picks a result, callback hooks it back into the
  // parent's selection state (pre-selects the scope in the checklist).
  onPick?: (match: EndpointMatch, kind: ScopeKind) => void;
}

// Query the Graph permissions catalog by endpoint: user types a Graph URL
// and an HTTP method, and the panel lists every permission that grants
// access. RSC-style scopes are tagged so admins can distinguish tenant-wide
// consent from container-scoped (Teams/Groups/Chats) RSC.
export function ScopeLookupPanel({ catalog, onPick }: Props) {
  const [method, setMethod] = useState<(typeof METHODS)[number]>('GET');
  const [path, setPath] = useState('');
  const [submitted, setSubmitted] = useState(false);

  const matches = useMemo(() => {
    if (!submitted) return [] as EndpointMatch[];
    return findScopesForEndpoint(catalog, method, path);
  }, [catalog, method, path, submitted]);

  const delegatedMatches = matches.filter((m) => m.kinds.includes('delegated'));
  const applicationMatches = matches.filter((m) =>
    m.kinds.includes('application'),
  );

  return (
    <div>
      <div className="row" style={{ gap: 8, alignItems: 'flex-end' }}>
        <label className="field" style={{ flex: '0 0 110px', margin: 0 }}>
          <span>Method</span>
          <select
            value={method}
            onChange={(e) => setMethod(e.target.value as typeof METHODS[number])}
          >
            {METHODS.map((m) => (
              <option key={m} value={m}>
                {m}
              </option>
            ))}
          </select>
        </label>
        <label className="field" style={{ flex: 1, margin: 0 }}>
          <span>Graph endpoint</span>
          <input
            placeholder="/users/{id}/messages"
            value={path}
            onChange={(e) => setPath(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') setSubmitted(true);
            }}
          />
        </label>
        <button
          className="primary"
          disabled={!catalog || !path.trim()}
          onClick={() => setSubmitted(true)}
        >
          Find scopes
        </button>
      </div>
      <p className="muted" style={{ fontSize: 12, marginTop: 6 }}>
        Omit the <span className="mono">/v1.0</span> or{' '}
        <span className="mono">/beta</span> prefix. Placeholders like{' '}
        <span className="mono">{'{id}'}</span> match any single path segment.
      </p>

      {!catalog && (
        <div className="muted">
          <span className="spinner" /> Loading Microsoft Graph permissions
          catalog…
        </div>
      )}

      {submitted && catalog && (
        <>
          {matches.length === 0 ? (
            <div className="empty" style={{ marginTop: 12 }}>
              No scopes matched{' '}
              <span className="mono">
                {method} {path}
              </span>
              . Double-check the path template — segments like{' '}
              <span className="mono">{'{id}'}</span> must align exactly with
              Microsoft's documented paths.
            </div>
          ) : (
            <>
              <ResultsBlock
                title="Delegated permissions"
                kind="delegated"
                matches={delegatedMatches}
                onPick={onPick}
              />
              <ResultsBlock
                title="Application permissions"
                kind="application"
                matches={applicationMatches}
                onPick={onPick}
              />
            </>
          )}
        </>
      )}
    </div>
  );
}

function ResultsBlock({
  title,
  kind,
  matches,
  onPick,
}: {
  title: string;
  kind: ScopeKind;
  matches: EndpointMatch[];
  onPick?: (match: EndpointMatch, kind: ScopeKind) => void;
}) {
  if (matches.length === 0) return null;
  return (
    <div style={{ marginTop: 16 }}>
      <h4 style={{ margin: '8px 0' }}>
        {title} ({matches.length})
      </h4>
      <table className="table">
        <thead>
          <tr>
            <th>Scope</th>
            <th>Privilege</th>
            <th>Tags</th>
            {onPick && <th style={{ width: 80 }}></th>}
          </tr>
        </thead>
        <tbody>
          {matches.map((m) => {
            const scheme =
              kind === 'delegated' ? m.meta.delegated : m.meta.application;
            const { label, tone } = privilegeLabel(scheme?.privilegeLevel);
            return (
              <tr key={`${kind}:${m.scope}`}>
                <td>
                  <div className="mono" style={{ fontWeight: 600 }}>
                    {m.scope}
                  </div>
                  <div className="muted" style={{ fontSize: 11 }}>
                    matched{' '}
                    <span className="mono">{m.matchedPath}</span>
                  </div>
                </td>
                <td>
                  {scheme ? (
                    <span className={`priv-badge priv-${tone}`}>{label}</span>
                  ) : (
                    <span className="muted">—</span>
                  )}
                </td>
                <td>
                  <div className="row" style={{ gap: 4, flexWrap: 'wrap' }}>
                    {m.isRsc && (
                      <span
                        className="badge"
                        title="Resource-Specific Consent — consent is granted per container (Team/Chat/Group), not tenant-wide"
                        style={{
                          background: 'rgba(111, 66, 193, 0.15)',
                          color: '#c296f5',
                          border: '1px solid rgba(111, 66, 193, 0.45)',
                        }}
                      >
                        RSC
                      </span>
                    )}
                    {scheme?.requiresAdminConsent && (
                      <span className="badge" title="Requires admin consent">
                        Admin
                      </span>
                    )}
                  </div>
                </td>
                {onPick && (
                  <td style={{ textAlign: 'right' }}>
                    <button onClick={() => onPick(m, kind)}>Select</button>
                  </td>
                )}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
