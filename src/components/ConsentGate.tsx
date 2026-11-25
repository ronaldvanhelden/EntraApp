import { useEffect, useState, type ReactNode } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import type { GraphTokenFn } from '../graph/client';

// Wraps content that needs a permission outside the read set. The token is
// tried silently; when consent is still missing a button asks for it instead
// of a popup opening on page load (which the browser would block).
export function ConsentGate({
  scopes,
  what,
  inline = false,
  children,
}: {
  scopes: string[];
  // What the permission is for, e.g. "directory roles".
  what: string;
  // Inside a details row rather than filling a tab or page.
  inline?: boolean;
  children: (token: GraphTokenFn) => ReactNode;
}) {
  const silent = useGraphToken(scopes, false);
  const interactive = useGraphToken(scopes);
  const [state, setState] = useState<'checking' | 'ok' | 'needed'>('checking');
  const [error, setError] = useState<string | null>(null);
  useEffect(() => {
    let cancelled = false;
    silent()
      .then(() => !cancelled && setState('ok'))
      .catch(() => !cancelled && setState((s) => (s === 'ok' ? s : 'needed')));
    return () => {
      cancelled = true;
    };
  }, [silent]);

  if (state === 'ok') return <>{children(interactive)}</>;
  const box = inline ? undefined : 'related-note';
  if (state === 'checking')
    return (
      <div className={box}>
        <span className="spinner" />
      </div>
    );
  return (
    <div className={box}>
      <div className="row" style={{ gap: 12, alignItems: 'center' }}>
        <span className="muted">
          Showing {what} needs the <span className="mono">{scopes.join(', ')}</span> permission.
        </span>
        <button
          onClick={() =>
            interactive()
              .then(() => setState('ok'))
              .catch((e: unknown) => setError(e instanceof Error ? e.message : String(e)))
          }
        >
          Grant access
        </button>
      </div>
      {error && (
        <p className="error" style={{ marginBottom: 0 }}>
          {error}
        </p>
      )}
    </div>
  );
}
