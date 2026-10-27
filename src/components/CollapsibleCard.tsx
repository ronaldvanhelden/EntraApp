import { useState, type ReactNode } from 'react';

const STORAGE_PREFIX = 'entraapp.collapsed.';

function readCollapsed(storageKey: string | undefined, fallback: boolean) {
  if (!storageKey) return fallback;
  try {
    const raw = localStorage.getItem(STORAGE_PREFIX + storageKey);
    return raw === null ? fallback : raw === '1';
  } catch {
    return fallback;
  }
}

function writeCollapsed(storageKey: string | undefined, value: boolean) {
  if (!storageKey) return;
  try {
    localStorage.setItem(STORAGE_PREFIX + storageKey, value ? '1' : '0');
  } catch {
    /* storage unavailable — state just won't persist */
  }
}

// A .card whose body can be folded away by clicking its header. When a
// storageKey is given the open/closed state is remembered across visits.
export function CollapsibleCard({
  title,
  subtitle,
  action,
  storageKey,
  defaultCollapsed = false,
  className,
  children,
}: {
  title: ReactNode;
  subtitle?: ReactNode;
  action?: ReactNode;
  storageKey?: string;
  defaultCollapsed?: boolean;
  className?: string;
  children?: ReactNode;
}) {
  const [collapsed, setCollapsed] = useState(() =>
    readCollapsed(storageKey, defaultCollapsed),
  );
  const toggle = () =>
    setCollapsed((prev) => {
      writeCollapsed(storageKey, !prev);
      return !prev;
    });

  return (
    <div
      className={`card collapsible${collapsed ? ' collapsed' : ''}${className ? ` ${className}` : ''}`}
    >
      <div
        className="card-header"
        onClick={toggle}
        role="button"
        aria-expanded={!collapsed}
      >
        <div className="row" style={{ gap: 8, alignItems: 'flex-start' }}>
          <span className="chevron" aria-hidden>
            ▾
          </span>
          <div>
            <h3>{title}</h3>
            {subtitle && (
              <div className="muted" style={{ fontSize: 12, marginTop: 4 }}>
                {subtitle}
              </div>
            )}
          </div>
        </div>
        {action && (
          <div onClick={(e) => e.stopPropagation()}>{action}</div>
        )}
      </div>
      {!collapsed && <div className="card-body">{children}</div>}
    </div>
  );
}
