import { useEffect, useState } from 'react';
import {
  getPrivilegeLevel,
  getScopeMeta,
  loadScopeCatalog,
  privilegeLabel,
  type ScopeKind,
  type ScopeMeta,
} from '../graph/scopeCatalog';

// Shared hook: loads the Microsoft Graph permissions catalog once per session
// and returns the reduced map. All consumers read from the same memory cache
// after the first resolve.
export function useScopeCatalog() {
  const [catalog, setCatalog] = useState<Record<string, ScopeMeta> | null>(null);
  useEffect(() => {
    let cancelled = false;
    loadScopeCatalog()
      .then((c) => {
        if (!cancelled) setCatalog(c);
      })
      .catch(() => {
        // Fetching the catalog is best-effort; without it we simply render no
        // privilege info. No point surfacing the error to the user.
      });
    return () => {
      cancelled = true;
    };
  }, []);
  return catalog;
}

export function PrivilegeBadge({
  catalog,
  resourceAppId,
  scope,
  kind,
}: {
  catalog: Record<string, ScopeMeta> | null;
  resourceAppId?: string;
  scope?: string;
  kind: ScopeKind;
}) {
  const meta = getScopeMeta(catalog, resourceAppId, scope);
  const level = getPrivilegeLevel(meta, kind);
  if (level === undefined) return null;
  const { label, tone } = privilegeLabel(level);
  return (
    <span
      className={`priv-badge priv-${tone}`}
      title={`Microsoft privilege level: ${label}${
        meta?.[kind]?.requiresAdminConsent ? ' · requires admin consent' : ''
      }`}
    >
      {label}
    </span>
  );
}
