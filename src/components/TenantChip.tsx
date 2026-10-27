import { useEffect, useState } from 'react';
import { Link } from 'react-router-dom';
import { useAuthConfig } from '../auth/context';
import { useCurrentTenantId } from '../auth/useCurrentTenantId';
import { useGraphToken } from '../auth/useGraphToken';
import { findTenantInformation } from '../graph/organization';

// Header indicator for the directory the session is actually pointed at. It
// matters most when the URL pins a tenant (?tenant=debble.com) and the signed-in
// account lives somewhere else — the pin icon makes that explicit.
export function TenantChip() {
  const { config, tenantOverride } = useAuthConfig();
  const tenantId = useCurrentTenantId();
  const token = useGraphToken();
  const [name, setName] = useState<string | null>(null);
  const [domain, setDomain] = useState<string | null>(null);

  useEffect(() => {
    if (!tenantId) return;
    let cancelled = false;
    setName(null);
    setDomain(null);
    findTenantInformation(token, tenantId)
      .then((info) => {
        if (cancelled) return;
        setName(info.displayName ?? null);
        setDomain(info.defaultDomainName ?? null);
      })
      .catch(() => {
        // CrossTenantInformation.ReadBasic.All may not be consented; the
        // configured tenant value below is still a useful label.
      });
    return () => {
      cancelled = true;
    };
  }, [token, tenantId]);

  const label = name ?? domain ?? tenantOverride ?? tenantId ?? config.tenantId;
  const title = [
    name && `Tenant: ${name}`,
    domain && `Domain: ${domain}`,
    tenantId && `Tenant ID: ${tenantId}`,
    tenantOverride
      ? `Pinned by the URL (?tenant=${tenantOverride})`
      : `Default tenant: ${config.tenantId}`,
  ]
    .filter(Boolean)
    .join('\n');

  return (
    <Link to="/settings" className={`tenant-chip${tenantOverride ? ' pinned' : ''}`} title={title}>
      {tenantOverride && (
        <svg
          viewBox="0 0 16 16"
          width="12"
          height="12"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.5"
          strokeLinecap="round"
          strokeLinejoin="round"
          aria-hidden
        >
          <path d="M8 10.5V14" />
          <path d="M4.5 6.5h7l1 4h-9z" />
          <path d="M6.5 6.5V2.5h3v4" />
        </svg>
      )}
      <span className="tenant-chip-label">{label}</span>
    </Link>
  );
}
