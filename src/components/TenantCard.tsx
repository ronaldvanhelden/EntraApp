import { useEffect, useState, type ReactNode } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import { entraPlan, getOrganization, type Organization } from '../graph/organization';
import { countCollection } from '../graph/overview';
import { CopyButton } from './CopyButton';

// The directory at a glance, on top of the Overview: which tenant this is,
// which Entra ID edition it has (that decides whether sign-in logs, PIM and
// risk data exist at all) and how big it is.

const PLAN_HINT = {
  Free: 'No Entra ID P1 or P2: no sign-in logs or Conditional Access, no PIM.',
  P1: 'Entra ID P1: Conditional Access and sign-in logs. PIM and risk detection need P2.',
  P2: 'Entra ID P2: includes Privileged Identity Management and Identity Protection.',
};

function Fact({ label, children }: { label: string; children: ReactNode }) {
  return (
    <>
      <div className="k">{label}</div>
      <div>{children}</div>
    </>
  );
}

function Count({ load }: { load: () => Promise<number> }) {
  const [n, setN] = useState<number | null | undefined>(undefined);
  useEffect(() => {
    let cancelled = false;
    load()
      .then((v) => !cancelled && setN(v))
      .catch(() => !cancelled && setN(null));
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  if (n === undefined) return <span className="spinner" />;
  return n === null ? <span className="muted">—</span> : <>{n.toLocaleString()}</>;
}

export function TenantCard() {
  const token = useGraphToken();
  const [org, setOrg] = useState<Organization | null | undefined>(undefined);
  useEffect(() => {
    let cancelled = false;
    getOrganization(token)
      .then((o) => !cancelled && setOrg(o))
      .catch(() => !cancelled && setOrg(null));
    return () => {
      cancelled = true;
    };
  }, [token]);

  // Without the organisation record there is nothing worth a card.
  if (org === null) return null;
  if (org === undefined)
    return (
      <div className="card tenant-card">
        <span className="spinner" />
      </div>
    );

  const plan = entraPlan(org);
  const domains = org.verifiedDomains ?? [];
  const primary = domains.find((d) => d.isDefault)?.name ?? domains[0]?.name;
  const initial = domains.find((d) => d.isInitial)?.name;
  return (
    <div className="card tenant-card">
      <div className="row" style={{ gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
        <h3 style={{ margin: 0 }}>{org.displayName ?? 'This tenant'}</h3>
        <span
          className={`badge ${plan.edition === 'Free' ? 'pending' : 'granted'}`}
          title={PLAN_HINT[plan.edition]}
        >
          Entra ID {plan.edition}
        </span>
        {plan.governance && (
          <span className="badge granted" title="Microsoft Entra ID Governance add-on">
            ID Governance
          </span>
        )}
        {org.tenantType && org.tenantType !== 'AAD' && (
          <span className="badge" title="Tenant type">
            {org.tenantType === 'CIAM' ? 'External (customer) tenant' : org.tenantType}
          </span>
        )}
      </div>
      <div className="kv compact tenant-facts">
        <Fact label="Tenant ID">
          <span className="mono">{org.id}</span>
          <CopyButton value={org.id} label="tenant ID" />
        </Fact>
        <Fact label="Primary domain">
          {primary ?? '—'}
          {initial && initial !== primary && <span className="muted"> · {initial}</span>}
        </Fact>
        <Fact label="Verified domains">
          <span title={domains.map((d) => d.name).join('\n')}>{domains.length}</span>
        </Fact>
        <Fact label="Users">
          <Count load={() => countCollection(token, 'users')} />
        </Fact>
        <Fact label="Groups">
          <Count load={() => countCollection(token, 'groups')} />
        </Fact>
        <Fact label="Directory sync">
          {org.onPremisesSyncEnabled ? (
            <>
              On-premises AD
              {org.onPremisesLastSyncDateTime && (
                <span className="muted">
                  {' · last '}
                  {new Date(org.onPremisesLastSyncDateTime).toLocaleString()}
                </span>
              )}
            </>
          ) : (
            'Cloud only'
          )}
        </Fact>
        <Fact label="Country">{org.countryLetterCode ?? '—'}</Fact>
        <Fact label="Created">
          {org.createdDateTime ? new Date(org.createdDateTime).toLocaleDateString() : '—'}
        </Fact>
      </div>
    </div>
  );
}
