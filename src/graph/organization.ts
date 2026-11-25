import { graph } from './client';

type TokenFn = () => Promise<string>;

export interface TenantInformation {
  tenantId: string;
  displayName?: string;
  defaultDomainName?: string;
  federationBrandName?: string;
}

// Resolve a tenant id to its basic public information. Requires the
// CrossTenantInformation.ReadBasic.All delegated scope.
export function findTenantInformation(token: TokenFn, tenantId: string) {
  return graph<TenantInformation>(
    token,
    `/tenantRelationships/findTenantInformationByTenantId(tenantId='${tenantId}')`,
  );
}

interface AssignedPlan {
  capabilityStatus?: string;
  service?: string;
  servicePlanId?: string;
}

export interface Organization {
  id: string;
  displayName?: string;
  verifiedDomains?: { name: string; isDefault?: boolean; isInitial?: boolean }[];
  createdDateTime?: string | null;
  countryLetterCode?: string | null;
  // AAD for a workforce tenant, CIAM for an external (customer) tenant.
  tenantType?: string | null;
  onPremisesSyncEnabled?: boolean | null;
  onPremisesLastSyncDateTime?: string | null;
  technicalNotificationMails?: string[];
  assignedPlans?: AssignedPlan[];
}

// The signed-in tenant's own record.
export async function getOrganization(token: TokenFn): Promise<Organization | null> {
  const res = await graph<{ value: Organization[] }>(token, '/organization', {
    query: {
      $select:
        'id,displayName,verifiedDomains,createdDateTime,countryLetterCode,tenantType,' +
        'onPremisesSyncEnabled,onPremisesLastSyncDateTime,technicalNotificationMails,assignedPlans',
    },
  });
  return res.value[0] ?? null;
}

// Service plan ids as listed in Microsoft's licensing service plan reference.
const PLAN_P1 = '41781fb2-bc02-4b7c-bd55-b576c07bb09d';
const PLAN_P2 = 'eec0eb4f-6444-4f95-aba0-50c24d67f998';
const PLAN_GOVERNANCE = 'e866a266-3cff-43a3-acca-0c90a7e00c8b';

export interface EntraPlan {
  // The highest Entra ID edition any licence in the tenant brings.
  edition: 'Free' | 'P1' | 'P2';
  // The Entra ID Governance add-on.
  governance: boolean;
}

// Which Entra ID edition the tenant has, read from the service plans its
// licences provide. One P2 licence already switches the P2 features on for
// the tenant, so this says what works — not how many users are licensed.
export function entraPlan(org: Organization): EntraPlan {
  const active = new Set(
    (org.assignedPlans ?? [])
      .filter((p) => p.capabilityStatus === 'Enabled' || p.capabilityStatus === 'Warning')
      .map((p) => p.servicePlanId?.toLowerCase()),
  );
  return {
    edition: active.has(PLAN_P2) ? 'P2' : active.has(PLAN_P1) ? 'P1' : 'Free',
    governance: active.has(PLAN_GOVERNANCE),
  };
}
