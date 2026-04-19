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
