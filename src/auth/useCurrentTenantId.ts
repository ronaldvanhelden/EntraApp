import { useMsal } from '@azure/msal-react';

export function useCurrentTenantId(): string | undefined {
  const { instance, accounts } = useMsal();
  const account = instance.getActiveAccount() ?? accounts[0];
  return account?.tenantId;
}
