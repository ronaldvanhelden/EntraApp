// Deep links into the Microsoft Entra admin center. The tenant id goes in
// the path so the link opens in the right directory even when the user's
// browser session defaults to another one.

const ENTRA = 'https://entra.microsoft.com/';

function base(tenantId?: string): string {
  return tenantId ? `${ENTRA}${tenantId}` : ENTRA;
}

export function entraAppRegistrationUrl(appId: string, tenantId?: string): string {
  return `${base(tenantId)}#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/${appId}`;
}

export function entraEnterpriseAppUrl(
  spObjectId: string,
  appId: string,
  tenantId?: string,
): string {
  return `${base(tenantId)}#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId/${spObjectId}/appId/${appId}`;
}
