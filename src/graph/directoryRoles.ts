import { graphAll } from './client';

type TokenFn = () => Promise<string>;

// Built-in role template IDs — identical in every tenant.
export const ROLE_GLOBAL_ADMIN = '62e90394-69f5-4237-9190-012177145e10';
export const ROLE_PRIVILEGED_ROLE_ADMIN = 'e8611ab8-c189-46e8-94e1-60213ab1f814';
export const ROLE_CLOUD_APP_ADMIN = '158c047a-c907-4556-b7ef-446551a6b5f7';
export const ROLE_APP_ADMIN = '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3';

// APIs whose application permissions (app roles) only Global Admin or
// Privileged Role Admin may consent to.
export const PRIVILEGED_CONSENT_RESOURCES = new Set([
  '00000003-0000-0000-c000-000000000000', // Microsoft Graph
  '00000002-0000-0000-c000-000000000000', // Azure AD Graph
]);

interface DirectoryRole {
  id: string;
  displayName?: string;
  roleTemplateId: string;
}

// Directory roles the signed-in user currently holds. Only *active*
// assignments show up here — a PIM-eligible role that hasn't been activated
// is absent, which is exactly what we want for gating admin actions.
export async function listMyActiveRoleTemplateIds(token: TokenFn) {
  const roles = await graphAll<DirectoryRole>(
    token,
    '/me/transitiveMemberOf/microsoft.graph.directoryRole',
    { query: { $select: 'id,displayName,roleTemplateId' } },
  );
  return new Set(roles.map((r) => r.roleTemplateId));
}
