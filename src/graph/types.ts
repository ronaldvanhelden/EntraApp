// Minimal Graph entity shapes — only fields we consume.

export interface Application {
  id: string;
  appId: string;
  displayName: string;
  createdDateTime?: string;
  signInAudience?: string;
  publisherDomain?: string;
  notes?: string;
  identifierUris?: string[];
  requiredResourceAccess?: RequiredResourceAccess[];
  passwordCredentials?: PasswordCredential[];
  keyCredentials?: KeyCredential[];
}

export interface KeyCredential {
  keyId: string;
  type?: string;
  usage?: string;
  displayName?: string | null;
  startDateTime?: string;
  endDateTime?: string;
  customKeyIdentifier?: string | null;
}

export interface PasswordCredential {
  keyId: string;
  displayName?: string | null;
  hint?: string | null;
  startDateTime?: string;
  endDateTime?: string;
  secretText?: string | null;
}

export interface FederatedIdentityCredential {
  id: string;
  name: string;
  issuer: string;
  subject: string;
  audiences: string[];
  description?: string | null;
}

export interface AppCredentialSignInActivity {
  keyId: string;
  keyType: 'clientSecret' | 'certificate';
  keyUsage?: 'Sign' | 'Verify';
  createdDateTime?: string;
  signInActivity?: {
    lastSignInDateTime?: string;
    requestId?: string;
    resourceId?: string;
    resourceDisplayName?: string;
  };
}

export interface ServicePrincipal {
  id: string;
  appId: string;
  displayName: string;
  servicePrincipalType?: string;
  accountEnabled?: boolean;
  appRoles?: AppRole[];
  oauth2PermissionScopes?: OAuth2PermissionScope[];
  tags?: string[];
  publisherName?: string;
  appOwnerOrganizationId?: string;
}

export interface AppRole {
  id: string;
  value: string;
  displayName: string;
  description: string;
  isEnabled: boolean;
  allowedMemberTypes: string[];
}

export interface OAuth2PermissionScope {
  id: string;
  value: string;
  adminConsentDisplayName: string;
  adminConsentDescription: string;
  userConsentDisplayName?: string;
  userConsentDescription?: string;
  type: 'User' | 'Admin';
  isEnabled: boolean;
}

export interface RequiredResourceAccess {
  resourceAppId: string;
  resourceAccess: ResourceAccess[];
}

export interface ResourceAccess {
  id: string;
  type: 'Scope' | 'Role';
}

export interface OAuth2PermissionGrant {
  id: string;
  clientId: string;
  consentType: 'AllPrincipals' | 'Principal';
  principalId?: string | null;
  resourceId: string;
  scope: string;
}

export interface AppRoleAssignment {
  id: string;
  appRoleId: string;
  createdDateTime?: string;
  principalDisplayName?: string;
  principalId: string;
  principalType?: string;
  resourceDisplayName?: string;
  resourceId: string;
}
