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
  federatedIdentityCredentials?: FederatedIdentityCredential[];
  info?: InformationalUrl;
  // Auth / platform configuration — which OAuth 2.0 flows the app supports.
  web?: WebApplication;
  spa?: SpaApplication;
  publicClient?: PublicClientApplication;
  isFallbackPublicClient?: boolean;
}

export interface WebApplication {
  redirectUris?: string[];
  homePageUrl?: string | null;
  logoutUrl?: string | null;
  implicitGrantSettings?: ImplicitGrantSettings;
}
export interface SpaApplication {
  redirectUris?: string[];
}
export interface PublicClientApplication {
  redirectUris?: string[];
}
export interface ImplicitGrantSettings {
  enableIdTokenIssuance?: boolean;
  enableAccessTokenIssuance?: boolean;
}

export interface DirectoryObjectLite {
  id: string;
  displayName?: string | null;
  userPrincipalName?: string | null;
}

export interface InformationalUrl {
  logoUrl?: string | null;
  marketingUrl?: string | null;
  privacyStatementUrl?: string | null;
  supportUrl?: string | null;
  termsOfServiceUrl?: string | null;
}

export interface KeyCredential {
  keyId: string;
  type?: string;
  usage?: string;
  displayName?: string | null;
  startDateTime?: string;
  endDateTime?: string;
  customKeyIdentifier?: string | null;
  // Write-only: base64-encoded DER certificate (public X.509). Graph never
  // returns this on reads, but including it in a PATCH adds a new key.
  key?: string | null;
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

// /beta/reports/servicePrincipalSignInActivities/{appId}
export interface SignInActivityDetail {
  lastSignInDateTime?: string;
  lastSignInRequestId?: string;
}
export interface ServicePrincipalSignInActivity {
  id?: string;
  appId?: string;
  lastSignInActivity?: SignInActivityDetail;
  delegatedClientSignInActivity?: SignInActivityDetail;
  delegatedResourceSignInActivity?: SignInActivityDetail;
  applicationAuthenticationClientSignInActivity?: SignInActivityDetail;
  applicationAuthenticationResourceSignInActivity?: SignInActivityDetail;
}

// /beta/auditLogs/signIns — one record per sign-in event.
export type SignInClientCredentialType =
  | 'none'
  | 'clientSecret'
  | 'clientAssertion'
  | 'federatedIdentityCredential'
  | 'managedIdentity'
  | 'unknown'
  | string;

export interface SignInKeyValue {
  key?: string;
  value?: string;
}

export interface SignIn {
  id: string;
  createdDateTime?: string;
  appId?: string;
  appDisplayName?: string;
  ipAddress?: string;
  clientAppUsed?: string;
  authenticationProtocol?: string;
  // User sign-ins
  userDisplayName?: string;
  userPrincipalName?: string;
  userId?: string;
  // App-only sign-ins
  servicePrincipalId?: string;
  servicePrincipalName?: string;
  // Resource targeted
  resourceDisplayName?: string;
  resourceId?: string;
  // Event classification and credentials (beta)
  signInEventTypes?: string[];
  clientCredentialType?: SignInClientCredentialType;
  federatedCredentialId?: string;
  // Dedicated fields populated on app-only (service principal) sign-ins.
  // Preferred over digging through authenticationProcessingDetails.
  servicePrincipalCredentialKeyId?: string;
  servicePrincipalCredentialThumbprint?: string;
  authenticationProcessingDetails?: SignInKeyValue[];
  authenticationMethodsUsed?: string[];
  isInteractive?: boolean;
  status?: {
    errorCode?: number;
    failureReason?: string;
    additionalDetails?: string;
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
  info?: InformationalUrl;
  keyCredentials?: KeyCredential[];
  passwordCredentials?: PasswordCredential[];
  createdDateTime?: string;
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
