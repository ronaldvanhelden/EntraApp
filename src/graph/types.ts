// Minimal Graph entity shapes — only fields we consume.

export interface Application {
  id: string;
  appId: string;
  displayName: string;
  createdDateTime?: string;
  // Set only on items returned from the directory recycle bin.
  deletedDateTime?: string | null;
  signInAudience?: string;
  publisherDomain?: string;
  notes?: string;
  // Alternate key. Writable only while null — immutable once set.
  uniqueName?: string | null;
  // Free-form directory-side labels. Entra copies these onto the associated
  // service principal, so they show up on the enterprise app too.
  tags?: string[];
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
  // What this app exposes to *other* clients: delegated scopes live under
  // api, application roles directly on the app. Both are replaced wholesale
  // on PATCH, and an entry must be disabled before it can be removed.
  api?: ApiApplication;
  appRoles?: AppRole[];
}

export interface ApiApplication {
  oauth2PermissionScopes?: OAuth2PermissionScope[];
  requestedAccessTokenVersion?: number | null;
  knownClientApplications?: string[];
  preAuthorizedApplications?: { appId: string; delegatedPermissionIds: string[] }[];
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
  // Where from
  location?: {
    city?: string | null;
    state?: string | null;
    countryOrRegion?: string | null;
  };
  autonomousSystemNumber?: number;
  userAgent?: string;
  deviceDetail?: {
    deviceId?: string | null;
    displayName?: string | null;
    operatingSystem?: string | null;
    browser?: string | null;
    isCompliant?: boolean | null;
    isManaged?: boolean | null;
    trustType?: string | null;
  };
  networkLocationDetails?: { networkType?: string; networkNames?: string[] }[];
  // Correlation / troubleshooting
  correlationId?: string;
  originalRequestId?: string;
  sessionId?: string;
  uniqueTokenIdentifier?: string;
  processingTimeInMilliseconds?: number;
  // Tenants involved
  homeTenantId?: string;
  homeTenantName?: string;
  resourceTenantId?: string;
  resourceServicePrincipalId?: string;
  resourceOwnerTenantId?: string;
  appOwnerTenantId?: string;
  crossTenantAccessType?: string;
  userType?: string;
  // Policy / risk
  conditionalAccessStatus?: string;
  authenticationRequirement?: string;
  riskLevelDuringSignIn?: string;
  riskLevelAggregated?: string;
  riskState?: string;
  riskDetail?: string;
  tokenIssuerName?: string;
  tokenIssuerType?: string;
  incomingTokenType?: string;
  managedServiceIdentity?: {
    msiType?: string;
    associatedResourceId?: string;
    federatedTokenId?: string;
    federatedTokenIssuer?: string;
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
  // Set only on items returned from the directory recycle bin.
  deletedDateTime?: string | null;
}

export interface AppRole {
  id: string;
  value: string;
  displayName: string;
  description: string;
  isEnabled: boolean;
  allowedMemberTypes: string[];
  // 'Application' for roles defined on the app registration; Entra also
  // synthesises roles (e.g. the default access role) with other origins.
  origin?: string;
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
