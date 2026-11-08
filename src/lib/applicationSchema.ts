// JSON Schema for the Microsoft Graph v1.0 `application` resource, used by the
// manifest editor for validation, hover docs and completion. Hand-maintained
// from https://learn.microsoft.com/graph/api/resources/application — only the
// properties a GET /applications/{id} returns are covered. Read-only
// properties are marked with `readOnly` and are never sent on save.

const guid = {
  type: 'string',
  pattern: '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$',
};

const uriList = (description: string) => ({
  type: 'array',
  description,
  items: { type: 'string', format: 'uri' },
  uniqueItems: true,
});

const nullableString = (description: string) => ({ type: ['string', 'null'], description });

const keyCredential = {
  type: 'object',
  description: 'A certificate (public key) credential.',
  properties: {
    customKeyIdentifier: nullableString('Base64 thumbprint of the certificate.'),
    displayName: nullableString('Friendly name for the key.'),
    endDateTime: nullableString('When the credential expires (ISO 8601, UTC).'),
    key: nullableString(
      'Base64 raw certificate data. Always null on GET; only supplied when adding a certificate.',
    ),
    keyId: { ...guid, description: 'Unique identifier of the key.' },
    startDateTime: nullableString('When the credential becomes valid (ISO 8601, UTC).'),
    type: {
      type: 'string',
      enum: ['AsymmetricX509Cert', 'X509CertAndPassword', 'Symmetric'],
      description: 'Credential type. AsymmetricX509Cert for app certificates.',
    },
    usage: {
      type: 'string',
      enum: ['Verify', 'Sign'],
      description: 'Verify for app authentication certificates; Sign for SAML signing keys.',
    },
  },
};

const passwordCredential = {
  type: 'object',
  description: 'A client secret. Secret values are never returned after creation.',
  properties: {
    customKeyIdentifier: nullableString('Not used for secrets.'),
    displayName: nullableString('Description of the secret.'),
    endDateTime: nullableString('Expiry (ISO 8601, UTC).'),
    hint: nullableString('First three characters of the secret.'),
    keyId: { ...guid, description: 'Unique identifier of the secret.' },
    secretText: nullableString('Only returned once, when the secret is created.'),
    startDateTime: nullableString('When the secret becomes valid.'),
  },
};

const permissionScope = {
  type: 'object',
  description: 'A delegated permission this API exposes.',
  required: ['id', 'value', 'type', 'isEnabled'],
  properties: {
    id: { ...guid, description: 'Unique scope id. Never reuse an id.' },
    value: {
      type: 'string',
      description:
        'The scp claim value, e.g. "Files.Read". Clients request "<App ID URI>/<value>".',
    },
    type: {
      type: 'string',
      enum: ['User', 'Admin'],
      description: 'User: users may consent. Admin: only an admin may consent.',
    },
    isEnabled: {
      type: 'boolean',
      description: 'Must be false before the scope can be removed.',
    },
    adminConsentDisplayName: nullableString('Shown to admins on the consent screen.'),
    adminConsentDescription: nullableString('Shown to admins on the consent screen.'),
    userConsentDisplayName: nullableString('Shown to users on the consent screen.'),
    userConsentDescription: nullableString('Shown to users on the consent screen.'),
    origin: nullableString('Read-only marker set by Entra.'),
  },
};

const appRole = {
  type: 'object',
  description: 'A role that can be assigned to users, groups or applications.',
  required: ['id', 'allowedMemberTypes', 'displayName', 'isEnabled'],
  properties: {
    id: { ...guid, description: 'Unique role id. Never reuse an id.' },
    allowedMemberTypes: {
      type: 'array',
      items: { type: 'string', enum: ['User', 'Application'] },
      description:
        'User: assignable to users and groups (roles claim in user tokens). Application: assignable to apps (application permission).',
    },
    displayName: { type: 'string' },
    description: { type: 'string' },
    value: nullableString('The roles claim value, e.g. "Reports.Read.All".'),
    isEnabled: { type: 'boolean', description: 'Must be false before the role can be removed.' },
    origin: nullableString('Read-only marker set by Entra.'),
  },
};

const resourceAccess = {
  type: 'object',
  description: 'A permission this app requests from a resource (API permissions).',
  required: ['resourceAppId', 'resourceAccess'],
  properties: {
    resourceAppId: {
      ...guid,
      description:
        'appId of the resource, e.g. Microsoft Graph 00000003-0000-0000-c000-000000000000.',
    },
    resourceAccess: {
      type: 'array',
      items: {
        type: 'object',
        required: ['id', 'type'],
        properties: {
          id: { ...guid, description: 'Id of the scope or app role on the resource.' },
          type: {
            type: 'string',
            enum: ['Scope', 'Role'],
            description: 'Scope = delegated permission, Role = application permission.',
          },
        },
      },
    },
  },
};

const optionalClaim = {
  type: 'object',
  required: ['name'],
  properties: {
    name: { type: 'string', description: 'Claim name, e.g. "email", "groups", "upn".' },
    source: nullableString('null for built-in claims, "user" for directory extension attributes.'),
    essential: { type: 'boolean' },
    additionalProperties: {
      type: 'array',
      items: {
        type: 'string',
        enum: [
          'include_externally_authenticated_upn',
          'include_externally_authenticated_upn_without_hash',
          'sam_account_name',
          'dns_domain_and_sam_account_name',
          'netbios_domain_and_sam_account_name',
          'emit_as_roles',
          'cloud_displayname',
          'use_guid',
        ],
      },
    },
  },
};

export const APPLICATION_SCHEMA = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  title: 'Microsoft Graph application',
  type: 'object',
  properties: {
    '@odata.context': { type: 'string', readOnly: true },
    id: { ...guid, readOnly: true, description: 'Object id of the app registration. Read-only.' },
    appId: {
      ...guid,
      readOnly: true,
      description: 'Application (client) id. Read-only. In v2 access tokens this is the aud claim.',
    },
    deletedDateTime: { type: ['string', 'null'], readOnly: true },
    createdDateTime: { type: ['string', 'null'], readOnly: true },
    publisherDomain: { type: ['string', 'null'], readOnly: true },
    disabledByMicrosoftStatus: {
      type: ['string', 'null'],
      readOnly: true,
      description: 'Set when Microsoft has disabled the app. Read-only.',
    },
    applicationTemplateId: { type: ['string', 'null'], readOnly: true },
    serviceManagementReference: nullableString('Reference to an ITSM / CMDB record.'),
    verifiedPublisher: {
      type: 'object',
      readOnly: true,
      description: 'Set through the verified publisher process, not by PATCH.',
    },
    certification: { type: ['object', 'null'], readOnly: true },
    displayName: { type: 'string', description: 'Display name of the app.' },
    description: nullableString('Free-text description (max 1024 chars).'),
    notes: nullableString('Internal notes about the app.'),
    uniqueName: nullableString(
      'Alternate key. Can be set once while null and is immutable afterwards.',
    ),
    groupMembershipClaims: {
      type: ['string', 'null'],
      enum: [
        null,
        'None',
        'SecurityGroup',
        'All',
        'DirectoryRole',
        'ApplicationGroup',
        'SecurityGroup, DirectoryRole',
        'SecurityGroup, ApplicationGroup',
      ],
      description: 'Which groups are emitted in the groups claim of user tokens.',
    },
    identifierUris: uriList(
      'App ID URIs, e.g. "api://<appId>". Clients request scopes as "<App ID URI>/<scope>"; v1 access tokens usually carry it as aud.',
    ),
    isDeviceOnlyAuthSupported: { type: ['boolean', 'null'] },
    isFallbackPublicClient: {
      type: ['boolean', 'null'],
      description:
        'Allow public client flows: device code, ROPC and Windows integrated auth without a client secret.',
    },
    nativeAuthenticationApisEnabled: {
      type: ['string', 'null'],
      enum: [null, 'none', 'all'],
      description: 'Native authentication APIs (External ID).',
    },
    oauth2RequirePostResponse: { type: 'boolean' },
    defaultRedirectUri: nullableString('Default redirect URI.'),
    signInAudience: {
      type: 'string',
      enum: [
        'AzureADMyOrg',
        'AzureADMultipleOrgs',
        'AzureADandPersonalMicrosoftAccount',
        'PersonalMicrosoftAccount',
      ],
      description:
        'Who can sign in. AzureADandPersonalMicrosoftAccount and PersonalMicrosoftAccount require api.requestedAccessTokenVersion = 2.',
    },
    tags: {
      type: 'array',
      items: { type: 'string' },
      uniqueItems: true,
      description: 'Free-form labels. EntraApp manages tags prefixed "debble:".',
    },
    tokenEncryptionKeyId: {
      type: ['string', 'null'],
      description: 'keyId of a keyCredential used to encrypt tokens.',
    },
    samlMetadataUrl: nullableString('URL of SAML metadata.'),
    addIns: { type: 'array', items: { type: 'object' } },
    api: {
      type: 'object',
      description: 'Settings for when this app is an API that other apps call.',
      properties: {
        acceptMappedClaims: { type: ['boolean', 'null'] },
        knownClientApplications: {
          type: 'array',
          items: guid,
          description: 'Client appIds bundled for combined consent.',
        },
        oauth2PermissionScopes: { type: 'array', items: permissionScope },
        preAuthorizedApplications: {
          type: 'array',
          items: {
            type: 'object',
            required: ['appId', 'delegatedPermissionIds'],
            properties: {
              appId: { ...guid, description: 'Client app id that needs no consent.' },
              delegatedPermissionIds: { type: 'array', items: guid },
            },
          },
        },
        requestedAccessTokenVersion: {
          type: ['integer', 'null'],
          enum: [null, 1, 2],
          description:
            'Access token format issued FOR this API. null or 1 = v1 (iss sts.windows.net, aud usually the App ID URI). 2 = v2 (iss login.microsoftonline.com/…/v2.0, aud = appId).',
        },
      },
    },
    appRoles: { type: 'array', items: appRole },
    info: {
      type: 'object',
      properties: {
        logoUrl: { type: ['string', 'null'], readOnly: true },
        marketingUrl: nullableString('Marketing page.'),
        privacyStatementUrl: nullableString('Privacy statement, shown on the consent screen.'),
        supportUrl: nullableString('Support page.'),
        termsOfServiceUrl: nullableString('Terms of service, shown on the consent screen.'),
      },
    },
    keyCredentials: {
      type: 'array',
      items: keyCredential,
      description: 'Certificates. Use the Credentials tab to add one safely.',
    },
    passwordCredentials: {
      type: 'array',
      items: passwordCredential,
      readOnly: true,
      description: 'Client secrets. Cannot be changed by PATCH; use the Credentials tab.',
    },
    optionalClaims: {
      type: ['object', 'null'],
      properties: {
        idToken: { type: 'array', items: optionalClaim },
        accessToken: { type: 'array', items: optionalClaim },
        saml2Token: { type: 'array', items: optionalClaim },
      },
    },
    parentalControlSettings: {
      type: 'object',
      properties: {
        countriesBlockedForMinors: { type: 'array', items: { type: 'string' } },
        legalAgeGroupRule: {
          type: 'string',
          enum: [
            'Allow',
            'RequireConsentForPrivacyServices',
            'RequireConsentForMinors',
            'RequireConsentForKids',
            'BlockMinors',
          ],
        },
      },
    },
    publicClient: {
      type: 'object',
      description: 'Mobile & desktop (public client) platform.',
      properties: {
        redirectUris: uriList(
          'Native redirect URIs, e.g. http://localhost, msal<appId>://auth or https://login.microsoftonline.com/common/oauth2/nativeclient.',
        ),
      },
    },
    requestSignatureVerification: { type: ['object', 'null'] },
    requiredResourceAccess: { type: 'array', items: resourceAccess },
    servicePrincipalLockConfiguration: { type: ['object', 'null'] },
    spa: {
      type: 'object',
      description: 'Single-page application platform: auth code + PKCE from the browser.',
      properties: {
        redirectUris: uriList('SPA redirect URIs. The token endpoint allows CORS only for these.'),
      },
    },
    web: {
      type: 'object',
      description: 'Web platform: confidential, server-side clients.',
      properties: {
        homePageUrl: nullableString('Home page of the app.'),
        logoutUrl: nullableString('Front-channel logout URL.'),
        redirectUris: uriList(
          'Web redirect URIs (server-side, code redeemed with a secret or certificate).',
        ),
        redirectUriSettings: { type: 'array', items: { type: 'object' } },
        implicitGrantSettings: {
          type: 'object',
          description: 'Legacy implicit grant. Prefer auth code (+ PKCE).',
          properties: {
            enableIdTokenIssuance: {
              type: 'boolean',
              description:
                'Allow response_type=id_token (also needed for the "code id_token" hybrid flow).',
            },
            enableAccessTokenIssuance: {
              type: 'boolean',
              description: 'Allow response_type=token: access token in the URL fragment. Legacy.',
            },
          },
        },
      },
    },
  },
} as const;

// Properties never sent on save: read-only, or managed through dedicated
// endpoints (secrets via addPassword/removePassword).
export const READ_ONLY_PROPERTIES = new Set(
  Object.entries(APPLICATION_SCHEMA.properties)
    .filter(([, v]) => (v as { readOnly?: boolean }).readOnly)
    .map(([k]) => k),
);
