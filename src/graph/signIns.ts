import { graph, graphAll, GraphError } from './client';
import type { SignIn } from './types';

type TokenFn = () => Promise<string>;

// /auditLogs/signIns rejects most $select combinations with "Unsupported
// Query." We therefore take the default projection — which on beta already
// includes the dedicated service-principal credential fields — and let any
// tenant-specific omissions surface via the detail modal's raw dump.

// Sign-in event classes in the Entra sign-in log. IMPORTANT: without an
// explicit signInEventTypes filter the API returns *only* interactiveUser
// events, so every other class has to be requested on its own.
export type SignInEventType =
  'interactiveUser' | 'nonInteractiveUser' | 'servicePrincipal' | 'managedIdentity';

export const SIGN_IN_EVENT_TYPES: SignInEventType[] = [
  'interactiveUser',
  'nonInteractiveUser',
  'servicePrincipal',
  'managedIdentity',
];

function q(s: string): string {
  return s.replace(/'/g, "''");
}

// Sign-ins where the app acted as the *client* (it obtained a token).
export function listSignInsForApp(
  token: TokenFn,
  appId: string,
  eventType: SignInEventType,
  top = 50,
  maxPages = 2,
) {
  return graphAll<SignIn>(
    token,
    '/auditLogs/signIns',
    {
      api: 'beta',
      query: {
        $filter: `appId eq '${q(appId)}' and signInEventTypes/any(t:t eq '${eventType}')`,
        $top: top,
      },
    },
    maxPages,
  );
}

// Sign-ins where the app acted as the *resource* — some user or other app
// obtained a token *for* this app's API. resourceId in the sign-in log is the
// resource's appId (e.g. Microsoft Graph shows up as 00000003-…-c000-…).
// A resource that still has callers is not safe to delete.
export function listSignInsForResource(
  token: TokenFn,
  resourceAppId: string,
  eventType: SignInEventType,
  top = 50,
  maxPages = 2,
) {
  return graphAll<SignIn>(
    token,
    '/auditLogs/signIns',
    {
      api: 'beta',
      query: {
        $filter: `resourceId eq '${q(resourceAppId)}' and signInEventTypes/any(t:t eq '${eventType}')`,
        $top: top,
      },
    },
    maxPages,
  );
}

// Interactive user sign-ins only. Non-interactive (token refresh, SSO) are a
// separate event class — see listSignInsForApp(…, 'nonInteractiveUser').
export function listUserSignInsForApp(token: TokenFn, appId: string, top = 50) {
  return listSignInsForApp(token, appId, 'interactiveUser', top);
}

// Service-principal / app-only sign-ins are a distinct event type and must be
// requested explicitly via the signInEventTypes filter.
export function listAppOnlySignInsForApp(token: TokenFn, appId: string, top = 50) {
  return listSignInsForApp(token, appId, 'servicePrincipal', top);
}

// App-only sign-ins for the app, narrowed to a single credential. Graph beta's
// servicePrincipalCredentialKeyId is the GUID that matches
// PasswordCredential.keyId and KeyCredential.keyId, so it covers both secrets
// and certificates. Filtering on that property server-side is *not*
// documented as supported, so if Graph rejects the query (400) we fall back
// to pulling the app's recent app-only sign-ins and matching client-side on
// keyId or certificate thumbprint.
export async function listAppOnlySignInsForCredential(
  token: TokenFn,
  appId: string,
  keyId: string,
  thumbprint?: string,
  top = 50,
): Promise<SignIn[]> {
  try {
    return await graphAll<SignIn>(
      token,
      '/auditLogs/signIns',
      {
        api: 'beta',
        query: {
          $filter: `appId eq '${q(appId)}' and signInEventTypes/any(t:t eq 'servicePrincipal') and servicePrincipalCredentialKeyId eq '${q(keyId)}'`,
          $top: top,
        },
      },
      2,
    );
  } catch (e) {
    if (!(e instanceof GraphError && e.status === 400)) throw e;
    const all = await listSignInsForApp(token, appId, 'servicePrincipal', top, 2);
    return all.filter((s) => signInUsesCredential(s, keyId, thumbprint));
  }
}

// Fetch a single sign-in event by id (== signInActivity.lastSignInRequestId
// from the credential activity report). Requires AuditLog.Read.All.
export function getSignIn(token: TokenFn, id: string) {
  return graph<SignIn>(token, `/auditLogs/signIns/${id}`, {
    api: 'beta',
  });
}

export function isAppOnlySignIn(s: SignIn): boolean {
  return (s.signInEventTypes ?? []).some(
    (t) => t === 'servicePrincipal' || t === 'managedIdentity',
  );
}

// Does this app-only sign-in reference the given secret/certificate? Matches
// on keyId (secrets and certs) or SHA-1 thumbprint (certs), case-insensitive.
export function signInUsesCredential(s: SignIn, keyId: string, thumbprint?: string): boolean {
  const k = keyId.toLowerCase();
  const t = thumbprint?.toLowerCase();
  const candidates = [
    s.servicePrincipalCredentialKeyId,
    s.servicePrincipalCredentialThumbprint,
    extractKeyIdFromSignIn(s),
  ]
    .filter((v): v is string => !!v)
    .map((v) => v.toLowerCase());
  return candidates.some((c) => c === k || (t !== undefined && c === t));
}

// Best-effort extraction of the credential used for an app-only sign-in.
// Microsoft Graph beta exposes two dedicated fields —
// `servicePrincipalCredentialKeyId` (GUID, matches KeyCredential.keyId and
// PasswordCredential.keyId) and `servicePrincipalCredentialThumbprint` (hex
// SHA-1 of the cert DER, matches customKeyIdentifier) — which are
// authoritative. Older events may only carry this info inside the free-form
// authenticationProcessingDetails blob, so we fall back to scanning that.
export function extractKeyIdFromSignIn(s: SignIn): string | undefined {
  if (s.servicePrincipalCredentialKeyId) return s.servicePrincipalCredentialKeyId;
  if (s.servicePrincipalCredentialThumbprint) return s.servicePrincipalCredentialThumbprint;
  const details = s.authenticationProcessingDetails ?? [];
  for (const d of details) {
    if (!d.key || !d.value) continue;
    const k = d.key.toLowerCase().replace(/[\s_-]+/g, '');
    if (k.includes('keyid') || k.includes('thumbprint')) return d.value;
  }
  return undefined;
}
