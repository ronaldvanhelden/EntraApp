import { graph, graphAll } from './client';
import type { SignIn } from './types';

type TokenFn = () => Promise<string>;

// /auditLogs/signIns rejects most $select combinations with "Unsupported
// Query." We therefore take the default projection — which on beta already
// includes the dedicated service-principal credential fields — and let any
// tenant-specific omissions surface via the detail modal's raw dump.

export function listUserSignInsForApp(token: TokenFn, appId: string, top = 50) {
  return graphAll<SignIn>(
    token,
    '/auditLogs/signIns',
    {
      api: 'beta',
      query: {
        $filter: `appId eq '${appId.replace(/'/g, "''")}'`,
        $top: top,
      },
    },
    2,
  );
}

// Service-principal / app-only sign-ins are a distinct event type and must be
// requested explicitly via the signInEventTypes filter.
export function listAppOnlySignInsForApp(
  token: TokenFn,
  appId: string,
  top = 50,
) {
  const safe = appId.replace(/'/g, "''");
  return graphAll<SignIn>(
    token,
    '/auditLogs/signIns',
    {
      api: 'beta',
      query: {
        $filter: `appId eq '${safe}' and signInEventTypes/any(t:t eq 'servicePrincipal')`,
        $top: top,
      },
    },
    2,
  );
}

// Fetch a single sign-in event by id (== signInActivity.lastSignInRequestId
// from the credential activity report). Requires AuditLog.Read.All.
export function getSignIn(token: TokenFn, id: string) {
  return graph<SignIn>(token, `/auditLogs/signIns/${id}`, {
    api: 'beta',
  });
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
  if (s.servicePrincipalCredentialThumbprint)
    return s.servicePrincipalCredentialThumbprint;
  const details = s.authenticationProcessingDetails ?? [];
  for (const d of details) {
    if (!d.key || !d.value) continue;
    const k = d.key.toLowerCase().replace(/[\s_-]+/g, '');
    if (k.includes('keyid') || k.includes('thumbprint')) return d.value;
  }
  return undefined;
}
