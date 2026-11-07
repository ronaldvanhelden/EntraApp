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

export type CredentialKind = 'secret' | 'certificate' | 'federated';

// Does the recorded credential *type* fit this kind of credential? Entra
// records the type for user sign-ins through confidential clients too, but
// the key id only for service-principal sign-ins — so for delegated flows
// this is often the only link between an event and a secret/certificate.
export function credentialTypeMatches(s: SignIn, kind: CredentialKind): boolean {
  const t = s.clientCredentialType;
  if (!t) return false;
  switch (kind) {
    case 'secret':
      return t === 'clientSecret';
    case 'certificate':
      return t === 'clientAssertion' || t === 'certificate';
    case 'federated':
      return t === 'federatedIdentityCredential';
  }
}

export interface CredentialRef {
  kind: CredentialKind;
  keyId?: string;
  thumbprint?: string;
  ficId?: string;
}

// Exact attribution: the event names this key (secrets, certificates) or
// this federated credential.
export function signInAttributedTo(s: SignIn, ref: CredentialRef): boolean {
  if (ref.kind === 'federated') {
    return !!ref.ficId && s.federatedCredentialId?.toLowerCase() === ref.ficId.toLowerCase();
  }
  return !!ref.keyId && signInUsesCredential(s, ref.keyId, ref.thumbprint);
}

export interface CredentialSignIns {
  // Events that name this exact credential (any flow).
  attributed: SignIn[];
  // Delegated events through this app that presented a credential of the
  // same *type* but where Entra didn't record which key — can't be pinned
  // to this credential if the app has several of that type.
  sameType: SignIn[];
  // Per-query failures, e.g. a 403 for AuditLog.Read.All.
  errors: { flow: SignInEventType; message: string; forbidden: boolean }[];
}

// Everything the sign-in log has on one credential in the last 30 days,
// across app-only and delegated flows. Queries run in parallel; a failing
// query is reported, not fatal.
export async function listSignInsUsingCredential(
  token: TokenFn,
  appId: string,
  ref: CredentialRef,
  top = 50,
  maxPages = 2,
): Promise<CredentialSignIns> {
  const flows: SignInEventType[] = ['servicePrincipal', 'interactiveUser', 'nonInteractiveUser'];
  const results = await Promise.all(
    flows.map(async (flow) => {
      try {
        const rows =
          flow === 'servicePrincipal' && ref.kind !== 'federated' && ref.keyId
            ? await listAppOnlySignInsForCredential(token, appId, ref.keyId, ref.thumbprint, top)
            : await listSignInsForApp(token, appId, flow, top, maxPages);
        return { flow, rows, error: undefined };
      } catch (e) {
        return {
          flow,
          rows: [] as SignIn[],
          error: {
            flow,
            message: e instanceof Error ? e.message : String(e),
            forbidden: e instanceof GraphError && e.status === 403,
          },
        };
      }
    }),
  );
  const attributed = new Map<string, SignIn>();
  const sameType = new Map<string, SignIn>();
  for (const { rows } of results) {
    for (const s of rows) {
      if (signInAttributedTo(s, ref)) attributed.set(s.id, s);
      else if (!isAppOnlySignIn(s) && credentialTypeMatches(s, ref.kind)) sameType.set(s.id, s);
    }
  }
  const byTime = (a: SignIn, b: SignIn) =>
    (b.createdDateTime ?? '').localeCompare(a.createdDateTime ?? '');
  return {
    attributed: [...attributed.values()].sort(byTime),
    sameType: [...sameType.values()].sort(byTime),
    errors: results.flatMap((r) => (r.error ? [r.error] : [])),
  };
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
