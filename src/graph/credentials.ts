import { graphAll } from './client';
import type {
  AppCredentialSignInActivity,
  FederatedIdentityCredential,
} from './types';

type TokenFn = () => Promise<string>;

export function listFederatedIdentityCredentials(
  token: TokenFn,
  applicationObjectId: string,
) {
  return graphAll<FederatedIdentityCredential>(
    token,
    `/applications/${applicationObjectId}/federatedIdentityCredentials`,
  );
}

// appCredentialSignInActivities is only available on the beta endpoint.
// Requires AuditLog.Read.All.
export function listAppCredentialSignInActivities(
  token: TokenFn,
  appId: string,
) {
  return graphAll<AppCredentialSignInActivity>(
    token,
    '/reports/appCredentialSignInActivities',
    {
      api: 'beta',
      query: { $filter: `appId eq '${appId.replace(/'/g, "''")}'` },
    },
  );
}

export function buildCredentialActivityMap(
  activities: AppCredentialSignInActivity[],
): Map<string, AppCredentialSignInActivity> {
  const map = new Map<string, AppCredentialSignInActivity>();
  for (const a of activities) {
    if (!a.keyId) continue;
    const existing = map.get(a.keyId);
    if (!existing) {
      map.set(a.keyId, a);
      continue;
    }
    const existingTime = existing.signInActivity?.lastSignInDateTime ?? '';
    const candidateTime = a.signInActivity?.lastSignInDateTime ?? '';
    if (candidateTime > existingTime) map.set(a.keyId, a);
  }
  return map;
}
