import {
  InteractionRequiredAuthError,
  type AccountInfo,
  type IPublicClientApplication,
} from '@azure/msal-browser';
import { legacyScopes } from './config';

// One place where Graph tokens are acquired, for the hook and for sign-in.
//
// Tenants that consented before reading and writing were split only hold the
// read-write permissions. Asking them for the new read-only ones would mean a
// fresh consent prompt — for non-admins a dead end — so a scope set that isn't
// consented is retried with its older, broader equivalent before anyone is
// asked anything. Which of the two worked is remembered for the page's
// lifetime (a tenant switch reloads the page), so the detour is taken once
// per scope set and not on every Graph call.
const resolved = new Map<string, string[]>();

export async function acquireGraphToken(
  instance: IPublicClientApplication,
  account: AccountInfo,
  scopes: string[],
  // Off: only silent attempts; a missing consent throws instead of opening a
  // popup. For calls made on page load, where a popup would be blocked.
  interactive: boolean,
): Promise<string> {
  const key = `${account.homeAccountId}|${account.tenantId}|${[...scopes].sort().join(' ')}`;
  const known = resolved.get(key);
  const legacy = legacyScopes(scopes);
  const variants = known ? [known] : legacy ? [scopes, legacy] : [scopes];
  let failure: unknown;
  for (const variant of variants) {
    try {
      const result = await instance.acquireTokenSilent({ account, scopes: variant });
      resolved.set(key, variant);
      return result.accessToken;
    } catch (err) {
      if (!(err instanceof InteractionRequiredAuthError)) throw err;
      failure = err;
    }
  }
  if (!interactive) throw failure;
  const result = await instance.acquireTokenPopup({ account, scopes: known ?? scopes });
  resolved.set(key, known ?? scopes);
  return result.accessToken;
}
