import { useMsal } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { useCallback, useMemo } from 'react';

// Returns a stable callback that acquires a Graph access token for the given
// scope set. Falls back to an interactive popup when MSAL signals that user
// (or admin) consent is required for any of the requested scopes — this is
// what makes incremental consent work: the prompt only appears when a feature
// asks for a scope the user hasn't consented to yet.
//
// `scopes` is treated as a stable identity if its contents don't change
// between renders, so consumers can pass module-level constants from `SCOPES`
// without re-triggering effects.
export function useGraphToken(
  scopes: readonly string[],
): () => Promise<string> {
  const { instance, accounts } = useMsal();
  const key = useMemo(() => [...scopes].sort().join(' '), [scopes]);

  return useCallback(
    async (): Promise<string> => {
      const account = instance.getActiveAccount() ?? accounts[0];
      if (!account) throw new Error('Not signed in');
      const requestScopes = [...scopes];
      try {
        const result = await instance.acquireTokenSilent({
          account,
          scopes: requestScopes,
        });
        return result.accessToken;
      } catch (err) {
        if (err instanceof InteractionRequiredAuthError) {
          const result = await instance.acquireTokenPopup({
            account,
            scopes: requestScopes,
          });
          return result.accessToken;
        }
        throw err;
      }
    },
    // `scopes` is captured via the stable `key`; eslint can't see through the
    // indirection.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [instance, accounts, key],
  );
}
