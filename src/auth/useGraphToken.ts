import { useMsal } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { useCallback } from 'react';
import { GRAPH_SCOPES } from './config';

// Pass `scopes` to ask for more than the sign-in set — e.g. the write scopes
// in config.ts, which are only consented to when a write is first attempted.
// With `interactive` off the token is only taken silently: if the scopes still
// need consent this throws instead of opening a popup — for calls made on page
// load, where a popup would be unexpected (and blocked by the browser).
export function useGraphToken(scopes: string[] = GRAPH_SCOPES, interactive = true) {
  const { instance, accounts } = useMsal();

  return useCallback(async (): Promise<string> => {
    const account = instance.getActiveAccount() ?? accounts[0];
    if (!account) throw new Error('Not signed in');

    try {
      const result = await instance.acquireTokenSilent({
        account,
        scopes,
      });
      return result.accessToken;
    } catch (err) {
      if (interactive && err instanceof InteractionRequiredAuthError) {
        const result = await instance.acquireTokenPopup({
          account,
          scopes,
        });
        return result.accessToken;
      }
      throw err;
    }
  }, [instance, accounts, scopes, interactive]);
}
