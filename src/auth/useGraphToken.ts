import { useMsal } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { useCallback } from 'react';
import { GRAPH_SCOPES } from './config';

export function useGraphToken() {
  const { instance, accounts } = useMsal();

  return useCallback(async (): Promise<string> => {
    const account = instance.getActiveAccount() ?? accounts[0];
    if (!account) throw new Error('Not signed in');

    try {
      const result = await instance.acquireTokenSilent({
        account,
        scopes: GRAPH_SCOPES,
      });
      return result.accessToken;
    } catch (err) {
      if (err instanceof InteractionRequiredAuthError) {
        const result = await instance.acquireTokenPopup({
          account,
          scopes: GRAPH_SCOPES,
        });
        return result.accessToken;
      }
      throw err;
    }
  }, [instance, accounts]);
}
