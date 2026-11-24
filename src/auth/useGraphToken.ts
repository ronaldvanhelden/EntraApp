import { useMsal } from '@azure/msal-react';
import { useCallback } from 'react';
import { READ_SCOPES } from './config';
import { acquireGraphToken } from './acquireToken';

// Returns the token getter the Graph client calls for every request.
//
// Without `scopes` it follows the request: the client passes the permission a
// write needs (see scopesForRequest in config.ts) and everything else gets
// the read-only set, so the write permissions are only consented to when a
// change is first made. Pass `scopes` to pin the getter to one set — the
// user/group write scopes and the optional read extras in config.ts.
//
// With `interactive` off the token is only taken silently: if the scopes still
// need consent this throws instead of opening a popup — for calls made on page
// load, where a popup would be unexpected (and blocked by the browser).
export function useGraphToken(scopes?: string[], interactive = true) {
  const { instance, accounts } = useMsal();

  return useCallback(
    async (needed?: string[]): Promise<string> => {
      const account = instance.getActiveAccount() ?? accounts[0];
      if (!account) throw new Error('Not signed in');
      return acquireGraphToken(instance, account, scopes ?? needed ?? READ_SCOPES, interactive);
    },
    [instance, accounts, scopes, interactive],
  );
}
