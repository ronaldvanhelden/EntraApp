import { createContext, useContext } from 'react';
import type { AuthConfig } from './config';

interface AuthConfigContextValue {
  /** Effective configuration: the saved config with any URL tenant applied. */
  config: AuthConfig;
  /** Saved configuration, ignoring the URL tenant. */
  savedConfig: AuthConfig;
  /** Tenant this tab is pinned to via ?tenant=…, or null. */
  tenantOverride: string | null;
  /** Persists a new configuration and drops any URL tenant pin. */
  setConfig: (config: AuthConfig) => void;
}

export const AuthConfigContext = createContext<AuthConfigContextValue | null>(null);

export function useAuthConfig(): AuthConfigContextValue {
  const value = useContext(AuthConfigContext);
  if (!value) throw new Error('AuthConfigContext missing');
  return value;
}
