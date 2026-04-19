import { createContext, useContext } from 'react';
import type { AuthConfig } from './config';

interface AuthConfigContextValue {
  config: AuthConfig;
  setConfig: (config: AuthConfig) => void;
}

export const AuthConfigContext = createContext<AuthConfigContextValue | null>(
  null,
);

export function useAuthConfig(): AuthConfigContextValue {
  const value = useContext(AuthConfigContext);
  if (!value) throw new Error('AuthConfigContext missing');
  return value;
}
