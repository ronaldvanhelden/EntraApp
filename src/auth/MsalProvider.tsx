import { useEffect, useMemo, useState, type ReactNode } from 'react';
import {
  PublicClientApplication,
  EventType,
  type AccountInfo,
} from '@azure/msal-browser';
import { MsalProvider as InnerProvider } from '@azure/msal-react';
import {
  buildMsalConfig,
  isConfigured,
  loadAuthConfig,
  type AuthConfig,
} from './config';
import { AuthConfigContext } from './context';
import { SetupScreen } from '../components/SetupScreen';

interface Props {
  children: ReactNode;
}

export function AppMsalProvider({ children }: Props) {
  const [config, setConfig] = useState<AuthConfig>(() => loadAuthConfig());
  const [ready, setReady] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const instance = useMemo(() => {
    if (!isConfigured(config)) return null;
    return new PublicClientApplication(buildMsalConfig(config));
  }, [config]);

  useEffect(() => {
    if (!instance) {
      setReady(false);
      return;
    }
    let cancelled = false;
    setReady(false);
    setError(null);

    instance
      .initialize()
      .then(() => instance.handleRedirectPromise())
      .then((result) => {
        if (cancelled) return;
        const account: AccountInfo | null =
          result?.account ?? instance.getAllAccounts()[0] ?? null;
        if (account) instance.setActiveAccount(account);
        setReady(true);
      })
      .catch((e) => {
        if (!cancelled) setError(e?.message ?? String(e));
      });

    const cbId = instance.addEventCallback((event) => {
      if (
        event.eventType === EventType.LOGIN_SUCCESS &&
        event.payload &&
        'account' in event.payload
      ) {
        const acc = (event.payload as { account?: AccountInfo }).account;
        if (acc) instance.setActiveAccount(acc);
      }
    });

    return () => {
      cancelled = true;
      if (cbId) instance.removeEventCallback(cbId);
    };
  }, [instance]);

  const contextValue = useMemo(
    () => ({ config, setConfig }),
    [config],
  );

  if (!instance || !isConfigured(config)) {
    return (
      <AuthConfigContext.Provider value={contextValue}>
        <SetupScreen />
      </AuthConfigContext.Provider>
    );
  }

  if (error) {
    return (
      <div className="center">
        <h2>Authentication error</h2>
        <p className="error mono">{error}</p>
      </div>
    );
  }

  if (!ready) {
    return (
      <div className="center">
        <span className="spinner" />
        <p className="muted">Initializing authentication…</p>
      </div>
    );
  }

  return (
    <AuthConfigContext.Provider value={contextValue}>
      <InnerProvider instance={instance}>{children}</InnerProvider>
    </AuthConfigContext.Provider>
  );
}
