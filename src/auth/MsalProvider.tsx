import {
  useCallback,
  useEffect,
  useMemo,
  useState,
  type ReactNode,
} from 'react';
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
  saveAuthConfig,
  type AuthConfig,
} from './config';
import { clearTenantOverride, resolveTenantOverride } from './tenantParam';
import { AuthConfigContext } from './context';
import { SetupScreen } from '../components/SetupScreen';

interface Props {
  children: ReactNode;
}

// Tenant the tab's MSAL cache was populated for. When the effective tenant
// changes — e.g. a ?tenant=… shortcut opened in a tab that already signed in
// against another tenant — the cached accounts belong to the wrong authority,
// so we drop them and let the user sign in against the new tenant.
const CACHE_TENANT_KEY = 'entraapp.authTenant';

export function AppMsalProvider({ children }: Props) {
  const [savedConfig, setSavedConfig] = useState<AuthConfig>(() =>
    loadAuthConfig(),
  );
  const [tenantOverride, setTenantOverride] = useState<string | null>(() =>
    resolveTenantOverride(),
  );
  const [ready, setReady] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const config = useMemo<AuthConfig>(
    () =>
      tenantOverride
        ? { ...savedConfig, tenantId: tenantOverride }
        : savedConfig,
    [savedConfig, tenantOverride],
  );

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
      .then(async (result) => {
        if (cancelled) return;
        // Without a redirect response, whatever sits in the cache predates this
        // page load — discard it if it was acquired for a different tenant.
        if (!result) {
          const previous = sessionStorage.getItem(CACHE_TENANT_KEY);
          if (previous && previous !== config.tenantId) {
            await instance.clearCache();
          }
        }
        sessionStorage.setItem(CACHE_TENANT_KEY, config.tenantId);
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
  }, [instance, config.tenantId]);

  // Saving configuration explicitly beats the URL pin: drop the override so the
  // chosen tenant actually takes effect.
  const setConfig = useCallback((next: AuthConfig) => {
    saveAuthConfig(next);
    clearTenantOverride();
    setTenantOverride(null);
    setSavedConfig(next);
  }, []);

  const contextValue = useMemo(
    () => ({ config, savedConfig, tenantOverride, setConfig }),
    [config, savedConfig, tenantOverride, setConfig],
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
