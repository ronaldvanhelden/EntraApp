import { useCallback, useEffect, useMemo, useState, type ReactNode } from 'react';
import {
  PublicClientApplication,
  EventType,
  type AccountInfo,
  type IPublicClientApplication,
} from '@azure/msal-browser';
import { MsalProvider as InnerProvider } from '@azure/msal-react';
import {
  buildMsalConfig,
  LOGIN_SCOPES,
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

// Switching tenants (a ?tenant=… shortcut, the sidebar's tenant list) keeps
// the tab's MSAL cache, so every account that signed in here stays available.
// What changes per tenant is which account is the active one: we remember the
// account last used for each tenant and, on arriving at a tenant, quietly get
// a token for it there. That works without a prompt when the same account was
// used for that tenant before, or is a guest in it; only when nothing in the
// cache gets in does the sign-in screen appear — the other tenants' accounts
// stay signed in meanwhile.
const TENANT_ACCOUNTS_KEY = 'entraapp.tenantAccounts';
const GUID_RE = /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i;

function tenantAccounts(): Record<string, string> {
  try {
    const parsed = JSON.parse(sessionStorage.getItem(TENANT_ACCOUNTS_KEY) ?? '{}') as unknown;
    return parsed && typeof parsed === 'object' ? (parsed as Record<string, string>) : {};
  } catch {
    return {};
  }
}

function rememberTenantAccount(tenant: string, account: AccountInfo): void {
  try {
    sessionStorage.setItem(
      TENANT_ACCOUNTS_KEY,
      JSON.stringify({ ...tenantAccounts(), [tenant.toLowerCase()]: account.homeAccountId }),
    );
  } catch {
    /* private mode: the next switch just tries every cached account */
  }
}

// The cached account that can work in this tenant, as seen from this tenant
// (its tenantId is this tenant's, also for a guest), or null when a sign-in
// is needed. Tries the account remembered for the tenant first, then one
// that lives in it, then the others — the first that gets a token wins.
async function restoreAccount(
  instance: IPublicClientApplication,
  tenant: string,
): Promise<AccountInfo | null> {
  const key = tenant.toLowerCase();
  const remembered = tenantAccounts()[key];
  const all = instance.getAllAccounts();
  const ordered = [
    ...all.filter((a) => a.homeAccountId === remembered),
    ...all.filter((a) => a.tenantId?.toLowerCase() === key),
    ...all,
  ].filter((a, i, list) => list.findIndex((b) => b.homeAccountId === a.homeAccountId) === i);
  const wanted = GUID_RE.test(tenant) ? key : null;
  for (const account of ordered.slice(0, 4)) {
    try {
      // forceRefresh: MSAL's cache lookup goes by the account's own tenant, so
      // without it an account from another tenant would hand back its cached
      // token for *that* tenant — and the app would show the wrong directory.
      // Redeeming the refresh token at this tenant's authority gives a token
      // (and an account profile) that really belongs here.
      const result = await instance.acquireTokenSilent({
        account,
        scopes: LOGIN_SCOPES,
        forceRefresh: true,
      });
      if (wanted && result.tenantId?.toLowerCase() !== wanted) continue;
      return result.account ?? account;
    } catch {
      // Not usable here without interaction: try the next one.
    }
  }
  return null;
}

export function AppMsalProvider({ children }: Props) {
  const [savedConfig, setSavedConfig] = useState<AuthConfig>(() => loadAuthConfig());
  const [tenantOverride, setTenantOverride] = useState<string | null>(() =>
    resolveTenantOverride(),
  );
  const [ready, setReady] = useState(false);
  const [error, setError] = useState<string | null>(null);
  // Whether an account is signed in for the current tenant.
  const [signedIn, setSignedIn] = useState(false);

  const config = useMemo<AuthConfig>(
    () => (tenantOverride ? { ...savedConfig, tenantId: tenantOverride } : savedConfig),
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
        // A redirect response is a fresh sign-in for this tenant; otherwise
        // see which cached account gets in here without asking.
        const account: AccountInfo | null =
          result?.account ?? (await restoreAccount(instance, config.tenantId));
        if (cancelled) return;
        instance.setActiveAccount(account);
        if (account) rememberTenantAccount(config.tenantId, account);
        setSignedIn(Boolean(account));
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
        if (acc) {
          instance.setActiveAccount(acc);
          rememberTenantAccount(config.tenantId, acc);
          setSignedIn(true);
        }
      } else if (event.eventType === EventType.LOGOUT_SUCCESS) {
        setSignedIn(false);
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
    () => ({ config, savedConfig, tenantOverride, signedIn, setConfig }),
    [config, savedConfig, tenantOverride, signedIn, setConfig],
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
