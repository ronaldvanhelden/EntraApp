import type { ReactNode } from 'react';
import { useEffect, useState } from 'react';
import { NavLink, useLocation } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import { LOGIN_SCOPES, READ_SCOPES } from '../auth/config';
import { useGraphToken } from '../auth/useGraphToken';
import { useAuthConfig } from '../auth/context';
import { useCurrentTenantId } from '../auth/useCurrentTenantId';
import {
  buildTenantUrl,
  exitTenantOverride,
  loadTenantShortcuts,
  sectionHash,
} from '../auth/tenantParam';
import { TenantChip } from './TenantChip';
import {
  AppRegistrationIcon,
  EnterpriseAppIcon,
  GroupIcon,
  OverviewIcon,
  SettingsIcon,
  ShieldIcon,
  TenantIcon,
  UserIcon,
} from './ButtonIcons';
import { scrollToTop } from '../lib/scrollMemory';

interface Props {
  children: ReactNode;
}

// Phones, plus portrait screens up to 920px (e.g. a Galaxy Z Fold inner screen
// held upright). Keep in sync with the mobile @media block in index.css.
const MOBILE_QUERY = '(max-width: 640px), (max-width: 920px) and (orientation: portrait)';
const COLLAPSE_KEY = 'entraapp:nav-collapsed';

function useIsMobile(): boolean {
  const [isMobile, setIsMobile] = useState(() =>
    typeof window !== 'undefined' ? window.matchMedia(MOBILE_QUERY).matches : false,
  );
  useEffect(() => {
    const mql = window.matchMedia(MOBILE_QUERY);
    const handler = (e: MediaQueryListEvent) => setIsMobile(e.matches);
    mql.addEventListener('change', handler);
    return () => mql.removeEventListener('change', handler);
  }, []);
  return isMobile;
}

export function Layout({ children }: Props) {
  const isMobile = useIsMobile();
  // Mobile drawer open/closed. Default: closed (mobile starts collapsed).
  const [navOpen, setNavOpen] = useState(false);
  // Desktop collapsed state, persisted across reloads.
  const [desktopCollapsed, setDesktopCollapsed] = useState<boolean>(() => {
    if (typeof window === 'undefined') return false;
    return window.localStorage.getItem(COLLAPSE_KEY) === '1';
  });
  const location = useLocation();

  // Close the mobile drawer whenever navigation happens.
  useEffect(() => {
    setNavOpen(false);
  }, [location.pathname]);

  // Prevent body scroll while the drawer is open on mobile.
  useEffect(() => {
    if (!navOpen) return;
    const previous = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      document.body.style.overflow = previous;
    };
  }, [navOpen]);

  useEffect(() => {
    window.localStorage.setItem(COLLAPSE_KEY, desktopCollapsed ? '1' : '0');
  }, [desktopCollapsed]);

  const toggleMenu = () => {
    if (isMobile) setNavOpen((v) => !v);
    else setDesktopCollapsed((v) => !v);
  };

  const showingMenu = isMobile ? navOpen : !desktopCollapsed;
  const layoutClasses = [
    'layout',
    navOpen ? 'nav-open' : '',
    !isMobile && desktopCollapsed ? 'nav-collapsed' : '',
  ]
    .filter(Boolean)
    .join(' ');

  // Signed in for *this* tenant: after a tenant switch the cache can still
  // hold accounts of other tenants that don't get in here.
  const { signedIn } = useAuthConfig();
  if (!signedIn) return <SignInScreen />;

  return (
    <ReadAccessGate>
      <div className={layoutClasses}>
        <Header onMenuClick={toggleMenu} navOpen={showingMenu} />
        <Sidebar />
        {navOpen && (
          <div className="nav-backdrop" onClick={() => setNavOpen(false)} aria-hidden="true" />
        )}
        <main className="main">
          <ScrollReset />
          {children}
        </main>
      </div>
    </ReadAccessGate>
  );
}

function Header({ onMenuClick, navOpen }: { onMenuClick: () => void; navOpen: boolean }) {
  const { instance, accounts } = useMsal();
  const active = instance.getActiveAccount() ?? accounts[0];
  const signOut = () => {
    // Come back to this exact URL so a ?tenant= pin (and any base path)
    // survives the round trip. Popup logout isn't supported on mobile
    // browsers; fall back to redirect, which uses postLogoutRedirectUri.
    instance
      .logoutPopup({ mainWindowRedirectUri: window.location.href })
      .catch(() => instance.logoutRedirect().catch(() => {}));
  };
  return (
    <header className="header">
      <div className="header-left">
        <button
          type="button"
          className="ghost menu-toggle"
          aria-label={navOpen ? 'Close navigation' : 'Open navigation'}
          aria-expanded={navOpen}
          onClick={onMenuClick}
        >
          <span className="menu-icon" aria-hidden="true">
            <span />
            <span />
            <span />
          </span>
        </button>
        <div className="brand">EntraApp</div>
        <TenantChip />
      </div>
      <div className="user">
        <span className="user-email" title={active?.username}>
          {active?.username}
        </span>
        <button className="ghost" onClick={signOut}>
          Sign out
        </button>
      </div>
    </header>
  );
}

function Sidebar() {
  const nav = ({ isActive }: { isActive: boolean }) => (isActive ? 'active' : undefined);
  return (
    <nav className="sidebar">
      <h4>Directory</h4>
      <NavLink to="/" end className={nav}>
        <OverviewIcon />
        Overview
      </NavLink>
      <NavLink to="/applications" className={nav}>
        <AppRegistrationIcon />
        App registrations
      </NavLink>
      <NavLink to="/enterprise-apps" className={nav}>
        <EnterpriseAppIcon />
        Enterprise apps
      </NavLink>
      <NavLink to="/users" className={nav}>
        <UserIcon />
        Users
      </NavLink>
      <NavLink to="/groups" className={nav}>
        <GroupIcon />
        Groups
      </NavLink>
      <NavLink to="/pim" className={nav}>
        <ShieldIcon />
        Privileged roles
      </NavLink>
      <h4>Tools</h4>
      <NavLink to="/settings" className={nav}>
        <SettingsIcon />
        Settings
      </NavLink>
      <TenantLinks />
    </nav>
  );
}

// The tenant shortcuts from Settings, as a switcher. A click moves this tab
// to that tenant (a reload on its ?tenant= URL) and stays in the same section
// of the app; accounts that already signed in are reused, so going back and
// forth doesn't ask for a sign-in each time.
function TenantLinks() {
  const { tenantOverride } = useAuthConfig();
  const tenantId = useCurrentTenantId();
  const { pathname } = useLocation();
  // Read on every render: the list is edited on the Settings page.
  const shortcuts = loadTenantShortcuts();
  if (shortcuts.length === 0) return null;
  const section = sectionHash(pathname);
  const pinned = tenantOverride?.toLowerCase();
  return (
    <>
      <h4>Tenants</h4>
      <a
        href="#/"
        className={pinned ? undefined : 'active'}
        title="The tenant from Settings, without a pin"
        onClick={(e) => {
          e.preventDefault();
          if (pinned) exitTenantOverride(section);
        }}
      >
        <TenantIcon />
        <span className="sidebar-label">Default tenant</span>
      </a>
      {shortcuts.map((s) => {
        const id = s.tenantId.toLowerCase();
        const active = Boolean(pinned) && (id === pinned || id === tenantId?.toLowerCase());
        return (
          <a
            key={s.tenantId}
            href={buildTenantUrl(s.tenantId, section)}
            className={active ? 'active' : undefined}
            title={`Switch this tab to ${s.tenantId}`}
            onClick={(e) => active && e.preventDefault()}
          >
            <TenantIcon />
            <span className="sidebar-label">{s.label}</span>
          </a>
        );
      })}
    </>
  );
}

// Sign-in only asks for the user's own profile. Reading the directory takes
// permissions an admin has to consent to; they are checked for silently, and
// when they are missing this screen asks for them from a click — a consent
// popup opened on page load would be blocked by the browser.
function ReadAccessGate({ children }: { children: ReactNode }) {
  const { instance, accounts } = useMsal();
  const silent = useGraphToken(READ_SCOPES, false);
  const interactive = useGraphToken(READ_SCOPES);
  const [state, setState] = useState<'checking' | 'ok' | 'needed'>('checking');
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    silent()
      .then(() => !cancelled && setState('ok'))
      .catch(() => !cancelled && setState((s) => (s === 'ok' ? s : 'needed')));
    return () => {
      cancelled = true;
    };
  }, [silent]);

  if (state === 'ok') return <>{children}</>;
  if (state === 'checking')
    return (
      <div className="center">
        <span className="spinner" />
      </div>
    );

  const active = instance.getActiveAccount() ?? accounts[0];
  const grant = () => {
    setError(null);
    interactive()
      .then(() => setState('ok'))
      .catch((e: unknown) => setError(e instanceof Error ? e.message : String(e)));
  };
  return (
    <div className="center">
      <h1>EntraApp</h1>
      <p className="muted" style={{ maxWidth: 520 }}>
        Signed in as <strong>{active?.username}</strong>. To show this directory, EntraApp needs
        permission to read its apps, users, groups, permission grants and audit logs. Nothing is
        changed with it — permission to make changes is asked for separately, the first time you
        make one.
      </p>
      <p className="mono muted" style={{ fontSize: 12, maxWidth: 520 }}>
        {READ_SCOPES.join(' · ')}
      </p>
      <div className="row">
        <button className="primary" onClick={grant}>
          Grant read access
        </button>
        <button onClick={() => instance.logoutRedirect().catch(() => {})}>Sign out</button>
      </div>
      {error && (
        <p className="error" style={{ maxWidth: 520 }}>
          {error}
        </p>
      )}
      <p className="muted" style={{ fontSize: 12, maxWidth: 520 }}>
        These permissions need an administrator's consent. If you are not one, ask an admin to grant
        them for the organisation, then reload this page.
      </p>
    </div>
  );
}

function SignInScreen() {
  const { instance } = useMsal();
  const { config, tenantOverride } = useAuthConfig();
  const signIn = () => {
    instance
      .loginPopup({ scopes: LOGIN_SCOPES, prompt: 'select_account' })
      .catch(() =>
        instance.loginRedirect({ scopes: LOGIN_SCOPES, prompt: 'select_account' }).catch(() => {}),
      );
  };
  return (
    <div className="center">
      <h1>EntraApp</h1>
      <p className="muted">Sign in with your Entra ID account to continue.</p>
      {tenantOverride ? (
        <p className="muted" style={{ fontSize: 13 }}>
          Signing in to tenant <strong className="mono">{tenantOverride}</strong> (pinned by this
          URL). Pick the account that has access there — a guest account works.
        </p>
      ) : (
        <p className="muted" style={{ fontSize: 13 }}>
          Tenant: <span className="mono">{config.tenantId}</span>
        </p>
      )}
      <button className="primary" onClick={signIn}>
        Sign in with Microsoft
      </button>
    </div>
  );
}

// Open every page at the top. Without this the shared <main> scroller keeps the
// previous page's offset; list pages restore their own position afterwards.
function ScrollReset() {
  const { pathname } = useLocation();
  useEffect(() => {
    scrollToTop();
  }, [pathname]);
  return null;
}
