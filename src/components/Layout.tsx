import type { ReactNode } from 'react';
import { useEffect, useState } from 'react';
import { NavLink, useLocation } from 'react-router-dom';
import {
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
  useMsal,
} from '@azure/msal-react';
import { GRAPH_SCOPES } from '../auth/config';

interface Props {
  children: ReactNode;
}

export function Layout({ children }: Props) {
  const [navOpen, setNavOpen] = useState(false);
  const location = useLocation();

  // Close the drawer whenever navigation happens.
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

  return (
    <>
      <UnauthenticatedTemplate>
        <SignInScreen />
      </UnauthenticatedTemplate>
      <AuthenticatedTemplate>
        <div className={`layout${navOpen ? ' nav-open' : ''}`}>
          <Header onMenuClick={() => setNavOpen((v) => !v)} navOpen={navOpen} />
          <Sidebar />
          {navOpen && (
            <div
              className="nav-backdrop"
              onClick={() => setNavOpen(false)}
              aria-hidden="true"
            />
          )}
          <main className="main">{children}</main>
        </div>
      </AuthenticatedTemplate>
    </>
  );
}

function Header({
  onMenuClick,
  navOpen,
}: {
  onMenuClick: () => void;
  navOpen: boolean;
}) {
  const { instance, accounts } = useMsal();
  const active = instance.getActiveAccount() ?? accounts[0];
  const signOut = () => {
    // Popup logout isn't supported on mobile browsers; fall back to redirect.
    instance
      .logoutPopup({ mainWindowRedirectUri: '/' })
      .catch(() =>
        instance.logoutRedirect({ postLogoutRedirectUri: '/' }).catch(() => {}),
      );
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
  const nav = ({ isActive }: { isActive: boolean }) =>
    isActive ? 'active' : undefined;
  return (
    <nav className="sidebar">
      <h4>Directory</h4>
      <NavLink to="/" end className={nav}>
        Overview
      </NavLink>
      <NavLink to="/applications" className={nav}>
        App registrations
      </NavLink>
      <NavLink to="/enterprise-apps" className={nav}>
        Enterprise apps
      </NavLink>
      <h4>Tools</h4>
      <NavLink to="/settings" className={nav}>
        Settings
      </NavLink>
    </nav>
  );
}

function SignInScreen() {
  const { instance } = useMsal();
  const signIn = () => {
    instance
      .loginPopup({ scopes: GRAPH_SCOPES, prompt: 'select_account' })
      .catch(() =>
        instance
          .loginRedirect({ scopes: GRAPH_SCOPES, prompt: 'select_account' })
          .catch(() => {}),
      );
  };
  return (
    <div className="center">
      <h1>EntraApp</h1>
      <p className="muted">Sign in with your Entra ID account to continue.</p>
      <button className="primary" onClick={signIn}>
        Sign in with Microsoft
      </button>
    </div>
  );
}
