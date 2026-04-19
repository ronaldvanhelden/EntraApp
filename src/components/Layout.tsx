import type { ReactNode } from 'react';
import { NavLink } from 'react-router-dom';
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
  return (
    <>
      <UnauthenticatedTemplate>
        <SignInScreen />
      </UnauthenticatedTemplate>
      <AuthenticatedTemplate>
        <div className="layout">
          <Header />
          <Sidebar />
          <main className="main">{children}</main>
        </div>
      </AuthenticatedTemplate>
    </>
  );
}

function Header() {
  const { instance, accounts } = useMsal();
  const active = instance.getActiveAccount() ?? accounts[0];
  return (
    <header className="header">
      <div className="brand">EntraApp</div>
      <div className="user">
        <span>{active?.username}</span>
        <button
          className="ghost"
          onClick={() =>
            instance.logoutPopup({ mainWindowRedirectUri: '/' }).catch(() => {})
          }
        >
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
  return (
    <div className="center">
      <h1>EntraApp</h1>
      <p className="muted">Sign in with your Entra ID account to continue.</p>
      <button
        className="primary"
        onClick={() =>
          instance
            .loginPopup({ scopes: GRAPH_SCOPES, prompt: 'select_account' })
            .catch(() => {})
        }
      >
        Sign in with Microsoft
      </button>
    </div>
  );
}
