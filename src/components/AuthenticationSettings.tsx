import { useState, type ReactNode } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import { updateApplication, type UpdateApplicationPatch } from '../graph/applications';
import type { Application } from '../graph/types';
import { CopyButton } from './CopyButton';
import { Modal } from './Modal';
import { useConfirm } from './ConfirmDialog';

type Platform = 'web' | 'spa' | 'publicClient';
type ClientKind = 'confidential' | 'public';
type Status = 'recommended' | 'current' | 'legacy' | 'avoid';

const STATUS: Record<Status, { label: string; cls: string; title: string }> = {
  recommended: {
    label: 'Recommended',
    cls: 'badge granted',
    title: 'Current best practice from Microsoft and OAuth 2.1',
  },
  current: { label: 'Current', cls: 'badge delegated', title: 'Supported; use when it fits' },
  legacy: {
    label: 'Legacy',
    cls: 'badge pending',
    title: 'Kept for older apps; do not use for new ones',
  },
  avoid: {
    label: 'Not recommended',
    cls: 'badge expired',
    title: 'Security risk; only for apps that cannot be changed',
  },
};

const CLIENT: Record<ClientKind, { label: string; cls: string; title: string }> = {
  confidential: {
    label: 'Confidential client · backend',
    cls: 'badge app',
    title: 'Runs on a server you control and can keep a secret or certificate',
  },
  public: {
    label: 'Public client · frontend',
    cls: 'badge delegated',
    title: 'Runs on the user’s device (browser, phone, desktop, CLI) and cannot keep a secret',
  },
};

// The Authentication tab: one card per way an app can sign users in, each
// labelled with the kind of client it is for and whether it is current or
// legacy, plus the access token version this app gets as an API.
export function AuthenticationSettings({
  app,
  onAppChange,
  ficCount = 0,
}: {
  app: Application;
  onAppChange: (next: Application) => void;
  // Federated credentials are loaded separately by the page.
  ficCount?: number;
}) {
  const token = useGraphToken();
  const { confirm, confirmDialog } = useConfirm();
  const [saving, setSaving] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [editPlatform, setEditPlatform] = useState<Platform | null>(null);

  const webUris = app.web?.redirectUris ?? [];
  const spaUris = app.spa?.redirectUris ?? [];
  const publicUris = app.publicClient?.redirectUris ?? [];
  const idTokenIssuance = !!app.web?.implicitGrantSettings?.enableIdTokenIssuance;
  const accessTokenIssuance = !!app.web?.implicitGrantSettings?.enableAccessTokenIssuance;
  const fallbackPublic = !!app.isFallbackPublicClient;
  const hasCredential =
    (app.passwordCredentials?.length ?? 0) > 0 ||
    (app.keyCredentials?.length ?? 0) > 0 ||
    (app.federatedIdentityCredentials?.length ?? 0) > 0 ||
    ficCount > 0;

  const patch = async (body: UpdateApplicationPatch, key: string) => {
    setSaving(key);
    setError(null);
    try {
      await updateApplication(token, app.id, body);
      onAppChange({ ...app, ...body });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      throw e;
    } finally {
      setSaving(null);
    }
  };

  const setImplicit = (next: {
    enableIdTokenIssuance?: boolean;
    enableAccessTokenIssuance?: boolean;
  }) =>
    patch(
      {
        web: {
          ...app.web,
          implicitGrantSettings: {
            enableIdTokenIssuance: idTokenIssuance,
            enableAccessTokenIssuance: accessTokenIssuance,
            ...next,
          },
        },
      },
      'implicit',
    ).catch(() => {});

  const saveRedirectUris = async (platform: Platform, uris: string[]) => {
    const body: UpdateApplicationPatch =
      platform === 'web'
        ? { web: { ...app.web, redirectUris: uris } }
        : platform === 'spa'
          ? { spa: { ...app.spa, redirectUris: uris } }
          : { publicClient: { ...app.publicClient, redirectUris: uris } };
    await patch(body, `uris:${platform}`);
    setEditPlatform(null);
  };

  const busy = saving !== null;

  return (
    <>
      <div className="card">
        <h3>Authentication</h3>
        <p className="muted" style={{ marginTop: 0, fontSize: 13 }}>
          An app registration can serve several kinds of client at once. Each card below is one way
          to sign in, and says which kind of client it is for:
        </p>
        <div className="auth-legend">
          <div>
            <ClientBadge kind="confidential" />
            <p>
              Code runs on a server you control (ASP.NET, Node, Java, a daemon, a web API calling
              another API). It proves who it is with a <strong>client secret</strong>,{' '}
              <strong>certificate</strong> or <strong>federated credential</strong> when it redeems
              a code or asks for a token.
            </p>
          </div>
          <div>
            <ClientBadge kind="public" />
            <p>
              Code runs on the user’s device: browser SPA, mobile app, desktop app, CLI. Anything
              shipped there can be read, so it <strong>cannot keep a secret</strong>. Instead it
              protects the code exchange with <strong>PKCE</strong>.
            </p>
          </div>
        </div>
        {error && <p className="error">{error}</p>}
      </div>

      <AuthCard
        title="Web application (server-side)"
        icon="server"
        use="Server-side web apps and backends-for-frontend (ASP.NET, Node, Java, Python) that sign users in and keep the tokens on the server."
        client="confidential"
        status="recommended"
        summary="Server-side sign-in with the authorization code flow. The browser only carries a one-time code; your server swaps it for tokens, proving itself with a secret, certificate or federated credential."
        can={[
          'Authorization code grant: /authorize returns the code to a Web redirect URI (query or form_post), never to browser JavaScript',
          'Code redemption at /oauth2/v2.0/token authenticated with a client_secret, or a client_assertion JWT signed by a certificate or issued by a federated identity',
          'ID token, access token and, with offline_access, a refresh token with a sliding lifetime of up to 90 days',
          'With the same credential: the on-behalf-of grant (jwt-bearer) for downstream APIs and client credentials for app-only tokens',
        ]}
        footnote={
          webUris.length > 0 && !hasCredential
            ? '⚠ This app has Web redirect URIs but no secret, certificate or federated credential, so codes cannot be redeemed.'
            : undefined
        }
      >
        <RedirectUris uris={webUris} onEdit={() => setEditPlatform('web')} disabled={busy} />
      </AuthCard>

      <AuthCard
        title="Single-page application (browser)"
        icon="browser"
        use="Browser apps (React, Angular, Vue, Blazor WebAssembly) with no backend of their own that sign users in and call APIs straight from JavaScript."
        client="public"
        status="recommended"
        summary="Browser JavaScript (MSAL.js 2+, React, Angular, Vue) signs in with authorization code + PKCE and redeems the code itself, directly from the browser."
        can={[
          'Authorization code grant with PKCE (S256): the browser redeems the code itself with a CORS request to /oauth2/v2.0/token, allowed only for the Origin of an SPA redirect URI',
          'ID token, access token and a refresh token that is single-use and expires 24 hours after sign-in',
          'Silent token renewal with the refresh token instead of a hidden iframe with prompt=none, so it works when third-party cookies are blocked',
        ]}
      >
        <RedirectUris uris={spaUris} onEdit={() => setEditPlatform('spa')} disabled={busy} />
      </AuthCard>

      <AuthCard
        title="Mobile & desktop application"
        icon="device"
        use="Apps installed on the user’s device — iOS, Android, Windows, macOS, Electron, or a CLI that opens a browser — that sign users in interactively or through a broker."
        client="public"
        status="recommended"
        summary="Native apps (WPF/WinForms, MAUI, iOS, Android, Electron, CLIs that open a browser) sign in with authorization code + PKCE and receive the code on a loopback or custom-scheme URI."
        can={[
          'Authorization code grant with PKCE, code delivered to a loopback (http://localhost:port), custom scheme (msal{client-id}://auth) or nativeclient URI',
          'Code redemption at /oauth2/v2.0/token without client authentication: the app is a public client',
          'Brokered sign-in through Authenticator, Company Portal or Windows WAM using the device’s Primary Refresh Token, so device-based Conditional Access and SSO across apps apply',
          'Refresh tokens with a sliding lifetime of up to 90 days, kept in the OS token cache',
        ]}
        footnote={
          <>
            Typical URIs: <code>http://localhost</code> (loopback, any port),{' '}
            <code>msal{app.appId}://auth</code> (MSAL mobile),{' '}
            <code>https://login.microsoftonline.com/common/oauth2/nativeclient</code> (embedded web
            view).
          </>
        }
      >
        <RedirectUris
          uris={publicUris}
          onEdit={() => setEditPlatform('publicClient')}
          disabled={busy}
        />
      </AuthCard>

      <AuthCard
        title="Sign-in without a redirect (public client flows)"
        icon="terminal"
        use="Only apps that truly can’t show a browser or receive a redirect, such as Teams Rooms and other shared or input-constrained devices, or legacy tools that can’t be updated yet. For CLIs and scripts, prefer browser-based or brokered sign-in for users, and managed identities or workload identity federation for automation."
        client="public"
        status={fallbackPublic ? 'legacy' : 'current'}
        statusOverride={[
          { status: 'avoid', text: 'Device code' },
          { status: 'avoid', text: 'Username/password (ROPC)' },
          { status: 'legacy', text: 'Windows integrated auth' },
        ]}
        summary="Flows that never use a redirect URI. With this on, Entra treats the app as a public client for them, so no secret is needed."
        can={[
          'Device authorization grant: /devicecode returns a user_code and device_code; the app polls /token with grant_type=urn:ietf:params:oauth:grant-type:device_code while the user signs in elsewhere',
          'Resource owner password credentials (ROPC): grant_type=password posts the username and password straight to /token (no MFA, no interactive Conditional Access)',
          'Integrated Windows Authentication: a Kerberos ticket or WS-Trust token from a domain-joined machine is exchanged for Entra tokens',
          'All of these at /token without client_secret or client_assertion: isFallbackPublicClient=true makes Entra treat the app as public for flows without a redirect URI',
        ]}
        footnote={
          <>
            ⚠ Device code flow is a common phishing vector: an attacker starts the flow and gets a
            user to enter the code on the real Microsoft sign-in page, and the tokens go to the
            attacker. Microsoft recommends blocking it wherever possible. Security defaults block
            it, and from 1 July 2026 every new tenant does too. Other tenants can turn on the
            Microsoft-managed Conditional Access policy “Block device code flow”. See{' '}
            <a
              href="https://learn.microsoft.com/entra/identity/conditional-access/policy-block-authentication-flows"
              target="_blank"
              rel="noreferrer"
            >
              Block authentication flows
            </a>
            .
          </>
        }
      >
        <Toggle
          checked={fallbackPublic}
          disabled={busy}
          onChange={(v) => patch({ isFallbackPublicClient: v }, 'fallback').catch(() => {})}
          on="Yes, public client flows allowed"
          off="No"
        />
      </AuthCard>

      <AuthCard
        title="Implicit grant (legacy browser sign-in)"
        icon="link"
        use="Only older browser apps that can’t be updated yet: MSAL.js 1.x / ADAL.js, or ASP.NET / OWIN apps using the “code id_token” hybrid flow. New apps should use the SPA or Web platform."
        client="public"
        status="legacy"
        summary="The pre-PKCE way for browser apps to get tokens: the token is returned directly in the redirect URL fragment instead of a code. Kept for apps that cannot move to the SPA platform yet."
        can={[
          'ID tokens: /authorize with response_type=id_token (or the hybrid code id_token) returns the ID token directly in the redirect URI fragment or form_post',
          'Access tokens: /authorize with response_type=token returns the access token in the URL fragment, with no call to the token endpoint',
          'No refresh tokens: renewal is a hidden iframe with prompt=none that relies on the Entra session cookie',
        ]}
      >
        <div className="stack" style={{ gap: 6 }}>
          <Toggle
            checked={idTokenIssuance}
            disabled={busy}
            onChange={(v) => setImplicit({ enableIdTokenIssuance: v })}
            on="ID tokens issued (response_type=id_token)"
            off="ID tokens: off"
          />
          <Toggle
            checked={accessTokenIssuance}
            disabled={busy}
            onChange={(v) => setImplicit({ enableAccessTokenIssuance: v })}
            on="Access tokens issued (response_type=token)"
            off="Access tokens: off"
          />
        </div>
      </AuthCard>

      <TokenVersionCard
        app={app}
        busy={busy}
        confirm={confirm}
        onSave={(body) => patch(body, 'tokenver')}
      />

      {editPlatform && (
        <RedirectUriEditModal
          title={
            editPlatform === 'web'
              ? 'Edit Web redirect URIs'
              : editPlatform === 'spa'
                ? 'Edit SPA redirect URIs'
                : 'Edit mobile & desktop redirect URIs'
          }
          initial={editPlatform === 'web' ? webUris : editPlatform === 'spa' ? spaUris : publicUris}
          onClose={() => setEditPlatform(null)}
          onSave={(uris) => saveRedirectUris(editPlatform, uris)}
        />
      )}
      {confirmDialog}
    </>
  );
}

function AuthCard({
  title,
  icon,
  client,
  status,
  statusOverride,
  use,
  summary,
  can,
  footnote,
  children,
}: {
  title: string;
  icon: AuthIconName;
  client: ClientKind;
  status: Status;
  // For a card covering several flows with different standing.
  statusOverride?: Array<{ status: Status; text: string }>;
  // Who this card is for, in plain words: the kind of app that should use it.
  use: string;
  summary: string;
  can: string[];
  footnote?: ReactNode;
  children?: ReactNode;
}) {
  return (
    <div className="card auth-card">
      <AuthCardHeader icon={icon} tone={client} title={title}>
        {statusOverride ? (
          statusOverride.map((s) => (
            <span key={s.text} className={STATUS[s.status].cls} title={STATUS[s.status].title}>
              {s.text}: {STATUS[s.status].label}
            </span>
          ))
        ) : (
          <span className={STATUS[status].cls} title={STATUS[status].title}>
            {STATUS[status].label}
          </span>
        )}
      </AuthCardHeader>
      <IntendedUse>{use}</IntendedUse>
      <p style={{ margin: '0 0 10px', fontSize: 13 }}>{summary}</p>
      <div className="auth-can">
        <div className="auth-can-title">What it enables</div>
        <ul>
          {can.map((x) => (
            <li key={x}>{x}</li>
          ))}
        </ul>
      </div>
      {footnote && (
        <div className="muted" style={{ fontSize: 12, marginTop: 8 }}>
          {footnote}
        </div>
      )}
      {children && <div style={{ marginTop: 12 }}>{children}</div>}
    </div>
  );
}

// Client-kind badge. Public clients lead with a JS mark: the code ships to the
// user's device, where anyone can read it.
function ClientBadge({ kind }: { kind: ClientKind }) {
  const c = CLIENT[kind];
  return (
    <span className={`${c.cls} client-badge`} title={c.title}>
      {kind === 'public' && (
        <span className="js-icon" aria-hidden>
          JS
        </span>
      )}
      {c.label}
    </span>
  );
}

type AuthIconName = 'server' | 'browser' | 'device' | 'terminal' | 'link' | 'token';

// 24px line icons, drawn in the card's tone colour.
const AUTH_ICONS: Record<AuthIconName, ReactNode> = {
  server: (
    <>
      <rect x="3" y="4" width="18" height="7" rx="1.5" />
      <rect x="3" y="13" width="18" height="7" rx="1.5" />
      <path d="M7 7.5h.01M7 16.5h.01M11 7.5h6M11 16.5h6" />
    </>
  ),
  browser: (
    <>
      <rect x="3" y="4" width="18" height="16" rx="2" />
      <path d="M3 9h18M6.5 6.5h.01M9 6.5h.01" />
      <path d="M10 13l-2 2 2 2M14 13l2 2-2 2" />
    </>
  ),
  device: (
    <>
      <rect x="2" y="4" width="14" height="10" rx="1.5" />
      <path d="M6 18h6M9 14v4" />
      <rect x="17" y="8" width="5" height="12" rx="1" />
      <path d="M19.5 17.5h.01" />
    </>
  ),
  terminal: (
    <>
      <rect x="3" y="4" width="18" height="16" rx="2" />
      <path d="M7 9l3 3-3 3M12 15h5" />
    </>
  ),
  link: (
    <>
      <path d="M10 14a4 4 0 0 0 5.66 0l3-3a4 4 0 0 0-5.66-5.66l-1 1" />
      <path d="M14 10a4 4 0 0 0-5.66 0l-3 3a4 4 0 0 0 5.66 5.66l1-1" />
    </>
  ),
  token: (
    <>
      <circle cx="8" cy="15" r="4" />
      <path d="M10.85 12.15L19 4M16 7l2 2M14 9l2 2" />
    </>
  ),
};

// Card title row: a JS mark first for public clients, then the card icon, title, badges, then a separator line.
function AuthCardHeader({
  icon,
  tone,
  title,
  children,
}: {
  icon: AuthIconName;
  tone: ClientKind | 'api';
  title: string;
  children?: ReactNode;
}) {
  return (
    <div className="auth-card-header">
      {tone === 'public' && (
        <span
          className="js-icon js-icon-lg"
          role="img"
          aria-label={CLIENT.public.label}
          title={CLIENT.public.title}
        >
          JS
        </span>
      )}
      <span className={`auth-icon ${tone}`} aria-hidden>
        <svg
          width="18"
          height="18"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.8"
          strokeLinecap="round"
          strokeLinejoin="round"
        >
          {AUTH_ICONS[icon]}
        </svg>
      </span>
      <h3>{title}</h3>
      <div className="auth-card-badges">{children}</div>
    </div>
  );
}

function IntendedUse({ children }: { children: ReactNode }) {
  return (
    <div className="auth-use">
      <strong>Intended use:</strong> {children}
    </div>
  );
}

function Toggle({
  checked,
  disabled,
  onChange,
  on,
  off,
}: {
  checked: boolean;
  disabled?: boolean;
  onChange: (next: boolean) => void;
  on: string;
  off: string;
}) {
  return (
    <label className="row" style={{ gap: 6 }}>
      <input
        type="checkbox"
        style={{ width: 'auto' }}
        checked={checked}
        disabled={disabled}
        onChange={(e) => onChange(e.target.checked)}
      />
      {checked ? on : off}
    </label>
  );
}

function RedirectUris({
  uris,
  onEdit,
  disabled,
}: {
  uris: string[];
  onEdit: () => void;
  disabled?: boolean;
}) {
  return (
    <div>
      <div className="row" style={{ justifyContent: 'space-between' }}>
        <strong style={{ fontSize: 13 }}>Redirect URIs ({uris.length})</strong>
        <button onClick={onEdit} disabled={disabled}>
          Edit
        </button>
      </div>
      {uris.length === 0 ? (
        <div className="muted" style={{ fontSize: 13, marginTop: 4 }}>
          None: this platform is not in use.
        </div>
      ) : (
        <ul className="mono" style={{ fontSize: 12, margin: '6px 0 0', paddingLeft: 20 }}>
          {uris.map((u) => (
            <li key={u}>
              {u}
              <CopyButton value={u} label="redirect URI" />
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

// requestedAccessTokenVersion lives on the API (resource) side: it decides
// the format of access tokens other clients receive *for this app*, whichever
// endpoint (v1 /oauth2/token or v2 /oauth2/v2.0/token) they used.
function TokenVersionCard({
  app,
  busy,
  confirm,
  onSave,
}: {
  app: Application;
  busy: boolean;
  confirm: ReturnType<typeof useConfirm>['confirm'];
  onSave: (body: UpdateApplicationPatch) => Promise<void>;
}) {
  const raw = app.api?.requestedAccessTokenVersion;
  const version = raw === 2 ? 2 : 1;
  const tid = '{tenant-id}';
  const appIdUri = app.identifierUris?.[0];
  const personal =
    app.signInAudience === 'AzureADandPersonalMicrosoftAccount' ||
    app.signInAudience === 'PersonalMicrosoftAccount';

  const choose = async (next: 1 | 2) => {
    if (next === version) return;
    const ok = await confirm({
      title: `Switch to v${next}.0 access tokens?`,
      message: (
        <>
          <p style={{ marginTop: 0 }}>
            New access tokens for this API will have a different <code>iss</code> and usually a
            different <code>aud</code>:
          </p>
          <p className="mono" style={{ fontSize: 12 }}>
            iss →{' '}
            {next === 2
              ? `https://login.microsoftonline.com/${tid}/v2.0`
              : `https://sts.windows.net/${tid}/`}
            <br />
            aud → {next === 2 ? app.appId : (appIdUri ?? 'the App ID URI the client requested')}
          </p>
          <p style={{ marginBottom: 0 }}>
            If the API validates issuer or audience strictly, it must accept the new values first,
            or callers get 401s. Tokens already issued keep their old format until they expire.
          </p>
        </>
      ),
      confirmLabel: `Switch to v${next}.0`,
    });
    if (!ok) return;
    await onSave({
      api: { ...app.api, requestedAccessTokenVersion: next === 2 ? 2 : null },
    }).catch(() => {});
  };

  const rows: Array<[string, string, string]> = [
    ['ver', '"1.0"', '"2.0"'],
    ['iss', `https://sts.windows.net/${tid}/`, `https://login.microsoftonline.com/${tid}/v2.0`],
    [
      'aud',
      appIdUri
        ? `${appIdUri} (App ID URI, as requested)`
        : 'App ID URI, as requested (e.g. api://…)',
      `${app.appId} (client id GUID)`,
    ],
    ['Calling app', 'appid, appidacr', 'azp, azpacr'],
    ['User name', 'unique_name, upn', 'preferred_username (upn is optional)'],
    ['Personal accounts', 'Not supported', 'Supported'],
  ];

  return (
    <div className="card auth-card">
      <AuthCardHeader icon="token" tone="api" title="Access token version (this app as an API)">
        <span className={STATUS.recommended.cls}>v2.0: Recommended</span>
        <span className={STATUS.legacy.cls}>v1.0: Legacy default</span>
      </AuthCardHeader>
      <IntendedUse>
        Set this when the app is an API that other apps call. Choose v2.0 for new APIs; keep v1.0
        only when existing consumers or token validation code expect v1.0 tokens.
      </IntendedUse>
      <p style={{ margin: '0 0 8px', fontSize: 13 }}>
        Decides the format of access tokens that <em>other</em> apps get to call this app’s API. The
        resource decides, not the client: a client using the v2.0 endpoint still gets a v1.0 token
        for an API set to v1. ID tokens are not affected; their version follows the endpoint the
        client signed in with.
      </p>

      <div className="row" style={{ gap: 12, marginBottom: 12 }}>
        {([1, 2] as const).map((v) => (
          <label key={v} className="row" style={{ gap: 6 }}>
            <input
              type="radio"
              name="tokenver"
              style={{ width: 'auto' }}
              checked={version === v}
              disabled={busy || (v === 1 && personal)}
              onChange={() => choose(v)}
            />
            v{v}.0{v === 1 && raw == null ? ' (default, not set)' : ''}
          </label>
        ))}
        {personal && (
          <span className="muted" style={{ fontSize: 12 }}>
            v1.0 isn’t possible: the sign-in audience includes personal Microsoft accounts.
          </span>
        )}
      </div>

      <div style={{ overflowX: 'auto' }}>
        <table className="table token-compare">
          <thead>
            <tr>
              <th>Claim</th>
              <th className={version === 1 ? 'current' : undefined}>v1.0</th>
              <th className={version === 2 ? 'current' : undefined}>v2.0</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(([k, v1, v2]) => (
              <tr key={k}>
                <td>{k}</td>
                <td className={`mono${version === 1 ? ' current' : ''}`}>{v1}</td>
                <td className={`mono${version === 2 ? ' current' : ''}`}>{v2}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="muted" style={{ fontSize: 12, marginTop: 8 }}>
        How to check: decode an access token for this API (e.g. on jwt.ms) and look at{' '}
        <code>ver</code>, <code>iss</code> and <code>aud</code>. In v1.0, <code>aud</code> is
        whatever the client asked for, so a request for{' '}
        <code>{appIdUri ?? 'api://…'}/.default</code> gives that URI and a request by client id
        gives the GUID. In v2.0 it is always the client id GUID. Your API’s token validation must
        accept the matching issuer and audience.
      </div>
    </div>
  );
}

function RedirectUriEditModal({
  title,
  initial,
  onClose,
  onSave,
}: {
  title: string;
  initial: string[];
  onClose: () => void;
  onSave: (uris: string[]) => Promise<void>;
}) {
  const [text, setText] = useState(initial.join('\n'));
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const save = async () => {
    const uris = text
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);
    setSaving(true);
    setError(null);
    try {
      await onSave(uris);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setSaving(false);
    }
  };

  return (
    <Modal
      title={title}
      onClose={() => !saving && onClose()}
      footer={
        <>
          <button disabled={saving} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={saving} onClick={save}>
            {saving ? 'Saving…' : 'Save'}
          </button>
        </>
      }
    >
      <p className="muted" style={{ fontSize: 13, marginTop: 0 }}>
        One URI per line. Must be an HTTPS URL (or <code>http://localhost</code>), except for mobile
        & desktop URIs which also allow custom schemes.
      </p>
      <textarea
        value={text}
        onChange={(e) => setText(e.target.value)}
        className="mono"
        rows={Math.max(6, Math.min(16, text.split('\n').length + 2))}
        style={{ width: '100%', fontSize: 12 }}
        placeholder="https://contoso.example.com/auth/callback"
        autoFocus
      />
      {error && <p className="error">{error}</p>}
    </Modal>
  );
}
