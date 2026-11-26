# EntraApp

[WebApp](https://entra.2cloud.nl/)

A dark-mode, dedicated React SPA for managing **Microsoft Entra ID** app registrations
and enterprise apps, with a focus on **managing API permissions** (both application
and delegated) on enterprise apps. Authenticates with MSAL using delegated permissions
against Microsoft Graph.

## Features

- Sign in with MSAL (delegated, browser SPA / PKCE)
- Browse **app registrations** and view metadata & manifest-required resource access
- Browse **enterprise apps** (service principals) with search and type filters
- **Manage API permissions** on enterprise apps:
  - Add application (app-only) permissions via `appRoleAssignments`
  - Add delegated permissions via `oauth2PermissionGrants` (admin consent by
    default, single-user consent supported)
  - Revoke either kind inline
  - Pick common APIs (Microsoft Graph, Exchange, SharePoint, Azure Management)
    or search all service principals
- **Risky permissions**: every service principal holding a dangerous
  application (app-only) permission on Microsoft Graph, Exchange Online,
  SharePoint Online or the legacy Azure AD Graph, worst first. Permissions are
  rated _critical_ (a route to Global Administrator, such as
  `RoleManagement.ReadWrite.Directory` or `AppRoleAssignment.ReadWrite.All`) or
  _high_ (tenant-wide mail, files and chats, or write access to the directory);
  the list is in `src/lib/dangerousPermissions.ts`. Needs no permission beyond
  the reading set.
- **Tags & favorites** on app registrations, stored in the directory
  (`application.tags`) so they're shared with everyone managing the tenant. Tags
  written here are namespaced `debble:<name>`; tags set by Entra or other tooling
  (`HideApp`, …) are shown read-only and preserved on every write.
  `debble:Favorite` is a reserved tag rendered as a star; the list filters by tag.
- **Unique name** (`application.uniqueName`) shown on the list and detail pages,
  searchable, and settable once for apps that don't have one yet
- **Tenant shortcuts**: pin a tab to a specific tenant with `?tenant=…`

## Tenant shortcuts (`?tenant=`)

Add `?tenant=<tenant id or domain>` to the URL to point sign-in — and therefore
every Graph call — at that directory instead of the saved default. Because the
app uses a `HashRouter`, the parameter sits in front of the hash and survives all
in-app navigation, so a single bookmark keeps you in that tenant:

```
https://entra.2cloud.nl/?tenant=debble.com#/enterprise-apps
```

- Works with a tenant GUID, a verified domain, or `common` / `organizations`.
- Useful when your home tenant differs from the one you manage: sign in with the
  account (guest accounts included) that has access to the pinned tenant.
- The pin is remembered per browser tab, so a redirect sign-in doesn't lose it
  and a second tab on the plain URL is back on the default tenant.
- The header chip shows the active directory and marks it when pinned by the URL.
- **Settings → Tenant shortcuts** stores named links per tenant, ready to copy or
  bookmark. Saving a new default tenant in Settings releases the pin.

## Stack

- React 18 + TypeScript + Vite
- `@azure/msal-browser` + `@azure/msal-react`
- Microsoft Graph v1.0 via `fetch`
- `HashRouter` for zero-config GitHub Pages hosting

## Run locally

```bash
npm install
npm run dev
```

Open the URL Vite prints, then either accept the default app registration
(`Debble EntraApp`) or go to **Settings** and set your own client/tenant.

## App registration requirements

The default client is the multi-tenant app registration
`438ccf20-d1a2-4027-baf3-0104d018357e` (`Debble EntraApp`) with SPA redirect
URI `https://entra.2cloud.nl/`. To use a different client you
must register an application in Entra ID with:

- **Platform**: Single-page application (SPA)
- **Redirect URI**: the exact URL the SPA is served from, including the
  trailing slash (e.g. `https://entra.2cloud.nl/`).
  The Setup screen computes and shows this URL — copy it into the app
  registration.
- **Supported account types**: any (single-tenant or multi-tenant)
- **API permissions (Microsoft Graph, Delegated)** — declare these on the app
  registration. The app asks for them in steps, so nobody consents to more
  than they use:
  - **Sign-in**: `User.Read`
  - **Reading** (asked once, right after sign-in): `User.Read.All`,
    `Group.Read.All`, `Application.Read.All`,
    `DelegatedPermissionGrant.Read.All`, `CrossTenantInformation.ReadBasic.All`,
    `AuditLog.Read.All`
  - **Changing things** (each asked the first time such a change is made):
    - `Application.ReadWrite.All` — app registrations and enterprise apps
    - `DelegatedPermissionGrant.ReadWrite.All` — delegated permission grants
    - `AppRoleAssignment.ReadWrite.All` — application permissions and
      user/group assignments
    - `User.ReadWrite.All`, `Group.ReadWrite.All` — users and groups
    - `User-PasswordProfile.ReadWrite.All` — resetting a password
    - `User.Invite.All` — resending a guest invitation
    - `RoleAssignmentSchedule.ReadWrite.Directory` — activating your own
      role (PIM); with `RoleEligibilitySchedule.ReadWrite.Directory` also
      assigning and removing roles for others
  - **Optional extras** (asked from the screen that shows them):
    `RoleManagement.Read.Directory` (roles and PIM),
    `IdentityRiskyUser.Read.All`, `Sites.Read.All`
- **Admin consent**: the `.All` scopes require it. Granting it for the tenant
  up front skips the prompts; a read-only admin can instead consent to the
  reading set alone. Tenants that consented before reading and writing were
  split keep working: the read-write permissions they hold cover the
  read-only ones.

## Deploy to GitHub Pages

Pushes to `main` automatically build and deploy `dist/` via
`.github/workflows/static.yml`. The site is published on the custom domain
`https://entra.2cloud.nl/` and built with base path `/`.

Custom domain setup:

1. DNS (TransIP): a `CNAME` record `entra` → `ronaldvanhelden.github.io.`
2. Repo **Settings → Pages → Custom domain**: `entra.2cloud.nl`, then enable
   **Enforce HTTPS**. No `CNAME` file is needed with Actions-based deploys.
3. App registration: add SPA redirect URI `https://entra.2cloud.nl/`.

Forks without a custom domain are served from `https://<owner>.github.io/<repo>/`;
build those with `VITE_BASE=/<repo>/` (set it on the Build step in the workflow).

## How permission grants work (mental model)

- **Delegated** permissions are stored on the **client** service principal as
  `oauth2PermissionGrants` rows, one row per (clientId, resourceId, consentType).
  Multiple scopes live in the same row as a space-separated `scope` string.
  Adding a scope PATCHes the row; the last scope removed DELETEs the row.
- **Application (app-only)** permissions are stored as `appRoleAssignments`
  under the client service principal, one row per granted role. Adding POSTs
  a new row; revoking DELETEs by assignment id.

## Security notes

- No secrets are stored; all calls go directly from the browser to Graph with
  a user access token acquired via PKCE.
- Configuration (client/tenant IDs) lives in `localStorage` under
  `entraapp.authConfig` so each device can point at a different client.
- The same permissions that let you grant/revoke API permissions let you do
  serious damage to the tenant — only use admin accounts that understand this.
