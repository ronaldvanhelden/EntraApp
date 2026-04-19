# EntraApp

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
(`Debble Permission Manager`) or go to **Settings** and set your own client/tenant.

## App registration requirements

The default client is the multi-tenant app registration
`285e9c19-6642-4f20-af26-489b47636cc9` (`Debble EntraApp`) with SPA redirect
URI `https://ronaldvanhelden.github.io/EntraApp`. To use a different client you
must register an application in Entra ID with:

- **Platform**: Single-page application (SPA)
- **Redirect URI**: the exact URL the SPA is served from (no trailing slash).
  The Setup screen computes and shows this URL — copy it into the app
  registration.
- **Supported account types**: any (single-tenant or multi-tenant)
- **API permissions (Microsoft Graph, Delegated)**:
  - `User.Read`
  - `Application.ReadWrite.All`
  - `Directory.ReadWrite.All`
  - `DelegatedPermissionGrant.ReadWrite.All`
- **Admin consent** granted for the tenant (the `.All` scopes all require it)

## Deploy to GitHub Pages

Pushes to `main` automatically build and deploy `dist/` via
`.github/workflows/static.yml`. The workflow derives the correct base path
(`/<repo>/`) so routing and asset URLs work without edits.

Published URL format: `https://<owner>.github.io/<repo>/`

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
