import { useCallback, useRef, useState } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import { SITE_READ_SCOPES } from '../auth/config';
import { listGroupSites, type GroupSite } from '../graph/quickSearch';
import { CopyButton } from './CopyButton';

// SharePoint site URLs for Microsoft 365 (unified) groups. Reading them takes
// Sites.Read.All, which isn't part of sign-in: it is tried silently, and when
// consent is still missing the cells offer a button that asks for it.

type TokenFn = () => Promise<string>;

export function useGroupSites() {
  const silent = useGraphToken(SITE_READ_SCOPES, false);
  const interactive = useGraphToken(SITE_READ_SCOPES);
  const [sites, setSites] = useState<Map<string, GroupSite>>(() => new Map());
  const [needsConsent, setNeedsConsent] = useState(false);
  // What the page currently wants, so granting consent can load all of it.
  const wanted = useRef<string[]>([]);

  const run = useCallback(async (token: TokenFn, ids: string[]) => {
    try {
      const found = await listGroupSites(token, ids);
      setNeedsConsent(false);
      setSites((prev) => {
        const next = new Map(prev);
        // 0: the lookup itself failed, shown as "couldn't be read".
        for (const id of ids) next.set(id, found.get(id) ?? { status: 0 });
        return next;
      });
    } catch {
      setNeedsConsent(true);
    }
  }, []);

  // Ask for the sites of these groups (unified ones only make sense).
  const request = useCallback(
    (ids: string[]) => {
      wanted.current = ids;
      if (ids.length > 0) void run(silent, ids);
    },
    [run, silent],
  );
  // Must be called from a click: may open the consent popup.
  const grant = useCallback(() => void run(interactive, wanted.current), [run, interactive]);

  return { sites, needsConsent, request, grant };
}

export function GroupSiteLink({
  site,
  needsConsent,
  onGrant,
  full = false,
}: {
  // undefined while loading.
  site: GroupSite | undefined;
  needsConsent: boolean;
  onGrant: () => void;
  // The whole URL with a copy button (detail page) instead of just the path.
  full?: boolean;
}) {
  if (site?.url) {
    let label = site.url;
    if (!full) {
      try {
        label = decodeURIComponent(new URL(site.url).pathname) || site.url;
      } catch {
        // Keep the raw value.
      }
    }
    return (
      <>
        <a
          href={site.url}
          target="_blank"
          rel="noreferrer"
          title={`Open ${site.url}`}
          onClick={(e) => e.stopPropagation()}
        >
          {label}
        </a>
        {full && <CopyButton value={site.url} label="SharePoint URL" />}
      </>
    );
  }
  if (site)
    return (
      <span
        className="muted"
        title={
          site.status === 403
            ? "Your account isn't allowed to open this group's site"
            : site.status === 404
              ? 'This group has no SharePoint site (yet)'
              : "The site couldn't be read"
        }
      >
        {site.status === 403 ? 'No access' : site.status === 404 ? 'No site' : '—'}
      </span>
    );
  if (needsConsent)
    return (
      <button
        type="button"
        className="ghost site-consent"
        title="Reading SharePoint sites needs the Sites.Read.All permission — you're asked to consent once"
        onClick={(e) => {
          e.stopPropagation();
          onGrant();
        }}
      >
        Show site
      </button>
    );
  return <span className="spinner" />;
}
