// Remembers how far a list page was scrolled so returning to it (back button
// or browser back) lands where the user left off. The page content scrolls
// inside <main class="main"> on desktop and the window on mobile, so both are
// saved. Stored per list URL in sessionStorage; failures are ignored.

const PREFIX = 'entraapp:scroll:';

function mainEl(): HTMLElement | null {
  return document.querySelector<HTMLElement>('.layout > .main');
}

export function saveScroll(key: string) {
  try {
    const pos = { main: mainEl()?.scrollTop ?? 0, win: window.scrollY };
    sessionStorage.setItem(PREFIX + key, JSON.stringify(pos));
  } catch {
    // storage unavailable: nothing to restore later
  }
}

// Restores and forgets the saved position for `key`. Call once the list has
// rendered its rows, or the browser clamps the offset to the short page.
export function restoreScroll(key: string) {
  let pos: { main?: number; win?: number } | null = null;
  try {
    const raw = sessionStorage.getItem(PREFIX + key);
    sessionStorage.removeItem(PREFIX + key);
    pos = raw ? JSON.parse(raw) : null;
  } catch {
    return;
  }
  if (!pos) return;
  requestAnimationFrame(() => {
    const el = mainEl();
    if (el && pos?.main) el.scrollTop = pos.main;
    if (pos?.win) window.scrollTo(0, pos.win);
  });
}

// Each page starts at the top; restoreScroll runs later, once data is in.
export function scrollToTop() {
  const el = mainEl();
  if (el) el.scrollTop = 0;
  window.scrollTo(0, 0);
}
