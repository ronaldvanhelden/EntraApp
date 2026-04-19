import { useMemo } from 'react';

// Deterministic blocky identicon (GitHub-style 5x5 mirrored grid). Given the
// same input string it always renders the same SVG, so every resource gets a
// stable visual identifier derived from its id.

interface Props {
  // Any unique identifier — resource GUID, appId, UPN, etc.
  id: string;
  size?: number;
  title?: string;
}

// FNV-1a 32-bit hash — small, fast, no crypto dependency needed.
function hash32(s: string): number {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return h >>> 0;
}

// Derive a full palette from the hash so the foreground and background feel
// intentional rather than random. HSL keeps brightness consistent across
// different hues.
function colorsFromHash(h: number): { fg: string; bg: string } {
  const hue = h % 360;
  // Dark background that harmonizes with our dark theme.
  const bg = `hsl(${hue}, 35%, 18%)`;
  const fg = `hsl(${(hue + 30) % 360}, 65%, 60%)`;
  return { fg, bg };
}

export function Identicon({ id, size = 28, title }: Props) {
  const svg = useMemo(() => buildSvg(id || 'unknown', size), [id, size]);
  return (
    <span
      role="img"
      aria-label={title ?? 'identicon'}
      title={title ?? id}
      style={{
        display: 'inline-block',
        width: size,
        height: size,
        lineHeight: 0,
        flexShrink: 0,
      }}
      // The SVG is a pure-function output of `id`, so this is safe. Rendering
      // via innerHTML avoids a per-cell React reconciliation overhead.
      dangerouslySetInnerHTML={{ __html: svg }}
    />
  );
}

function buildSvg(id: string, size: number): string {
  const grid = 5;
  const cells: boolean[] = [];
  // Use successive bits of a rehashed id to fill the left three columns of a
  // 5-column grid, then mirror them. This is the canonical 5x5 identicon.
  let h = hash32(id);
  for (let i = 0; i < grid * 3; i++) {
    cells.push((h & 1) === 1);
    h = (h >>> 1) | ((h & 1) << 31);
    if (i % 4 === 3) h = hash32(`${id}:${i}`);
  }
  const { fg, bg } = colorsFromHash(hash32(`${id}#c`));
  const cellSize = size / grid;
  const rects: string[] = [];
  for (let y = 0; y < grid; y++) {
    for (let x = 0; x < grid; x++) {
      const srcX = x < 3 ? x : grid - 1 - x;
      const on = cells[y * 3 + srcX];
      if (!on) continue;
      rects.push(
        `<rect x="${x * cellSize}" y="${y * cellSize}" width="${cellSize}" height="${cellSize}" fill="${fg}"/>`,
      );
    }
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" style="border-radius:4px;display:block;"><rect width="${size}" height="${size}" fill="${bg}"/>${rects.join('')}</svg>`;
}
