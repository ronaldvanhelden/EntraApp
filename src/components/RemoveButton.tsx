import type { MouseEvent } from 'react';

interface Props {
  /** Accessible name and tooltip, e.g. "Remove resource". */
  label: string;
  onClick: (e: MouseEvent<HTMLButtonElement>) => void;
  disabled?: boolean;
  /** Shows an ellipsis instead of the icon while the action runs. */
  busy?: boolean;
  className?: string;
}

// Icon-only danger button with a recycle-bin glyph. Used for per-item remove
// actions in tables and group headers where a text label is just noise.
export function RemoveButton({ label, onClick, disabled, busy, className }: Props) {
  return (
    <button
      type="button"
      className={`danger icon-btn${className ? ` ${className}` : ''}`}
      aria-label={label}
      title={label}
      disabled={disabled || busy}
      onClick={onClick}
    >
      {busy ? (
        '…'
      ) : (
        <svg
          viewBox="0 0 16 16"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.4"
          strokeLinecap="round"
          strokeLinejoin="round"
          aria-hidden
        >
          <path d="M2.5 4.5h11" />
          <path d="M6 4.5V3a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v1.5" />
          <path d="M4 4.5l.7 8.2a1 1 0 0 0 1 .9h4.6a1 1 0 0 0 1-.9l.7-8.2" />
          <path d="M6.5 7.5v4M9.5 7.5v4" />
        </svg>
      )}
    </button>
  );
}
