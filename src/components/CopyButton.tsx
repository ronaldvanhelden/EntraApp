import { useState } from 'react';

interface Props {
  value: string;
  label?: string;
  className?: string;
}

// Low-profile clipboard button. Renders an outline clipboard icon by default;
// swaps to a checkmark for a moment after a successful copy.
export function CopyButton({ value, label, className }: Props) {
  const [copied, setCopied] = useState(false);

  const copy = async (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    try {
      await navigator.clipboard.writeText(value);
      setCopied(true);
      setTimeout(() => setCopied(false), 1200);
    } catch {
      // Clipboard might be blocked (insecure context, permission prompt).
      // No-op: the raw value is still visible for manual selection.
    }
  };

  return (
    <button
      type="button"
      className={`copy-btn${copied ? ' copied' : ''}${className ? ` ${className}` : ''}`}
      onClick={copy}
      aria-label={copied ? 'Copied' : `Copy ${label ?? 'value'}`}
      title={copied ? 'Copied' : `Copy ${label ?? 'value'}`}
    >
      {copied ? (
        <svg
          viewBox="0 0 16 16"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.8"
          strokeLinecap="round"
          strokeLinejoin="round"
          aria-hidden
        >
          <path d="M3 8.5l3.2 3 6.8-7" />
        </svg>
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
          <rect x="4.5" y="4.5" width="8" height="9" rx="1.2" />
          <path d="M6.5 4.5V3a1 1 0 0 1 1-1h3a1 1 0 0 1 1 1v1.5" />
        </svg>
      )}
    </button>
  );
}
