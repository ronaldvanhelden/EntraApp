import { useState } from 'react';
import { Identicon } from './Identicon';

interface Props {
  id: string;
  logoUrl?: string | null;
  size?: number;
  title?: string;
}

// Prefers a real app logo (from servicePrincipal.info.logoUrl or
// application.info.logoUrl) and falls back to an identicon if the logo fails
// to load or isn't available.
export function AppIcon({ id, logoUrl, size = 28, title }: Props) {
  const [broken, setBroken] = useState(false);
  if (!logoUrl || broken) {
    return <Identicon id={id} size={size} title={title} />;
  }
  return (
    <img
      src={logoUrl}
      alt={title ?? ''}
      title={title}
      width={size}
      height={size}
      loading="lazy"
      onError={() => setBroken(true)}
      style={{
        width: size,
        height: size,
        borderRadius: 4,
        objectFit: 'contain',
        background: 'var(--bg-hover)',
        flexShrink: 0,
        display: 'block',
      }}
    />
  );
}
