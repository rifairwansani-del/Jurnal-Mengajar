import { useMemo, useState } from 'react';
import { getDriveImageCandidates } from '../utils/portalHelpers';

interface PersonAvatarProps {
  src?: string;
  fallbackSrc: string;
  alt: string;
  className?: string;
}

export function PersonAvatar({ src, fallbackSrc, alt, className = '' }: PersonAvatarProps) {
  const candidates = useMemo(() => getDriveImageCandidates(src), [src]);
  const srcKey = src || '';
  const [errorState, setErrorState] = useState<{ key: string; index: number }>({
    key: srcKey,
    index: 0,
  });

  const activeIndex = errorState.key === srcKey ? errorState.index : 0;
  const resolvedSrc = candidates[activeIndex] || fallbackSrc;

  const handleError = () => {
    if (activeIndex < candidates.length) {
      setErrorState({ key: srcKey, index: activeIndex + 1 });
    }
  };

  return (
    <img
      src={resolvedSrc}
      alt={alt}
      referrerPolicy="no-referrer"
      onError={handleError}
      className={className}
    />
  );
}
