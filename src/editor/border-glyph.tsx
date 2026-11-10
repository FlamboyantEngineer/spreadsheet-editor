import type { CSSProperties } from 'react';
/** Dedicated table-border glyphs share the controls' inherited stroke and focus shell. */
export function BorderGlyph({ side }: { side: string }) {
  const style: CSSProperties = { width: 18, height: 18 };
  return (
    <svg
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.5"
      style={style}
      aria-hidden="true"
    >
      <path d="M3 3h18v18H3zM3 12h18M12 3v18" strokeDasharray="1 3" opacity=".3" />
      {side === 'top' ? (
        <path d="M3 3h18" />
      ) : side === 'bottom' ? (
        <path d="M3 21h18" />
      ) : side === 'left' ? (
        <path d="M3 3v18" />
      ) : side === 'right' ? (
        <path d="M21 3v18" />
      ) : side === 'all' ? (
        <path d="M3 3h18v18H3zM3 12h18M12 3v18" />
      ) : (
        <path d="m7 7 10 10M17 7 7 17" />
      )}
    </svg>
  );
}
