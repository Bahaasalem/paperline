import { useId } from 'react';
import type { DocKind } from '../types';

/**
 * The Paperline mark: a sheet of paper with a folded corner, two lines of text,
 * and a signature line that runs off the page. Paper + line.
 */
export function Logo({ size = 28 }: { size?: number }) {
  const id = useId();
  return (
    <svg width={size} height={size} viewBox="0 0 64 64" aria-hidden className="logo">
      <defs>
        <linearGradient id={`${id}g`} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stopColor="#5b4cf0" />
          <stop offset="1" stopColor="#0ea5a4" />
        </linearGradient>
      </defs>
      <rect x="2" y="2" width="60" height="60" rx="17" fill={`url(#${id}g)`} />
      <path d="M21 11h17l11 11v29a4 4 0 0 1-4 4H21a4 4 0 0 1-4-4V15a4 4 0 0 1 4-4z" fill="#fff" />
      <path d="M38 11v7a4 4 0 0 0 4 4h7z" fill="#c7d2fe" />
      <path d="M23 27h12M23 34h18" stroke="#5b4cf0" strokeWidth="4" strokeLinecap="round" />
      <path d="M23 44h33" stroke="#f5a524" strokeWidth="5" strokeLinecap="round" />
    </svg>
  );
}

/** Logo plus the product name. */
export function Wordmark({ size = 28 }: { size?: number }) {
  return (
    <span className="wordmark" style={{ fontSize: size * 0.68 }}>
      <Logo size={size} />
      <span>
        paper<b>line</b>
      </span>
    </span>
  );
}

const KIND_COLOR: Record<DocKind, [string, string]> = {
  doc: ['#4f46e5', '#7c74ff'],
  sheet: ['#0e9f6e', '#34c38f'],
  slides: ['#e0436b', '#ff7a95'],
  pdf: ['#dc6a1f', '#f59e4b'],
};

/** One icon family for the apps: a colored tile, a folded sheet, and the app's own mark. */
export function AppIcon({ kind, size = 56 }: { kind: DocKind; size?: number }) {
  const id = useId();
  const [c1, c2] = KIND_COLOR[kind];
  return (
    <svg width={size} height={size} viewBox="0 0 64 64" aria-hidden className="app-icon">
      <defs>
        <linearGradient id={`${id}k`} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stopColor={c2} />
          <stop offset="1" stopColor={c1} />
        </linearGradient>
      </defs>
      <rect x="2" y="2" width="60" height="60" rx="17" fill={`url(#${id}k)`} />
      <path d="M20 12h17l10 10v28a4 4 0 0 1-4 4H20a4 4 0 0 1-4-4V16a4 4 0 0 1 4-4z" fill="#fff" />
      <path d="M37 12v6a4 4 0 0 0 4 4h6z" fill="#fff" opacity=".55" />
      {kind === 'doc' && <path d="M21.5 27h15M21.5 33h20M21.5 39h20M21.5 45h11" stroke={c1} strokeWidth="3.2" strokeLinecap="round" />}
      {kind === 'sheet' && (
        <g fill="none" stroke={c1} strokeWidth="2.6">
          <rect x="21" y="26" width="21" height="21" rx="2.5" />
          <path d="M21 33h21M21 40h21M28 26v21M35 26v21" />
        </g>
      )}
      {kind === 'slides' && (
        <g>
          <rect x="21" y="26" width="21" height="15" rx="2.5" fill="none" stroke={c1} strokeWidth="2.6" />
          <path d="M26 37v-4M31 37v-7M36 37v-5" stroke={c1} strokeWidth="2.6" strokeLinecap="round" />
          <path d="M31.5 41v5M27 47h9" stroke={c1} strokeWidth="2.6" strokeLinecap="round" />
        </g>
      )}
      {kind === 'pdf' && (
        <g>
          <path d="M21.5 27h14M21.5 32.5h18" stroke={c1} strokeWidth="3" strokeLinecap="round" opacity=".45" />
          <path d="M21 45c3-6 5-7 6-4s1 4 3 0 3-4 4-1 2 3 7-1" fill="none" stroke={c1} strokeWidth="2.8" strokeLinecap="round" strokeLinejoin="round" />
        </g>
      )}
    </svg>
  );
}
