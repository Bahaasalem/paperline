export function Logo({ size = 28 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 64 64" aria-hidden>
      <defs>
        <linearGradient id="pl-g" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stopColor="#6366f1" />
          <stop offset="1" stopColor="#14b8a6" />
        </linearGradient>
      </defs>
      <rect x="4" y="4" width="56" height="56" rx="16" fill="url(#pl-g)" />
      <path d="M20 18h17a9 9 0 0 1 0 18H26v10" fill="none" stroke="#fff" strokeWidth="6" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M20 46h24" fill="none" stroke="#fff" strokeOpacity=".55" strokeWidth="6" strokeLinecap="round" />
    </svg>
  );
}
