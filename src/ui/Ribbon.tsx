import { useEffect, useRef, useState, type ReactNode, useLayoutEffect, type CSSProperties } from 'react';
import { ChevronDown } from 'lucide-react';

export interface RibbonTab {
  id: string;
  label: string;
  content: ReactNode;
}

export function Ribbon({ tabs, extra }: { tabs: RibbonTab[]; extra?: ReactNode }) {
  const [active, setActive] = useState(tabs[0]?.id);
  const current = tabs.find((t) => t.id === active) ?? tabs[0];
  return (
    <div className="ribbon">
      <div className="ribbon-tabs" role="tablist">
        {tabs.map((t) => (
          <button key={t.id} role="tab" aria-selected={t.id === current.id} className={`ribbon-tab ${t.id === current.id ? 'active' : ''}`} onClick={() => setActive(t.id)}>
            {t.label}
          </button>
        ))}
        <div className="ribbon-extra">{extra}</div>
      </div>
      <div className="ribbon-body" role="toolbar">
        {current.content}
      </div>
    </div>
  );
}

export function Group({ label, children }: { label?: string; children: ReactNode }) {
  return (
    <div className="r-group">
      <div className="r-group-items">{children}</div>
      {label && <div className="r-group-label">{label}</div>}
    </div>
  );
}

interface BtnProps {
  icon: ReactNode;
  label: string;
  onClick?: () => void;
  active?: boolean;
  disabled?: boolean;
  big?: boolean;
  showLabel?: boolean;
}

/** Keeps the editor's selection: we prevent mousedown focus theft. */
export function Btn({ icon, label, onClick, active, disabled, big, showLabel }: BtnProps) {
  return (
    <button
      type="button"
      className={`r-btn ${big ? 'big' : ''} ${active ? 'on' : ''}`}
      title={label}
      aria-label={label}
      aria-pressed={active}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {icon}
      {(big || showLabel) && <span>{label}</span>}
    </button>
  );
}

export function Select<T extends string | number>({
  value,
  options,
  onChange,
  label,
  width,
}: {
  value: T;
  options: { value: T; label: string }[];
  onChange: (v: T) => void;
  label: string;
  width?: number;
}) {
  return (
    <select
      className="r-select"
      title={label}
      aria-label={label}
      style={{ width }}
      value={String(value)}
      onChange={(e) => {
        const o = options.find((x) => String(x.value) === e.target.value);
        if (o) onChange(o.value);
      }}
    >
      {!options.some((o) => String(o.value) === String(value)) && <option value={String(value)}>{String(value)}</option>}
      {options.map((o) => (
        <option key={String(o.value)} value={String(o.value)}>
          {o.label}
        </option>
      ))}
    </select>
  );
}

const PALETTE = [
  '#000000', '#434343', '#666666', '#999999', '#cccccc', '#ffffff',
  '#b91c1c', '#ea580c', '#ca8a04', '#16a34a', '#0d9488', '#2563eb',
  '#4f46e5', '#9333ea', '#db2777', '#fecaca', '#fed7aa', '#fef08a',
  '#bbf7d0', '#99f6e4', '#bfdbfe', '#c7d2fe', '#e9d5ff', '#fbcfe8',
];

/**
 * Popups open under their trigger with fixed positioning, so the ribbon's
 * horizontal scrolling never clips them. Flips to the other edge near the window side.
 */
function usePopStyle(ref: React.RefObject<HTMLDivElement | null>, open: boolean, align: 'start' | 'end' = 'start', width = 260): CSSProperties | undefined {
  const [st, setSt] = useState<CSSProperties>();
  useLayoutEffect(() => {
    if (!open || !ref.current) return;
    const place = () => {
      const r = ref.current!.getBoundingClientRect();
      const rtl = getComputedStyle(ref.current!).direction === 'rtl';
      const fromRight = (align === 'end') !== rtl;
      let left = fromRight ? r.right - width : r.left;
      left = Math.max(8, Math.min(left, window.innerWidth - width - 8));
      setSt({ position: 'fixed', top: r.bottom + 6, left, insetInlineStart: 'auto', insetInlineEnd: 'auto', maxHeight: window.innerHeight - r.bottom - 16, overflowY: 'auto' });
    };
    place();
    window.addEventListener('resize', place);
    return () => window.removeEventListener('resize', place);
  }, [open, ref, align, width]);
  return open ? st : undefined;
}

export function ColorBtn({ icon, label, color, onPick, allowNone }: { icon: ReactNode; label: string; color?: string; onPick: (c: string | null) => void; allowNone?: boolean }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const popStyle = usePopStyle(ref, open, 'start', 196);
  useEffect(() => {
    if (!open) return;
    const h = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener('mousedown', h);
    return () => document.removeEventListener('mousedown', h);
  }, [open]);
  return (
    <div className="color-btn" ref={ref}>
      <button type="button" className="r-btn" title={label} aria-label={label} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen((o) => !o)}>
        <span className="color-icon">
          {icon}
          <i style={{ background: color ?? 'transparent' }} />
        </span>
        <ChevronDown size={12} />
      </button>
      {open && (
        <div className="color-pop" style={popStyle} onMouseDown={(e) => e.preventDefault()}>
          <div className="swatches">
            {PALETTE.map((c) => (
              <button
                key={c}
                type="button"
                className="swatch"
                style={{ background: c }}
                title={c}
                onClick={() => {
                  onPick(c);
                  setOpen(false);
                }}
              />
            ))}
          </div>
          <div className="color-pop-row">
            {allowNone && (
              <button
                type="button"
                className="link-btn"
                onClick={() => {
                  onPick(null);
                  setOpen(false);
                }}
              >
                ∅
              </button>
            )}
            <input
              type="color"
              aria-label={label}
              onChange={(e) => {
                onPick(e.target.value);
              }}
            />
          </div>
        </div>
      )}
    </div>
  );
}

export function Menu({ trigger, children, align = 'start' }: { trigger: (open: boolean, toggle: () => void) => ReactNode; children: (close: () => void) => ReactNode; align?: 'start' | 'end' }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const popStyle = usePopStyle(ref, open, align);
  useEffect(() => {
    if (!open) return;
    const h = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) setOpen(false);
    };
    const k = (e: KeyboardEvent) => e.key === 'Escape' && setOpen(false);
    document.addEventListener('mousedown', h);
    document.addEventListener('keydown', k);
    return () => {
      document.removeEventListener('mousedown', h);
      document.removeEventListener('keydown', k);
    };
  }, [open]);
  return (
    <div className="menu-wrap" ref={ref}>
      {trigger(open, () => setOpen((o) => !o))}
      {open && <div className={`menu-pop ${align}`} style={popStyle}>{children(() => setOpen(false))}</div>}
    </div>
  );
}

export function MenuItem({ icon, label, onClick, hint }: { icon?: ReactNode; label: string; onClick: () => void; hint?: string }) {
  return (
    <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={onClick}>
      <span className="menu-icon">{icon}</span>
      <span className="menu-label">{label}</span>
      {hint && <span className="menu-hint">{hint}</span>}
    </button>
  );
}
