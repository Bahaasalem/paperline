import { useEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import { Check, ChevronDown, MonitorDown, Search } from 'lucide-react';
import { useI18n } from '../i18n';
import { BUNDLED_FONTS, canListAllFonts, familyCss, listAllFonts, rememberFont, useDeviceFonts } from '../fonts';
import { usePopStyle } from './Ribbon';

interface Extra {
  value: string;
  label: string;
  css: string;
}

/**
 * Font menu for every editor: Paperline's own fonts, the fonts on this device
 * (each shown in its own typeface), recent picks and a search box.
 */
export function FontPicker({ value, onChange, width = 150, extra, extraLabel }: { value: string; onChange: (f: string) => void; width?: number; extra?: Extra[]; extraLabel?: string }) {
  const { t } = useI18n();
  const { device, complete, recent } = useDeviceFonts();
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState('');
  const [asking, setAsking] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const pop = usePopStyle(ref, open, 'start', 280);

  useEffect(() => {
    if (!open) return;
    const h = (e: MouseEvent) => { if (!ref.current?.contains(e.target as Node)) setOpen(false); };
    const k = (e: KeyboardEvent) => e.key === 'Escape' && setOpen(false);
    document.addEventListener('mousedown', h);
    document.addEventListener('keydown', k);
    return () => { document.removeEventListener('mousedown', h); document.removeEventListener('keydown', k); };
  }, [open]);

  const extraValues = new Set(extra?.map((e) => e.value));
  const shown = extra?.find((e) => e.value === value);
  const match = (f: string) => !q || f.toLowerCase().includes(q.trim().toLowerCase());
  const groups = useMemo(() => {
    const own = BUNDLED_FONTS.filter(match);
    const dev = device.filter((f) => !BUNDLED_FONTS.includes(f)).filter(match);
    const rec = q ? [] : recent.filter((f) => !extraValues.has(f));
    return { own, dev, rec };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [device, recent, q]);

  const pick = (f: string) => {
    if (!extraValues.has(f)) rememberFont(f);
    onChange(f);
    setOpen(false);
    setQ('');
  };

  const item = (f: string, label = f, css = familyCss(f)) => (
    <button key={f + label} type="button" role="option" aria-selected={f === value} className={`font-item ${f === value ? 'on' : ''}`} onMouseDown={(e) => e.preventDefault()} onClick={() => pick(f)}>
      <span style={{ fontFamily: css } as CSSProperties}>{label}</span>
      {f === value && <Check size={14} />}
    </button>
  );

  return (
    <div className="font-picker" ref={ref}>
      <button type="button" className="font-btn" style={{ width }} title={t('font')} aria-label={t('font')} aria-expanded={open} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen((o) => !o)}>
        <span style={{ fontFamily: shown ? shown.css : familyCss(value) }}>{shown ? shown.label : value}</span>
        <ChevronDown size={14} />
      </button>
      {open && (
        <div className="font-pop" style={pop} role="listbox">
          <label className="font-search">
            <Search size={14} />
            <input autoFocus dir="auto" value={q} placeholder={t('searchFonts')} onChange={(e) => setQ(e.target.value)} onKeyDown={(e) => {
              e.stopPropagation();
              if (e.key === 'Enter') {
                const first = [...groups.own, ...groups.dev][0];
                if (first) pick(first);
                else if (q.trim()) pick(q.trim()); // any family name the user types
              }
            }} />
          </label>
          <div className="font-list">
            {!!extra?.length && !q && (
              <>
                <div className="font-group">{extraLabel}</div>
                {extra.map((e) => item(e.value, e.label, e.css))}
              </>
            )}
            {groups.rec.length > 0 && (
              <>
                <div className="font-group">{t('recentFonts')}</div>
                {groups.rec.map((f) => item(f))}
              </>
            )}
            {groups.own.length > 0 && (
              <>
                <div className="font-group">{t('appFonts')}</div>
                {groups.own.map((f) => item(f))}
              </>
            )}
            <div className="font-group">{t('deviceFonts')} {device.length > 0 && <span>· {groups.dev.length}</span>}</div>
            {groups.dev.map((f) => item(f))}
            {q && !groups.own.length && !groups.dev.length && item(q.trim(), `${t('useFont')} “${q.trim()}”`)}
          </div>
          {canListAllFonts() && !complete && (
            <button type="button" className="font-all" disabled={asking} onClick={async () => { setAsking(true); await listAllFonts(); setAsking(false); }}>
              <MonitorDown size={15} /> {t('showAllFonts')}
            </button>
          )}
        </div>
      )}
    </div>
  );
}
