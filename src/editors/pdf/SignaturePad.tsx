import { useEffect, useRef, useState } from 'react';
import { Eraser, PenLine, Type, Upload, X } from 'lucide-react';
import { useI18n } from '../../i18n';

const STORE = 'paperline.signatures';

export function savedSignatures(): string[] {
  try {
    return JSON.parse(localStorage.getItem(STORE) ?? '[]');
  } catch {
    return [];
  }
}
function remember(src: string) {
  try {
    const all = [src, ...savedSignatures().filter((s) => s !== src)].slice(0, 4);
    localStorage.setItem(STORE, JSON.stringify(all));
  } catch {
    /* storage unavailable */
  }
}

/** Crop transparent edges so the signature is tight. */
function trim(canvas: HTMLCanvasElement): string {
  const ctx = canvas.getContext('2d')!;
  const { width: w, height: h } = canvas;
  const data = ctx.getImageData(0, 0, w, h).data;
  let x1 = w, y1 = h, x2 = 0, y2 = 0;
  for (let y = 0; y < h; y++)
    for (let x = 0; x < w; x++)
      if (data[(y * w + x) * 4 + 3] > 8) {
        if (x < x1) x1 = x;
        if (x > x2) x2 = x;
        if (y < y1) y1 = y;
        if (y > y2) y2 = y;
      }
  if (x2 <= x1 || y2 <= y1) return '';
  const pad = 6;
  const out = document.createElement('canvas');
  out.width = x2 - x1 + pad * 2;
  out.height = y2 - y1 + pad * 2;
  out.getContext('2d')!.drawImage(canvas, x1 - pad, y1 - pad, out.width, out.height, 0, 0, out.width, out.height);
  return out.toDataURL('image/png');
}

const INKS = ['#111827', '#1d4ed8', '#7c3aed'];
// Handwriting faces (Latin and Arabic), with system script fonts as fallbacks.
const SCRIPT_FONTS = [
  '"Dancing Script", "Segoe Script", "Snell Roundhand", cursive',
  '"Great Vibes", "Brush Script MT", cursive',
  '"Aref Ruqaa", "Diwani Letter", "IBM Plex Sans Arabic", serif',
];

export function SignaturePad({ onDone, onClose }: { onDone: (src: string) => void; onClose: () => void }) {
  const { t } = useI18n();
  const [mode, setMode] = useState<'draw' | 'type' | 'saved'>(savedSignatures().length ? 'saved' : 'draw');
  const [ink, setInk] = useState(INKS[0]);
  const [name, setName] = useState('');
  const [font, setFont] = useState(0);
  const canvas = useRef<HTMLCanvasElement>(null);
  const drawing = useRef<{ x: number; y: number } | null>(null);
  const [empty, setEmpty] = useState(true);

  useEffect(() => {
    const c = canvas.current;
    if (!c) return;
    const dpr = window.devicePixelRatio || 1;
    c.width = c.clientWidth * dpr;
    c.height = c.clientHeight * dpr;
    c.getContext('2d')!.scale(dpr, dpr);
  }, [mode]);

  const pos = (e: React.PointerEvent) => {
    const r = canvas.current!.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  };

  const finishTyped = async () => {
    if (!name.trim()) return;
    const f = `64px ${SCRIPT_FONTS[font]}`;
    await document.fonts?.load(f, name).catch(() => null);
    const c = document.createElement('canvas');
    const ctx = c.getContext('2d')!;
    ctx.font = f;
    c.width = Math.ceil(ctx.measureText(name).width + 60);
    c.height = 110;
    ctx.font = f;
    ctx.fillStyle = ink;
    ctx.textBaseline = 'middle';
    ctx.direction = 'ltr';
    ctx.textAlign = 'left';
    ctx.fillText(name, 30, 58);
    const src = trim(c);
    if (src) { remember(src); onDone(src); }
  };

  const finishDrawn = () => {
    const src = trim(canvas.current!);
    if (src) { remember(src); onDone(src); }
  };

  const upload = (f: File) => {
    const r = new FileReader();
    r.onload = () => onDone(String(r.result));
    r.readAsDataURL(f);
  };

  return (
    <div className="modal-back" onPointerDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal sig-modal" role="dialog" aria-label={t('signature')}>
        <header>
          <h3>{t('signature')}</h3>
          <button className="icon-btn" onClick={onClose} aria-label={t('close')}><X size={18} /></button>
        </header>
        <div className="seg">
          {savedSignatures().length > 0 && <button className={mode === 'saved' ? 'on' : ''} onClick={() => setMode('saved')}>{t('saved')}</button>}
          <button className={mode === 'draw' ? 'on' : ''} onClick={() => setMode('draw')}><PenLine size={14} /> {t('sigDraw')}</button>
          <button className={mode === 'type' ? 'on' : ''} onClick={() => setMode('type')}><Type size={14} /> {t('sigType')}</button>
          <label className="seg-file"><Upload size={14} /> {t('sigUpload')}<input type="file" accept="image/*" hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) upload(f); }} /></label>
        </div>
        {mode === 'saved' && (
          <div className="sig-saved">
            {savedSignatures().map((s) => (
              <button key={s} onClick={() => onDone(s)}><img src={s} alt="" /></button>
            ))}
          </div>
        )}
        {mode === 'draw' && (
          <>
            <canvas
              ref={canvas}
              className="sig-canvas"
              onPointerDown={(e) => { e.currentTarget.setPointerCapture(e.pointerId); drawing.current = pos(e); }}
              onPointerMove={(e) => {
                if (!drawing.current) return;
                const p = pos(e);
                const ctx = canvas.current!.getContext('2d')!;
                ctx.strokeStyle = ink;
                ctx.lineWidth = 2.6;
                ctx.lineCap = 'round';
                ctx.lineJoin = 'round';
                ctx.beginPath();
                ctx.moveTo(drawing.current.x, drawing.current.y);
                ctx.lineTo(p.x, p.y);
                ctx.stroke();
                drawing.current = p;
                setEmpty(false);
              }}
              onPointerUp={() => (drawing.current = null)}
            />
            <div className="sig-row">
              {INKS.map((c) => <button key={c} className={`ink-dot ${ink === c ? 'on' : ''}`} style={{ background: c }} onClick={() => setInk(c)} aria-label={c} />)}
              <span className="grow" />
              <button className="ghost-btn" onClick={() => { const c = canvas.current!; c.getContext('2d')!.clearRect(0, 0, c.width, c.height); setEmpty(true); }}><Eraser size={14} /> {t('clear')}</button>
              <button className="primary-btn" disabled={empty} onClick={finishDrawn}>{t('insert')}</button>
            </div>
          </>
        )}
        {mode === 'type' && (
          <>
            <input className="sig-name" dir="auto" autoFocus placeholder={t('sigName')} value={name} onChange={(e) => setName(e.target.value)} />
            <div className="sig-fonts">
              {SCRIPT_FONTS.map((f, i) => (
                <button key={f} className={font === i ? 'on' : ''} style={{ fontFamily: f, color: ink }} onClick={() => setFont(i)}>{name || 'Signature'}</button>
              ))}
            </div>
            <div className="sig-row">
              {INKS.map((c) => <button key={c} className={`ink-dot ${ink === c ? 'on' : ''}`} style={{ background: c }} onClick={() => setInk(c)} aria-label={c} />)}
              <span className="grow" />
              <button className="primary-btn" disabled={!name.trim()} onClick={finishTyped}>{t('insert')}</button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
