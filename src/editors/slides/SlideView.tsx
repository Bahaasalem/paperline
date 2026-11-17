import type { CSSProperties, ReactNode } from 'react';
import { ChartSvg } from '../sheet/Chart';
import type { ShapeEl, ShapeKind, Slide, SlideEl, TableEl, TextEl, TextStyle } from './model';

export function textCss(st: TextStyle): CSSProperties {
  return {
    fontSize: `${st.fontSize}pt`,
    fontFamily: st.fontFace ? `"${st.fontFace}", Inter, "IBM Plex Sans Arabic", sans-serif` : undefined,
    color: st.color,
    fontWeight: st.bold ? 700 : 400,
    fontStyle: st.italic ? 'italic' : undefined,
    textDecoration: [st.underline && 'underline', st.strike && 'line-through'].filter(Boolean).join(' ') || undefined,
    textAlign: st.align,
    lineHeight: st.lineSpacing ? 1.2 * st.lineSpacing : undefined,
    justifyContent: st.valign === 'middle' ? 'center' : st.valign === 'bottom' ? 'flex-end' : 'flex-start',
  };
}

type Pt = [number, number];
const poly = (pts: Pt[]) => `M ${pts.map((p) => `${p[0].toFixed(2)} ${p[1].toFixed(2)}`).join(' L ')} Z`;

/** Fit points (in any space) to the box. */
function fit(pts: Pt[], w: number, h: number): Pt[] {
  const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
  const x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys);
  return pts.map(([x, y]) => [((x - x0) / (x1 - x0 || 1)) * w, ((y - y0) / (y1 - y0 || 1)) * h]);
}
const regular = (n: number, w: number, h: number) =>
  fit(Array.from({ length: n }, (_, i) => { const a = -Math.PI / 2 + (i * 2 * Math.PI) / n; return [Math.cos(a), Math.sin(a)] as Pt; }), w, h);
const star = (n: number, inner: number, w: number, h: number) =>
  fit(Array.from({ length: n * 2 }, (_, i) => { const a = -Math.PI / 2 + (i * Math.PI) / n; const r = i % 2 ? inner : 1; return [r * Math.cos(a), r * Math.sin(a)] as Pt; }), w, h);

/** A right-pointing block arrow in a length × thickness box. */
function arrow(len: number, thick: number): Pt[] {
  const hy = thick * 0.25, hx = len - Math.min(len * 0.4, thick * 0.6);
  return [[0, hy], [hx, hy], [hx, 0], [len, thick / 2], [hx, thick], [hx, thick - hy], [0, thick - hy]];
}

/** SVG path data for a shape in a w × h box. Lines are handled separately. */
export function shapePath(kind: ShapeKind, w: number, h: number): string {
  const m = Math.min(w, h);
  switch (kind) {
    case 'triangle': return poly([[w / 2, 0], [w, h], [0, h]]);
    case 'rtTriangle': return poly([[0, 0], [w, h], [0, h]]);
    case 'diamond': return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'pentagon': return poly(regular(5, w, h));
    case 'hexagon': return poly([[0, h / 2], [w * 0.25, 0], [w * 0.75, 0], [w, h / 2], [w * 0.75, h], [w * 0.25, h]]);
    case 'octagon': { const c = m * 0.29; return poly([[c, 0], [w - c, 0], [w, c], [w, h - c], [w - c, h], [c, h], [0, h - c], [0, c]]); }
    case 'parallelogram': { const o = Math.min(w * 0.25, h * 0.6); return poly([[o, 0], [w, 0], [w - o, h], [0, h]]); }
    case 'trapezoid': { const o = Math.min(w * 0.25, h * 0.6); return poly([[o, 0], [w - o, 0], [w, h], [0, h]]); }
    case 'star4': return poly(star(4, 0.38, w, h));
    case 'star5': return poly(star(5, 0.382, w, h));
    case 'star8': return poly(star(8, 0.62, w, h));
    case 'plus': return poly([[w * 0.34, 0], [w * 0.66, 0], [w * 0.66, h * 0.34], [w, h * 0.34], [w, h * 0.66], [w * 0.66, h * 0.66], [w * 0.66, h], [w * 0.34, h], [w * 0.34, h * 0.66], [0, h * 0.66], [0, h * 0.34], [w * 0.34, h * 0.34]]);
    case 'heart':
      return `M ${w / 2} ${h} C ${w * 0.18} ${h * 0.78} 0 ${h * 0.55} 0 ${h * 0.3} C 0 ${h * 0.1} ${w * 0.16} 0 ${w * 0.28} 0 C ${w * 0.4} 0 ${w * 0.47} ${h * 0.08} ${w / 2} ${h * 0.18} C ${w * 0.53} ${h * 0.08} ${w * 0.6} 0 ${w * 0.72} 0 C ${w * 0.84} 0 ${w} ${h * 0.1} ${w} ${h * 0.3} C ${w} ${h * 0.55} ${w * 0.82} ${h * 0.78} ${w / 2} ${h} Z`;
    case 'callout': {
      const b = h * 0.78, r = Math.min(w, b) * 0.14;
      return `M ${r} 0 H ${w - r} Q ${w} 0 ${w} ${r} V ${b - r} Q ${w} ${b} ${w - r} ${b} H ${w * 0.42} L ${w * 0.2} ${h} L ${w * 0.26} ${b} H ${r} Q 0 ${b} 0 ${b - r} V ${r} Q 0 0 ${r} 0 Z`;
    }
    case 'cloud': {
      const n = 10;
      const pts: Pt[] = Array.from({ length: n }, (_, i) => { const a = (i * 2 * Math.PI) / n; return [w / 2 + w * 0.4 * Math.cos(a), h / 2 + h * 0.36 * Math.sin(a)]; });
      let d = `M ${pts[0][0]} ${pts[0][1]}`;
      for (let i = 0; i < n; i++) {
        const [x, y] = pts[(i + 1) % n];
        const [px, py] = pts[i];
        const r = Math.hypot(x - px, y - py) * 0.62;
        d += ` A ${r} ${r} 0 0 1 ${x} ${y}`;
      }
      return d + ' Z';
    }
    case 'rightArrow': return poly(arrow(w, h));
    case 'leftArrow': return poly(arrow(w, h).map(([x, y]) => [w - x, y]));
    case 'upArrow': return poly(arrow(h, w).map(([x, y]) => [y, h - x]));
    case 'downArrow': return poly(arrow(h, w).map(([x, y]) => [y, x]));
    case 'leftRightArrow': {
      const hx = Math.min(w * 0.3, h * 0.6), hy = h * 0.25;
      return poly([[0, h / 2], [hx, 0], [hx, hy], [w - hx, hy], [w - hx, 0], [w, h / 2], [w - hx, h], [w - hx, h - hy], [hx, h - hy], [hx, h]]);
    }
    case 'chevron': { const d = Math.min(w * 0.5, h * 0.5); return poly([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h], [d, h / 2]]); }
    case 'homePlate': { const d = Math.min(w * 0.5, h * 0.5); return poly([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h]]); }
    default: return poly([[0, 0], [w, 0], [w, h], [0, h]]);
  }
}

const dashArray = (dash: ShapeEl['dash'], sw: number) => (dash === 'dash' ? `${sw * 3} ${sw * 2}` : dash === 'dot' ? `${sw * 0.1} ${sw * 2}` : undefined);

export function ShapeSvg({ el }: { el: Pick<ShapeEl, 'shape' | 'w' | 'h' | 'fill' | 'stroke' | 'strokeW' | 'dash' | 'flipH' | 'flipV' | 'id'> }) {
  const { w, h } = el;
  const isLine = el.shape === 'line' || el.shape === 'arrowLine';
  const sw = el.strokeW ?? (el.stroke ? 2 : 0);
  const lineW = Math.max(1, el.strokeW || 2);
  const p = {
    fill: isLine ? 'none' : el.fill,
    stroke: isLine ? el.stroke ?? el.fill : el.stroke ?? 'none',
    strokeWidth: isLine ? lineW : sw,
    strokeDasharray: dashArray(el.dash, isLine ? lineW : sw),
    strokeLinecap: el.dash === 'dot' ? ('round' as const) : undefined,
    strokeLinejoin: 'round' as const,
  };
  const flip = el.flipH || el.flipV ? { transform: `scale(${el.flipH ? -1 : 1}, ${el.flipV ? -1 : 1})` } : undefined;
  let shape: ReactNode;
  const markerId = `ah-${el.id}`;
  if (isLine) {
    shape = (
      <>
        {el.shape === 'arrowLine' && (
          <defs>
            <marker id={markerId} viewBox="0 0 10 10" refX="8" refY="5" markerWidth={Math.max(3, 12 / lineW)} markerHeight={Math.max(3, 12 / lineW)} markerUnits="strokeWidth" orient="auto-start-reverse">
              <path d="M 0 0 L 10 5 L 0 10 z" fill={p.stroke} />
            </marker>
          </defs>
        )}
        <line x1={0} y1={h / 2} x2={w} y2={h / 2} {...p} markerEnd={el.shape === 'arrowLine' ? `url(#${markerId})` : undefined} />
      </>
    );
  } else if (el.shape === 'ellipse') {
    shape = <ellipse cx={w / 2} cy={h / 2} rx={Math.max(0, w / 2 - sw / 2)} ry={Math.max(0, h / 2 - sw / 2)} {...p} />;
  } else if (el.shape === 'roundRect') {
    shape = <rect x={sw / 2} y={sw / 2} width={Math.max(0, w - sw)} height={Math.max(0, h - sw)} rx={Math.min(w, h) * 0.16} {...p} />;
  } else {
    // inset by half the stroke so outlines stay inside the box
    const i = sw / 2;
    shape = <path d={shapePath(el.shape, Math.max(1, w - sw), Math.max(1, h - sw))} transform={i ? `translate(${i} ${i})` : undefined} {...p} />;
  }
  return (
    <svg className="shape-svg" width={w} height={h} viewBox={`0 0 ${w} ${h}`} preserveAspectRatio="none" style={flip}>
      {shape}
    </svg>
  );
}

export function TextBody({ el, placeholder }: { el: TextEl | ShapeEl; placeholder?: string }) {
  const lines = el.text.split('\n');
  const bullets = el.type === 'text' && (el.bullets || el.numbered);
  if (!el.text) return placeholder ? <div className="slide-ph">{placeholder}</div> : null;
  if (bullets) {
    const items = lines.map((l, i) => (
      <li key={i} className={l.startsWith('  ') ? 'lvl2' : ''}>
        {l.trimStart() || ' '}
      </li>
    ));
    return el.numbered ? <ol className="slide-bullets numbered">{items}</ol> : <ul className="slide-bullets">{items}</ul>;
  }
  return <div className="slide-text-body">{el.text}</div>;
}

const tint = (hex: string, t = 0.86) => {
  const n = parseInt(hex.replace('#', '').padEnd(6, '0').slice(0, 6), 16);
  const m = (s: number) => Math.round(((n >> s) & 255) + (255 - ((n >> s) & 255)) * t);
  return `#${((m(16) << 16) | (m(8) << 8) | m(0)).toString(16).padStart(6, '0')}`;
};

export function TableView({ el, cell }: { el: TableEl; cell?: (r: number, c: number, text: string) => ReactNode }) {
  const band = tint(el.color);
  return (
    <table className="slide-table" style={{ ...textCss(el.style), borderColor: tint(el.color, 0.55) }} dir="auto">
      <tbody>
        {el.rows.map((row, r) => {
          const head = el.header && r === 0;
          const striped = el.banded && !head && (r - (el.header ? 1 : 0)) % 2 === 0;
          return (
            <tr key={r} style={{ height: `${100 / el.rows.length}%` }}>
              {row.map((text, c) => (
                <td key={c} style={{ background: head ? el.color : striped ? band : '#ffffff', color: head ? '#ffffff' : undefined, fontWeight: head ? 700 : undefined, borderColor: tint(el.color, 0.55) }}>
                  {cell ? cell(r, c, text) : text}
                </td>
              ))}
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

export function ElView({ el, placeholder, children, cell }: { el: SlideEl; placeholder?: string; children?: ReactNode; cell?: (r: number, c: number, text: string) => ReactNode }) {
  const base: CSSProperties = {
    left: el.x,
    top: el.y,
    width: el.w,
    height: el.h,
    transform: el.rot ? `rotate(${el.rot}deg)` : undefined,
    opacity: el.opacity !== undefined && el.opacity < 1 ? el.opacity : undefined,
    filter: el.shadow ? 'drop-shadow(0 6px 10px rgba(15, 23, 42, 0.28))' : undefined,
  };
  if (el.type === 'image') {
    const flip = el.flipH || el.flipV ? `scale(${el.flipH ? -1 : 1}, ${el.flipV ? -1 : 1})` : undefined;
    return (
      <div className="slide-el image" style={base}>
        <img src={el.src} alt="" draggable={false} style={{ transform: flip }} />
        {children}
      </div>
    );
  }
  if (el.type === 'table') {
    return (
      <div className="slide-el table" style={base}>
        <TableView el={el} cell={cell} />
        {children}
      </div>
    );
  }
  if (el.type === 'chart') {
    return (
      <div className="slide-el chart" style={base}>
        <ChartSvg type={el.chart} data={el.data} w={el.w} h={el.h} title={el.title} legend={el.legend !== false} />
        {children}
      </div>
    );
  }
  if (el.type === 'shape') {
    return (
      <div className="slide-el shape" style={base}>
        <ShapeSvg el={el} />
        <div className="slide-text" dir="auto" style={textCss(el.style)}>
          {children ?? <TextBody el={el} />}
        </div>
      </div>
    );
  }
  return (
    <div className="slide-el text" style={{ ...base, background: el.fill }}>
      <div className="slide-text" dir="auto" style={textCss(el.style)}>
        {children ?? <TextBody el={el} placeholder={placeholder} />}
      </div>
    </div>
  );
}

const isDark = (c: string) => {
  const h = c.replace('#', '');
  if (h.length < 6) return false;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
  return 0.299 * r + 0.587 * g + 0.114 * b < 128;
};

export interface SlideDeco {
  num?: number;
  footer?: string;
}

/** Slide number and footer, drawn the same in the editor, the show, print and PDF. */
export function Decorations({ slide, deco, w }: { slide: Slide; deco?: SlideDeco; w: number }) {
  if (!deco || (deco.num === undefined && !deco.footer)) return null;
  const color = isDark(slide.bg) ? 'rgba(255,255,255,0.7)' : 'rgba(30,41,59,0.6)';
  return (
    <>
      {deco.footer && <div className="slide-deco footer" dir="auto" style={{ color, width: w * 0.5, left: w * 0.25 }}>{deco.footer}</div>}
      {deco.num !== undefined && <div className="slide-deco num" style={{ color }}>{deco.num}</div>}
    </>
  );
}

/** A slide at its native size; scale it with a wrapper. */
export function SlideView({ slide, w, h, placeholders, deco }: { slide: Slide; w: number; h: number; placeholders?: (el: TextEl) => string | undefined; deco?: SlideDeco }) {
  return (
    <div className="slide-surface" style={{ width: w, height: h, background: slide.bg }}>
      {slide.els.map((el) => (
        <ElView key={el.id} el={el} placeholder={el.type === 'text' ? placeholders?.(el) : undefined} />
      ))}
      <Decorations slide={slide} deco={deco} w={w} />
    </div>
  );
}

export function ScaledSlide({ slide, w, h, width, placeholders, deco }: { slide: Slide; w: number; h: number; width: number; placeholders?: (el: TextEl) => string | undefined; deco?: SlideDeco }) {
  const s = width / w;
  return (
    <div className="scaled-slide" style={{ width, height: h * s }}>
      <div style={{ transform: `scale(${s})`, transformOrigin: '0 0', width: w, height: h }}>
        <SlideView slide={slide} w={w} h={h} placeholders={placeholders} deco={deco} />
      </div>
    </div>
  );
}

/** Slide number and footer for slide i of a deck. */
export const decoFor = (deck: { slideNumbers?: boolean; footer?: string; slides: Slide[] }, i: number): SlideDeco | undefined =>
  deck.slideNumbers || deck.footer ? { num: deck.slideNumbers ? i + 1 : undefined, footer: deck.footer } : undefined;
