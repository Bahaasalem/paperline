import type { CSSProperties, ReactNode } from 'react';
import type { ShapeEl, Slide, SlideEl, TextEl, TextStyle } from './model';

export function textCss(st: TextStyle): CSSProperties {
  return {
    fontSize: `${st.fontSize}pt`,
    fontFamily: st.fontFace ? `"${st.fontFace}", Inter, "IBM Plex Sans Arabic", sans-serif` : undefined,
    color: st.color,
    fontWeight: st.bold ? 700 : 400,
    fontStyle: st.italic ? 'italic' : undefined,
    textDecoration: st.underline ? 'underline' : undefined,
    textAlign: st.align,
    justifyContent: st.valign === 'middle' ? 'center' : st.valign === 'bottom' ? 'flex-end' : 'flex-start',
  };
}

export function ShapeSvg({ el }: { el: ShapeEl }) {
  const { w, h } = el;
  const sw = el.strokeW ?? (el.stroke ? 2 : 0);
  const p = { fill: el.shape === 'line' ? 'none' : el.fill, stroke: el.shape === 'line' ? el.stroke ?? el.fill : el.stroke ?? 'none', strokeWidth: el.shape === 'line' ? Math.max(2, sw) : sw };
  let shape: ReactNode;
  switch (el.shape) {
    case 'ellipse':
      shape = <ellipse cx={w / 2} cy={h / 2} rx={Math.max(0, w / 2 - sw / 2)} ry={Math.max(0, h / 2 - sw / 2)} {...p} />;
      break;
    case 'roundRect':
      shape = <rect x={sw / 2} y={sw / 2} width={Math.max(0, w - sw)} height={Math.max(0, h - sw)} rx={Math.min(w, h) * 0.16} {...p} />;
      break;
    case 'triangle':
      shape = <polygon points={`${w / 2},${sw} ${w - sw},${h - sw} ${sw},${h - sw}`} {...p} />;
      break;
    case 'rightArrow': {
      const hy = h * 0.25, hx = w - Math.min(w * 0.4, h * 0.6);
      shape = <polygon points={`0,${hy} ${hx},${hy} ${hx},0 ${w},${h / 2} ${hx},${h} ${hx},${h - hy} 0,${h - hy}`} {...p} />;
      break;
    }
    case 'line':
      shape = <line x1={0} y1={h / 2} x2={w} y2={h / 2} {...p} />;
      break;
    default:
      shape = <rect x={sw / 2} y={sw / 2} width={Math.max(0, w - sw)} height={Math.max(0, h - sw)} {...p} />;
  }
  return (
    <svg className="shape-svg" width={w} height={h} viewBox={`0 0 ${w} ${h}`} preserveAspectRatio="none">
      {shape}
    </svg>
  );
}

export function TextBody({ el, placeholder }: { el: TextEl | ShapeEl; placeholder?: string }) {
  const lines = el.text.split('\n');
  const bullets = el.type === 'text' && el.bullets;
  if (!el.text) return placeholder ? <div className="slide-ph">{placeholder}</div> : null;
  if (bullets)
    return (
      <ul className="slide-bullets">
        {lines.map((l, i) => (
          <li key={i} className={l.startsWith('  ') ? 'lvl2' : ''}>
            {l.trimStart() || ' '}
          </li>
        ))}
      </ul>
    );
  return <div className="slide-text-body">{el.text}</div>;
}

export function ElView({ el, placeholder, children }: { el: SlideEl; placeholder?: string; children?: ReactNode }) {
  const base: CSSProperties = { left: el.x, top: el.y, width: el.w, height: el.h, transform: el.rot ? `rotate(${el.rot}deg)` : undefined };
  if (el.type === 'image') {
    return (
      <div className="slide-el image" style={base}>
        <img src={el.src} alt="" draggable={false} />
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

/** A slide at its native size; scale it with a wrapper. */
export function SlideView({ slide, w, h, placeholders }: { slide: Slide; w: number; h: number; placeholders?: (el: TextEl) => string | undefined }) {
  return (
    <div className="slide-surface" style={{ width: w, height: h, background: slide.bg }}>
      {slide.els.map((el) => (
        <ElView key={el.id} el={el} placeholder={el.type === 'text' ? placeholders?.(el) : undefined} />
      ))}
    </div>
  );
}

export function ScaledSlide({ slide, w, h, width, placeholders }: { slide: Slide; w: number; h: number; width: number; placeholders?: (el: TextEl) => string | undefined }) {
  const s = width / w;
  return (
    <div className="scaled-slide" style={{ width, height: h * s }}>
      <div style={{ transform: `scale(${s})`, transformOrigin: '0 0', width: w, height: h }}>
        <SlideView slide={slide} w={w} h={h} placeholders={placeholders} />
      </div>
    </div>
  );
}
