export type ShapeKind = 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'rightArrow' | 'line';

export interface TextStyle {
  fontSize: number; // pt
  fontFace?: string;
  color: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  align: 'left' | 'center' | 'right';
  valign: 'top' | 'middle' | 'bottom';
}

export interface BaseEl {
  id: string;
  x: number;
  y: number;
  w: number;
  h: number;
  rot?: number;
}

export interface TextEl extends BaseEl {
  type: 'text';
  text: string;
  style: TextStyle;
  placeholder?: 'title' | 'subtitle' | 'body';
  fill?: string;
  bullets?: boolean;
}

export interface ShapeEl extends BaseEl {
  type: 'shape';
  shape: ShapeKind;
  fill: string;
  stroke?: string;
  strokeW?: number;
  text: string;
  style: TextStyle;
}

export interface ImageEl extends BaseEl {
  type: 'image';
  src: string;
}

export type SlideEl = TextEl | ShapeEl | ImageEl;

export interface Slide {
  id: string;
  bg: string;
  els: SlideEl[];
  notes?: string;
}

export interface DeckTheme {
  id: string;
  bg: string;
  title: string;
  text: string;
  accent: string;
  font: string;
}

export interface Deck {
  w: number;
  h: number;
  theme: string;
  slides: Slide[];
}

export const THEMES: DeckTheme[] = [
  { id: 'paper', bg: '#ffffff', title: '#1e1b4b', text: '#334155', accent: '#6366f1', font: 'Inter' },
  { id: 'mint', bg: '#f0fdfa', title: '#134e4a', text: '#1f2937', accent: '#14b8a6', font: 'Inter' },
  { id: 'sunset', bg: '#fff7ed', title: '#7c2d12', text: '#431407', accent: '#f97316', font: 'Georgia' },
  { id: 'night', bg: '#0f172a', title: '#f8fafc', text: '#cbd5e1', accent: '#38bdf8', font: 'Inter' },
  { id: 'berry', bg: '#fdf2f8', title: '#831843', text: '#3f3f46', accent: '#db2777', font: 'Inter' },
  { id: 'forest', bg: '#14532d', title: '#ecfccb', text: '#d9f99d', accent: '#facc15', font: 'Georgia' },
];

export const themeOf = (d: Deck) => THEMES.find((t) => t.id === d.theme) ?? THEMES[0];

let seq = 0;
export const uid = () => `e${Date.now().toString(36)}${(seq++).toString(36)}${Math.random().toString(36).slice(2, 6)}`;

export const textStyle = (p: Partial<TextStyle> = {}): TextStyle => ({ fontSize: 18, color: '#334155', align: 'left', valign: 'top', ...p });

export type Layout = 'title' | 'content' | 'two' | 'blank';

export function makeSlide(layout: Layout, deck: Pick<Deck, 'w' | 'h' | 'theme'>): Slide {
  const th = THEMES.find((t) => t.id === deck.theme) ?? THEMES[0];
  const { w, h } = deck;
  const m = w * 0.06;
  const els: SlideEl[] = [];
  const title = (y: number, hh: number, size: number, align: TextStyle['align'] = 'left'): TextEl => ({
    id: uid(), type: 'text', placeholder: 'title', text: '', x: m, y, w: w - 2 * m, h: hh,
    style: textStyle({ fontSize: size, bold: true, color: th.title, align, valign: 'middle', fontFace: th.font }),
  });
  if (layout === 'title') {
    els.push(title(h * 0.3, h * 0.22, 40, 'center'));
    els.push({ id: uid(), type: 'text', placeholder: 'subtitle', text: '', x: m * 2, y: h * 0.54, w: w - 4 * m, h: h * 0.14, style: textStyle({ fontSize: 20, color: th.text, align: 'center', fontFace: th.font }) });
  } else if (layout === 'content') {
    els.push(title(m * 0.8, h * 0.16, 32));
    els.push({ id: uid(), type: 'text', placeholder: 'body', bullets: true, text: '', x: m, y: h * 0.27, w: w - 2 * m, h: h * 0.64, style: textStyle({ fontSize: 20, color: th.text, fontFace: th.font }) });
  } else if (layout === 'two') {
    els.push(title(m * 0.8, h * 0.16, 32));
    const cw = (w - 3 * m) / 2;
    for (let i = 0; i < 2; i++)
      els.push({ id: uid(), type: 'text', placeholder: 'body', bullets: true, text: '', x: m + i * (cw + m), y: h * 0.27, w: cw, h: h * 0.64, style: textStyle({ fontSize: 18, color: th.text, fontFace: th.font }) });
  }
  return { id: uid(), bg: th.bg, els };
}

export function newDeck(): Deck {
  const d: Deck = { w: 960, h: 540, theme: 'paper', slides: [] };
  d.slides.push(makeSlide('title', d));
  return d;
}

/** Re-color existing slides when switching theme. */
export function applyTheme(deck: Deck, themeId: string): Deck {
  const old = themeOf(deck);
  const th = THEMES.find((t) => t.id === themeId) ?? THEMES[0];
  const map = (c: string, fallback: string) => (c.toLowerCase() === old.title.toLowerCase() ? th.title : c.toLowerCase() === old.text.toLowerCase() ? th.text : c.toLowerCase() === old.accent.toLowerCase() ? th.accent : fallback);
  return {
    ...deck,
    theme: th.id,
    slides: deck.slides.map((s) => ({
      ...s,
      bg: s.bg.toLowerCase() === old.bg.toLowerCase() ? th.bg : s.bg,
      els: s.els.map((e) => {
        if (e.type === 'image') return e;
        const style = { ...e.style, color: map(e.style.color, e.style.color), fontFace: e.style.fontFace === old.font ? th.font : e.style.fontFace };
        if (e.type === 'shape') return { ...e, style, fill: map(e.fill, e.fill) };
        return { ...e, style };
      }),
    })),
  };
}
