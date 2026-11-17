import type { ChartType } from '../sheet/model';

export type ShapeKind =
  | 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'rtTriangle' | 'diamond' | 'pentagon' | 'hexagon' | 'octagon'
  | 'parallelogram' | 'trapezoid' | 'star4' | 'star5' | 'star8' | 'heart' | 'plus' | 'callout' | 'cloud'
  | 'rightArrow' | 'leftArrow' | 'upArrow' | 'downArrow' | 'leftRightArrow' | 'chevron' | 'homePlate'
  | 'line' | 'arrowLine';

/** Shapes in the order the Insert gallery shows them. */
export const SHAPE_GROUPS: ShapeKind[][] = [
  ['line', 'arrowLine', 'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'pentagon', 'hexagon', 'octagon', 'parallelogram', 'trapezoid'],
  ['rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'chevron', 'homePlate'],
  ['star4', 'star5', 'star8', 'heart', 'plus', 'callout', 'cloud'],
];

export type Transition = 'none' | 'fade' | 'push' | 'wipe' | 'zoom' | 'cover';

export interface TextStyle {
  fontSize: number; // pt
  fontFace?: string;
  color: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Line height multiplier. */
  lineSpacing?: number;
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
  flipH?: boolean;
  flipV?: boolean;
  /** 0–1 */
  opacity?: number;
  shadow?: boolean;
}

export interface TextEl extends BaseEl {
  type: 'text';
  text: string;
  style: TextStyle;
  placeholder?: 'title' | 'subtitle' | 'body';
  fill?: string;
  bullets?: boolean;
  numbered?: boolean;
}

export interface ShapeEl extends BaseEl {
  type: 'shape';
  shape: ShapeKind;
  fill: string;
  stroke?: string;
  strokeW?: number;
  dash?: 'solid' | 'dash' | 'dot';
  text: string;
  style: TextStyle;
}

export interface TableEl extends BaseEl {
  type: 'table';
  rows: string[][];
  /** Header and banding color. */
  color: string;
  header: boolean;
  banded: boolean;
  style: TextStyle;
}

export interface ChartData {
  cats: string[];
  series: { name: string; values: (number | null)[] }[];
}

export interface ChartEl extends BaseEl {
  type: 'chart';
  chart: ChartType;
  data: ChartData;
  title?: string;
  legend?: boolean;
}

export interface ImageEl extends BaseEl {
  type: 'image';
  src: string;
}

export type SlideEl = TextEl | ShapeEl | ImageEl | TableEl | ChartEl;
/** Elements that hold editable text with a text style. */
export type TextualEl = TextEl | ShapeEl;
export const isTextual = (e: SlideEl | null | undefined): e is TextualEl => !!e && (e.type === 'text' || e.type === 'shape');

export interface Slide {
  id: string;
  bg: string;
  els: SlideEl[];
  notes?: string;
  hidden?: boolean;
  transition?: Transition;
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
  slideNumbers?: boolean;
  footer?: string;
}

export const SLIDE_SIZES = [
  { id: '16:9', w: 960, h: 540 },
  { id: '4:3', w: 960, h: 720 },
  { id: '16:10', w: 960, h: 600 },
  { id: 'A4', w: 960, h: 679 },
] as const;

/** Resize the slides, moving and scaling every element to keep the layout. */
export function resizeDeck(deck: Deck, w: number, h: number): Deck {
  const sx = w / deck.w, sy = h / deck.h;
  const k = Math.min(sx, sy);
  return {
    ...deck,
    w,
    h,
    slides: deck.slides.map((s) => ({
      ...s,
      els: s.els.map((e) => {
        // keep proportions: scale by the smaller factor and re-center on the new axis
        const nw = e.w * k, nh = e.h * k;
        const cx = (e.x + e.w / 2) * sx, cy = (e.y + e.h / 2) * sy;
        return { ...e, x: Math.round(cx - nw / 2), y: Math.round(cy - nh / 2), w: Math.round(nw), h: Math.round(nh) };
      }),
    })),
  };
}

export function newTable(rows: number, cols: number, x: number, y: number, w: number, color: string, font?: string): TableEl {
  return {
    id: uid(), type: 'table', x, y, w, h: rows * 40,
    rows: Array.from({ length: rows }, () => Array.from({ length: cols }, () => '')),
    color, header: true, banded: true,
    style: textStyle({ fontSize: 14, color: '#1f2937', fontFace: font }),
  };
}

export function newChart(type: ChartType, x: number, y: number): ChartEl {
  return {
    id: uid(), type: 'chart', chart: type, x, y, w: 480, h: 300, legend: true, title: '',
    data: {
      cats: ['Q1', 'Q2', 'Q3', 'Q4'],
      series: [
        { name: 'Series 1', values: [4.3, 2.5, 3.5, 4.5] },
        { name: 'Series 2', values: [2.4, 4.4, 1.8, 2.8] },
      ],
    },
  };
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
        if (e.type === 'image' || e.type === 'chart') return e;
        if (e.type === 'table') return { ...e, color: map(e.color, e.color) === e.color && e.color.toLowerCase() === old.accent.toLowerCase() ? th.accent : map(e.color, e.color), style: { ...e.style, fontFace: e.style.fontFace === old.font ? th.font : e.style.fontFace } };
        const style = { ...e.style, color: map(e.style.color, e.style.color), fontFace: e.style.fontFace === old.font ? th.font : e.style.fontFace };
        if (e.type === 'shape') return { ...e, style, fill: map(e.fill, e.fill) };
        return { ...e, style };
      }),
    })),
  };
}
