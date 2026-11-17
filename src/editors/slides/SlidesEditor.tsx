import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import {
  AlignCenter, AlignLeft, AlignRight, BringToFront, Copy, ImagePlus, Italic, Bold, Underline, LayoutTemplate,
  PaintBucket, Play, Plus, Presentation, Redo2, SendToBack, SquareDashed, Trash2,
  Type, Undo2, ChevronUp, ChevronDown, Palette, List, Columns2, AlignVerticalJustifyCenter, AlignVerticalJustifyStart,
  AlignVerticalJustifyEnd, PenLine, Shapes, Layers, AlignStartVertical, AlignCenterVertical, AlignEndVertical,
  AlignStartHorizontal, AlignCenterHorizontal, AlignEndHorizontal, AlignHorizontalDistributeCenter, AlignVerticalDistributeCenter,
  ArrowUpToLine, ArrowDownToLine, RotateCw, RotateCcw, FlipHorizontal2, FlipVertical2, Blend, Table2, ChartColumn, ChartBar,
  ChartLine, ChartArea, ChartPie, ChartScatter, Donut, Pencil, EyeOff, Eye, Hash, PanelBottom, MonitorPlay, Scissors,
  ClipboardPaste, Strikethrough, ListOrdered, AArrowUp, AArrowDown, Sparkles, Ratio, Rows3, Columns3, BoxSelect, Timer,
  ChevronLeft, ChevronRight, X,
} from 'lucide-react';
import { useI18n, type Key } from '../../i18n';
import { Btn, ColorBtn, Group, Menu, MenuItem, Ribbon, Select } from '../../ui/Ribbon';
import { AskDialog, type Ask } from '../../ui/AskDialog';
import { useHistory } from '../useHistory';
import { FontPicker } from '../../ui/FontPicker';
import type { ChartType } from '../sheet/model';
import {
  applyTheme, isTextual, makeSlide, newChart, newTable, resizeDeck, SHAPE_GROUPS, SLIDE_SIZES, textStyle, THEMES, themeOf, uid,
  type ChartData, type Deck, type Layout, type ShapeEl, type ShapeKind, type Slide, type SlideEl, type TableEl, type TextEl,
  type TextStyle, type Transition,
} from './model';
import { decoFor, Decorations, ElView, ScaledSlide, ShapeSvg, SlideView, TextBody } from './SlideView';

interface Props {
  data: Deck;
  onChange: (d: Deck) => void;
}

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';
const HANDLES: Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const CLIP_MIME = 'application/x-paperline-slide-els';
const TRANSITIONS: Transition[] = ['none', 'fade', 'push', 'wipe', 'zoom', 'cover'];
const CHART_TYPES: [ChartType, typeof ChartColumn][] = [
  ['column', ChartColumn], ['bar', ChartBar], ['line', ChartLine], ['area', ChartArea], ['pie', ChartPie], ['doughnut', Donut], ['scatter', ChartScatter],
];
const SNAP = 6;

interface Guides { v: number[]; h: number[] }
type Box = { x: number; y: number; w: number; h: number };

const boundsOf = (els: Box[]): Box => {
  const x0 = Math.min(...els.map((e) => e.x)), y0 = Math.min(...els.map((e) => e.y));
  const x1 = Math.max(...els.map((e) => e.x + e.w)), y1 = Math.max(...els.map((e) => e.y + e.h));
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
};

/** Fresh ids for pasted or duplicated elements. */
const cloneEls = (els: SlideEl[], dx: number, dy: number) => els.map((e) => ({ ...e, id: uid(), x: e.x + dx, y: e.y + dy }) as SlideEl);

function chartToText(d: ChartData) {
  const head = ['', ...d.series.map((s) => s.name)].join('\t');
  return [head, ...d.cats.map((c, i) => [c, ...d.series.map((s) => (s.values[i] ?? '') as string | number)].join('\t'))].join('\n');
}
function textToChart(text: string): ChartData | null {
  const rows = text.split(/\r?\n/).filter((l) => l.trim()).map((l) => l.split(l.includes('\t') ? '\t' : ',').map((c) => c.trim()));
  if (rows.length < 2 || rows[0].length < 2) return null;
  const names = rows[0].slice(1);
  return {
    cats: rows.slice(1).map((r) => r[0] ?? ''),
    series: names.map((name, si) => ({
      name: name || `Series ${si + 1}`,
      values: rows.slice(1).map((r) => { const v = parseFloat((r[si + 1] ?? '').replace(/,/g, '')); return Number.isFinite(v) ? v : null; }),
    })),
  };
}

export function SlidesEditor({ data, onChange }: Props) {
  const { t } = useI18n();
  const { commit, undo, redo } = useHistory(data, onChange);
  const [cur, setCur] = useState(0);
  const [selIds, setSelIds] = useState<string[]>([]);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [cell, setCell] = useState<{ r: number; c: number } | null>(null);
  const [presenting, setPresenting] = useState<{ start: number; presenter: boolean } | null>(null);
  const [fit, setFit] = useState(1);
  const [guides, setGuides] = useState<Guides>({ v: [], h: [] });
  const [marquee, setMarquee] = useState<Box | null>(null);
  const [ask, setAsk] = useState<Ask | null>(null);
  const [ctx, setCtx] = useState<{ x: number; y: number } | null>(null);
  const [dropAt, setDropAt] = useState<number | null>(null);
  const [preview, setPreview] = useState<{ key: number; tr: Transition }>({ key: 0, tr: 'none' });
  const stageRef = useRef<HTMLDivElement>(null);
  const surfaceRef = useRef<HTMLDivElement>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const clip = useRef<{ els: SlideEl[]; text: string; n: number } | null>(null);
  const editBefore = useRef<Deck | null>(null);
  const dragSlide = useRef<number | null>(null);
  const live = useRef(data);
  live.current = data;

  const idx = Math.min(cur, data.slides.length - 1);
  const slide = data.slides[idx];
  const sels = slide ? slide.els.filter((e) => selIds.includes(e.id)) : [];
  const sel = sels.length === 1 ? sels[0] : null;
  const theme = themeOf(data);

  const placeholderFor = useCallback(
    (el: TextEl) => (el.placeholder === 'title' ? t('clickTitle') : el.placeholder === 'subtitle' ? t('clickSubtitle') : el.placeholder === 'body' ? t('clickText') : undefined),
    [t],
  );

  useLayoutEffect(() => {
    const el = stageRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => {
      const pad = 56;
      setFit(Math.max(0.1, Math.min((el.clientWidth - pad) / data.w, (el.clientHeight - pad) / data.h)));
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, [data.w, data.h]);

  useEffect(() => {
    if (!ctx) return;
    const h = (e: MouseEvent) => { if (!(e.target as HTMLElement).closest('.ctx-menu')) setCtx(null); };
    document.addEventListener('mousedown', h);
    return () => document.removeEventListener('mousedown', h);
  }, [ctx]);

  const focusStage = () => stageRef.current?.focus({ preventScroll: true });
  const select = (ids: string[]) => { setSelIds(ids); setCtx(null); };

  // ----- mutations -----
  const setSlides = (slides: Slide[]) => commit({ ...live.current, slides });
  const updateSlide = (fn: (s: Slide) => Slide, i = idx) => {
    const slides = live.current.slides.slice();
    slides[i] = fn(slides[i]);
    setSlides(slides);
  };
  const mapEls = (ids: string[], fn: (e: SlideEl) => SlideEl, record = true) => {
    const slides = live.current.slides.slice();
    slides[idx] = { ...slides[idx], els: slides[idx].els.map((e) => (ids.includes(e.id) ? fn(e) : e)) };
    if (record) commit({ ...live.current, slides });
    else onChange({ ...live.current, slides });
  };
  const updateEl = (id: string, patch: Partial<SlideEl>, record = true) => mapEls([id], (e) => ({ ...e, ...patch }) as SlideEl, record);
  const patchSel = (patch: Partial<SlideEl> | ((e: SlideEl) => Partial<SlideEl>)) =>
    mapEls(selIds, (e) => ({ ...e, ...(typeof patch === 'function' ? patch(e) : patch) }) as SlideEl);
  const styled = sels.filter((e): e is TextEl | ShapeEl | TableEl => isTextual(e) || e.type === 'table');
  const updateStyle = (patch: Partial<TextStyle> | ((s: TextStyle) => Partial<TextStyle>)) => {
    if (!styled.length) return;
    mapEls(styled.map((e) => e.id), (e) => (isTextual(e) || e.type === 'table' ? ({ ...e, style: { ...e.style, ...(typeof patch === 'function' ? patch(e.style) : patch) } } as SlideEl) : e));
  };
  const addEls = (els: SlideEl[]) => {
    updateSlide((s) => ({ ...s, els: [...s.els, ...els] }));
    select(els.map((e) => e.id));
  };
  const addEl = (el: SlideEl) => addEls([el]);
  const deleteSel = () => {
    if (!sels.length) return;
    updateSlide((s) => ({ ...s, els: s.els.filter((e) => !selIds.includes(e.id)) }));
    select([]);
  };

  /** Start a run of live edits (typing in a table) that undo as one step. */
  const beginEdit = () => { editBefore.current = live.current; };
  const endEdit = () => {
    const before = editBefore.current;
    editBefore.current = null;
    if (before && before !== live.current) { const after = live.current; onChange(before); commit(after); }
  };

  // ----- slides -----
  const addSlide = (layout: Layout) => {
    const slides = data.slides.slice();
    slides.splice(idx + 1, 0, makeSlide(layout, data));
    setSlides(slides);
    setCur(idx + 1);
    select([]);
  };
  const dupSlide = (i = idx) => {
    const slides = data.slides.slice();
    const s = data.slides[i];
    slides.splice(i + 1, 0, { ...s, id: uid(), els: s.els.map((e) => ({ ...e, id: uid() })) });
    setSlides(slides);
    setCur(i + 1);
  };
  const delSlide = (i = idx) => {
    if (data.slides.length <= 1) return;
    setSlides(data.slides.filter((_, j) => j !== i));
    setCur(Math.max(0, Math.min(i, data.slides.length - 2)));
    select([]);
  };
  const moveSlideTo = (from: number, to: number) => {
    if (from === to || from + 1 === to) return;
    const slides = data.slides.slice();
    const [s] = slides.splice(from, 1);
    const at = to > from ? to - 1 : to;
    slides.splice(at, 0, s);
    setSlides(slides);
    setCur(at);
  };
  const toggleHidden = (i = idx) => updateSlide((s) => ({ ...s, hidden: !s.hidden }), i);

  // ----- insert -----
  const center = (w: number, h: number) => ({ x: Math.round((data.w - w) / 2), y: Math.round((data.h - h) / 2) });
  const addText = () =>
    addEl({ id: uid(), type: 'text', text: '', x: data.w * 0.3, y: data.h * 0.4, w: data.w * 0.4, h: 60, style: textStyle({ color: theme.text, fontFace: theme.font }) });
  const addShape = (shape: ShapeKind) => {
    const isLine = shape === 'line' || shape === 'arrowLine';
    const w = isLine ? 220 : ['rightArrow', 'leftArrow', 'leftRightArrow', 'chevron', 'homePlate'].includes(shape) ? 200 : shape === 'upArrow' || shape === 'downArrow' ? 90 : 160;
    const h = isLine ? 20 : shape === 'upArrow' || shape === 'downArrow' ? 170 : ['rightArrow', 'leftArrow', 'leftRightArrow'].includes(shape) ? 90 : shape === 'rect' || shape === 'roundRect' || shape === 'callout' || shape === 'cloud' ? 110 : 160;
    addEl({
      id: uid(), type: 'shape', shape, text: '', ...center(w, h), w, h,
      fill: theme.accent, stroke: isLine ? theme.accent : undefined, strokeW: isLine ? 3 : undefined,
      style: textStyle({ color: '#ffffff', align: 'center', valign: 'middle', fontFace: theme.font }),
    });
  };
  const addImageSrc = (src: string) => {
    const img = new Image();
    img.onload = () => {
      let w = img.naturalWidth, h = img.naturalHeight;
      const max = Math.min(data.w * 0.6, data.h * 0.7 * (w / h));
      if (w > max) { h = (h * max) / w; w = max; }
      addEl({ id: uid(), type: 'image', src, ...center(w, h), w, h });
    };
    img.src = src;
  };
  const addImage = (file: File) => {
    const r = new FileReader();
    r.onload = () => addImageSrc(String(r.result));
    r.readAsDataURL(file);
  };
  const addTable = (rows: number, cols: number) => {
    const w = Math.min(data.w * 0.8, cols * 150);
    const el = newTable(rows, cols, 0, 0, w, theme.accent, theme.font);
    Object.assign(el, center(el.w, el.h));
    addEl(el);
  };
  const addChart = (type: ChartType) => {
    const el = newChart(type, 0, 0);
    Object.assign(el, center(el.w, el.h));
    addEl(el);
  };

  // ----- clipboard -----
  const copySel = (cut = false) => {
    if (!sels.length) return null;
    const text = sels.filter(isTextual).map((e) => e.text).join('\n');
    clip.current = { els: sels, text, n: 0 };
    if (cut) deleteSel();
    return { json: JSON.stringify(sels), text };
  };
  const pasteEls = (els: SlideEl[]) => {
    const n = (clip.current?.n ?? 0) + 1;
    if (clip.current) clip.current.n = n;
    addEls(cloneEls(els, 16 * n, 16 * n));
  };
  const pasteText = (text: string) =>
    addEl({ id: uid(), type: 'text', text, x: data.w * 0.2, y: data.h * 0.35, w: data.w * 0.6, h: Math.min(data.h * 0.5, 40 + 28 * text.split('\n').length), style: textStyle({ color: theme.text, fontFace: theme.font }) });
  const pasteFromButton = async () => {
    if (clip.current) { pasteEls(clip.current.els); return; }
    try {
      const text = await navigator.clipboard.readText();
      if (text) pasteText(text);
    } catch { /* clipboard access denied */ }
  };
  const isEditingText = (e: { target: EventTarget | null }) => !!(e.target as HTMLElement | null)?.closest?.('input, textarea, select, [contenteditable="true"]');
  const onCopy = (e: React.ClipboardEvent, cut = false) => {
    if (isEditingText(e) || !sels.length) return;
    const out = copySel(cut);
    if (!out) return;
    e.preventDefault();
    e.clipboardData.setData('text/plain', out.text);
    e.clipboardData.setData(CLIP_MIME, out.json);
  };
  const onPaste = (e: React.ClipboardEvent) => {
    if (isEditingText(e) || presenting) return;
    const cd = e.clipboardData;
    const json = cd.getData(CLIP_MIME);
    const img = Array.from(cd.files).find((f) => f.type.startsWith('image/'));
    const text = cd.getData('text/plain');
    e.preventDefault();
    if (json) { try { pasteEls(JSON.parse(json)); return; } catch { /* fall through */ } }
    if (img) { addImage(img); return; }
    if (clip.current && text === clip.current.text) { pasteEls(clip.current.els); return; }
    if (text) pasteText(text);
  };

  // ----- arrange -----
  const reorder = (how: 'front' | 'back' | 'forward' | 'backward') => {
    if (!sels.length) return;
    updateSlide((s) => {
      const rest = s.els.filter((e) => !selIds.includes(e.id));
      const mine = s.els.filter((e) => selIds.includes(e.id));
      if (how === 'front') return { ...s, els: [...rest, ...mine] };
      if (how === 'back') return { ...s, els: [...mine, ...rest] };
      const els = s.els.slice();
      const order = how === 'forward' ? [...els.keys()].reverse() : [...els.keys()];
      for (const i of order) {
        if (!selIds.includes(els[i].id)) continue;
        const j = how === 'forward' ? i + 1 : i - 1;
        if (j < 0 || j >= els.length || selIds.includes(els[j].id)) continue;
        [els[i], els[j]] = [els[j], els[i]];
      }
      return { ...s, els };
    });
  };
  /** Align to each other when several are selected, to the slide when one is. */
  const align = (how: 'l' | 'c' | 'r' | 't' | 'm' | 'b') => {
    if (!sels.length) return;
    const b = sels.length > 1 ? boundsOf(sels) : { x: 0, y: 0, w: data.w, h: data.h };
    patchSel((e) => {
      switch (how) {
        case 'l': return { x: b.x };
        case 'c': return { x: Math.round(b.x + (b.w - e.w) / 2) };
        case 'r': return { x: b.x + b.w - e.w };
        case 't': return { y: b.y };
        case 'm': return { y: Math.round(b.y + (b.h - e.h) / 2) };
        default: return { y: b.y + b.h - e.h };
      }
    });
  };
  const distribute = (axis: 'x' | 'y') => {
    if (sels.length < 3) return;
    const dim = axis === 'x' ? 'w' : 'h';
    const sorted = sels.slice().sort((a, b) => a[axis] - b[axis]);
    const first = sorted[0], last = sorted[sorted.length - 1];
    const total = sorted.reduce((s, e) => s + e[dim], 0);
    const gap = (last[axis] + last[dim] - first[axis] - total) / (sorted.length - 1);
    const pos = new Map<string, number>();
    let p = first[axis];
    for (const e of sorted) { pos.set(e.id, Math.round(p)); p += e[dim] + gap; }
    patchSel((e) => ({ [axis]: pos.get(e.id) }));
  };
  const rotateBy = (deg: number) => patchSel((e) => ({ rot: (((e.rot ?? 0) + deg) % 360 + 360) % 360 || undefined }));

  // ----- pointer: move, resize, rotate, marquee -----
  const toSlide = (cx: number, cy: number) => {
    const r = surfaceRef.current!.getBoundingClientRect();
    return { x: (cx - r.left) / fit, y: (cy - r.top) / fit };
  };

  const startMove = (e: React.PointerEvent, el: SlideEl) => {
    if (e.button !== 0 || editingId === el.id) return;
    e.stopPropagation();
    e.preventDefault();
    focusStage();
    setCtx(null);
    if (editingId) { endEdit(); setEditingId(null); setCell(null); }
    if (e.shiftKey || e.ctrlKey || e.metaKey) {
      select(selIds.includes(el.id) ? selIds.filter((x) => x !== el.id) : [...selIds, el.id]);
      return;
    }
    const ids = selIds.includes(el.id) ? selIds : [el.id];
    if (!selIds.includes(el.id)) select(ids);
    const moving = live.current.slides[idx].els.filter((x) => ids.includes(x.id));
    const origin = new Map(moving.map((m) => [m.id, { x: m.x, y: m.y }]));
    const box0 = boundsOf(moving);
    const others = live.current.slides[idx].els.filter((x) => !ids.includes(x.id));
    const vx = [0, data.w / 2, data.w, ...others.flatMap((o) => [o.x, o.x + o.w / 2, o.x + o.w])];
    const hy = [0, data.h / 2, data.h, ...others.flatMap((o) => [o.y, o.y + o.h / 2, o.y + o.h])];
    const sx = e.clientX, sy = e.clientY;
    const before = live.current;
    let moved = false;
    const snap1 = (pos: number, size: number, lines: number[]) => {
      let best: { d: number; line: number } | null = null;
      for (const off of [0, size / 2, size]) for (const l of lines) {
        const d = l - (pos + off);
        if (Math.abs(d) <= SNAP && (!best || Math.abs(d) < Math.abs(best.d))) best = { d, line: l };
      }
      return best;
    };
    const move = (ev: PointerEvent) => {
      let dx = (ev.clientX - sx) / fit, dy = (ev.clientY - sy) / fit;
      if (!moved && Math.abs(dx) + Math.abs(dy) < 2) return;
      moved = true;
      if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
      const g: Guides = { v: [], h: [] };
      if (!ev.altKey) {
        const bx = snap1(box0.x + dx, box0.w, vx);
        const by = snap1(box0.y + dy, box0.h, hy);
        if (bx) { dx += bx.d; g.v.push(bx.line); }
        if (by) { dy += by.d; g.h.push(by.line); }
      }
      setGuides(g);
      mapEls(ids, (m) => { const o = origin.get(m.id)!; return { ...m, x: Math.round(o.x + dx), y: Math.round(o.y + dy) } as SlideEl; }, false);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      setGuides({ v: [], h: [] });
      if (moved) { const after = live.current; onChange(before); commit(after); }
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const startResize = (e: React.PointerEvent, el: SlideEl, handle: Handle) => {
    if (e.button !== 0) return;
    e.stopPropagation();
    e.preventDefault();
    const sx = e.clientX, sy = e.clientY;
    const o = { x: el.x, y: el.y, w: el.w, h: el.h };
    const rad = ((el.rot ?? 0) * Math.PI) / 180;
    const cos = Math.cos(rad), sin = Math.sin(rad);
    const before = live.current;
    let moved = false;
    const move = (ev: PointerEvent) => {
      const gx = (ev.clientX - sx) / fit, gy = (ev.clientY - sy) / fit;
      if (!moved && Math.abs(gx) + Math.abs(gy) < 2) return;
      moved = true;
      // work in the element's own (unrotated) axes so the opposite edge stays put
      const lx = gx * cos + gy * sin, ly = -gx * sin + gy * cos;
      let w = o.w, h = o.h;
      if (handle.includes('e')) w = o.w + lx;
      if (handle.includes('w')) w = o.w - lx;
      if (handle.includes('s')) h = o.h + ly;
      if (handle.includes('n')) h = o.h - ly;
      const isLine = el.type === 'shape' && (el.shape === 'line' || el.shape === 'arrowLine');
      w = Math.max(isLine ? 4 : 10, w);
      h = Math.max(isLine ? 2 : 10, h);
      if ((el.type === 'image' || ev.shiftKey) && handle.length === 2) {
        const r = o.w / o.h;
        if (w / h > r) w = h * r; else h = w / r;
      }
      const cx = (handle.includes('e') ? 1 : handle.includes('w') ? -1 : 0) * (w - o.w) / 2;
      const cy = (handle.includes('s') ? 1 : handle.includes('n') ? -1 : 0) * (h - o.h) / 2;
      const ncx = o.x + o.w / 2 + cx * cos - cy * sin;
      const ncy = o.y + o.h / 2 + cx * sin + cy * cos;
      updateEl(el.id, { x: Math.round(ncx - w / 2), y: Math.round(ncy - h / 2), w: Math.round(w), h: Math.round(h) }, false);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      if (moved) { const after = live.current; onChange(before); commit(after); }
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const startRotate = (e: React.PointerEvent, el: SlideEl) => {
    e.stopPropagation();
    e.preventDefault();
    const r = surfaceRef.current!.getBoundingClientRect();
    const cx = r.left + (el.x + el.w / 2) * fit, cy = r.top + (el.y + el.h / 2) * fit;
    const before = live.current;
    let moved = false;
    const move = (ev: PointerEvent) => {
      moved = true;
      let deg = (Math.atan2(ev.clientY - cy, ev.clientX - cx) * 180) / Math.PI + 90;
      deg = ev.shiftKey ? Math.round(deg / 15) * 15 : Math.round(deg);
      for (const s of [0, 90, 180, 270, 360]) if (Math.abs(deg - s) < 3) deg = s;
      deg = ((deg % 360) + 360) % 360;
      updateEl(el.id, { rot: deg || undefined }, false);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      if (moved) { const after = live.current; onChange(before); commit(after); }
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const startMarquee = (e: React.PointerEvent) => {
    if (e.button !== 0) return;
    focusStage();
    (document.activeElement as HTMLElement | null)?.blur?.();
    focusStage();
    if (editingId) { endEdit(); setEditingId(null); setCell(null); }
    setCtx(null);
    if (!surfaceRef.current) return;
    const p0 = toSlide(e.clientX, e.clientY);
    const additive = e.shiftKey || e.ctrlKey || e.metaKey;
    const base = additive ? selIds : [];
    if (!additive) setSelIds([]);
    let box: Box | null = null;
    const move = (ev: PointerEvent) => {
      const p = toSlide(ev.clientX, ev.clientY);
      box = { x: Math.min(p.x, p0.x), y: Math.min(p.y, p0.y), w: Math.abs(p.x - p0.x), h: Math.abs(p.y - p0.y) };
      if (box.w + box.h < 4) return;
      setMarquee(box);
      const b = box;
      const hit = live.current.slides[idx].els.filter((x) => x.x < b.x + b.w && x.x + x.w > b.x && x.y < b.y + b.h && x.y + x.h > b.y).map((x) => x.id);
      setSelIds([...new Set([...base, ...hit])]);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      setMarquee(null);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const startEditing = (el: SlideEl, target?: HTMLElement) => {
    if (el.type === 'image' || el.type === 'chart') {
      if (el.type === 'chart') editChartData(el);
      return;
    }
    select([el.id]);
    if (el.type === 'table') {
      const td = target?.closest('td') as HTMLTableCellElement | null;
      const tr = td?.parentElement as HTMLTableRowElement | null;
      setCell(td && tr ? { r: tr.rowIndex, c: td.cellIndex } : { r: 0, c: 0 });
      beginEdit();
    }
    setEditingId(el.id);
  };
  const stopEditing = () => { if (editingId) { endEdit(); setEditingId(null); setCell(null); focusStage(); } };

  // ----- tables -----
  const tableSel = sel?.type === 'table' ? sel : null;
  const editTable = (fn: (t: TableEl) => Partial<TableEl>) => { if (tableSel) updateEl(tableSel.id, fn(tableSel)); };
  const addRow = (below: boolean) =>
    editTable((tb) => {
      const at = (cell?.r ?? tb.rows.length - 1) + (below ? 1 : 0);
      const rows = tb.rows.slice();
      rows.splice(at, 0, tb.rows[0].map(() => ''));
      return { rows, h: Math.round(tb.h + tb.h / tb.rows.length) };
    });
  const addCol = (right: boolean) =>
    editTable((tb) => {
      const at = (cell?.c ?? tb.rows[0].length - 1) + (right ? 1 : 0);
      return { rows: tb.rows.map((r) => { const n = r.slice(); n.splice(at, 0, ''); return n; }) };
    });
  const delRow = () =>
    editTable((tb) => {
      if (tb.rows.length <= 1) return {};
      const at = cell?.r ?? tb.rows.length - 1;
      setCell(null);
      return { rows: tb.rows.filter((_, i) => i !== at), h: Math.round(tb.h - tb.h / tb.rows.length) };
    });
  const delCol = () =>
    editTable((tb) => {
      if (tb.rows[0].length <= 1) return {};
      const at = cell?.c ?? tb.rows[0].length - 1;
      setCell(null);
      return { rows: tb.rows.map((r) => r.filter((_, i) => i !== at)) };
    });
  const setCellText = (tb: TableEl, r: number, c: number, v: string) => {
    const rows = tb.rows.map((row) => row.slice());
    rows[r][c] = v;
    updateEl(tb.id, { rows }, false);
  };

  // ----- charts -----
  const editChartData = (el: Extract<SlideEl, { type: 'chart' }>) =>
    setAsk({
      title: t('p_editData'),
      wide: true,
      fields: [
        { key: 'title', label: t('s_chartTitle'), value: el.title ?? '' },
        { key: 'data', label: t('p_chartData'), value: chartToText(el.data), type: 'textarea', rows: 9, hint: t('p_chartDataHint') },
      ],
      onOk: (v) => {
        const d = textToChart(v.data);
        updateEl(el.id, { title: v.title, ...(d ? { data: d } : {}) } as Partial<SlideEl>);
      },
    });

  // ----- keyboard -----
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (presenting !== null || editingId || ask) return;
      const tgt = e.target as HTMLElement;
      if (tgt.closest('input, textarea, select, [contenteditable="true"]')) return;
      if (!tgt.closest('.slides-editor') && tgt !== document.body) return;
      const mod = e.ctrlKey || e.metaKey;
      const k = e.key.toLowerCase();
      if (mod && k === 'z') { e.preventDefault(); if (e.shiftKey) redo(); else undo(); return; }
      if (mod && k === 'y') { e.preventDefault(); redo(); return; }
      if (mod && k === 'a') { e.preventDefault(); select(slide.els.map((x) => x.id)); return; }
      if (mod && k === 'd') { e.preventDefault(); if (sels.length) addEls(cloneEls(sels, 20, 20)); else dupSlide(); return; }
      if (mod && k === 'm') { e.preventDefault(); addSlide('content'); return; }
      if (mod && (k === ']' || k === '[')) { e.preventDefault(); reorder(k === ']' ? (e.shiftKey ? 'front' : 'forward') : (e.shiftKey ? 'back' : 'backward')); return; }
      if (sels.length) {
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteSel(); return; }
        const step = e.shiftKey ? 10 : 1;
        const d = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key];
        if (d) { e.preventDefault(); patchSel((x) => ({ x: x.x + d[0], y: x.y + d[1] })); return; }
        if ((e.key === 'Enter' || e.key === 'F2') && sel) { e.preventDefault(); startEditing(sel); return; }
        if (e.key === 'Escape') { select([]); return; }
        if (e.key === 'Tab') {
          e.preventDefault();
          const i = slide.els.findIndex((x) => x.id === sels[0].id);
          const n = slide.els.length;
          select([slide.els[(i + (e.shiftKey ? -1 : 1) + n) % n].id]);
          return;
        }
      } else {
        if (e.key === 'PageDown' || e.key === 'ArrowDown') { e.preventDefault(); setCur((c) => Math.min(data.slides.length - 1, c + 1)); }
        if (e.key === 'PageUp' || e.key === 'ArrowUp') { e.preventDefault(); setCur((c) => Math.max(0, c - 1)); }
        if (e.key === 'Delete') { e.preventDefault(); delSlide(); }
        if (e.key === 'Tab' && slide.els.length) { e.preventDefault(); select([slide.els[e.shiftKey ? slide.els.length - 1 : 0].id]); }
      }
      if (e.key === 'F5') { e.preventDefault(); setPresenting({ start: e.shiftKey ? idx : 0, presenter: e.altKey }); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  if (!slide) return null;
  const st = styled[0]?.style;
  const hasStyled = styled.length > 0;
  const textEl = sel?.type === 'text' ? sel : null;
  const shapes = sels.filter((e): e is ShapeEl => e.type === 'shape');

  // ----- ribbon pieces -----
  const layoutMenu = (onPick: (l: Layout) => void) => (
    <Menu trigger={(_o, toggle) => <Btn big icon={<Plus size={22} />} label={t('newSlide')} onClick={toggle} />}>
      {(close) =>
        (
          [
            ['title', t('layoutTitle'), <Presentation size={16} />],
            ['content', t('layoutContent'), <List size={16} />],
            ['two', t('layoutTwo'), <Columns2 size={16} />],
            ['blank', t('layoutBlank'), <SquareDashed size={16} />],
          ] as const
        ).map(([l, label, icon]) => <MenuItem key={l} icon={icon} label={label} onClick={() => { onPick(l); close(); }} />)
      }
    </Menu>
  );

  const shapeGallery = (big = false) => (
    <Menu trigger={(_o, toggle) => <Btn big={big} icon={<Shapes size={big ? 22 : 18} />} label={t('p_shapes')} showLabel={!big} onClick={toggle} />}>
      {(close) => (
        <div className="shape-gallery">
          {SHAPE_GROUPS.map((g, gi) => (
            <div key={gi} className="shape-grid">
              {g.map((k) => (
                <button key={k} type="button" className="shape-pick" title={k} aria-label={k} onMouseDown={(e) => e.preventDefault()} onClick={() => { addShape(k); close(); }}>
                  <ShapeSvg el={{ id: `g-${k}`, shape: k, w: 26, h: k === 'line' || k === 'arrowLine' ? 20 : 20, fill: 'currentColor', stroke: k === 'line' || k === 'arrowLine' ? 'currentColor' : undefined, strokeW: 1.5 }} />
                </button>
              ))}
            </div>
          ))}
        </div>
      )}
    </Menu>
  );

  const tableMenu = (
    <Menu trigger={(_o, toggle) => <Btn big icon={<Table2 size={22} />} label={t('table')} onClick={toggle} />}>
      {(close) => <TablePicker label={t('p_insertTable')} onPick={(r, c) => { addTable(r, c); close(); }} />}
    </Menu>
  );

  const chartMenu = (
    <Menu trigger={(_o, toggle) => <Btn big icon={<ChartColumn size={22} />} label={t('s_charts')} onClick={toggle} />}>
      {(close) => CHART_TYPES.map(([ct, Icon]) => <MenuItem key={ct} icon={<Icon size={16} />} label={t(`s_chart_${ct}` as Key)} onClick={() => { addChart(ct); close(); }} />)}
    </Menu>
  );

  const arrangeMenu = (
    <Menu trigger={(_o, toggle) => <Btn icon={<Layers size={18} />} label={t('p_arrange')} showLabel onClick={toggle} />}>
      {(close) => {
        const it = (icon: React.ReactNode, label: string, fn: () => void, disabled = false) =>
          disabled ? null : <MenuItem key={label} icon={icon} label={label} onClick={() => { fn(); close(); }} />;
        return (
          <>
            {it(<BringToFront size={16} />, t('bringFront'), () => reorder('front'), !sels.length)}
            {it(<ArrowUpToLine size={16} />, t('p_bringForward'), () => reorder('forward'), !sels.length)}
            {it(<ArrowDownToLine size={16} />, t('p_sendBackward'), () => reorder('backward'), !sels.length)}
            {it(<SendToBack size={16} />, t('sendBack'), () => reorder('back'), !sels.length)}
            <div className="menu-sep" />
            {it(<AlignStartVertical size={16} />, t('p_alignLeft'), () => align('l'), !sels.length)}
            {it(<AlignCenterVertical size={16} />, t('p_alignCenter'), () => align('c'), !sels.length)}
            {it(<AlignEndVertical size={16} />, t('p_alignRight'), () => align('r'), !sels.length)}
            {it(<AlignStartHorizontal size={16} />, t('p_alignTop'), () => align('t'), !sels.length)}
            {it(<AlignCenterHorizontal size={16} />, t('p_alignMiddle'), () => align('m'), !sels.length)}
            {it(<AlignEndHorizontal size={16} />, t('p_alignBottom'), () => align('b'), !sels.length)}
            {sels.length >= 3 && <div className="menu-sep" />}
            {it(<AlignHorizontalDistributeCenter size={16} />, t('p_distH'), () => distribute('x'), sels.length < 3)}
            {it(<AlignVerticalDistributeCenter size={16} />, t('p_distV'), () => distribute('y'), sels.length < 3)}
            <div className="menu-sep" />
            {it(<RotateCw size={16} />, t('p_rotRight'), () => rotateBy(90), !sels.length)}
            {it(<RotateCcw size={16} />, t('p_rotLeft'), () => rotateBy(-90), !sels.length)}
            {it(<FlipHorizontal2 size={16} />, t('p_flipH'), () => patchSel((x) => ({ flipH: !x.flipH })), !sels.length)}
            {it(<FlipVertical2 size={16} />, t('p_flipV'), () => patchSel((x) => ({ flipV: !x.flipV })), !sels.length)}
            {!sels.length && <div className="menu-empty">{t('p_selectFirst')}</div>}
          </>
        );
      }}
    </Menu>
  );

  const clipboardTools = (
    <Group label={t('p_clipboard')}>
      <Btn big icon={<ClipboardPaste size={22} />} label={t('s_paste')} onClick={pasteFromButton} />
      <div className="r-col">
        <Btn icon={<Scissors size={16} />} label={t('s_cut')} disabled={!sels.length} onClick={() => copySel(true)} />
        <Btn icon={<Copy size={16} />} label={t('copy')} disabled={!sels.length} onClick={() => { const o = copySel(); if (o) navigator.clipboard?.writeText(o.text).catch(() => {}); }} />
      </div>
    </Group>
  );

  const fontTools = (
    <Group label={t('font')}>
      <div className="r-col">
        <div className="r-row">
          <FontPicker value={st?.fontFace ?? theme.font} onChange={(f) => updateStyle({ fontFace: f })} width={140} />
          <Select label={t('size')} value={st?.fontSize ?? 18} options={[8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96].map((x) => ({ value: x, label: String(x) }))} onChange={(v) => updateStyle({ fontSize: v })} width={60} />
          <Btn icon={<AArrowUp size={18} />} label={t('growFont')} disabled={!hasStyled} onClick={() => updateStyle((s) => ({ fontSize: Math.min(400, Math.round(s.fontSize * 1.1 + 1)) }))} />
          <Btn icon={<AArrowDown size={18} />} label={t('shrinkFont')} disabled={!hasStyled} onClick={() => updateStyle((s) => ({ fontSize: Math.max(6, Math.round((s.fontSize - 1) / 1.1)) }))} />
        </div>
        <div className="r-row">
          <Btn icon={<Bold size={18} />} label={t('bold')} active={!!st?.bold} disabled={!hasStyled} onClick={() => updateStyle({ bold: !st?.bold })} />
          <Btn icon={<Italic size={18} />} label={t('italic')} active={!!st?.italic} disabled={!hasStyled} onClick={() => updateStyle({ italic: !st?.italic })} />
          <Btn icon={<Underline size={18} />} label={t('underline')} active={!!st?.underline} disabled={!hasStyled} onClick={() => updateStyle({ underline: !st?.underline })} />
          <Btn icon={<Strikethrough size={18} />} label={t('strike')} active={!!st?.strike} disabled={!hasStyled} onClick={() => updateStyle({ strike: !st?.strike })} />
          <ColorBtn icon={<Type size={16} />} label={t('textColor')} color={st?.color} onPick={(c) => c && updateStyle({ color: c })} />
        </div>
      </div>
    </Group>
  );

  const paragraphTools = (
    <Group label={t('p_paragraph')}>
      <div className="r-col">
        <div className="r-row">
          <Btn icon={<List size={18} />} label={t('bullets')} active={!!textEl?.bullets && !textEl.numbered} disabled={!textEl} onClick={() => textEl && updateEl(textEl.id, { bullets: !(textEl.bullets && !textEl.numbered), numbered: false } as Partial<SlideEl>)} />
          <Btn icon={<ListOrdered size={18} />} label={t('numbering')} active={!!textEl?.numbered} disabled={!textEl} onClick={() => textEl && updateEl(textEl.id, { numbered: !textEl.numbered, bullets: false } as Partial<SlideEl>)} />
          <Select label={t('lineSpacing')} value={st?.lineSpacing ?? 1} options={[1, 1.15, 1.5, 2, 2.5, 3].map((x) => ({ value: x, label: `↕ ${x}` }))} onChange={(v) => updateStyle({ lineSpacing: v === 1 ? undefined : v })} width={64} />
        </div>
        <div className="r-row">
          <Btn icon={<AlignLeft size={18} />} label={t('alignLeft')} active={st?.align === 'left'} disabled={!hasStyled} onClick={() => updateStyle({ align: 'left' })} />
          <Btn icon={<AlignCenter size={18} />} label={t('alignCenter')} active={st?.align === 'center'} disabled={!hasStyled} onClick={() => updateStyle({ align: 'center' })} />
          <Btn icon={<AlignRight size={18} />} label={t('alignRight')} active={st?.align === 'right'} disabled={!hasStyled} onClick={() => updateStyle({ align: 'right' })} />
          <Btn icon={<AlignVerticalJustifyStart size={18} />} label={t('s_alignTop')} active={st?.valign === 'top'} disabled={!hasStyled} onClick={() => updateStyle({ valign: 'top' })} />
          <Btn icon={<AlignVerticalJustifyCenter size={18} />} label={t('s_alignMiddle')} active={st?.valign === 'middle'} disabled={!hasStyled} onClick={() => updateStyle({ valign: 'middle' })} />
          <Btn icon={<AlignVerticalJustifyEnd size={18} />} label={t('s_alignBottom')} active={st?.valign === 'bottom'} disabled={!hasStyled} onClick={() => updateStyle({ valign: 'bottom' })} />
        </div>
      </div>
    </Group>
  );

  const fillTarget = sels.filter((e): e is ShapeEl | TextEl => e.type === 'shape' || e.type === 'text');
  const drawingTools = (
    <Group label={t('p_drawing')}>
      {shapeGallery(true)}
      <div className="r-col">
        {arrangeMenu}
        <div className="r-row">
          <ColorBtn icon={<PaintBucket size={16} />} label={t('fillColor')} color={fillTarget[0]?.fill} allowNone
            onPick={(c) => mapEls(fillTarget.map((x) => x.id), (x) => ({ ...x, fill: c ?? (x.type === 'shape' ? 'transparent' : undefined) }) as SlideEl)} />
          <ColorBtn icon={<PenLine size={16} />} label={t('outline')} color={shapes[0]?.stroke} allowNone
            onPick={(c) => mapEls(shapes.map((x) => x.id), (x) => ({ ...x, stroke: c ?? undefined, strokeW: c ? (x as ShapeEl).strokeW || 2 : 0 }) as SlideEl)} />
        </div>
      </div>
    </Group>
  );

  const homeTab = (
    <>
      <Group>
        <Btn icon={<Undo2 size={18} />} label={t('undo')} onClick={undo} />
        <Btn icon={<Redo2 size={18} />} label={t('redo')} onClick={redo} />
      </Group>
      {clipboardTools}
      <Group label={t('slide')}>
        {layoutMenu(addSlide)}
        <div className="r-col">
          <Btn icon={<Copy size={16} />} label={t('duplicate')} onClick={() => dupSlide()} />
          <Btn icon={<Trash2 size={16} />} label={t('delete')} disabled={data.slides.length <= 1} onClick={() => delSlide()} />
        </div>
      </Group>
      {fontTools}
      {paragraphTools}
      {drawingTools}
      <Group label={t('editing')}>
        <Btn icon={<BoxSelect size={18} />} label={t('s_selectAll')} onClick={() => select(slide.els.map((x) => x.id))} />
      </Group>
    </>
  );

  const insertTab = (
    <>
      <Group label={t('slide')}>{layoutMenu(addSlide)}</Group>
      <Group label={t('insert')}>
        {tableMenu}
        <Btn big icon={<ImagePlus size={22} />} label={t('image')} onClick={() => fileRef.current?.click()} />
        {shapeGallery(true)}
        {chartMenu}
      </Group>
      <Group label={t('p_text')}>
        <Btn big icon={<Type size={22} />} label={t('textBox')} onClick={addText} />
        <div className="r-col">
          <Btn icon={<Hash size={16} />} label={t('p_slideNumbers')} showLabel active={!!data.slideNumbers} onClick={() => commit({ ...data, slideNumbers: !data.slideNumbers })} />
          <Btn icon={<PanelBottom size={16} />} label={t('footer')} showLabel active={!!data.footer}
            onClick={() => setAsk({ title: t('footer'), fields: [{ key: 'f', label: t('p_footerText'), value: data.footer ?? '' }], onOk: (v) => commit({ ...live.current, footer: v.f.trim() || undefined }) })} />
        </div>
      </Group>
    </>
  );

  const designTab = (
    <>
      <Group label={t('theme')}>
        <div className="theme-row">
          {THEMES.map((th) => (
            <button key={th.id} type="button" className={`theme-chip ${data.theme === th.id ? 'on' : ''}`} style={{ background: th.bg, color: th.title, fontFamily: th.font }} title={th.id} onClick={() => commit(applyTheme(data, th.id))}>
              Aa
              <i style={{ background: th.accent }} />
            </button>
          ))}
        </div>
      </Group>
      <Group label={t('p_slideSize')}>
        <Ratio size={18} className="r-icon" />
        <Select label={t('p_slideSize')} value={SLIDE_SIZES.find((s) => s.w === data.w && s.h === data.h)?.id ?? `${data.w}×${data.h}`}
          options={SLIDE_SIZES.map((s) => ({ value: s.id as string, label: s.id }))}
          onChange={(v) => { const s = SLIDE_SIZES.find((x) => x.id === v); if (s) commit(resizeDeck(data, s.w, s.h)); }} width={80} />
      </Group>
      <Group label={t('background')}>
        <ColorBtn icon={<Palette size={16} />} label={t('background')} color={slide.bg} onPick={(c) => c && updateSlide((s) => ({ ...s, bg: c }))} />
        <Btn icon={<LayoutTemplate size={18} />} label={t('p_applyAll')} showLabel onClick={() => setSlides(data.slides.map((s) => ({ ...s, bg: slide.bg })))} />
      </Group>
    </>
  );

  const playPreview = (tr: Transition) => setPreview((p) => ({ key: p.key + 1, tr }));
  const transitionsTab = (
    <>
      <Group label={t('p_transitions')}>
        <div className="tr-row">
          {TRANSITIONS.map((tr) => (
            <button key={tr} type="button" className={`tr-chip ${(slide.transition ?? 'none') === tr ? 'on' : ''}`} onMouseDown={(e) => e.preventDefault()}
              onClick={() => { updateSlide((s) => ({ ...s, transition: tr === 'none' ? undefined : tr })); playPreview(tr); }}>
              <span className={`tr-demo tr-${tr}`}><i /></span>
              {t(`tr_${tr}` as Key)}
            </button>
          ))}
        </div>
      </Group>
      <Group>
        <Btn big icon={<Sparkles size={22} />} label={t('p_preview')} onClick={() => playPreview(slide.transition ?? 'fade')} />
        <Btn big icon={<LayoutTemplate size={22} />} label={t('p_applyAll')} onClick={() => setSlides(data.slides.map((s) => ({ ...s, transition: slide.transition })))} />
      </Group>
    </>
  );

  const showTab = (
    <>
      <Group label={t('present')}>
        <Btn big icon={<Play size={22} />} label={t('fromStart')} onClick={() => setPresenting({ start: 0, presenter: false })} />
        <Btn big icon={<Presentation size={22} />} label={t('fromCurrent')} onClick={() => setPresenting({ start: idx, presenter: false })} />
        <Btn big icon={<MonitorPlay size={22} />} label={t('p_presenterView')} onClick={() => setPresenting({ start: idx, presenter: true })} />
      </Group>
      <Group label={t('slide')}>
        <Btn big icon={slide.hidden ? <Eye size={22} /> : <EyeOff size={22} />} label={slide.hidden ? t('p_showSlide') : t('p_hideSlide')} active={!!slide.hidden} onClick={() => toggleHidden()} />
      </Group>
    </>
  );

  // contextual tab for the selection
  const formatTab = (() => {
    if (!sels.length) return null;
    const allImg = sels.every((e) => e.type === 'image');
    const opacity = Math.round((sels[0].opacity ?? 1) * 100);
    const common = (
      <>
        <Group label={t('p_effects')}>
          <div className="r-col">
            <div className="r-row">
              <Blend size={16} className="r-icon" />
              <Select label={t('opacity')} value={opacity} options={[100, 90, 80, 70, 60, 50, 40, 30, 20, 10].map((x) => ({ value: x, label: `${x}%` }))} onChange={(v) => patchSel({ opacity: v === 100 ? undefined : v / 100 })} width={70} />
            </div>
            <Btn icon={<Layers size={16} />} label={t('p_shadow')} showLabel active={!!sels[0].shadow} onClick={() => patchSel({ shadow: !sels[0].shadow || undefined })} />
          </div>
        </Group>
        <Group label={t('p_arrange')}>
          {arrangeMenu}
          <div className="r-col">
            <div className="r-row">
              <Btn icon={<RotateCcw size={16} />} label={t('p_rotLeft')} onClick={() => rotateBy(-90)} />
              <Btn icon={<RotateCw size={16} />} label={t('p_rotRight')} onClick={() => rotateBy(90)} />
              <Btn icon={<FlipHorizontal2 size={16} />} label={t('p_flipH')} active={!!sels[0].flipH} onClick={() => patchSel((x) => ({ flipH: !x.flipH || undefined }))} />
              <Btn icon={<FlipVertical2 size={16} />} label={t('p_flipV')} active={!!sels[0].flipV} onClick={() => patchSel((x) => ({ flipV: !x.flipV || undefined }))} />
            </div>
            {sel && (
              <div className="r-row size-fields">
                <NumField label="W" value={sel.w} onChange={(v) => updateEl(sel.id, { w: Math.max(4, v) })} />
                <NumField label="H" value={sel.h} onChange={(v) => updateEl(sel.id, { h: Math.max(2, v) })} />
                <NumField label="°" value={sel.rot ?? 0} onChange={(v) => updateEl(sel.id, { rot: ((v % 360) + 360) % 360 || undefined })} />
              </div>
            )}
          </div>
        </Group>
      </>
    );
    if (tableSel) {
      return {
        label: t('p_tableDesign'),
        content: (
          <>
            <Group label={t('p_tableStyle')}>
              <div className="r-col">
                <Btn icon={<Rows3 size={16} />} label={t('p_headerRow')} showLabel active={tableSel.header} onClick={() => editTable((tb) => ({ header: !tb.header }))} />
                <Btn icon={<Rows3 size={16} />} label={t('p_banded')} showLabel active={tableSel.banded} onClick={() => editTable((tb) => ({ banded: !tb.banded }))} />
              </div>
              <ColorBtn icon={<PaintBucket size={16} />} label={t('p_tableColor')} color={tableSel.color} onPick={(c) => c && editTable(() => ({ color: c }))} />
            </Group>
            <Group label={t('p_rowsCols')}>
              <div className="r-col">
                <div className="r-row">
                  <Btn icon={<Rows3 size={16} />} label={t('p_rowAbove')} showLabel onClick={() => addRow(false)} />
                  <Btn icon={<Rows3 size={16} />} label={t('p_rowBelow')} showLabel onClick={() => addRow(true)} />
                </div>
                <div className="r-row">
                  <Btn icon={<Columns3 size={16} />} label={t('p_colLeft')} showLabel onClick={() => addCol(false)} />
                  <Btn icon={<Columns3 size={16} />} label={t('p_colRight')} showLabel onClick={() => addCol(true)} />
                </div>
              </div>
              <div className="r-col">
                <Btn icon={<Trash2 size={16} />} label={t('deleteRow')} showLabel disabled={tableSel.rows.length <= 1} onClick={delRow} />
                <Btn icon={<Trash2 size={16} />} label={t('deleteCol')} showLabel disabled={tableSel.rows[0].length <= 1} onClick={delCol} />
              </div>
            </Group>
            {common}
          </>
        ),
      };
    }
    if (sel?.type === 'chart') {
      const ch = sel;
      return {
        label: t('p_chartDesign'),
        content: (
          <>
            <Group label={t('p_chartType')}>
              {CHART_TYPES.map(([ct, Icon]) => <Btn key={ct} icon={<Icon size={18} />} label={t(`s_chart_${ct}` as Key)} active={ch.chart === ct} onClick={() => updateEl(ch.id, { chart: ct } as Partial<SlideEl>)} />)}
            </Group>
            <Group label={t('s_charts')}>
              <Btn big icon={<Pencil size={22} />} label={t('p_editData')} onClick={() => editChartData(ch)} />
              <Btn icon={<List size={16} />} label={t('s_legend')} showLabel active={ch.legend !== false} onClick={() => updateEl(ch.id, { legend: ch.legend === false } as Partial<SlideEl>)} />
            </Group>
            {common}
          </>
        ),
      };
    }
    return {
      label: allImg ? t('p_pictureFormat') : t('p_shapeFormat'),
      content: (
        <>
          {!allImg && (
            <Group label={t('p_shapeStyle')}>
              {shapeGallery(true)}
              <div className="r-col">
                <div className="r-row">
                  <ColorBtn icon={<PaintBucket size={16} />} label={t('fillColor')} color={fillTarget[0]?.fill} allowNone
                    onPick={(c) => mapEls(fillTarget.map((x) => x.id), (x) => ({ ...x, fill: c ?? (x.type === 'shape' ? 'transparent' : undefined) }) as SlideEl)} />
                  <ColorBtn icon={<PenLine size={16} />} label={t('outline')} color={shapes[0]?.stroke} allowNone
                    onPick={(c) => mapEls(shapes.map((x) => x.id), (x) => ({ ...x, stroke: c ?? undefined, strokeW: c ? (x as ShapeEl).strokeW || 2 : 0 }) as SlideEl)} />
                </div>
                {shapes.length > 0 && (
                  <div className="r-row">
                    <Select label={t('p_outlineW')} value={shapes[0].strokeW ?? 0} options={[0, 0.5, 1, 1.5, 2, 3, 4, 6, 8, 12].map((x) => ({ value: x, label: `${x} pt` }))}
                      onChange={(v) => mapEls(shapes.map((x) => x.id), (x) => ({ ...x, strokeW: v, stroke: (x as ShapeEl).stroke ?? (v ? '#1f2937' : undefined) }) as SlideEl)} width={70} />
                    <Select label={t('p_dash')} value={shapes[0].dash ?? 'solid'} options={(['solid', 'dash', 'dot'] as const).map((x) => ({ value: x, label: t(`p_dash_${x}` as Key) }))}
                      onChange={(v) => mapEls(shapes.map((x) => x.id), (x) => ({ ...x, dash: v === 'solid' ? undefined : v }) as SlideEl)} width={80} />
                  </div>
                )}
              </div>
            </Group>
          )}
          {common}
        </>
      ),
    };
  })();

  const ctxItems: ([React.ReactNode, string, () => void, string?] | null)[] = sels.length
    ? [
        [<Scissors size={16} />, t('s_cut'), () => copySel(true), 'Ctrl+X'],
        [<Copy size={16} />, t('copy'), () => copySel(), 'Ctrl+C'],
        [<ClipboardPaste size={16} />, t('s_paste'), pasteFromButton, 'Ctrl+V'],
        [<Copy size={16} />, t('duplicate'), () => addEls(cloneEls(sels, 20, 20)), 'Ctrl+D'],
        [<Trash2 size={16} />, t('delete'), deleteSel, 'Del'],
        null,
        [<BringToFront size={16} />, t('bringFront'), () => reorder('front'), 'Ctrl+Shift+]'],
        [<ArrowUpToLine size={16} />, t('p_bringForward'), () => reorder('forward'), 'Ctrl+]'],
        [<ArrowDownToLine size={16} />, t('p_sendBackward'), () => reorder('backward'), 'Ctrl+['],
        [<SendToBack size={16} />, t('sendBack'), () => reorder('back'), 'Ctrl+Shift+['],
        ...(sel?.type === 'chart' ? [null, [<Pencil size={16} />, t('p_editData'), () => editChartData(sel)] as [React.ReactNode, string, () => void]] : []),
      ]
    : [
        [<ClipboardPaste size={16} />, t('s_paste'), pasteFromButton, 'Ctrl+V'],
        [<BoxSelect size={16} />, t('s_selectAll'), () => select(slide.els.map((x) => x.id)), 'Ctrl+A'],
        null,
        [<Plus size={16} />, t('newSlide'), () => addSlide('content'), 'Ctrl+M'],
        [<Copy size={16} />, t('duplicate'), () => dupSlide()],
        [slide.hidden ? <Eye size={16} /> : <EyeOff size={16} />, slide.hidden ? t('p_showSlide') : t('p_hideSlide'), () => toggleHidden()],
      ];

  const tabs = [
    { id: 'home', label: t('tab_home'), content: homeTab },
    { id: 'insert', label: t('tab_insert'), content: insertTab },
    { id: 'design', label: t('tab_design'), content: designTab },
    { id: 'transitions', label: t('p_transitions'), content: transitionsTab },
    { id: 'show', label: t('tab_slideshow'), content: showTab },
  ];
  if (formatTab) tabs.push({ id: 'format', label: formatTab.label, content: formatTab.content });

  return (
    <div className="editor slides-editor" onCopy={(e) => onCopy(e)} onCut={(e) => onCopy(e, true)} onPaste={onPaste}>
      <Ribbon
        tabs={tabs}
        extra={
          <button className="present-btn" onClick={() => setPresenting({ start: idx, presenter: false })}>
            <Play size={14} /> {t('present')}
          </button>
        }
      />
      <input ref={fileRef} type="file" accept="image/*" hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) addImage(f); e.target.value = ''; }} />
      <div className="slides-body">
        <aside className="slide-rail" aria-label="Slides" onDragLeave={(e) => { if (!e.currentTarget.contains(e.relatedTarget as Node)) setDropAt(null); }}>
          {data.slides.map((s, i) => (
            <div
              key={s.id}
              className={`rail-item ${i === idx ? 'active' : ''} ${s.hidden ? 'is-hidden' : ''} ${dropAt === i ? 'drop-before' : ''} ${dropAt === i + 1 && i === data.slides.length - 1 ? 'drop-after' : ''}`}
              onClick={() => { setCur(i); select([]); stopEditing(); }}
              tabIndex={0}
              draggable
              onDragStart={(e) => { dragSlide.current = i; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', String(i)); }}
              onDragOver={(e) => {
                if (dragSlide.current === null) return;
                e.preventDefault();
                const r = e.currentTarget.getBoundingClientRect();
                setDropAt(e.clientY < r.top + r.height / 2 ? i : i + 1);
              }}
              onDrop={(e) => { e.preventDefault(); if (dragSlide.current !== null && dropAt !== null) moveSlideTo(dragSlide.current, dropAt); dragSlide.current = null; setDropAt(null); }}
              onDragEnd={() => { dragSlide.current = null; setDropAt(null); }}
            >
              <span className="rail-num">{i + 1}{s.hidden && <EyeOff size={11} />}</span>
              <div className="rail-thumb">
                <ScaledSlide slide={s} w={data.w} h={data.h} width={168} deco={decoFor(data, i)} />
                {s.transition && <span className="rail-tr" title={t(`tr_${s.transition}` as Key)}><Sparkles size={10} /></span>}
              </div>
              {i === idx && (
                <div className="rail-actions">
                  <button title={t('moveUp')} aria-label={t('moveUp')} onClick={(e) => { e.stopPropagation(); moveSlideTo(i, i - 1); }} disabled={i === 0}><ChevronUp size={14} /></button>
                  <button title={t('moveDown')} aria-label={t('moveDown')} onClick={(e) => { e.stopPropagation(); moveSlideTo(i, i + 2); }} disabled={i === data.slides.length - 1}><ChevronDown size={14} /></button>
                  <button title={s.hidden ? t('p_showSlide') : t('p_hideSlide')} aria-label={s.hidden ? t('p_showSlide') : t('p_hideSlide')} onClick={(e) => { e.stopPropagation(); toggleHidden(i); }}>{s.hidden ? <Eye size={13} /> : <EyeOff size={13} />}</button>
                  <button title={t('duplicate')} aria-label={t('duplicate')} onClick={(e) => { e.stopPropagation(); dupSlide(i); }}><Copy size={13} /></button>
                  <button title={t('delete')} aria-label={t('delete')} onClick={(e) => { e.stopPropagation(); delSlide(i); }}><Trash2 size={13} /></button>
                </div>
              )}
            </div>
          ))}
          <button className="rail-add" onClick={() => addSlide('content')}>
            <Plus size={16} /> {t('newSlide')}
          </button>
        </aside>
        <div className="slide-main">
          <div
            className="slide-stage"
            ref={stageRef}
            tabIndex={-1}
            onPointerDown={startMarquee}
            onContextMenu={(e) => { e.preventDefault(); setCtx({ x: e.clientX, y: e.clientY }); }}
          >
            <div className="slide-frame" style={{ width: data.w * fit, height: data.h * fit }}>
              <div key={preview.key} className={`slide-scale ${preview.key ? `tr-play tr-${preview.tr}` : ''}`} style={{ transform: `scale(${fit})`, width: data.w, height: data.h }}>
                <div ref={surfaceRef} className="slide-surface editing" style={{ width: data.w, height: data.h, background: slide.bg }}>
                  {slide.els.map((el) => {
                    const isEditing = el.id === editingId;
                    return (
                      <div
                        key={el.id}
                        className={`el-wrap ${selIds.includes(el.id) ? 'selected' : ''} ${isEditing ? 'is-editing' : ''}`}
                        onPointerDown={(e) => startMove(e, el)}
                        onDoubleClick={(e) => { e.stopPropagation(); startEditing(el, e.target as HTMLElement); }}
                        onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); if (!selIds.includes(el.id)) select([el.id]); setCtx({ x: e.clientX, y: e.clientY }); }}
                      >
                        <ElView
                          el={el}
                          placeholder={el.type === 'text' ? placeholderFor(el) : undefined}
                          cell={isEditing && el.type === 'table' ? (r, c, text) => (
                            <CellInput
                              value={text}
                              focused={cell?.r === r && cell?.c === c}
                              onFocus={() => setCell({ r, c })}
                              onChange={(v) => setCellText(live.current.slides[idx].els.find((x) => x.id === el.id) as TableEl, r, c, v)}
                              onNav={(dir) => {
                                const tb = live.current.slides[idx].els.find((x) => x.id === el.id) as TableEl;
                                const cols = tb.rows[0].length;
                                let n = r * cols + c + dir;
                                if (n >= tb.rows.length * cols) {
                                  const rows = [...tb.rows, tb.rows[0].map(() => '')];
                                  updateEl(tb.id, { rows, h: Math.round(tb.h + tb.h / tb.rows.length) }, false);
                                }
                                n = Math.max(0, n);
                                setCell({ r: Math.floor(n / cols), c: n % cols });
                              }}
                              onExit={stopEditing}
                            />
                          ) : undefined}
                        >
                          {isEditing && isTextual(el) ? (
                            <InlineText
                              value={el.text}
                              onDone={(text) => { setEditingId(null); focusStage(); if (text !== el.text) updateEl(el.id, { text } as Partial<SlideEl>); }}
                            />
                          ) : el.type === 'shape' ? <TextBody el={el} /> : undefined}
                        </ElView>
                      </div>
                    );
                  })}
                  <Decorations slide={slide} deco={decoFor(data, idx)} w={data.w} />
                  {sels.map((el) => {
                    const single = sels.length === 1 && el.id !== editingId;
                    return (
                      <div key={el.id} className={`sel-frame ${single ? '' : 'multi'}`} style={{ left: el.x, top: el.y, width: el.w, height: el.h, transform: el.rot ? `rotate(${el.rot}deg)` : undefined, ['--inv' as string]: 1 / fit }}>
                        {single && HANDLES.map((h) => <span key={h} className={`handle ${h}`} onPointerDown={(e) => startResize(e, el, h)} />)}
                        {single && <span className="rot-handle" title={t('rotate')} onPointerDown={(e) => startRotate(e, el)} />}
                      </div>
                    );
                  })}
                  {guides.v.map((x, i) => <div key={`v${i}`} className="guide v" style={{ left: x, ['--inv' as string]: 1 / fit }} />)}
                  {guides.h.map((y, i) => <div key={`h${i}`} className="guide h" style={{ top: y, ['--inv' as string]: 1 / fit }} />)}
                  {marquee && <div className="marquee" style={{ left: marquee.x, top: marquee.y, width: marquee.w, height: marquee.h, ['--inv' as string]: 1 / fit }} />}
                </div>
              </div>
            </div>
          </div>
          <textarea
            className="notes"
            dir="auto"
            placeholder={t('notesPlaceholder')}
            aria-label={t('notes')}
            value={slide.notes ?? ''}
            onFocus={beginEdit}
            onBlur={endEdit}
            onChange={(e) => {
              const slides = live.current.slides.slice();
              slides[idx] = { ...slides[idx], notes: e.target.value };
              onChange({ ...live.current, slides });
            }}
          />
        </div>
      </div>
      <div className="status-bar">
        <span>{t('slide')} {idx + 1} / {data.slides.length}{slide.hidden ? ` · ${t('p_hidden')}` : ''}</span>
        {sels.length > 1 && <span>{t('p_selected').replace('{n}', String(sels.length))}</span>}
        {sel && <span>{Math.round(sel.x)}, {Math.round(sel.y)} · {Math.round(sel.w)} × {Math.round(sel.h)}{sel.rot ? ` · ${sel.rot}°` : ''}</span>}
        <span className="grow" />
        <span>{Math.round(fit * 100)}%</span>
      </div>
      {ctx && (
        <div className="menu-pop ctx-menu" style={{ position: 'fixed', left: Math.max(8, Math.min(ctx.x, window.innerWidth - 270)), top: Math.max(8, Math.min(ctx.y, window.innerHeight - 420)) }} onContextMenu={(e) => e.preventDefault()}>
          {ctxItems.map((it, i) => (it ? <MenuItem key={i} icon={it[0]} label={it[1]} hint={it[3]} onClick={() => { setCtx(null); it[2](); focusStage(); }} /> : <div key={i} className="menu-sep" />))}
        </div>
      )}
      {ask && <AskDialog ask={ask} onClose={() => { setAsk(null); focusStage(); }} />}
      {presenting !== null && <Presenter deck={data} start={presenting.start} presenterView={presenting.presenter} onExit={(i) => { setPresenting(null); setCur(i); }} />}
    </div>
  );
}

function NumField({ label, value, onChange }: { label: string; value: number; onChange: (v: number) => void }) {
  const [v, setV] = useState(String(Math.round(value)));
  useEffect(() => setV(String(Math.round(value))), [value]);
  const done = () => { const n = parseFloat(v); if (Number.isFinite(n) && Math.round(n) !== Math.round(value)) onChange(Math.round(n)); else setV(String(Math.round(value))); };
  return (
    <label className="num-field" title={label}>
      <span>{label}</span>
      <input value={v} inputMode="numeric" onChange={(e) => setV(e.target.value)} onBlur={done} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); done(); } }} />
    </label>
  );
}

/** Hover a grid to choose the table size, like Office's table picker. */
function TablePicker({ label, onPick }: { label: string; onPick: (rows: number, cols: number) => void }) {
  const [h, setH] = useState({ r: 3, c: 4 });
  return (
    <div className="table-picker">
      <div className="tp-label">{label}: {h.r} × {h.c}</div>
      <div className="tp-grid" onMouseDown={(e) => e.preventDefault()}>
        {Array.from({ length: 8 }, (_, r) =>
          Array.from({ length: 10 }, (_, c) => (
            <button key={`${r}-${c}`} type="button" className={r < h.r && c < h.c ? 'on' : ''} aria-label={`${r + 1} × ${c + 1}`} onMouseEnter={() => setH({ r: r + 1, c: c + 1 })} onFocus={() => setH({ r: r + 1, c: c + 1 })} onClick={() => onPick(r + 1, c + 1)} />
          )),
        )}
      </div>
    </div>
  );
}

function CellInput({ value, focused, onFocus, onChange, onNav, onExit }: { value: string; focused: boolean; onFocus: () => void; onChange: (v: string) => void; onNav: (dir: number) => void; onExit: () => void }) {
  const ref = useRef<HTMLTextAreaElement>(null);
  useEffect(() => {
    if (focused && ref.current && document.activeElement !== ref.current) { ref.current.focus(); ref.current.select(); }
  }, [focused]);
  return (
    <textarea
      ref={ref}
      className="cell-input"
      dir="auto"
      rows={1}
      value={value}
      onFocus={onFocus}
      onChange={(e) => onChange(e.target.value)}
      onPointerDown={(e) => e.stopPropagation()}
      onKeyDown={(e) => {
        if (e.key === 'Tab') { e.preventDefault(); onNav(e.shiftKey ? -1 : 1); }
        else if (e.key === 'Escape') { e.preventDefault(); onExit(); }
      }}
    />
  );
}

function InlineText({ value, onDone }: { value: string; onDone: (v: string) => void }) {
  const ref = useRef<HTMLTextAreaElement>(null);
  const [v, setV] = useState(value);
  useEffect(() => {
    const el = ref.current;
    if (el) { el.focus(); el.select(); }
  }, []);
  return (
    <textarea
      ref={ref}
      className="inline-text"
      dir="auto"
      value={v}
      onChange={(e) => setV(e.target.value)}
      onPointerDown={(e) => e.stopPropagation()}
      onBlur={() => onDone(v)}
      onKeyDown={(e) => {
        if (e.key === 'Escape') { e.preventDefault(); onDone(v); }
        if (e.key === 'Tab') {
          // indent / outdent bullet line
          e.preventDefault();
          const el = e.currentTarget;
          const start = v.lastIndexOf('\n', el.selectionStart - 1) + 1;
          const next = e.shiftKey ? (v.slice(start).startsWith('  ') ? v.slice(0, start) + v.slice(start + 2) : v) : v.slice(0, start) + '  ' + v.slice(start);
          setV(next);
        }
      }}
    />
  );
}

const fmtTime = (s: number) => `${Math.floor(s / 3600) ? `${Math.floor(s / 3600)}:` : ''}${String(Math.floor(s / 60) % 60).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`;

/** The slide show. Hidden slides are skipped; the presenter view adds the next slide, notes and a timer. */
function Presenter({ deck, start, presenterView, onExit }: { deck: Deck; start: number; presenterView: boolean; onExit: (i: number) => void }) {
  const { t } = useI18n();
  const order = deck.slides.map((s, i) => (s.hidden && i !== start ? -1 : i)).filter((i) => i >= 0);
  const [pos, setPos] = useState(Math.max(0, order.indexOf(start)));
  const [ended, setEnded] = useState(false);
  const [black, setBlack] = useState(false);
  const [pv, setPv] = useState(presenterView);
  const [secs, setSecs] = useState(0);
  const [size, setSize] = useState({ w: window.innerWidth, h: window.innerHeight });
  const ref = useRef<HTMLDivElement>(null);
  const state = useRef({ pos, ended });
  state.current = { pos, ended };
  const i = order[pos];
  const exitAt = () => onExit(order[state.current.pos]);

  const next = () => { if (state.current.ended) return exitFs(); if (state.current.pos >= order.length - 1) setEnded(true); else setPos((p) => p + 1); setBlack(false); };
  const prev = () => { if (state.current.ended) setEnded(false); else setPos((p) => Math.max(0, p - 1)); setBlack(false); };
  const exitFs = () => { if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); else exitAt(); };

  useEffect(() => {
    ref.current?.requestFullscreen?.().catch(() => {});
    const onResize = () => setSize({ w: window.innerWidth, h: window.innerHeight });
    const onKey = (e: KeyboardEvent) => {
      if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n'].includes(e.key)) { e.preventDefault(); next(); }
      else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p'].includes(e.key) && !(e.key === 'p' && e.shiftKey)) { e.preventDefault(); prev(); }
      else if (e.key === 'Home') { setEnded(false); setPos(0); }
      else if (e.key === 'End') { setEnded(false); setPos(order.length - 1); }
      else if (e.key === 'b' || e.key === 'B' || e.key === '.') setBlack((b) => !b);
      else if (e.key === 'P') setPv((v) => !v);
      else if (e.key === 'Escape') exitFs();
    };
    const onFs = () => { if (!document.fullscreenElement) exitAt(); };
    const timer = window.setInterval(() => setSecs((s) => s + 1), 1000);
    window.addEventListener('resize', onResize);
    window.addEventListener('keydown', onKey);
    document.addEventListener('fullscreenchange', onFs);
    return () => {
      window.clearInterval(timer);
      window.removeEventListener('resize', onResize);
      window.removeEventListener('keydown', onKey);
      document.removeEventListener('fullscreenchange', onFs);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const slideAt = (k: number, w: number, h: number, animate: boolean) => {
    const s = Math.min(w / deck.w, h / deck.h);
    const sl = deck.slides[k];
    return (
      <div className="show-slide" style={{ width: deck.w * s, height: deck.h * s }}>
        <div key={k} className={animate && sl.transition ? `tr-play tr-${sl.transition}` : ''} style={{ transform: `scale(${s})`, transformOrigin: '0 0', width: deck.w, height: deck.h }}>
          <SlideView slide={sl} w={deck.w} h={deck.h} deco={decoFor(deck, k)} />
        </div>
      </div>
    );
  };
  const endCard = <div className="show-end">{t('p_endShow')}</div>;

  if (pv) {
    const nextI = order[pos + 1];
    const mainW = size.w * 0.62, mainH = size.h - 150;
    return (
      <div className="presenter pv" ref={ref}>
        <header className="pv-bar">
          <span className="pv-timer"><Timer size={16} /> {fmtTime(secs)}</span>
          <span>{t('slide')} {pos + 1} / {order.length}</span>
          <span className="pv-clock">{new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
          <span className="grow" />
          <button onClick={() => setPv(false)}>{t('present')}</button>
          <button onClick={exitFs} aria-label={t('close')}><X size={16} /></button>
        </header>
        <div className="pv-main">
          <div className="pv-current">{ended ? endCard : black ? <div className="show-black" /> : slideAt(i, mainW, mainH, true)}</div>
          <div className="pv-side">
            <div className="pv-label">{t('p_next')}</div>
            <div className="pv-next">{nextI !== undefined ? slideAt(nextI, size.w * 0.3, size.h * 0.3, false) : endCard}</div>
            <div className="pv-label">{t('notes')}</div>
            <div className="pv-notes" dir="auto">{deck.slides[i]?.notes || <span className="muted">{t('p_noNotes')}</span>}</div>
          </div>
        </div>
        <footer className="pv-nav">
          <button onClick={prev} aria-label={t('p_prev')}><ChevronLeft size={20} /></button>
          <button onClick={next} aria-label={t('p_next')}><ChevronRight size={20} /></button>
        </footer>
      </div>
    );
  }

  return (
    <div className="presenter" ref={ref} onClick={(e) => (e.clientX < size.w / 4 ? prev() : next())}>
      {ended ? endCard : black ? <div className="show-black" /> : slideAt(i, size.w, size.h, true)}
      <div className="presenter-hint">{pos + 1} / {order.length} · {t('p_showHint')}</div>
    </div>
  );
}
