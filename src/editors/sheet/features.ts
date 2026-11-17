/**
 * Spreadsheet tools that do not need the UI: merged cells, borders, autofill series,
 * conditional formatting, filters, find, and chart data. Kept pure so they can be tested.
 */
import { formatValue, type Scalar } from './formula';
import { addr, colName, inRange, parseAddr, type Borders, type Cell, type CellStyle, type Range, type Sheet } from './model';

export type ValueAt = (r: number, c: number) => Scalar;

// ---------- merged cells ----------

export const mergeAt = (sh: Sheet, r: number, c: number): Range | undefined => sh.merges?.find((m) => inRange(m, r, c));

/** Grow a selection until every merge it touches lies fully inside it. */
export function expandToMerges(sh: Sheet, sel: Range): Range {
  if (!sh.merges?.length) return sel;
  const out = { ...sel };
  let grew = true;
  while (grew) {
    grew = false;
    for (const m of sh.merges) {
      if (m.r1 <= out.r2 && out.r1 <= m.r2 && m.c1 <= out.c2 && out.c1 <= m.c2) {
        if (m.r1 < out.r1) { out.r1 = m.r1; grew = true; }
        if (m.c1 < out.c1) { out.c1 = m.c1; grew = true; }
        if (m.r2 > out.r2) { out.r2 = m.r2; grew = true; }
        if (m.c2 > out.c2) { out.c2 = m.c2; grew = true; }
      }
    }
  }
  return out;
}

/** Merge the selection: keeps the top-left value (like Excel) and drops merges it covers. */
export function mergeCells(sh: Sheet, sel: Range, center = true): Sheet {
  if (sel.r1 === sel.r2 && sel.c1 === sel.c2) return sh;
  const cells = { ...sh.cells };
  for (let r = sel.r1; r <= sel.r2; r++)
    for (let c = sel.c1; c <= sel.c2; c++) {
      if (r === sel.r1 && c === sel.c1) continue;
      const k = addr(r, c);
      if (cells[k]?.st) cells[k] = { v: '', st: cells[k].st };
      else delete cells[k];
    }
  const k0 = addr(sel.r1, sel.c1);
  if (center) cells[k0] = { ...cells[k0], v: cells[k0]?.v ?? '', st: { ...cells[k0]?.st, align: 'center', va: 'middle' } };
  const merges = (sh.merges ?? []).filter((m) => !(m.r1 <= sel.r2 && sel.r1 <= m.r2 && m.c1 <= sel.c2 && sel.c1 <= m.c2));
  return { ...sh, cells, merges: [...merges, { ...sel }] };
}

export function unmergeCells(sh: Sheet, sel: Range): Sheet {
  const merges = (sh.merges ?? []).filter((m) => !(m.r1 <= sel.r2 && sel.r1 <= m.r2 && m.c1 <= sel.c2 && sel.c1 <= m.c2));
  return { ...sh, merges: merges.length ? merges : undefined };
}

// ---------- borders ----------

export type BorderMode = 'all' | 'outer' | 'thickOuter' | 'top' | 'bottom' | 'thickBottom' | 'left' | 'right' | 'inner' | 'none';

export function bordersOf(st?: CellStyle): Borders | undefined {
  if (st?.bd) return st.bd;
  if (st?.border) return { t: 1, r: 1, b: 1, l: 1 };
  return undefined;
}

/** Inset box-shadows draw borders inside the cell without changing its size. */
export function borderShadow(bd?: Borders, color = '#1f2937'): string | undefined {
  if (!bd) return undefined;
  const s: string[] = [];
  if (bd.t) s.push(`inset 0 ${bd.t}px 0 ${color}`);
  if (bd.b) s.push(`inset 0 -${bd.b}px 0 ${color}`);
  if (bd.l) s.push(`inset ${bd.l}px 0 0 ${color}`);
  if (bd.r) s.push(`inset -${bd.r}px 0 0 ${color}`);
  return s.length ? s.join(', ') : undefined;
}

export function applyBorders(sh: Sheet, sel: Range, mode: BorderMode): Sheet {
  const cells = { ...sh.cells };
  for (let r = sel.r1; r <= sel.r2; r++)
    for (let c = sel.c1; c <= sel.c2; c++) {
      const k = addr(r, c);
      const cell = cells[k];
      const prev: Borders = { ...bordersOf(cell?.st) };
      const top = r === sel.r1, bottom = r === sel.r2, left = c === sel.c1, right = c === sel.c2;
      let bd: Borders = prev;
      switch (mode) {
        case 'none': bd = {}; break;
        case 'all': bd = { t: 1, r: 1, b: 1, l: 1 }; break;
        case 'outer':
        case 'thickOuter': {
          const w = mode === 'outer' ? 1 : 2;
          bd = { ...prev, ...(top && { t: w }), ...(bottom && { b: w }), ...(left && { l: w }), ...(right && { r: w }) };
          break;
        }
        case 'inner':
          bd = { ...prev, ...(!top && { t: 1 }), ...(!bottom && { b: 1 }), ...(!left && { l: 1 }), ...(!right && { r: 1 }) };
          break;
        case 'top': if (top) bd = { ...prev, t: 1 }; break;
        case 'bottom': if (bottom) bd = { ...prev, b: 1 }; break;
        case 'thickBottom': if (bottom) bd = { ...prev, b: 2 }; break;
        case 'left': if (left) bd = { ...prev, l: 1 }; break;
        case 'right': if (right) bd = { ...prev, r: 1 }; break;
      }
      const has = Object.values(bd).some(Boolean);
      const st: CellStyle = { ...cell?.st, bd: has ? bd : undefined, border: undefined };
      for (const key of Object.keys(st) as (keyof CellStyle)[]) if (st[key] === undefined) delete st[key];
      const next: Cell = { ...cell, v: cell?.v ?? '', st: Object.keys(st).length ? st : undefined };
      if (!next.st) delete next.st;
      if (next.v === '' && !next.st && !next.note && !next.list) delete cells[k];
      else cells[k] = next;
    }
  return { ...sh, cells };
}

// ---------- autofill ----------

const SERIES_LISTS: string[][] = [
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'],
  ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'],
  ['كانون الثاني', 'شباط', 'آذار', 'نيسان', 'أيار', 'حزيران', 'تموز', 'آب', 'أيلول', 'تشرين الأول', 'تشرين الثاني', 'كانون الأول'],
  ['Q1', 'Q2', 'Q3', 'Q4'],
];

function listOf(v: string): { list: string[]; i: number; upper: boolean; lower: boolean } | null {
  const t = v.trim();
  for (const list of SERIES_LISTS) {
    const i = list.findIndex((x) => x.toLowerCase() === t.toLowerCase());
    if (i >= 0) return { list, i, upper: t === t.toUpperCase() && t !== t.toLowerCase(), lower: t === t.toLowerCase() && t !== t.toUpperCase() };
  }
  return null;
}

const NUM_RE = /^[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?$/;
const TRAIL_RE = /^(.*?)(\d+)(\D*)$/;

/** Least-squares step for a numeric run, so 1,2 → 3,4 and 2,4,7 → a best-fit trend like Excel. */
function linear(vals: number[]) {
  const n = vals.length;
  if (n === 1) return { a: vals[0], b: 0 };
  const mx = (n - 1) / 2;
  const my = vals.reduce((s, v) => s + v, 0) / n;
  let num = 0, den = 0;
  vals.forEach((v, i) => { num += (i - mx) * (v - my); den += (i - mx) ** 2; });
  const b = num / den;
  return { a: my - b * mx, b };
}

const tidy = (x: number) => String(parseFloat(x.toPrecision(12)));

/**
 * Extend `src` (the cells being dragged, in fill order) by `n` cells.
 * `relocate(v, k)` moves a formula k cells along the fill direction.
 * Single numbers are copied (Excel's default), dates step by a day, words like
 * "Monday" or "Item 3" continue their series, and a mixed block repeats.
 */
export function fillSeries(src: (Cell | undefined)[], n: number, relocate: (v: string, k: number) => string): (Cell | undefined)[] {
  const k = src.length;
  const raw = src.map((c) => c?.v ?? '');
  const out: (Cell | undefined)[] = [];
  const copy = (i: number, v: string): Cell | undefined => {
    const s = src[i % k];
    if (!s && v === '') return undefined;
    return { ...s, v };
  };
  const allNum = raw.every((v) => NUM_RE.test(v.trim()));
  const isDate = src[0]?.st?.fmt === 'date' || src[0]?.st?.fmt === 'datetime';
  if (allNum && (k > 1 || isDate)) {
    const { a, b } = k > 1 ? linear(raw.map(Number)) : { a: Number(raw[0]), b: 1 };
    for (let i = 0; i < n; i++) out.push(copy(i, tidy(a + b * (k + i))));
    return out;
  }
  const lists = raw.map(listOf);
  if (lists.every((l) => l && l.list === lists[0]!.list)) {
    const L = lists[0]!.list;
    const step = k > 1 ? (((lists[1]!.i - lists[0]!.i) % L.length) + L.length) % L.length || 1 : 1;
    const last = lists[k - 1]!;
    for (let i = 0; i < n; i++) {
      let w = L[(last.i + step * (i + 1)) % L.length];
      if (last.upper) w = w.toUpperCase();
      else if (last.lower) w = w.toLowerCase();
      out.push(copy(i, w));
    }
    return out;
  }
  const trails = raw.map((v) => (v.startsWith('=') ? null : TRAIL_RE.exec(v)));
  if (!allNum && trails.every((m) => m && m[1] === trails[0]![1] && m[3] === trails[0]![3])) {
    const nums = trails.map((m) => Number(m![2]));
    const { a, b } = k > 1 ? linear(nums) : { a: nums[0], b: 1 };
    const width = trails[0]![2].length;
    for (let i = 0; i < n; i++) {
      const x = Math.round(a + b * (k + i));
      const digits = String(Math.abs(x)).padStart(trails[0]![2].startsWith('0') ? width : 0, '0');
      out.push(copy(i, `${trails[0]![1]}${x < 0 ? '-' : ''}${digits}${trails[0]![3]}`));
    }
    return out;
  }
  for (let i = 0; i < n; i++) {
    const j = i % k;
    const v = raw[j];
    out.push(copy(i, v.startsWith('=') ? relocate(v, k + i - j) : v));
  }
  return out;
}

// ---------- conditional formatting ----------

export interface CFOut {
  bg?: string;
  color?: string;
  /** Data bar: fraction of the cell width, and color. */
  bar?: { pct: number; color: string };
}

const num = (v: Scalar) => (typeof v === 'number' ? v : typeof v === 'string' && NUM_RE.test(v.trim()) ? Number(v) : null);

function mix(a: string, b: string, t: number) {
  const pa = parseInt(a.slice(1), 16), pb = parseInt(b.slice(1), 16);
  const ch = (p: number, s: number) => (p >> s) & 255;
  const m = (s: number) => Math.round(ch(pa, s) + (ch(pb, s) - ch(pa, s)) * t);
  return `#${((m(16) << 16) | (m(8) << 8) | m(0)).toString(16).padStart(6, '0')}`;
}

/** The last row/column that holds anything, to keep whole-column rules cheap. */
export function usedExtent(sh: Sheet) {
  let r = -1, c = -1;
  for (const k of Object.keys(sh.cells)) {
    const p = parseAddr(k);
    if (p) { r = Math.max(r, p.r); c = Math.max(c, p.c); }
  }
  return { r, c };
}

export function evalCF(sh: Sheet, valueAt: ValueAt): Map<string, CFOut> {
  const out = new Map<string, CFOut>();
  if (!sh.cf?.length) return out;
  const ext = usedExtent(sh);
  const put = (r: number, c: number, o: CFOut) => {
    const k = addr(r, c);
    out.set(k, { ...out.get(k), ...o });
  };
  for (const rule of sh.cf) {
    const rg = { ...rule.range, r2: Math.min(rule.range.r2, ext.r), c2: Math.min(rule.range.c2, ext.c) };
    const cells: { r: number; c: number; v: Scalar }[] = [];
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) cells.push({ r, c, v: valueAt(r, c) });
    const numbers = cells.map((x) => num(x.v)).filter((x): x is number => x !== null);
    const hl: CFOut = { bg: rule.bg ?? '#ffc7ce', color: rule.color ?? '#9c0006' };
    const a = num(rule.a ?? ''), b = num(rule.b ?? '');
    let test: ((v: Scalar) => boolean) | null = null;
    switch (rule.kind) {
      case 'gt': test = (v) => num(v) !== null && a !== null && num(v)! > a; break;
      case 'lt': test = (v) => num(v) !== null && a !== null && num(v)! < a; break;
      case 'between': test = (v) => num(v) !== null && a !== null && b !== null && num(v)! >= Math.min(a, b) && num(v)! <= Math.max(a, b); break;
      case 'eq': test = (v) => (a !== null && num(v) !== null ? num(v) === a : String(v ?? '').toLowerCase() === (rule.a ?? '').toLowerCase()); break;
      case 'text': test = (v) => !!rule.a && String(v ?? '').toLowerCase().includes(rule.a.toLowerCase()); break;
      case 'blank': test = (v) => v === null || v === ''; break;
      case 'dup': {
        const counts = new Map<string, number>();
        for (const x of cells) if (x.v !== null && x.v !== '') counts.set(String(x.v).toLowerCase(), (counts.get(String(x.v).toLowerCase()) ?? 0) + 1);
        test = (v) => v !== null && v !== '' && (counts.get(String(v).toLowerCase()) ?? 0) > 1;
        break;
      }
      case 'top':
      case 'bottom': {
        const sorted = [...numbers].sort((x, y) => (rule.kind === 'top' ? y - x : x - y));
        const cut = sorted[Math.max(0, Math.min(sorted.length, a ?? 10) - 1)];
        test = (v) => num(v) !== null && cut !== undefined && (rule.kind === 'top' ? num(v)! >= cut : num(v)! <= cut);
        break;
      }
      case 'aboveAvg': {
        const avg = numbers.reduce((s, x) => s + x, 0) / (numbers.length || 1);
        test = (v) => num(v) !== null && num(v)! > avg;
        break;
      }
      case 'bar': {
        if (!numbers.length) break;
        const lo = Math.min(0, ...numbers), hi = Math.max(...numbers);
        for (const x of cells) {
          const n = num(x.v);
          if (n !== null) put(x.r, x.c, { bar: { pct: hi === lo ? 1 : (n - lo) / (hi - lo), color: rule.c1 ?? '#638ec6' } });
        }
        break;
      }
      case 'scale': {
        if (!numbers.length) break;
        const lo = Math.min(...numbers), hi = Math.max(...numbers);
        const low = rule.c2 ?? '#f8696b', high = rule.c1 ?? '#63be7b';
        for (const x of cells) {
          const n = num(x.v);
          if (n === null) continue;
          const t = hi === lo ? 0.5 : (n - lo) / (hi - lo);
          const bg = rule.mid ? (t < 0.5 ? mix(low, rule.mid, t * 2) : mix(rule.mid, high, (t - 0.5) * 2)) : mix(low, high, t);
          put(x.r, x.c, { bg });
        }
        break;
      }
    }
    if (test) for (const x of cells) if (test(x.v)) put(x.r, x.c, hl);
  }
  return out;
}

// ---------- filter ----------

export const BLANK = '\u0000blank';

export function displayAt(sh: Sheet, valueAt: ValueAt, r: number, c: number) {
  const st = sh.cells[addr(r, c)]?.st;
  return formatValue(valueAt(r, c), st?.fmt, st?.dec);
}

/** Last row of the filter: grows to take in rows typed under it. */
export function filterEnd(sh: Sheet): number {
  const f = sh.filter;
  if (!f) return -1;
  let end = f.range.r2;
  for (const k of Object.keys(sh.cells)) {
    const p = parseAddr(k)!;
    if (p.c >= f.range.c1 && p.c <= f.range.c2 && p.r > end && sh.cells[k].v !== '') end = p.r;
  }
  return end;
}

/** Rows that the filter hides. */
export function filteredOut(sh: Sheet, valueAt: ValueAt): Set<number> {
  const hidden = new Set<number>();
  const f = sh.filter;
  if (!f) return hidden;
  const cols = Object.keys(f.show).map(Number);
  if (!cols.length) return hidden;
  const end = filterEnd(sh);
  const sets = new Map(cols.map((c) => [c, new Set(f.show[c])]));
  for (let r = f.range.r1 + 1; r <= end; r++)
    for (const c of cols) {
      const d = displayAt(sh, valueAt, r, c);
      if (!sets.get(c)!.has(d === '' ? BLANK : d)) { hidden.add(r); break; }
    }
  return hidden;
}

/** The distinct values in a filter column, numbers first in order, then text. */
export function filterValues(sh: Sheet, valueAt: ValueAt, c: number): string[] {
  const f = sh.filter;
  if (!f) return [];
  const end = filterEnd(sh);
  const seen = new Map<string, Scalar>();
  for (let r = f.range.r1 + 1; r <= end; r++) {
    const d = displayAt(sh, valueAt, r, c);
    const key = d === '' ? BLANK : d;
    if (!seen.has(key)) seen.set(key, valueAt(r, c));
  }
  return [...seen.entries()]
    .sort(([ka, a], [kb, b]) => {
      if (ka === BLANK) return 1;
      if (kb === BLANK) return -1;
      if (typeof a === 'number' && typeof b === 'number') return a - b;
      if (typeof a === 'number') return -1;
      if (typeof b === 'number') return 1;
      return ka.localeCompare(kb, undefined, { numeric: true });
    })
    .map(([k]) => k);
}

/** The block of filled cells around (r, c), like Excel's "current region". */
export function currentRegion(sh: Sheet, r: number, c: number): Range {
  const has = (rr: number, cc: number) => rr >= 0 && cc >= 0 && (sh.cells[addr(rr, cc)]?.v ?? '') !== '';
  const rg = { r1: r, c1: c, r2: r, c2: c };
  let grew = true;
  while (grew) {
    grew = false;
    const rowHas = (rr: number) => { for (let cc = rg.c1 - 1; cc <= rg.c2 + 1; cc++) if (has(rr, cc)) return true; return false; };
    const colHas = (cc: number) => { for (let rr = rg.r1 - 1; rr <= rg.r2 + 1; rr++) if (has(rr, cc)) return true; return false; };
    if (rg.r1 > 0 && rowHas(rg.r1 - 1)) { rg.r1--; grew = true; }
    if (rowHas(rg.r2 + 1)) { rg.r2++; grew = true; }
    if (rg.c1 > 0 && colHas(rg.c1 - 1)) { rg.c1--; grew = true; }
    if (colHas(rg.c2 + 1)) { rg.c2++; grew = true; }
  }
  return rg;
}

// ---------- find ----------

export interface FindOpts {
  matchCase?: boolean;
  whole?: boolean;
  /** Look in formulas (raw input) instead of shown values. */
  formulas?: boolean;
}

export function findCells(sh: Sheet, valueAt: ValueAt, q: string, o: FindOpts): { r: number; c: number }[] {
  if (!q) return [];
  const needle = o.matchCase ? q : q.toLowerCase();
  const hits: { r: number; c: number }[] = [];
  for (const k of Object.keys(sh.cells)) {
    const p = parseAddr(k)!;
    let text = o.formulas ? sh.cells[k].v : displayAt(sh, valueAt, p.r, p.c);
    if (!o.matchCase) text = text.toLowerCase();
    if (o.whole ? text === needle : text.includes(needle)) hits.push(p);
  }
  return hits.sort((a, b) => a.r - b.r || a.c - b.c);
}

export function replaceIn(v: string, q: string, rep: string, o: FindOpts): string {
  if (o.whole) return (o.matchCase ? v === q : v.toLowerCase() === q.toLowerCase()) ? rep : v;
  const esc = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return v.replace(new RegExp(esc, o.matchCase ? 'g' : 'gi'), () => rep);
}

// ---------- charts ----------

export interface ChartData {
  cats: string[];
  series: { name: string; values: (number | null)[] }[];
}

/**
 * Read a chart's range. A text first row names the series and a text first column
 * holds the categories. Series run down columns unless the block is wider than tall.
 */
export function chartData(rg: Range, valueAt: ValueAt, fmt: (r: number, c: number) => string): ChartData {
  const isText = (v: Scalar) => typeof v === 'string' && v !== '';
  const firstRowText = (() => { for (let c = rg.c1; c <= rg.c2; c++) if (isText(valueAt(rg.r1, c))) return true; return false; })();
  const firstColText = (() => { for (let r = rg.r1 + (firstRowText ? 1 : 0); r <= rg.r2; r++) if (isText(valueAt(r, rg.c1))) return true; return false; })();
  const hasHead = firstRowText;
  const hasCats = firstColText || (hasHead && !isText(valueAt(rg.r1, rg.c1)) && valueAt(rg.r1, rg.c1) === null && rg.c2 > rg.c1);
  const r0 = rg.r1 + (hasHead ? 1 : 0), c0 = rg.c1 + (hasCats ? 1 : 0);
  const rows = rg.r2 - r0 + 1, cols = rg.c2 - c0 + 1;
  const byCol = rows >= cols;
  const cats: string[] = [];
  const series: ChartData['series'] = [];
  const n = (v: Scalar) => (typeof v === 'number' ? v : typeof v === 'boolean' ? Number(v) : null);
  if (byCol) {
    for (let r = r0; r <= rg.r2; r++) cats.push(hasCats ? fmt(r, rg.c1) : String(r - r0 + 1));
    for (let c = c0; c <= rg.c2; c++) {
      const values: (number | null)[] = [];
      for (let r = r0; r <= rg.r2; r++) values.push(n(valueAt(r, c)));
      series.push({ name: (hasHead && fmt(rg.r1, c)) || colName(c), values });
    }
  } else {
    for (let c = c0; c <= rg.c2; c++) cats.push(hasHead ? fmt(rg.r1, c) : String(c - c0 + 1));
    for (let r = r0; r <= rg.r2; r++) {
      const values: (number | null)[] = [];
      for (let c = c0; c <= rg.c2; c++) values.push(n(valueAt(r, c)));
      series.push({ name: (hasCats && fmt(r, rg.c1)) || String(r + 1), values });
    }
  }
  return { cats, series };
}

export const CHART_COLORS = ['#4f46e5', '#0ea5a4', '#f5a524', '#e0436b', '#16a34a', '#7c3aed', '#0284c7', '#ea580c', '#64748b', '#db2777'];
