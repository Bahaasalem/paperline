export type NumFmt = 'general' | 'number' | 'currency' | 'percent' | 'date' | 'time' | 'datetime' | 'scientific' | 'text';
/** Border width per side: 1 thin, 2 thick. */
export interface Borders {
  t?: number;
  r?: number;
  b?: number;
  l?: number;
}

export interface CellStyle {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  s?: boolean;
  color?: string;
  bg?: string;
  align?: 'left' | 'center' | 'right';
  fmt?: NumFmt;
  dec?: number;
  fontSize?: number;
  /** Font family; Calibri (Excel's default) when unset. */
  font?: string;
  wrap?: boolean;
  /** Legacy: thin border on every side. New files use `bd`. */
  border?: boolean;
  bd?: Borders;
  va?: 'top' | 'middle' | 'bottom';
}

export interface Cell {
  /** Raw input: a literal ("12", "hello") or a formula ("=SUM(A1:A3)") */
  v: string;
  st?: CellStyle;
  /** A note shown when hovering the cell. */
  note?: string;
  /** Data validation: the cell offers this list as a dropdown. */
  list?: string[];
}

export type CFKind = 'gt' | 'lt' | 'between' | 'eq' | 'text' | 'dup' | 'top' | 'bottom' | 'aboveAvg' | 'blank' | 'bar' | 'scale';
export interface CFRule {
  id: string;
  range: Range;
  kind: CFKind;
  a?: string;
  b?: string;
  /** Highlight style for rule kinds that color matching cells. */
  bg?: string;
  color?: string;
  /** Bar color, or the high color of a color scale. */
  c1?: string;
  /** Low color of a color scale. */
  c2?: string;
  /** Optional middle color of a color scale. */
  mid?: string;
}

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';
export interface ChartSpec {
  id: string;
  type: ChartType;
  range: Range;
  title?: string;
  /** Position and size in unzoomed grid pixels (relative to the cell area). */
  x: number;
  y: number;
  w: number;
  h: number;
  legend?: boolean;
}

export interface SheetFilter {
  range: Range;
  /** Column index → display values to show. A column with no entry is not filtered. */
  show: Record<number, string[]>;
}

export interface Sheet {
  name: string;
  rows: number;
  cols: number;
  cells: Record<string, Cell>;
  colW: Record<number, number>;
  rowH: Record<number, number>;
  freezeTop?: boolean;
  merges?: Range[];
  hiddenR?: number[];
  hiddenC?: number[];
  filter?: SheetFilter;
  cf?: CFRule[];
  charts?: ChartSpec[];
}

export interface Workbook {
  sheets: Sheet[];
  active: number;
}

export const DEFAULT_COL_W = 96;
export const DEFAULT_ROW_H = 26;

export function colName(c: number): string {
  let s = '';
  c += 1;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

export function colIndex(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export const addr = (r: number, c: number) => `${colName(c)}${r + 1}`;

export function parseAddr(a: string): { r: number; c: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a);
  if (!m) return null;
  return { r: parseInt(m[2], 10) - 1, c: colIndex(m[1]) };
}

export function newSheet(name: string): Sheet {
  return { name, rows: 100, cols: 26, cells: {}, colW: {}, rowH: {} };
}

export function newWorkbook(sheetLabel = 'Sheet'): Workbook {
  return { sheets: [newSheet(`${sheetLabel}1`)], active: 0 };
}

export interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export function normRange(r: Range): Range {
  return {
    r1: Math.min(r.r1, r.r2),
    r2: Math.max(r.r1, r.r2),
    c1: Math.min(r.c1, r.c2),
    c2: Math.max(r.c1, r.c2),
  };
}

/** Shift every cell reference in a formula when rows/cols are inserted or removed. */
export function shiftFormula(f: string, axis: 'r' | 'c', at: number, delta: number): string {
  if (!f.startsWith('=')) return f;
  return f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])/g, (_m, d1, col, d2, row) => {
    let c = colIndex(col);
    let r = parseInt(row, 10) - 1;
    if (axis === 'c' && c >= at) c += delta;
    if (axis === 'r' && r >= at) r += delta;
    if (c < 0 || r < 0) return '#REF!';
    return `${d1}${colName(c)}${d2}${r + 1}`;
  });
}

/** Insert (delta>0) or delete (delta<0) rows/cols starting at index `at`. */
export function shiftSheet(sh: Sheet, axis: 'r' | 'c', at: number, delta: number): Sheet {
  const cells: Record<string, Cell> = {};
  for (const [k, cell] of Object.entries(sh.cells)) {
    const p = parseAddr(k)!;
    let { r, c } = p;
    const idx = axis === 'r' ? r : c;
    if (delta < 0 && idx >= at && idx < at - delta) continue; // deleted
    if (idx >= at) {
      if (axis === 'r') r += delta;
      else c += delta;
    }
    cells[addr(r, c)] = { ...cell, v: shiftFormula(cell.v, axis, at, delta) };
  }
  const shiftMap = (m: Record<number, number>) => {
    const out: Record<number, number> = {};
    for (const [k, v] of Object.entries(m)) {
      const i = Number(k);
      if (delta < 0 && i >= at && i < at - delta) continue;
      out[i >= at ? i + delta : i] = v;
    }
    return out;
  };
  const shiftIdx = (list?: number[]) => {
    if (!list) return list;
    const out = list.filter((i) => !(delta < 0 && i >= at && i < at - delta)).map((i) => (i >= at ? i + delta : i));
    return out.length ? out : undefined;
  };
  const merges = sh.merges?.map((m) => shiftRange(m, axis, at, delta)).filter((m): m is Range => !!m && (m.r2 > m.r1 || m.c2 > m.c1));
  const cf = sh.cf?.map((x) => ({ ...x, range: shiftRange(x.range, axis, at, delta) })).filter((x): x is CFRule => !!x.range);
  const charts = sh.charts?.map((x) => ({ ...x, range: shiftRange(x.range, axis, at, delta) })).filter((x): x is ChartSpec => !!x.range);
  const fr = sh.filter && shiftRange(sh.filter.range, axis, at, delta);
  let filter: SheetFilter | undefined;
  if (sh.filter && fr) {
    const show: Record<number, string[]> = {};
    for (const [k, v] of Object.entries(sh.filter.show)) {
      const i = Number(k);
      if (axis === 'c' && delta < 0 && i >= at && i < at - delta) continue;
      show[axis === 'c' && i >= at ? i + delta : i] = v;
    }
    filter = { range: fr, show };
  }
  return {
    ...sh,
    cells,
    merges: merges?.length ? merges : undefined,
    cf: cf?.length ? cf : undefined,
    charts: charts?.length ? charts : undefined,
    filter,
    hiddenR: axis === 'r' ? shiftIdx(sh.hiddenR) : sh.hiddenR,
    hiddenC: axis === 'c' ? shiftIdx(sh.hiddenC) : sh.hiddenC,
    rows: axis === 'r' ? Math.max(1, sh.rows + delta) : sh.rows,
    cols: axis === 'c' ? Math.max(1, sh.cols + delta) : sh.cols,
    rowH: axis === 'r' ? shiftMap(sh.rowH) : sh.rowH,
    colW: axis === 'c' ? shiftMap(sh.colW) : sh.colW,
  };
}

/** Move a range when rows/cols are inserted (delta>0) or deleted (delta<0) at `at`. Null when it was deleted entirely. */
export function shiftRange(rg: Range, axis: 'r' | 'c', at: number, delta: number): Range | null {
  const [lo, hi] = axis === 'r' ? [rg.r1, rg.r2] : [rg.c1, rg.c2];
  const move = (i: number, end: boolean) => {
    if (delta > 0) return i >= at ? i + delta : i;
    const n = -delta;
    if (i < at) return i;
    if (i >= at + n) return i - n;
    return end ? at - 1 : at; // inside the deleted block: clamp to its edge
  };
  let a = move(lo, false), b = move(hi, true);
  if (delta < 0 && lo >= at && hi < at - delta) return null;
  if (b < a) b = a;
  a = Math.max(0, a);
  return axis === 'r' ? { ...rg, r1: a, r2: b } : { ...rg, c1: a, c2: b };
}

export const inRange = (rg: Range, r: number, c: number) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2;
export const rangesOverlap = (a: Range, b: Range) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
export const rangeRef = (rg: Range) => (rg.r1 === rg.r2 && rg.c1 === rg.c2 ? addr(rg.r1, rg.c1) : `${addr(rg.r1, rg.c1)}:${addr(rg.r2, rg.c2)}`);
export function parseRange(s: string): Range | null {
  const [a, b] = s.trim().split(':');
  const p = parseAddr(a), q = b ? parseAddr(b) : p;
  return p && q ? normRange({ r1: p.r, c1: p.c, r2: q.r, c2: q.c }) : null;
}
