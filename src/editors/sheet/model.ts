export type NumFmt = 'general' | 'number' | 'currency' | 'percent' | 'date';
export type BorderStyle = 'none' | 'all' | 'outer' | 'bottom';

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
  border?: boolean;
}

export interface Cell {
  /** Raw input: a literal ("12", "hello") or a formula ("=SUM(A1:A3)") */
  v: string;
  st?: CellStyle;
}

export interface Sheet {
  name: string;
  rows: number;
  cols: number;
  cells: Record<string, Cell>;
  colW: Record<number, number>;
  rowH: Record<number, number>;
  freezeTop?: boolean;
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
  return {
    ...sh,
    cells,
    rows: axis === 'r' ? Math.max(1, sh.rows + delta) : sh.rows,
    cols: axis === 'c' ? Math.max(1, sh.cols + delta) : sh.cols,
    rowH: axis === 'r' ? shiftMap(sh.rowH) : sh.rowH,
    colW: axis === 'c' ? shiftMap(sh.colW) : sh.colW,
  };
}
