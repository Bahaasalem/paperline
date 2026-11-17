import { describe, expect, it } from 'vitest';
import { Engine } from './formula';
import { newSheet, shiftSheet, type Cell, type Sheet } from './model';
import { applyBorders, chartData, currentRegion, evalCF, expandToMerges, fillSeries, filteredOut, filterValues, findCells, mergeCells, replaceIn } from './features';

const fill = (vals: string[], n: number, fmt?: 'date') =>
  fillSeries(vals.map((v): Cell => ({ v, ...(fmt && { st: { fmt } }) })), n, (v, k) => `${v}+${k}`).map((c) => c?.v);

function sheetOf(cells: Record<string, string>): Sheet {
  const s = newSheet('S');
  for (const [k, v] of Object.entries(cells)) s.cells[k] = { v };
  return s;
}
const valueAt = (s: Sheet) => {
  const e = new Engine({ sheets: [s], active: 0 });
  return (r: number, c: number) => e.cellValue(s, r, c);
};

describe('autofill', () => {
  it('continues numbers, words and dates', () => {
    expect(fill(['1', '2'], 3)).toEqual(['3', '4', '5']);
    expect(fill(['10', '8'], 2)).toEqual(['6', '4']);
    expect(fill(['5'], 2)).toEqual(['5', '5']);
    expect(fill(['45000'], 2, 'date')).toEqual(['45001', '45002']);
    expect(fill(['Monday'], 3)).toEqual(['Tuesday', 'Wednesday', 'Thursday']);
    expect(fill(['DEC'], 2)).toEqual(['JAN', 'FEB']);
    expect(fill(['السبت'], 1)).toEqual(['الأحد']);
    expect(fill(['Item 1'], 2)).toEqual(['Item 2', 'Item 3']);
    expect(fill(['Q4'], 1)).toEqual(['Q1']);
    expect(fill(['a', 'b'], 3)).toEqual(['a', 'b', 'a']);
    expect(fill(['=A1'], 2)).toEqual(['=A1+1', '=A1+2']);
  });
});

describe('merge and borders', () => {
  it('merges keeping the first value', () => {
    const s = mergeCells(sheetOf({ A1: 'x', B1: 'y' }), { r1: 0, c1: 0, r2: 0, c2: 1 });
    expect(s.cells.A1.v).toBe('x');
    expect(s.cells.B1).toBeUndefined();
    expect(s.merges).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 1 }]);
    expect(expandToMerges(s, { r1: 0, c1: 1, r2: 0, c2: 1 })).toEqual({ r1: 0, c1: 0, r2: 0, c2: 1 });
    // inserting a column inside the merge widens it
    expect(shiftSheet(s, 'c', 1, 1).merges).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 2 }]);
  });
  it('draws outside borders on the edge only', () => {
    const s = applyBorders(sheetOf({}), { r1: 0, c1: 0, r2: 1, c2: 1 }, 'outer');
    expect(s.cells.A1.st?.bd).toEqual({ t: 1, l: 1 });
    expect(s.cells.B2.st?.bd).toEqual({ b: 1, r: 1 });
    expect(applyBorders(s, { r1: 0, c1: 0, r2: 1, c2: 1 }, 'none').cells.A1).toBeUndefined();
  });
});

describe('conditional formatting and filter', () => {
  it('highlights, bars and scales', () => {
    const s = sheetOf({ A1: '1', A2: '5', A3: '10', A4: '5' });
    s.cf = [
      { id: '1', range: { r1: 0, c1: 0, r2: 99, c2: 0 }, kind: 'gt', a: '4', bg: '#ff0000' },
      { id: '2', range: { r1: 0, c1: 0, r2: 99, c2: 0 }, kind: 'bar' },
    ];
    const out = evalCF(s, valueAt(s));
    expect(out.get('A1')?.bg).toBeUndefined();
    expect(out.get('A3')?.bg).toBe('#ff0000');
    expect(out.get('A3')?.bar?.pct).toBe(1);
    expect(out.get('A2')?.bar?.pct).toBe(0.5);
    s.cf = [{ id: '3', range: { r1: 0, c1: 0, r2: 3, c2: 0 }, kind: 'dup' }];
    expect([...evalCF(s, valueAt(s)).keys()].sort()).toEqual(['A2', 'A4']);
  });
  it('filters rows by shown values', () => {
    const s = sheetOf({ A1: 'Fruit', A2: 'apple', A3: 'pear', A4: 'apple', A5: 'kiwi' });
    s.filter = { range: { r1: 0, c1: 0, r2: 3, c2: 0 }, show: {} };
    expect(filterValues(s, valueAt(s), 0)).toEqual(['apple', 'kiwi', 'pear']);
    s.filter.show[0] = ['apple'];
    expect([...filteredOut(s, valueAt(s))].sort()).toEqual([2, 4]); // row under the range is included too
    expect(currentRegion(s, 2, 0)).toEqual({ r1: 0, c1: 0, r2: 4, c2: 0 });
  });
});

describe('find and charts', () => {
  it('finds and replaces', () => {
    const s = sheetOf({ A1: 'Apple pie', B2: 'apple', C3: '=1+1' });
    expect(findCells(s, valueAt(s), 'apple', {})).toEqual([{ r: 0, c: 0 }, { r: 1, c: 1 }]);
    expect(findCells(s, valueAt(s), 'apple', { whole: true })).toEqual([{ r: 1, c: 1 }]);
    expect(findCells(s, valueAt(s), '2', {})).toEqual([{ r: 2, c: 2 }]);
    expect(replaceIn('Apple apple', 'apple', 'x', {})).toBe('x x');
  });
  it('reads series and categories', () => {
    const s = sheetOf({ A1: '', B1: 'Sales', C1: 'Cost', A2: 'Jan', B2: '10', C2: '4', A3: 'Feb', B3: '12', C3: '5' });
    const d = chartData({ r1: 0, c1: 0, r2: 2, c2: 2 }, valueAt(s), (r, c) => String(valueAt(s)(r, c) ?? ''));
    expect(d.cats).toEqual(['Jan', 'Feb']);
    expect(d.series).toEqual([{ name: 'Sales', values: [10, 12] }, { name: 'Cost', values: [4, 5] }]);
  });
});
