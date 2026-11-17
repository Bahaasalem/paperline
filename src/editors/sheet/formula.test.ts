import { describe, expect, it } from 'vitest';
import { Engine, formatValue } from './formula';
import { newSheet, shiftSheet, type Workbook } from './model';

function book(cells: Record<string, string>, extra?: Record<string, Record<string, string>>): Workbook {
  const s = newSheet('Sheet1');
  for (const [k, v] of Object.entries(cells)) s.cells[k] = { v };
  const sheets = [s];
  for (const [name, cs] of Object.entries(extra ?? {})) {
    const o = newSheet(name);
    for (const [k, v] of Object.entries(cs)) o.cells[k] = { v };
    sheets.push(o);
  }
  return { sheets, active: 0 };
}

const val = (wb: Workbook, r: number, c: number) => new Engine(wb).cellValue(wb.sheets[0], r, c);
const calc = (formula: string, cells: Record<string, string> = {}) => val(book({ ...cells, Z1: formula }), 0, 25);

describe('formula engine', () => {
  it('does arithmetic with precedence', () => {
    expect(calc('=1+2*3')).toBe(7);
    expect(calc('=(1+2)*3')).toBe(9);
    expect(calc('=2^3^2')).toBe(64); // Excel is left-assoc: (2^3)^2
    expect(calc('=-2^2')).toBe(4);
    expect(calc('=10%')).toBeCloseTo(0.1);
    expect(calc('="a"&"b"&1')).toBe('ab1');
  });
  it('reads references and ranges', () => {
    const c = { A1: '1', A2: '2', A3: '3', B1: 'x' };
    expect(calc('=A1+A2', c)).toBe(3);
    expect(calc('=SUM(A1:A3)', c)).toBe(6);
    expect(calc('=AVERAGE(A1:A3)', c)).toBe(2);
    expect(calc('=COUNT(A1:B3)', c)).toBe(3);
    expect(calc('=COUNTA(A1:B3)', c)).toBe(4);
    expect(calc('=SUM(A:A)', c)).toBe(6);
    expect(calc('=MAX(A1:A3)-MIN(A1:A3)', c)).toBe(2);
  });
  it('handles logic and text functions', () => {
    expect(calc('=IF(1>2,"yes","no")')).toBe('no');
    expect(calc('=AND(TRUE,1)')).toBe(true);
    expect(calc('=UPPER(LEFT("hello",2))')).toBe('HE');
    expect(calc('=LEN("مرحبا")')).toBe(5);
    expect(calc('=ROUND(2.345,2)')).toBe(2.35);
    expect(calc('=IFERROR(1/0,"bad")')).toBe('bad');
  });
  it('supports conditional and lookup functions', () => {
    const c = { A1: 'apple', B1: '3', A2: 'pear', B2: '5', A3: 'apple', B3: '7' };
    expect(calc('=SUMIF(A1:A3,"apple",B1:B3)', c)).toBe(10);
    expect(calc('=COUNTIF(B1:B3,">4")', c)).toBe(2);
    expect(calc('=VLOOKUP("pear",A1:B3,2,FALSE)', c)).toBe(5);
    expect(calc('=INDEX(B1:B3,MATCH("pear",A1:A3,0))', c)).toBe(5);
  });
  it('reports errors', () => {
    expect(String(calc('=1/0'))).toBe('#DIV/0!');
    expect(String(calc('=NOPE(1)'))).toBe('#NAME?');
    const wb = book({ A1: '=B1', B1: '=A1' });
    expect(String(val(wb, 0, 0))).toBe('#CYCLE!');
  });
  it('reads other sheets', () => {
    const wb = book({ A1: "=Data!A1*2+'My Sheet'!B2" }, { Data: { A1: '21' }, 'My Sheet': { B2: '1' } });
    expect(val(wb, 0, 0)).toBe(43);
  });
});

describe('formatting', () => {
  it('formats numbers', () => {
    expect(formatValue(1234.5, 'currency')).toBe('$1,234.50');
    expect(formatValue(0.256, 'percent', 1)).toBe('25.6%');
    expect(formatValue(45658, 'date')).toBe('2025-01-01');
    expect(formatValue(1 / 3, undefined)).toBe('0.333333333333');
  });
});

describe('row and column shifting', () => {
  it('moves cells and rewrites formulas', () => {
    const wb = book({ A1: '1', A2: '2', A3: '=SUM(A1:A2)' });
    const sh = shiftSheet(wb.sheets[0], 'r', 1, 1); // insert a row above row 2
    expect(sh.cells.A3.v).toBe('2');
    expect(sh.cells.A4.v).toBe('=SUM(A1:A3)');
  });
});
