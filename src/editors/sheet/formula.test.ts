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

describe('extended functions', () => {
  const c = { A1: 'apple', B1: '3', C1: 'x', A2: 'pear', B2: '5', C2: 'y', A3: 'apple', B3: '7', C3: 'y' };
  it('does multi-criteria math', () => {
    expect(calc('=SUMIFS(B1:B3,A1:A3,"apple",C1:C3,"y")', c)).toBe(7);
    expect(calc('=COUNTIFS(A1:A3,"apple",B1:B3,">2")', c)).toBe(2);
    expect(calc('=AVERAGEIFS(B1:B3,A1:A3,"apple")', c)).toBe(5);
    expect(calc('=MAXIFS(B1:B3,A1:A3,"apple")', c)).toBe(7);
    expect(calc('=MINIFS(B1:B3,C1:C3,"y")', c)).toBe(5);
    expect(calc('=SUMPRODUCT(B1:B3,B1:B3)', c)).toBe(83);
  });
  it('branches and looks up', () => {
    expect(calc('=IFS(1>2,"a",2>1,"b")')).toBe('b');
    expect(calc('=SWITCH(2,1,"one",2,"two","other")')).toBe('two');
    expect(calc('=CHOOSE(3,"a","b","c")')).toBe('c');
    expect(calc('=XLOOKUP("pear",A1:A3,B1:B3)', c)).toBe(5);
    expect(calc('=XLOOKUP("kiwi",A1:A3,B1:B3,"none")', c)).toBe('none');
    expect(calc('=ROWS(A1:C3)*COLUMNS(A1:C3)', c)).toBe(9);
  });
  it('handles text', () => {
    expect(calc('=TEXTJOIN("-",TRUE,A1:A3)', c)).toBe('apple-pear-apple');
    expect(calc('=SUBSTITUTE("a-b-c","-","+")')).toBe('a+b+c');
    expect(calc('=REPLACE("abcdef",2,3,"X")')).toBe('aXef');
    expect(calc('=FIND("c","abc")')).toBe(3);
    expect(calc('=SEARCH("C","abc")')).toBe(3);
    expect(calc('=PROPER("hello world")')).toBe('Hello World');
    expect(calc('=REPT("ab",3)')).toBe('ababab');
    expect(calc('=VALUE("12.5")+1')).toBe(13.5);
    expect(calc('=EXACT("a","A")')).toBe(false);
  });
  it('handles info and math', () => {
    expect(calc('=ISBLANK(Q9)')).toBe(true);
    expect(calc('=ISNUMBER(5)')).toBe(true);
    expect(calc('=ISTEXT("a")')).toBe(true);
    expect(calc('=ISERROR(1/0)')).toBe(true);
    expect(calc('=CEILING(4.2,1)')).toBe(5);
    expect(calc('=FLOOR(4.8,1)')).toBe(4);
    expect(calc('=TRUNC(-4.8)')).toBe(-4);
    expect(calc('=LOG10(1000)')).toBe(3);
    expect(calc('=SIGN(-3)')).toBe(-1);
    const r = calc('=RANDBETWEEN(1,3)') as number;
    expect(r >= 1 && r <= 3).toBe(true);
  });
  it('handles statistics', () => {
    const n = { A1: '2', A2: '4', A3: '4', A4: '4', A5: '5', A6: '5', A7: '7', A8: '9' };
    expect(calc('=STDEVP(A1:A8)', n)).toBeCloseTo(2);
    expect(calc('=VAR(A1:A8)', n)).toBeCloseTo(32 / 7);
    expect(calc('=LARGE(A1:A8,2)', n)).toBe(7);
    expect(calc('=SMALL(A1:A8,1)', n)).toBe(2);
    expect(calc('=RANK(9,A1:A8)', n)).toBe(1);
  });
  it('handles dates and finance', () => {
    expect(calc('=WEEKDAY(DATE(2024,1,1))')).toBe(2); // Monday
    expect(calc('=DAY(EOMONTH(DATE(2024,1,15),1))')).toBe(29);
    expect(calc('=MONTH(EDATE(DATE(2024,1,31),1))')).toBe(2);
    expect(calc('=DATEDIF(DATE(2020,1,1),DATE(2024,6,1),"Y")')).toBe(4);
    expect(calc('=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7))')).toBe(5);
    expect(calc('=HOUR(TIME(13,30,0))')).toBe(13);
    expect(calc('=MINUTE(TIME(13,30,0))')).toBe(30);
    expect(calc('=PMT(0,10,1000)')).toBeCloseTo(-100);
    expect(calc('=PMT(0.05/12,360,200000)')).toBeCloseTo(-1073.64, 1);
    expect(calc('=NPV(0.1,110)')).toBeCloseTo(100);
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
