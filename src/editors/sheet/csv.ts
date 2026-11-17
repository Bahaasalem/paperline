import { addr, newSheet, parseAddr, type Sheet, type Workbook } from './model';
import { Engine } from './formula';

export function readCsv(text: string, name = 'Sheet1'): Workbook {
  const sh = newSheet(name);
  const rows = parseCsv(text);
  rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') sh.cells[addr(r, c)] = { v }; }));
  sh.rows = Math.max(100, rows.length + 20);
  sh.cols = Math.max(26, Math.max(0, ...rows.map((r) => r.length)) + 5);
  return { sheets: [sh], active: 0 };
}

export function parseCsv(text: string): string[][] {
  const delim = (text.split('\n')[0].match(/;/g)?.length ?? 0) > (text.split('\n')[0].match(/,/g)?.length ?? 0) ? ';' : text.includes('\t') && !text.includes(',') ? '\t' : ',';
  const rows: string[][] = [];
  let row: string[] = [];
  let f = '';
  let q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') {
        if (text[i + 1] === '"') { f += '"'; i++; }
        else q = false;
      } else f += ch;
    } else if (ch === '"') q = true;
    else if (ch === delim) { row.push(f); f = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(f); rows.push(row); row = []; f = '';
    } else f += ch;
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  return rows;
}

export function writeCsv(sh: Sheet, book: Workbook): string {
  const engine = new Engine(book);
  let maxR = -1, maxC = -1;
  for (const k of Object.keys(sh.cells)) {
    const p = parseAddr(k)!;
    if (sh.cells[k].v !== '') { maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c); }
  }
  const lines: string[] = [];
  for (let r = 0; r <= maxR; r++) {
    const vals: string[] = [];
    for (let c = 0; c <= maxC; c++) {
      const v = engine.cellValue(sh, r, c);
      let s = v === null ? '' : String(v);
      if (/[",\n]/.test(s)) s = `"${s.replace(/"/g, '""')}"`;
      vals.push(s);
    }
    lines.push(vals.join(','));
  }
  return lines.join('\r\n');
}

