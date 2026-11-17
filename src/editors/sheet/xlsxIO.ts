import ExcelJS from 'exceljs';
import { addr, newSheet, parseAddr, type Cell, type CellStyle, type Sheet, type Workbook, type NumFmt } from './model';
import { dateToSerial, Engine } from './formula';

const PX_PER_CHAR = 7.5;

function argbToHex(argb?: string): string | undefined {
  if (!argb || argb.length < 6) return undefined;
  const hex = argb.length === 8 ? argb.slice(2) : argb;
  return `#${hex.toLowerCase()}`;
}
const hexToArgb = (hex: string) => `FF${hex.replace('#', '').toUpperCase().padEnd(6, '0').slice(0, 6)}`;

function fmtFromNumFmt(nf?: string): { fmt?: NumFmt; dec?: number } {
  if (!nf || nf === 'General') return {};
  const dec = (nf.split('.')[1] ?? '').replace(/[^0]/g, '').length;
  if (/%/.test(nf)) return { fmt: 'percent', dec };
  if (/[$€£¥]|\[\$/.test(nf)) return { fmt: 'currency', dec };
  if (/[dmy]/i.test(nf) && !/0/.test(nf)) return { fmt: 'date' };
  if (/[0#]/.test(nf)) return { fmt: 'number', dec };
  return {};
}

function numFmtFor(st: CellStyle): string | undefined {
  const d = st.dec ?? (st.fmt === 'percent' ? 0 : 2);
  const z = d > 0 ? '.' + '0'.repeat(d) : '';
  switch (st.fmt) {
    case 'number':
      return `#,##0${z}`;
    case 'currency':
      return `"$"#,##0${z}`;
    case 'percent':
      return `0${z}%`;
    case 'date':
      return 'yyyy-mm-dd';
    default:
      return st.dec !== undefined ? `0${z}` : undefined;
  }
}

export async function readXlsx(buf: ArrayBuffer): Promise<Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const sheets: Sheet[] = [];
  wb.eachSheet((ws) => {
    const sh = newSheet(ws.name);
    let maxR = 0;
    let maxC = 0;
    ws.eachRow({ includeEmpty: false }, (row, rn) => {
      if (row.height) sh.rowH[rn - 1] = Math.round((row.height * 96) / 72);
      row.eachCell({ includeEmpty: false }, (c, cn) => {
        const r = rn - 1;
        const col = cn - 1;
        let v = '';
        const val = c.value as ExcelJS.CellValue;
        if (val === null || val === undefined) v = '';
        else if (typeof val === 'object' && 'formula' in val && val.formula) v = '=' + val.formula;
        else if (typeof val === 'object' && 'sharedFormula' in val) {
          // Shared formulas: take ExcelJS's translated formula when present, else cached result
          const f = (c as unknown as { formula?: string }).formula;
          v = f ? '=' + f : String((val as { result?: unknown }).result ?? '');
        } else if (val instanceof Date) v = String(dateToSerial(val));
        else if (typeof val === 'object' && 'richText' in val) v = val.richText.map((t) => t.text).join('');
        else if (typeof val === 'object' && 'text' in val) v = String((val as { text: unknown }).text);
        else if (typeof val === 'object' && 'error' in val) v = String((val as { error: unknown }).error);
        else if (typeof val === 'object' && 'result' in val) v = String((val as { result: unknown }).result ?? '');
        else v = String(val);

        const st: CellStyle = {};
        const f = c.font;
        if (f?.bold) st.b = true;
        if (f?.italic) st.i = true;
        if (f?.underline) st.u = true;
        if (f?.strike) st.s = true;
        if (f?.size && f.size !== 11) st.fontSize = f.size;
        if (f?.name && f.name !== 'Calibri') st.font = f.name;
        const fc = argbToHex(f?.color?.argb);
        if (fc && fc !== '#000000') st.color = fc;
        const fill = c.fill as ExcelJS.FillPattern | undefined;
        if (fill?.type === 'pattern' && fill.pattern === 'solid') {
          const bg = argbToHex(fill.fgColor?.argb);
          if (bg) st.bg = bg;
        }
        const h = c.alignment?.horizontal;
        if (h === 'center' || h === 'left' || h === 'right') st.align = h;
        if (c.alignment?.wrapText) st.wrap = true;
        if (c.border && (c.border.top || c.border.bottom || c.border.left || c.border.right)) st.border = true;
        Object.assign(st, fmtFromNumFmt(c.numFmt));
        if (val instanceof Date && !st.fmt) st.fmt = 'date';

        if (v === '' && Object.keys(st).length === 0) return;
        const cell: Cell = { v };
        if (Object.keys(st).length) cell.st = st;
        sh.cells[addr(r, col)] = cell;
        maxR = Math.max(maxR, r);
        maxC = Math.max(maxC, col);
      });
    });
    ws.columns?.forEach((col, i) => {
      if (col.width) sh.colW[i] = Math.round(col.width * PX_PER_CHAR);
    });
    if (ws.views?.[0] && (ws.views[0] as { state?: string; ySplit?: number }).state === 'frozen') sh.freezeTop = true;
    sh.rows = Math.max(100, maxR + 21);
    sh.cols = Math.max(26, maxC + 6);
    sheets.push(sh);
  });
  if (!sheets.length) sheets.push(newSheet('Sheet1'));
  return { sheets, active: 0 };
}

export async function writeXlsx(book: Workbook): Promise<Blob> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Paperline';
  wb.created = new Date();
  const engine = new Engine(book);
  for (const sh of book.sheets) {
    const ws = wb.addWorksheet(sh.name.slice(0, 31), sh.freezeTop ? { views: [{ state: 'frozen', ySplit: 1 }] } : {});
    for (const [k, w] of Object.entries(sh.colW)) ws.getColumn(Number(k) + 1).width = w / PX_PER_CHAR;
    for (const [k, h] of Object.entries(sh.rowH)) ws.getRow(Number(k) + 1).height = (h * 72) / 96;
    for (const [a, cell] of Object.entries(sh.cells)) {
      const p = parseAddr(a)!;
      const c = ws.getCell(p.r + 1, p.c + 1);
      if (cell.v.startsWith('=') && cell.v.length > 1) {
        const res = engine.cellValue(sh, p.r, p.c);
        const result = res === null || typeof res === 'object' ? undefined : res;
        c.value = { formula: cell.v.slice(1), result } as ExcelJS.CellFormulaValue;
      } else if (cell.v !== '') {
        const t = cell.v.trim();
        if (/^[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?$/.test(t)) c.value = parseFloat(t);
        else if (/^[-+]?\d+(\.\d+)?%$/.test(t)) c.value = parseFloat(t) / 100;
        else if (/^(true|false)$/i.test(t)) c.value = t.toLowerCase() === 'true';
        else c.value = cell.v;
      }
      const st = cell.st;
      if (!st) continue;
      c.font = {
        bold: st.b,
        italic: st.i,
        underline: st.u,
        strike: st.s,
        size: st.fontSize ?? 11,
        name: st.font ?? 'Calibri',
        color: st.color ? { argb: hexToArgb(st.color) } : undefined,
      };
      if (st.bg) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(st.bg) } };
      if (st.align || st.wrap) c.alignment = { horizontal: st.align, wrapText: st.wrap, vertical: st.wrap ? 'top' : undefined };
      if (st.border) {
        const b: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: 'FF94A3B8' } };
        c.border = { top: b, left: b, right: b, bottom: b };
      }
      const nf = numFmtFor(st);
      if (nf) c.numFmt = nf;
    }
  }
  const out = await wb.xlsx.writeBuffer();
  return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

