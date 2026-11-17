import ExcelJS from 'exceljs';
import { addr, newSheet, parseAddr, parseRange, rangeRef, type Borders, type Cell, type CellStyle, type CFRule, type Sheet, type Workbook, type NumFmt } from './model';
import { bordersOf } from './features';
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
  if (nf === '@') return { fmt: 'text' };
  if (/E[+-]/i.test(nf) && /0/.test(nf)) return { fmt: 'scientific', dec: (nf.split('.')[1] ?? '').replace(/E.*/i, '').replace(/[^0]/g, '').length };
  if (/h/i.test(nf) && /[dy]/i.test(nf)) return { fmt: 'datetime' };
  if (/h/i.test(nf) && /[ms]/i.test(nf)) return { fmt: 'time' };
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
    case 'time':
      return 'hh:mm:ss';
    case 'datetime':
      return 'yyyy-mm-dd hh:mm:ss';
    case 'scientific':
      return `0${z}E+00`;
    case 'text':
      return '@';
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
    const hiddenR: number[] = [];
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      if (row.hidden) hiddenR.push(rn - 1);
    });
    if (hiddenR.length) sh.hiddenR = hiddenR;
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
        const vert = c.alignment?.vertical;
        if (vert === 'top' || vert === 'middle') st.va = vert;
        const bd: Borders = {};
        const side = (b?: Partial<ExcelJS.Border>) => (!b?.style ? 0 : /thick|medium|double/.test(b.style) ? 2 : 1);
        if (side(c.border?.top)) bd.t = side(c.border?.top);
        if (side(c.border?.right)) bd.r = side(c.border?.right);
        if (side(c.border?.bottom)) bd.b = side(c.border?.bottom);
        if (side(c.border?.left)) bd.l = side(c.border?.left);
        if (Object.keys(bd).length) st.bd = bd;
        Object.assign(st, fmtFromNumFmt(c.numFmt));
        if (val instanceof Date && !st.fmt) st.fmt = 'date';

        const rawNote = c.note as string | { texts?: { text: string }[] } | undefined;
        const note = typeof rawNote === 'string' ? rawNote : rawNote?.texts?.map((x) => x.text).join('') ?? '';
        const dv = c.dataValidation;
        const list = dv?.type === 'list' && typeof dv.formulae?.[0] === 'string' && dv.formulae[0].startsWith('"')
          ? dv.formulae[0].slice(1, -1).split(',').map((x: string) => x.trim()).filter(Boolean)
          : undefined;
        if (v === '' && Object.keys(st).length === 0 && !note && !list) return;
        const cell: Cell = { v };
        if (Object.keys(st).length) cell.st = st;
        if (note) cell.note = note;
        if (list?.length) cell.list = list;
        sh.cells[addr(r, col)] = cell;
        maxR = Math.max(maxR, r);
        maxC = Math.max(maxC, col);
      });
    });
    const hiddenC: number[] = [];
    ws.columns?.forEach((col, i) => {
      if (col.width) sh.colW[i] = Math.round(col.width * PX_PER_CHAR);
      if (col.hidden) hiddenC.push(i);
    });
    if (hiddenC.length) sh.hiddenC = hiddenC;
    const merges = (ws.model as { merges?: string[] }).merges?.map(parseRange).filter((m) => !!m) ?? [];
    if (merges.length) sh.merges = merges as NonNullable<Sheet['merges']>;
    const af = ws.autoFilter as string | { from: string | { row: number; column: number }; to: string | { row: number; column: number } } | undefined;
    const afRange = typeof af === 'string' ? parseRange(af) : af ? parseRange(`${refOf(af.from)}:${refOf(af.to)}`) : null;
    if (afRange) sh.filter = { range: afRange, show: {} };
    const cf = readCF(ws);
    if (cf.length) sh.cf = cf;
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
    for (const r of sh.hiddenR ?? []) ws.getRow(r + 1).hidden = true;
    for (const c of sh.hiddenC ?? []) ws.getColumn(c + 1).hidden = true;
    for (const m of sh.merges ?? []) ws.mergeCells(rangeRef(m));
    if (sh.filter) ws.autoFilter = rangeRef(sh.filter.range);
    writeCF(ws, sh.cf ?? []);
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
      if (cell.note) c.note = cell.note;
      if (cell.list?.length) c.dataValidation = { type: 'list', allowBlank: true, formulae: [`"${cell.list.join(',').replace(/"/g, '').slice(0, 250)}"`] };
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
      if (st.align || st.wrap || st.va) c.alignment = { horizontal: st.align, wrapText: st.wrap, vertical: st.va ?? (st.wrap ? 'top' : undefined) };
      const bd = bordersOf(st);
      if (bd) {
        const b = (w?: number): Partial<ExcelJS.Border> | undefined => (w ? { style: w > 1 ? 'medium' : 'thin', color: { argb: 'FF1F2937' } } : undefined);
        c.border = { top: b(bd.t), left: b(bd.l), right: b(bd.r), bottom: b(bd.b) };
      }
      const nf = numFmtFor(st);
      if (nf) c.numFmt = nf;
    }
  }
  const out = await wb.xlsx.writeBuffer();
  return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}


const refOf = (p: string | { row: number; column: number }) => (typeof p === 'string' ? p : addr(p.row - 1, p.column - 1));
const argb = (hex?: string) => (hex ? { argb: hexToArgb(hex) } : undefined);

function writeCF(ws: ExcelJS.Worksheet, rules: CFRule[]) {
  let priority = 1;
  for (const r of rules) {
    const ref = rangeRef(r.range);
    const tl = addr(r.range.r1, r.range.c1);
    const abs = `$${ref.replace(/(\d+)/g, '$$$1').replace(':', ':$')}`;
    const style: Partial<ExcelJS.Style> = {
      ...(r.bg && { fill: { type: 'pattern', pattern: 'solid', bgColor: argb(r.bg) } as ExcelJS.FillPattern }),
      ...(r.color && { font: { color: argb(r.color) } }),
    };
    const lit = (x?: string) => (x !== undefined && x.trim() !== '' && !isNaN(Number(x)) ? Number(x) : `"${(x ?? '').replace(/"/g, '""')}"`);
    let rule: ExcelJS.ConditionalFormattingRule | null = null;
    const p = priority++;
    switch (r.kind) {
      case 'gt': rule = { type: 'cellIs', operator: 'greaterThan', formulae: [lit(r.a)], style, priority: p }; break;
      case 'lt': rule = { type: 'cellIs', operator: 'lessThan', formulae: [lit(r.a)], style, priority: p }; break;
      case 'between': rule = { type: 'cellIs', operator: 'between', formulae: [lit(r.a), lit(r.b)], style, priority: p }; break;
      case 'eq': rule = { type: 'cellIs', operator: 'equal', formulae: [lit(r.a)], style, priority: p }; break;
      case 'text': rule = { type: 'containsText', operator: 'containsText', text: r.a ?? '', style, priority: p }; break;
      case 'blank': rule = { type: 'expression', formulae: [`LEN(TRIM(${tl}))=0`], style, priority: p }; break;
      case 'dup': rule = { type: 'expression', formulae: [`AND(${tl}<>"",COUNTIF(${abs},${tl})>1)`], style, priority: p }; break;
      case 'top':
      case 'bottom': rule = { type: 'top10', rank: Number(r.a) || 10, percent: false, bottom: r.kind === 'bottom', style, priority: p }; break;
      case 'aboveAvg': rule = { type: 'aboveAverage', aboveAverage: true, style, priority: p }; break;
      case 'bar': rule = { type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: argb(r.c1 ?? '#638ec6'), priority: p } as ExcelJS.DataBarRuleType; break;
      case 'scale':
        rule = r.mid
          ? { type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }], color: [argb(r.c2)!, argb(r.mid)!, argb(r.c1)!], priority: p }
          : { type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'max' }], color: [argb(r.c2)!, argb(r.c1)!], priority: p };
        break;
    }
    if (rule) ws.addConditionalFormatting({ ref, rules: [rule] });
  }
}

function readCF(ws: ExcelJS.Worksheet): CFRule[] {
  const list = (ws as unknown as { conditionalFormattings?: { ref: string; rules: (ExcelJS.ConditionalFormattingRule & { color?: unknown })[] }[] }).conditionalFormattings ?? [];
  const out: CFRule[] = [];
  const hex = (c?: Partial<ExcelJS.Color>) => argbToHex(c?.argb);
  for (const cf of list) {
    const range = parseRange(cf.ref.split(' ')[0]);
    if (!range) continue;
    for (const rule of cf.rules) {
      const style = (rule as { style?: Partial<ExcelJS.Style> }).style;
      const fill = style?.fill as ExcelJS.FillPattern | undefined;
      const look = { bg: hex(fill?.bgColor) ?? hex(fill?.fgColor), color: hex(style?.font?.color) };
      const id = Math.random().toString(36).slice(2, 9);
      const f = (rule as { formulae?: unknown[] }).formulae?.map((x) => String(x).replace(/^"|"$/g, ''));
      const push = (r: Omit<CFRule, 'id' | 'range'>) => out.push({ id, range, ...r });
      switch (rule.type) {
        case 'cellIs': {
          const op = { greaterThan: 'gt', lessThan: 'lt', between: 'between', equal: 'eq' }[rule.operator ?? 'equal'] as CFRule['kind'] | undefined;
          if (op) push({ kind: op, a: f?.[0], b: f?.[1], ...look });
          break;
        }
        case 'containsText':
          if (rule.operator === 'containsBlanks') push({ kind: 'blank', ...look });
          else if (rule.text) push({ kind: 'text', a: rule.text, ...look });
          break;
        case 'expression':
          if (f?.[0] && /COUNTIF\(/i.test(f[0])) push({ kind: 'dup', ...look });
          else if (f?.[0] && /LEN\(TRIM/i.test(f[0])) push({ kind: 'blank', ...look });
          break;
        case 'top10': push({ kind: rule.bottom ? 'bottom' : 'top', a: String(rule.rank ?? 10), ...look }); break;
        case 'aboveAverage': push({ kind: 'aboveAvg', ...look }); break;
        case 'dataBar': push({ kind: 'bar', c1: hex(rule.color as Partial<ExcelJS.Color>) ?? '#638ec6' }); break;
        case 'colorScale': {
          const cs = (rule.color ?? []).map((c) => hex(c));
          if (cs.length >= 2) push({ kind: 'scale', c2: cs[0], mid: cs.length > 2 ? cs[1] : undefined, c1: cs[cs.length - 1] });
          break;
        }
      }
    }
  }
  return out;
}
