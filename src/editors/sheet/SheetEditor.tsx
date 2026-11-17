import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import {
  AlignCenter, AlignLeft, AlignRight, ArrowDownAZ, ArrowUpAZ, Bold, Columns3, Eraser, Grid3x3, Italic, Minus, PaintBucket,
  Pin, Plus, Redo2, Rows3, Sigma, Square, Strikethrough, Trash2, Type, Underline, Undo2, WrapText, ZoomIn, Calculator,
  Hash, ArrowLeftRight, Percent, DollarSign,
} from 'lucide-react';
import { useI18n } from '../../i18n';
import { Btn, ColorBtn, Group, Ribbon, Select } from '../../ui/Ribbon';
import { useHistory } from '../useHistory';
import {
  addr, colName, DEFAULT_COL_W, DEFAULT_ROW_H, newSheet, normRange, parseAddr, shiftSheet,
  type Cell, type CellStyle, type NumFmt, type Range, type Sheet, type Workbook,
} from './model';
import { Engine, formatValue, FUNCTION_NAMES } from './formula';

interface Props {
  data: Workbook;
  onChange: (wb: Workbook) => void;
}

const HEADER_W = 48;
const HEADER_H = 26;
const OVERSCAN = 8;

export function SheetEditor({ data, onChange }: Props) {
  const { t } = useI18n();
  const { commit, undo, redo } = useHistory(data, onChange);
  const sheet = data.sheets[data.active];
  const engine = useMemo(() => new Engine(data), [data]);

  const [anchor, setAnchor] = useState({ r: 0, c: 0 });
  const [focus, setFocus] = useState({ r: 0, c: 0 });
  const [editing, setEditing] = useState<{ r: number; c: number; value: string; fromBar?: boolean } | null>(null);
  const [zoom, setZoom] = useState(1);
  const [gridlines, setGridlines] = useState(true);
  const [renaming, setRenaming] = useState<number | null>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const editRef = useRef<HTMLInputElement>(null);
  const sinkRef = useRef<HTMLTextAreaElement>(null);
  /** Keyboard focus lives in a hidden textarea so IME and mobile keyboards can start editing. */
  const focusGrid = () => sinkRef.current?.focus({ preventScroll: true });
  const dragging = useRef(false);
  const [scroll, setScroll] = useState({ top: 0, height: 600 });

  const sel: Range = normRange({ r1: anchor.r, c1: anchor.c, r2: focus.r, c2: focus.c });
  const cellAt = (r: number, c: number): Cell | undefined => sheet.cells[addr(r, c)];
  const activeCell = cellAt(anchor.r, anchor.c);
  const activeStyle: CellStyle = activeCell?.st ?? {};

  // ----- geometry -----
  const colW = (c: number) => sheet.colW[c] ?? DEFAULT_COL_W;
  const rowH = (r: number) => sheet.rowH[r] ?? DEFAULT_ROW_H;
  const rowTops = useMemo(() => {
    const tops = new Array(sheet.rows + 1);
    tops[0] = 0;
    for (let r = 0; r < sheet.rows; r++) tops[r + 1] = tops[r] + (sheet.rowH[r] ?? DEFAULT_ROW_H);
    return tops as number[];
  }, [sheet.rows, sheet.rowH]);
  const colLefts = useMemo(() => {
    const l = new Array(sheet.cols + 1);
    l[0] = 0;
    for (let c = 0; c < sheet.cols; c++) l[c + 1] = l[c] + (sheet.colW[c] ?? DEFAULT_COL_W);
    return l as number[];
  }, [sheet.cols, sheet.colW]);

  const firstRow = Math.max(0, upperBound(rowTops, scroll.top / zoom - HEADER_H) - 1 - OVERSCAN);
  const lastRow = Math.min(sheet.rows - 1, upperBound(rowTops, (scroll.top + scroll.height) / zoom) + OVERSCAN);

  useLayoutEffect(() => {
    const el = gridRef.current;
    if (!el) return;
    const update = () => setScroll({ top: el.scrollTop, height: el.clientHeight });
    update();
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // ----- mutations -----
  const updateSheet = useCallback(
    (fn: (sh: Sheet) => Sheet) => {
      const sheets = data.sheets.slice();
      sheets[data.active] = fn(sheets[data.active]);
      commit({ ...data, sheets });
    },
    [data, commit],
  );

  const setCellValue = (r: number, c: number, v: string) => {
    updateSheet((sh) => {
      const cells = { ...sh.cells };
      const k = addr(r, c);
      const prev = cells[k];
      if (v === '' && !prev?.st) delete cells[k];
      else cells[k] = { ...prev, v };
      return { ...sh, cells, rows: Math.max(sh.rows, r + 2), cols: Math.max(sh.cols, c + 1) };
    });
  };

  const forSel = (fn: (cell: Cell | undefined, r: number, c: number) => Cell | undefined) => {
    updateSheet((sh) => {
      const cells = { ...sh.cells };
      for (let r = sel.r1; r <= sel.r2; r++)
        for (let c = sel.c1; c <= sel.c2; c++) {
          const k = addr(r, c);
          const out = fn(cells[k], r, c);
          if (!out || (out.v === '' && (!out.st || Object.keys(out.st).length === 0))) delete cells[k];
          else cells[k] = out;
        }
      return { ...sh, cells };
    });
  };

  const applyStyle = (patch: Partial<CellStyle>) => forSel((cell) => ({ v: cell?.v ?? '', st: clean({ ...cell?.st, ...patch }) }));
  const toggleStyle = (key: 'b' | 'i' | 'u' | 's' | 'wrap' | 'border') => applyStyle({ [key]: !activeStyle[key] });

  const commitEdit = (move?: 'down' | 'right' | 'up' | 'left') => {
    if (!editing) return;
    setCellValue(editing.r, editing.c, normalizeFormula(editing.value));
    setEditing(null);
    if (move) moveBy(move === 'down' ? 1 : move === 'up' ? -1 : 0, move === 'right' ? 1 : move === 'left' ? -1 : 0, false, editing);
    focusGrid();
  };

  const moveBy = (dr: number, dc: number, extend: boolean, from = focus) => {
    const base = extend ? focus : from;
    const r = clamp(base.r + dr, 0, sheet.rows - 1);
    const c = clamp(base.c + dc, 0, sheet.cols - 1);
    setFocus({ r, c });
    if (!extend) setAnchor({ r, c });
    ensureVisible(r, c);
  };

  const ensureVisible = (r: number, c: number) => {
    const el = gridRef.current;
    if (!el) return;
    const top = rowTops[r] * zoom;
    const bottom = rowTops[r + 1] * zoom + HEADER_H;
    const left = colLefts[c] * zoom;
    const right = colLefts[c + 1] * zoom + HEADER_W;
    if (top < el.scrollTop) el.scrollTop = top;
    else if (bottom > el.scrollTop + el.clientHeight) el.scrollTop = bottom - el.clientHeight;
    if (left < el.scrollLeft) el.scrollLeft = left;
    else if (right > el.scrollLeft + el.clientWidth) el.scrollLeft = right - el.clientWidth;
  };

  const startEdit = (initial?: string) => {
    setEditing({ r: anchor.r, c: anchor.c, value: initial ?? activeCell?.v ?? '' });
  };

  useEffect(() => {
    focusGrid();
  }, []);

  useEffect(() => {
    if (editing && !editing.fromBar) {
      const el = editRef.current;
      if (el && document.activeElement !== el) {
        el.focus();
        el.setSelectionRange(el.value.length, el.value.length);
      }
    }
  }, [editing]);

  // ----- keyboard -----
  const onGridKey = (e: React.KeyboardEvent) => {
    if (editing) return;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (mod && k.toLowerCase() === 'z') { e.preventDefault(); if (e.shiftKey) redo(); else undo(); return; }
    if (mod && k.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
    if (mod && k.toLowerCase() === 'b') { e.preventDefault(); toggleStyle('b'); return; }
    if (mod && k.toLowerCase() === 'i') { e.preventDefault(); toggleStyle('i'); return; }
    if (mod && k.toLowerCase() === 'u') { e.preventDefault(); toggleStyle('u'); return; }
    if (mod && k.toLowerCase() === 'a') { e.preventDefault(); setAnchor({ r: 0, c: 0 }); setFocus({ r: sheet.rows - 1, c: sheet.cols - 1 }); return; }
    if (mod) return;
    switch (k) {
      case 'ArrowDown': e.preventDefault(); moveBy(1, 0, e.shiftKey); return;
      case 'ArrowUp': e.preventDefault(); moveBy(-1, 0, e.shiftKey); return;
      case 'ArrowLeft': e.preventDefault(); moveBy(0, -1, e.shiftKey); return;
      case 'ArrowRight': e.preventDefault(); moveBy(0, 1, e.shiftKey); return;
      case 'Tab': e.preventDefault(); moveBy(0, e.shiftKey ? -1 : 1, false); return;
      case 'Enter': e.preventDefault(); moveBy(e.shiftKey ? -1 : 1, 0, false); return;
      case 'F2': e.preventDefault(); startEdit(); return;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        forSel((cell) => (cell?.st ? { v: '', st: cell.st } : undefined));
        return;
    }
    if (k.length === 1 && !e.altKey) {
      e.preventDefault();
      startEdit(k);
    }
  };

  const onEditKey = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') { e.preventDefault(); commitEdit(e.shiftKey ? 'up' : 'down'); }
    else if (e.key === 'Tab') { e.preventDefault(); commitEdit(e.shiftKey ? 'left' : 'right'); }
    else if (e.key === 'Escape') { e.preventDefault(); setEditing(null); focusGrid(); }
  };

  // ----- clipboard -----
  const onCopy = (e: React.ClipboardEvent, cut = false) => {
    if (editing) return;
    e.preventDefault();
    const rows: string[] = [];
    for (let r = sel.r1; r <= sel.r2; r++) {
      const vals: string[] = [];
      for (let c = sel.c1; c <= sel.c2; c++) {
        const v = engine.cellValue(sheet, r, c);
        vals.push(formatValue(v, cellAt(r, c)?.st?.fmt, cellAt(r, c)?.st?.dec));
      }
      rows.push(vals.join('\t'));
    }
    e.clipboardData.setData('text/plain', rows.join('\n'));
    // internal clipboard keeps formulas and styles
    const internal: (Cell | null)[][] = [];
    for (let r = sel.r1; r <= sel.r2; r++) {
      const row: (Cell | null)[] = [];
      for (let c = sel.c1; c <= sel.c2; c++) row.push(cellAt(r, c) ?? null);
      internal.push(row);
    }
    e.clipboardData.setData('application/x-paperline-cells', JSON.stringify({ src: { r: sel.r1, c: sel.c1 }, cells: internal }));
    if (cut) forSel(() => undefined);
  };

  const onPaste = (e: React.ClipboardEvent) => {
    if (editing) return;
    e.preventDefault();
    const internal = e.clipboardData.getData('application/x-paperline-cells');
    const updates: Record<string, Cell | null> = {};
    if (internal) {
      const { src, cells } = JSON.parse(internal) as { src: { r: number; c: number }; cells: (Cell | null)[][] };
      // tile the copied block over the selection when it fits a whole number of times (like Excel)
      const h = cells.length, w = cells[0]?.length ?? 1;
      const selH = sel.r2 - sel.r1 + 1, selW = sel.c2 - sel.c1 + 1;
      const reps = { r: selH % h === 0 ? selH / h : 1, c: selW % w === 0 ? selW / w : 1 };
      const origin = reps.r > 1 || reps.c > 1 ? { r: sel.r1, c: sel.c1 } : anchor;
      for (let ti = 0; ti < reps.r; ti++)
        for (let tj = 0; tj < reps.c; tj++) {
          const r0 = origin.r + ti * h, c0 = origin.c + tj * w;
          cells.forEach((row, i) => row.forEach((cell, j) => {
            updates[addr(r0 + i, c0 + j)] = cell ? { ...cell, v: relocate(cell.v, r0 - src.r, c0 - src.c) } : null;
          }));
        }
    } else {
      const text = e.clipboardData.getData('text/plain');
      text.replace(/\r?\n$/, '').split(/\r?\n/).forEach((line, i) => line.split('\t').forEach((v, j) => {
        const k = addr(anchor.r + i, anchor.c + j);
        updates[k] = { ...sheet.cells[k], v };
      }));
    }
    updateSheet((sh) => {
      const cells = { ...sh.cells };
      let maxR = sh.rows, maxC = sh.cols;
      for (const [k, v] of Object.entries(updates)) {
        if (v) cells[k] = v;
        else delete cells[k];
        const p = parseAddr(k)!;
        maxR = Math.max(maxR, p.r + 2);
        maxC = Math.max(maxC, p.c + 1);
      }
      return { ...sh, cells, rows: maxR, cols: maxC };
    });
  };

  // ----- mouse -----
  const onCellDown = (e: React.MouseEvent, r: number, c: number) => {
    if (e.button !== 0) return;
    if (editing && editing.value.startsWith('=') && (editing.r !== r || editing.c !== c)) {
      // point-mode: insert a reference into the formula being typed
      e.preventDefault();
      const ref = addr(r, c);
      setEditing({ ...editing, value: editing.value + (/[=(,+\-*/^&<>:]$/.test(editing.value) ? '' : '+') + ref });
      return;
    }
    if (editing) commitEdit();
    dragging.current = true;
    if (e.shiftKey) setFocus({ r, c });
    else {
      setAnchor({ r, c });
      setFocus({ r, c });
    }
    focusGrid();
  };
  const onCellEnter = (r: number, c: number) => {
    if (dragging.current) setFocus({ r, c });
  };
  useEffect(() => {
    const up = () => (dragging.current = false);
    window.addEventListener('mouseup', up);
    return () => window.removeEventListener('mouseup', up);
  }, []);

  const selectCol = (c: number, extend: boolean) => {
    if (!extend) setAnchor({ r: 0, c });
    setFocus({ r: sheet.rows - 1, c });
    if (!extend) setFocus({ r: sheet.rows - 1, c });
    focusGrid();
  };
  const selectRow = (r: number, extend: boolean) => {
    if (!extend) setAnchor({ r, c: 0 });
    setFocus({ r, c: sheet.cols - 1 });
    focusGrid();
  };

  const startColResize = (e: React.MouseEvent, c: number) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startW = colW(c);
    const rtl = getComputedStyle(gridRef.current!).direction === 'rtl';
    let w = startW;
    const guide = document.createElement('div');
    guide.className = 'resize-guide';
    document.body.appendChild(guide);
    const move = (ev: MouseEvent) => {
      w = Math.max(24, startW + ((rtl ? -1 : 1) * (ev.clientX - startX)) / zoom);
      guide.style.left = `${ev.clientX}px`;
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      guide.remove();
      if (w !== startW) updateSheet((sh) => ({ ...sh, colW: { ...sh.colW, [c]: Math.round(w) } }));
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  // ----- commands -----
  const insertFn = (name: string) => {
    const r = sel.r1, c = sel.c1;
    let arg = '';
    if (sel.r2 > sel.r1 || sel.c2 > sel.c1) {
      arg = `${addr(sel.r1, sel.c1)}:${addr(sel.r2, sel.c2)}`;
      // put the result under (or right of) the selection
      const target = sel.r2 > sel.r1 ? { r: sel.r2 + 1, c: sel.c1 } : { r: sel.r1, c: sel.c2 + 1 };
      setCellValue(target.r, target.c, `=${name}(${arg})`);
      setAnchor(target);
      setFocus(target);
      return;
    }
    // single cell: guess the range of numbers above (or to the left)
    let top = r - 1;
    while (top >= 0 && typeof engine.cellValue(sheet, top, c) === 'number') top--;
    if (top < r - 1) arg = `${addr(top + 1, c)}:${addr(r - 1, c)}`;
    else {
      let left = c - 1;
      while (left >= 0 && typeof engine.cellValue(sheet, r, left) === 'number') left--;
      if (left < c - 1) arg = `${addr(r, left + 1)}:${addr(r, c - 1)}`;
    }
    setEditing({ r, c, value: `=${name}(${arg}` + (arg ? ')' : '') });
  };

  const sortSel = (dir: 1 | -1) => {
    // sort the rows of the selection (or the used region) by the active column
    let { r1, r2, c1, c2 } = sel;
    if (r1 === r2) {
      const used = usedRange(sheet);
      if (!used) return;
      ({ r1, r2, c1, c2 } = used);
      if (cellAt(r1, anchor.c) && typeof engine.cellValue(sheet, r1, anchor.c) === 'string' && typeof engine.cellValue(sheet, r1 + 1, anchor.c) !== 'string') r1++; // skip header row
    }
    const key = anchor.c;
    const rows: { k: unknown; cells: (Cell | undefined)[] }[] = [];
    for (let r = r1; r <= r2; r++) {
      const cells: (Cell | undefined)[] = [];
      for (let c = c1; c <= c2; c++) cells.push(cellAt(r, c));
      rows.push({ k: engine.cellValue(sheet, r, key), cells });
    }
    rows.sort((a, b) => {
      if (a.k === null && b.k === null) return 0;
      if (a.k === null) return 1;
      if (b.k === null) return -1;
      if (typeof a.k === 'number' && typeof b.k === 'number') return (a.k - b.k) * dir;
      return String(a.k).localeCompare(String(b.k), undefined, { numeric: true }) * dir;
    });
    updateSheet((sh) => {
      const cells = { ...sh.cells };
      rows.forEach((row, i) => row.cells.forEach((cell, j) => {
        const k = addr(r1 + i, c1 + j);
        if (cell) cells[k] = cell;
        else delete cells[k];
      }));
      return { ...sh, cells };
    });
  };

  const insertRows = () => updateSheet((sh) => shiftSheet(sh, 'r', sel.r1, sel.r2 - sel.r1 + 1));
  const insertCols = () => updateSheet((sh) => shiftSheet(sh, 'c', sel.c1, sel.c2 - sel.c1 + 1));
  const deleteRows = () => updateSheet((sh) => shiftSheet(sh, 'r', sel.r1, -(sel.r2 - sel.r1 + 1)));
  const deleteCols = () => updateSheet((sh) => shiftSheet(sh, 'c', sel.c1, -(sel.c2 - sel.c1 + 1)));

  const setActiveSheet = (i: number) => {
    if (editing) commitEdit();
    onChange({ ...data, active: i });
    setAnchor({ r: 0, c: 0 });
    setFocus({ r: 0, c: 0 });
  };
  const addSheet = () => {
    let n = data.sheets.length + 1;
    while (data.sheets.some((s) => s.name === `${t('sheet')}${n}`)) n++;
    commit({ sheets: [...data.sheets, newSheet(`${t('sheet')}${n}`)], active: data.sheets.length });
  };
  const deleteSheet = (i: number) => {
    if (data.sheets.length <= 1) return;
    const sheets = data.sheets.filter((_, j) => j !== i);
    commit({ sheets, active: Math.min(data.active, sheets.length - 1) });
  };
  const renameSheet = (i: number, name: string) => {
    const n = name.trim().replace(/[\\/?*[\]:]/g, '').slice(0, 31);
    setRenaming(null);
    if (!n || data.sheets.some((s, j) => j !== i && s.name === n)) return;
    const sheets = data.sheets.slice();
    sheets[i] = { ...sheets[i], name: n };
    commit({ ...data, sheets });
  };

  // ----- status -----
  const stats = useMemo(() => {
    if (sel.r1 === sel.r2 && sel.c1 === sel.c2) return null;
    let sum = 0, n = 0, cnt = 0;
    const r2 = Math.min(sel.r2, sel.r1 + 5000);
    for (let r = sel.r1; r <= r2; r++)
      for (let c = sel.c1; c <= sel.c2; c++) {
        const v = engine.cellValue(sheet, r, c);
        if (v !== null && v !== '') cnt++;
        if (typeof v === 'number') { sum += v; n++; }
      }
    return { sum, avg: n ? sum / n : 0, cnt, n };
  }, [sel.r1, sel.r2, sel.c1, sel.c2, engine, sheet]);

  // ----- render -----
  const totalW = colLefts[sheet.cols];
  const totalH = rowTops[sheet.rows];
  const visibleRows: number[] = [];
  for (let r = firstRow; r <= lastRow; r++) visibleRows.push(r);
  const barValue = editing ? editing.value : activeCell?.v ?? '';

  const fmtOptions: { value: NumFmt; label: string }[] = [
    { value: 'general', label: t('formatGeneral') },
    { value: 'number', label: t('formatNumber') },
    { value: 'currency', label: t('formatCurrency') },
    { value: 'percent', label: t('formatPercent') },
    { value: 'date', label: t('formatDate') },
  ];
  const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36].map((s) => ({ value: s, label: String(s) }));

  const homeTab = (
    <>
      <Group>
        <Btn icon={<Undo2 size={18} />} label={t('undo')} onClick={undo} />
        <Btn icon={<Redo2 size={18} />} label={t('redo')} onClick={redo} />
      </Group>
      <Group label={t('font')}>
        <Select label={t('size')} value={activeStyle.fontSize ?? 11} options={sizes} onChange={(v) => applyStyle({ fontSize: v === 11 ? undefined : v })} width={64} />
        <Btn icon={<Bold size={18} />} label={t('bold')} active={!!activeStyle.b} onClick={() => toggleStyle('b')} />
        <Btn icon={<Italic size={18} />} label={t('italic')} active={!!activeStyle.i} onClick={() => toggleStyle('i')} />
        <Btn icon={<Underline size={18} />} label={t('underline')} active={!!activeStyle.u} onClick={() => toggleStyle('u')} />
        <Btn icon={<Strikethrough size={18} />} label={t('strike')} active={!!activeStyle.s} onClick={() => toggleStyle('s')} />
        <ColorBtn icon={<Type size={16} />} label={t('textColor')} color={activeStyle.color ?? '#000'} onPick={(c) => applyStyle({ color: c ?? undefined })} allowNone />
        <ColorBtn icon={<PaintBucket size={16} />} label={t('fillColor')} color={activeStyle.bg} onPick={(c) => applyStyle({ bg: c ?? undefined })} allowNone />
      </Group>
      <Group label={t('style')}>
        <Btn icon={<AlignLeft size={18} />} label={t('alignLeft')} active={activeStyle.align === 'left'} onClick={() => applyStyle({ align: activeStyle.align === 'left' ? undefined : 'left' })} />
        <Btn icon={<AlignCenter size={18} />} label={t('alignCenter')} active={activeStyle.align === 'center'} onClick={() => applyStyle({ align: activeStyle.align === 'center' ? undefined : 'center' })} />
        <Btn icon={<AlignRight size={18} />} label={t('alignRight')} active={activeStyle.align === 'right'} onClick={() => applyStyle({ align: activeStyle.align === 'right' ? undefined : 'right' })} />
        <Btn icon={<WrapText size={18} />} label={t('wrap')} active={!!activeStyle.wrap} onClick={() => toggleStyle('wrap')} />
        <Btn icon={<Square size={18} />} label={t('borders')} active={!!activeStyle.border} onClick={() => toggleStyle('border')} />
      </Group>
      <Group label={t('formatNumber')}>
        <Select label={t('formatNumber')} value={activeStyle.fmt ?? 'general'} options={fmtOptions} onChange={(v) => applyStyle({ fmt: v === 'general' ? undefined : v, dec: undefined })} width={110} />
        <Btn icon={<DollarSign size={18} />} label={t('formatCurrency')} onClick={() => applyStyle({ fmt: 'currency', dec: undefined })} />
        <Btn icon={<Percent size={18} />} label={t('formatPercent')} onClick={() => applyStyle({ fmt: 'percent', dec: undefined })} />
        <Btn icon={<Plus size={14} />} label={t('decInc')} onClick={() => applyStyle({ dec: Math.min(10, (activeStyle.dec ?? defaultDec(activeStyle.fmt)) + 1), fmt: activeStyle.fmt ?? 'number' })} />
        <Btn icon={<Minus size={14} />} label={t('decDec')} onClick={() => applyStyle({ dec: Math.max(0, (activeStyle.dec ?? defaultDec(activeStyle.fmt)) - 1), fmt: activeStyle.fmt ?? 'number' })} />
      </Group>
      <Group label={t('table')}>
        <Btn icon={<Rows3 size={18} />} label={t('insertRow')} onClick={insertRows} />
        <Btn icon={<Columns3 size={18} />} label={t('insertCol')} onClick={insertCols} />
        <Btn icon={<Trash2 size={16} />} label={t('deleteRow')} onClick={deleteRows} />
        <Btn icon={<span className="rot90"><Trash2 size={16} /></span>} label={t('deleteCol')} onClick={deleteCols} />
        <Btn icon={<Eraser size={18} />} label={t('clearFormat')} onClick={() => forSel((cell) => (cell?.v ? { v: cell.v } : undefined))} />
      </Group>
      <Group>
        <Btn big icon={<Sigma size={22} />} label={t('autoSum')} onClick={() => insertFn('SUM')} />
      </Group>
    </>
  );

  const formulasTab = (
    <>
      <Group label={t('functions')}>
        <Btn big icon={<Sigma size={22} />} label={t('sum')} onClick={() => insertFn('SUM')} />
        <Btn big icon={<Calculator size={22} />} label={t('average')} onClick={() => insertFn('AVERAGE')} />
        <Btn big icon={<Hash size={22} />} label={t('count')} onClick={() => insertFn('COUNT')} />
        <Btn big icon={<ArrowDownAZ size={22} />} label={t('min')} onClick={() => insertFn('MIN')} />
        <Btn big icon={<ArrowUpAZ size={22} />} label={t('max')} onClick={() => insertFn('MAX')} />
      </Group>
      <Group label={t('functions')}>
        <Select label={t('functions')} value={'' as string} options={[{ value: '', label: 'ƒx …' }, ...FUNCTION_NAMES.map((f) => ({ value: f, label: f }))]} onChange={(f) => f && insertFn(f)} width={140} />
      </Group>
    </>
  );

  const dataTab = (
    <>
      <Group label={t('tab_data')}>
        <Btn big icon={<ArrowDownAZ size={22} />} label={t('sortAsc')} onClick={() => sortSel(1)} />
        <Btn big icon={<ArrowUpAZ size={22} />} label={t('sortDesc')} onClick={() => sortSel(-1)} />
      </Group>
      <Group>
        <Btn big icon={<Pin size={22} />} label={t('freeze')} active={!!sheet.freezeTop} onClick={() => updateSheet((sh) => ({ ...sh, freezeTop: !sh.freezeTop }))} />
      </Group>
    </>
  );

  const viewTab = (
    <>
      <Group label={t('zoom')}>
        <Select label={t('zoom')} value={zoom} options={[0.5, 0.75, 0.9, 1, 1.25, 1.5, 2].map((z) => ({ value: z, label: `${Math.round(z * 100)}%` }))} onChange={setZoom} width={80} />
        <Btn icon={<ZoomIn size={18} />} label={t('zoom')} onClick={() => setZoom(1)} />
      </Group>
      <Group>
        <Btn big icon={<Grid3x3 size={22} />} label={t('gridlines')} active={gridlines} onClick={() => setGridlines((g) => !g)} />
        <Btn big icon={<Pin size={22} />} label={t('freeze')} active={!!sheet.freezeTop} onClick={() => updateSheet((sh) => ({ ...sh, freezeTop: !sh.freezeTop }))} />
      </Group>
    </>
  );

  const selBox = {
    left: colLefts[sel.c1],
    top: rowTops[sel.r1],
    width: colLefts[sel.c2 + 1] - colLefts[sel.c1],
    height: rowTops[sel.r2 + 1] - rowTops[sel.r1],
  };

  return (
    <div className="editor sheet-editor">
      <Ribbon
        tabs={[
          { id: 'home', label: t('tab_home'), content: homeTab },
          { id: 'formulas', label: t('tab_formulas'), content: formulasTab },
          { id: 'data', label: t('tab_data'), content: dataTab },
          { id: 'view', label: t('tab_view'), content: viewTab },
        ]}
      />
      <div className="formula-bar" dir="ltr">
        <div className="name-box">{sel.r1 === sel.r2 && sel.c1 === sel.c2 ? addr(anchor.r, anchor.c) : `${addr(sel.r1, sel.c1)}:${addr(sel.r2, sel.c2)}`}</div>
        <span className="fx">ƒx</span>
        <input
          className="formula-input"
          value={barValue}
          aria-label="Formula"
          onFocus={() => !editing && setEditing({ r: anchor.r, c: anchor.c, value: activeCell?.v ?? '', fromBar: true })}
          onChange={(e) => setEditing({ r: editing?.r ?? anchor.r, c: editing?.c ?? anchor.c, value: e.target.value, fromBar: true })}
          onKeyDown={onEditKey}
          onBlur={() => editing?.fromBar && commitEdit()}
        />
      </div>
      <div
        className={`grid-scroll ${gridlines ? '' : 'no-grid'}`}
        dir="ltr"
        ref={gridRef}
        tabIndex={-1}
        onFocus={(e) => { if (e.target === e.currentTarget) focusGrid(); }}
        onKeyDown={onGridKey}
        onCopy={(e) => onCopy(e)}
        onCut={(e) => onCopy(e, true)}
        onPaste={onPaste}
        onScroll={(e) => setScroll({ top: e.currentTarget.scrollTop, height: e.currentTarget.clientHeight })}
      >
        <div className="grid-canvas" style={{ width: (totalW + HEADER_W) * zoom, height: (totalH + HEADER_H) * zoom }}>
          <div className="grid-zoom" style={{ transform: `scale(${zoom})`, width: totalW + HEADER_W, height: totalH + HEADER_H }}>
            {/* column headers */}
            <div className="col-headers" style={{ width: totalW + HEADER_W, height: HEADER_H }}>
              <div className="corner" style={{ width: HEADER_W, height: HEADER_H }} onMouseDown={() => { setAnchor({ r: 0, c: 0 }); setFocus({ r: sheet.rows - 1, c: sheet.cols - 1 }); }} />
              {Array.from({ length: sheet.cols }, (_, c) => (
                <div
                  key={c}
                  className={`col-h ${c >= sel.c1 && c <= sel.c2 ? 'hl' : ''}`}
                  style={{ left: HEADER_W + colLefts[c], width: colW(c), height: HEADER_H }}
                  onMouseDown={(e) => selectCol(c, e.shiftKey)}
                >
                  {colName(c)}
                  <span className="col-resize" onMouseDown={(e) => startColResize(e, c)} />
                </div>
              ))}
            </div>
            {/* rows */}
            {visibleRows.map((r) => (
              <div key={r} className={`g-row ${sheet.freezeTop && r === 0 ? 'frozen' : ''}`} style={{ top: HEADER_H + rowTops[r], height: rowH(r), width: totalW + HEADER_W }}>
                <div className={`row-h ${r >= sel.r1 && r <= sel.r2 ? 'hl' : ''}`} style={{ width: HEADER_W, height: rowH(r) }} onMouseDown={(e) => selectRow(r, e.shiftKey)}>
                  {r + 1}
                </div>
                {Array.from({ length: sheet.cols }, (_, c) => {
                  const cell = sheet.cells[addr(r, c)];
                  if (!cell) return <div key={c} className="g-cell" style={{ left: HEADER_W + colLefts[c], width: colW(c), height: rowH(r) }} onMouseDown={(e) => onCellDown(e, r, c)} onMouseEnter={() => onCellEnter(r, c)} onDoubleClick={() => startEdit()} />;
                  const v = engine.cellValue(sheet, r, c);
                  const st = cell.st ?? {};
                  const text = formatValue(v, st.fmt, st.dec);
                  const isNum = typeof v === 'number';
                  const style: CSSProperties = {
                    left: HEADER_W + colLefts[c],
                    width: colW(c),
                    height: rowH(r),
                    fontWeight: st.b ? 700 : undefined,
                    fontStyle: st.i ? 'italic' : undefined,
                    textDecoration: [st.u && 'underline', st.s && 'line-through'].filter(Boolean).join(' ') || undefined,
                    color: typeof v === 'object' && v !== null ? '#dc2626' : st.color,
                    background: st.bg,
                    textAlign: st.align ?? (isNum ? 'right' : typeof v === 'boolean' || (typeof v === 'object' && v !== null) ? 'center' : 'start'),
                    fontSize: st.fontSize ? `${st.fontSize}pt` : undefined,
                    whiteSpace: st.wrap ? 'normal' : 'nowrap',
                    zIndex: !st.wrap && !isNum && text && !st.bg ? 1 : undefined,
                  };
                  return (
                    <div
                      key={c}
                      className={`g-cell ${st.border ? 'bordered' : ''} ${!st.wrap && !isNum && !st.bg && !st.align ? 'overflow' : ''}`}
                      style={style}
                      dir="auto"
                      onMouseDown={(e) => onCellDown(e, r, c)}
                      onMouseEnter={() => onCellEnter(r, c)}
                      onDoubleClick={() => startEdit()}
                    >
                      {text}
                    </div>
                  );
                })}
              </div>
            ))}
            <textarea
              ref={sinkRef}
              className="key-sink"
              aria-label="Cell input"
              style={{ left: HEADER_W + colLefts[anchor.c], top: HEADER_H + rowTops[anchor.r] }}
              onInput={(e) => {
                const v = e.currentTarget.value;
                e.currentTarget.value = '';
                if (!editing && v && !(e.nativeEvent as InputEvent).isComposing) startEdit(v);
              }}
              onCompositionEnd={(e) => {
                const v = e.currentTarget.value;
                e.currentTarget.value = '';
                if (!editing && v) startEdit(v);
              }}
            />
            {/* selection */}
            <div className="sel-box" style={{ left: HEADER_W + selBox.left, top: HEADER_H + selBox.top, width: selBox.width, height: selBox.height }} />
            <div className="sel-active" style={{ left: HEADER_W + colLefts[anchor.c], top: HEADER_H + rowTops[anchor.r], width: colW(anchor.c), height: rowH(anchor.r) }} />
            {editing && !editing.fromBar && (
              <input
                ref={editRef}
                className="cell-editor"
                dir="auto"
                style={{
                  left: HEADER_W + colLefts[editing.c],
                  top: HEADER_H + rowTops[editing.r],
                  minWidth: colW(editing.c),
                  height: rowH(editing.r),
                  fontWeight: cellAt(editing.r, editing.c)?.st?.b ? 700 : undefined,
                }}
                value={editing.value}
                onChange={(e) => setEditing({ ...editing, value: e.target.value })}
                onKeyDown={onEditKey}
                onBlur={(e) => {
                  // clicks into the grid in point-mode keep editing
                  if (editing.value.startsWith('=') && gridRef.current?.contains(e.relatedTarget as Node)) return;
                  commitEdit();
                }}
              />
            )}
          </div>
        </div>
      </div>
      <div className="sheet-footer">
        <div className="sheet-tabs">
          <button className="sheet-add" title={t('addSheet')} aria-label={t('addSheet')} onClick={addSheet}>
            <Plus size={16} />
          </button>
          {data.sheets.map((s, i) =>
            renaming === i ? (
              <input key={i} className="sheet-tab-input" autoFocus defaultValue={s.name} onBlur={(e) => renameSheet(i, e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') renameSheet(i, e.currentTarget.value); if (e.key === 'Escape') setRenaming(null); }} />
            ) : (
              <button
                key={i}
                className={`sheet-tab ${i === data.active ? 'active' : ''}`}
                onClick={() => setActiveSheet(i)}
                onDoubleClick={() => setRenaming(i)}
                title={`${t('renameSheet')}: double-click`}
              >
                {s.name}
                {i === data.active && data.sheets.length > 1 && (
                  <span className="sheet-del" title={t('deleteSheet')} role="button" onClick={(e) => { e.stopPropagation(); deleteSheet(i); }}>
                    ×
                  </span>
                )}
              </button>
            ),
          )}
        </div>
        <div className="sheet-stats">
          {stats && (
            <>
              {stats.n > 0 && <span>{t('average')}: {formatValue(stats.avg, undefined)}</span>}
              <span>{t('count')}: {stats.cnt}</span>
              {stats.n > 0 && <span>{t('sum')}: {formatValue(stats.sum, undefined)}</span>}
            </>
          )}
          <span className="muted"><ArrowLeftRight size={12} /> {Math.round(zoom * 100)}%</span>
        </div>
      </div>
    </div>
  );
}

// ----- helpers -----
function upperBound(arr: number[], x: number) {
  let lo = 0, hi = arr.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (arr[mid] <= x) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}
const clamp = (v: number, a: number, b: number) => Math.max(a, Math.min(b, v));
const defaultDec = (f?: NumFmt) => (f === 'percent' ? 0 : f === 'number' || f === 'currency' ? 2 : 0);
function clean(st: CellStyle): CellStyle | undefined {
  const out: CellStyle = {};
  for (const [k, v] of Object.entries(st)) if (v !== undefined && v !== false && v !== null) (out as Record<string, unknown>)[k] = v;
  return Object.keys(out).length ? out : undefined;
}
function normalizeFormula(v: string) {
  if (!v.startsWith('=')) return v;
  // auto-close parentheses and upper-case function names
  const open = (v.match(/\(/g)?.length ?? 0) - (v.match(/\)/g)?.length ?? 0);
  let out = v + ')'.repeat(Math.max(0, open));
  out = out.replace(/("[^"]*")|([A-Za-z_][\w.]*)(?=\()/g, (m, s) => (s ? m : m.toUpperCase()));
  out = out.replace(/("[^"]*")|\b([a-z]{1,3}\d+)\b/g, (m, s) => (s ? m : m.toUpperCase()));
  return out;
}
/** Move relative references in a copied formula by (dr, dc). */
function relocate(v: string, dr: number, dc: number) {
  if (!v.startsWith('=') || (dr === 0 && dc === 0)) return v;
  return v.replace(/("[^"]*")|(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])/g, (m, s, d1, col, d2, row) => {
    if (s) return m;
    let c = colIndexOf(col);
    let r = parseInt(row, 10) - 1;
    if (!d1) c += dc;
    if (!d2) r += dr;
    if (c < 0 || r < 0) return '#REF!';
    return `${d1}${colName(c)}${d2}${r + 1}`;
  });
}
function colIndexOf(name: string) {
  let n = 0;
  for (const ch of name) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
function usedRange(sh: Sheet): Range | null {
  let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
  for (const k of Object.keys(sh.cells)) {
    if (sh.cells[k].v === '') continue;
    const p = parseAddr(k)!;
    r1 = Math.min(r1, p.r); c1 = Math.min(c1, p.c); r2 = Math.max(r2, p.r); c2 = Math.max(c2, p.c);
  }
  return r2 < 0 ? null : { r1, c1, r2, c2 };
}

