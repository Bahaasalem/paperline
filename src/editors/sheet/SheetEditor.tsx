import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import {
  AlignCenter, AlignLeft, AlignRight, AlignVerticalJustifyCenter, AlignVerticalJustifyEnd, AlignVerticalJustifyStart, ArrowDownAZ,
  ArrowLeftRight, ArrowUpAZ, Bold, Calculator, ChartArea, ChartColumn, ChartLine, ChartPie, ChartScatter, ChevronDown, ChevronUp,
  ClipboardPaste, Columns3, Copy, DollarSign, Donut, Eraser, Eye, EyeOff, Filter, FilterX, Grid2x2, Grid3x3, Hash, Highlighter, Italic,
  ListChecks, Merge, Minus, MoveHorizontal, MoveVertical, PaintBucket, PanelBottom, PanelLeft, PanelRight, PanelTop, Pencil, Percent,
  Pin, Plus, Redo2, Replace, Rows3, Scissors, Search, Sigma, Square, SquareDashed, Strikethrough, StickyNote, StickyNoteX, Table2,
  Trash2, Type, Underline, Undo2, WrapText, X, ZoomIn, ChartBar, Braces, CopyX, Tags,
} from 'lucide-react';
import { useI18n, type Key } from '../../i18n';
import { Btn, ColorBtn, Group, Menu, MenuItem, Ribbon, Select } from '../../ui/Ribbon';
import { FontPicker } from '../../ui/FontPicker';
import { familyCss } from '../../fonts';
import { useHistory } from '../useHistory';
import {
  addr, colName, DEFAULT_COL_W, DEFAULT_ROW_H, inRange, newSheet, normRange, parseAddr, rangeRef, shiftSheet,
  type Cell, type CellStyle, type CFKind, type CFRule, type ChartSpec, type ChartType, type NumFmt, type Range, type Sheet, type Workbook,
} from './model';
import { Engine, formatValue, FUNCTION_NAMES } from './formula';
import {
  applyBorders, BLANK, borderShadow, bordersOf, chartData, currentRegion, displayAt, evalCF, expandToMerges, fillSeries, filterEnd,
  filteredOut, filterValues, findCells, mergeAt, mergeCells, replaceIn, unmergeCells, type BorderMode,
} from './features';
import { ChartSvg } from './Chart';

interface Props {
  data: Workbook;
  onChange: (wb: Workbook) => void;
}

const HEADER_W = 48;
const HEADER_H = 26;
const OVERSCAN = 8;

/** Copies made from menus can't put our rich payload on the system clipboard, so remember the last one. */
let lastCopy: { text: string; payload: string } | null = null;

interface AskField { key: string; label: string; value: string; type?: 'text' | 'number' | 'textarea' | 'select'; options?: { value: string; label: string }[] }
interface Ask { title: string; fields: AskField[]; onOk: (v: Record<string, string>) => void }

const HL_STYLES: Record<string, { bg?: string; color?: string }> = {
  red: { bg: '#ffc7ce', color: '#9c0006' },
  yellow: { bg: '#ffeb9c', color: '#9c5700' },
  green: { bg: '#c6efce', color: '#006100' },
  redText: { color: '#c00000' },
  blue: { bg: '#dbeafe', color: '#1e3a8a' },
};

const TABLE_COLORS = ['#4f46e5', '#0e9f6e', '#0ea5a4', '#e0436b', '#dc6a1f', '#475569'];

export function SheetEditor({ data, onChange }: Props) {
  const { t } = useI18n();
  const { commit, undo, redo } = useHistory(data, onChange);
  const sheet = data.sheets[data.active];
  const engine = useMemo(() => new Engine(data), [data]);
  const valueAt = useCallback((r: number, c: number) => engine.cellValue(sheet, r, c), [engine, sheet]);

  const [anchor, setAnchor] = useState({ r: 0, c: 0 });
  const [focus, setFocus] = useState({ r: 0, c: 0 });
  const [editing, setEditing] = useState<{ r: number; c: number; value: string; fromBar?: boolean } | null>(null);
  const [zoom, setZoom] = useState(1);
  const [gridlines, setGridlines] = useState(true);
  const [showFormulas, setShowFormulas] = useState(false);
  const [renaming, setRenaming] = useState<number | null>(null);
  const [find, setFind] = useState<{ replace: boolean } | null>(null);
  const [findQ, setFindQ] = useState('');
  const [replaceQ, setReplaceQ] = useState('');
  const [findOpts, setFindOpts] = useState({ matchCase: false, whole: false, formulas: false });
  const [hitIdx, setHitIdx] = useState(0);
  const [filterPop, setFilterPop] = useState<{ c: number; x: number; y: number } | null>(null);
  const [ask, setAsk] = useState<Ask | null>(null);
  const [ctx, setCtx] = useState<{ x: number; y: number } | null>(null);
  const [chartSel, setChartSel] = useState<string | null>(null);
  const [fillTo, setFillTo] = useState<{ r: number; c: number } | null>(null);
  const [hoverNote, setHoverNote] = useState<{ r: number; c: number } | null>(null);
  const [listOpen, setListOpen] = useState(false);
  const [flash, setFlash] = useState<string | null>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const editRef = useRef<HTMLInputElement>(null);
  const sinkRef = useRef<HTMLTextAreaElement>(null);
  /** Keyboard focus lives in a hidden textarea so IME and mobile keyboards can start editing. */
  const focusGrid = () => sinkRef.current?.focus({ preventScroll: true });
  const dragging = useRef(false);
  const filling = useRef(false);
  const applyFillRef = useRef<() => void>(() => {});
  const [scroll, setScroll] = useState({ top: 0, height: 600 });

  useEffect(() => {
    if (!flash) return;
    const id = setTimeout(() => setFlash(null), 4000);
    return () => clearTimeout(id);
  }, [flash]);

  const sel: Range = expandToMerges(sheet, normRange({ r1: anchor.r, c1: anchor.c, r2: focus.r, c2: focus.c }));
  const anchorMerge = mergeAt(sheet, anchor.r, anchor.c);
  const single = (sel.r1 === sel.r2 && sel.c1 === sel.c2) || (!!anchorMerge && rangeRef(anchorMerge) === rangeRef(sel));
  const cellAt = (r: number, c: number): Cell | undefined => sheet.cells[addr(r, c)];
  const activeCell = cellAt(anchor.r, anchor.c);
  const activeStyle: CellStyle = activeCell?.st ?? {};

  // ----- hidden rows/cols, filter, conditional formats -----
  const filtered = useMemo(() => filteredOut(sheet, valueAt), [sheet, valueAt]);
  const hiddenR = useMemo(() => new Set([...(sheet.hiddenR ?? []), ...filtered]), [sheet.hiddenR, filtered]);
  const hiddenC = useMemo(() => new Set(sheet.hiddenC ?? []), [sheet.hiddenC]);
  const cfMap = useMemo(() => evalCF(sheet, valueAt), [sheet, valueAt]);

  // ----- geometry -----
  const colW = (c: number) => (hiddenC.has(c) ? 0 : sheet.colW[c] ?? DEFAULT_COL_W);
  const rowH = (r: number) => (hiddenR.has(r) ? 0 : sheet.rowH[r] ?? DEFAULT_ROW_H);
  const rowTops = useMemo(() => {
    const tops = new Array(sheet.rows + 1);
    tops[0] = 0;
    for (let r = 0; r < sheet.rows; r++) tops[r + 1] = tops[r] + (hiddenR.has(r) ? 0 : sheet.rowH[r] ?? DEFAULT_ROW_H);
    return tops as number[];
  }, [sheet.rows, sheet.rowH, hiddenR]);
  const colLefts = useMemo(() => {
    const l = new Array(sheet.cols + 1);
    l[0] = 0;
    for (let c = 0; c < sheet.cols; c++) l[c + 1] = l[c] + (hiddenC.has(c) ? 0 : sheet.colW[c] ?? DEFAULT_COL_W);
    return l as number[];
  }, [sheet.cols, sheet.colW, hiddenC]);

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
      if (v === '' && !prev?.st && !prev?.note && !prev?.list) delete cells[k];
      else cells[k] = { ...prev, v };
      return { ...sh, cells, rows: Math.max(sh.rows, r + 2), cols: Math.max(sh.cols, c + 1) };
    });
  };

  const forRange = (rg: Range, fn: (cell: Cell | undefined, r: number, c: number) => Cell | undefined) => {
    updateSheet((sh) => {
      const cells = { ...sh.cells };
      for (let r = rg.r1; r <= rg.r2; r++)
        for (let c = rg.c1; c <= rg.c2; c++) {
          const k = addr(r, c);
          const out = fn(cells[k], r, c);
          if (!out || (out.v === '' && (!out.st || Object.keys(out.st).length === 0) && !out.note && !out.list)) delete cells[k];
          else cells[k] = out;
        }
      return { ...sh, cells };
    });
  };
  const forSel = (fn: (cell: Cell | undefined, r: number, c: number) => Cell | undefined) => forRange(sel, fn);

  const applyStyle = (patch: Partial<CellStyle>) => forSel((cell) => ({ ...cell, v: cell?.v ?? '', st: clean({ ...cell?.st, ...patch }) }));
  const toggleStyle = (key: 'b' | 'i' | 'u' | 's' | 'wrap') => applyStyle({ [key]: !activeStyle[key] });
  const clearContents = () => forSel((cell) => (cell?.st || cell?.note || cell?.list ? { ...cell, v: '' } : undefined));

  const commitEdit = (move?: 'down' | 'right' | 'up' | 'left') => {
    if (!editing) return;
    const value = normalizeFormula(editing.value);
    const list = cellAt(editing.r, editing.c)?.list;
    if (list && value !== '' && !list.some((x) => x.toLowerCase() === value.toLowerCase())) {
      setFlash(t('s_notInList'));
      setEditing(null);
      focusGrid();
      return;
    }
    setCellValue(editing.r, editing.c, value);
    setEditing(null);
    if (move) moveBy(move === 'down' ? 1 : move === 'up' ? -1 : 0, move === 'right' ? 1 : move === 'left' ? -1 : 0, false, editing);
    focusGrid();
  };

  /** Step along an axis, skipping hidden rows/cols. */
  const stepIdx = (from: number, d: number, max: number, hidden: Set<number>) => {
    let v = from;
    for (let i = 0; i < Math.abs(d); i++) {
      let nv = v + Math.sign(d);
      while (nv >= 0 && nv < max && hidden.has(nv)) nv += Math.sign(d);
      if (nv < 0 || nv >= max) break;
      v = nv;
    }
    return v;
  };

  const moveBy = (dr: number, dc: number, extend: boolean, from = focus) => {
    let base = extend ? focus : from;
    const m = !extend ? mergeAt(sheet, base.r, base.c) : undefined;
    if (m) base = { r: dr > 0 ? m.r2 : m.r1, c: dc > 0 ? m.c2 : m.c1 };
    let r = stepIdx(base.r, dr, sheet.rows, hiddenR);
    let c = stepIdx(base.c, dc, sheet.cols, hiddenC);
    const land = !extend ? mergeAt(sheet, r, c) : undefined;
    if (land) ({ r1: r, c1: c } = land);
    r = clamp(r, 0, sheet.rows - 1);
    c = clamp(c, 0, sheet.cols - 1);
    setFocus({ r, c });
    if (!extend) setAnchor({ r, c });
    setChartSel(null);
    setListOpen(false);
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

  const goTo = (r: number, c: number) => {
    setAnchor({ r, c });
    setFocus({ r, c });
    ensureVisible(r, c);
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

  // ----- find -----
  const hits = useMemo(() => (find ? findCells(sheet, valueAt, findQ, findOpts) : []), [find, sheet, valueAt, findQ, findOpts]);
  const gotoHit = (d: number) => {
    if (!hits.length) return;
    const i = (((hitIdx + d) % hits.length) + hits.length) % hits.length;
    setHitIdx(i);
    goTo(hits[i].r, hits[i].c);
  };
  useEffect(() => {
    setHitIdx(0);
    if (hits.length) goTo(hits[0].r, hits[0].c);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [findQ, findOpts]);
  const replaceOne = () => {
    const h = hits[Math.min(hitIdx, hits.length - 1)];
    if (!h) return;
    const cell = cellAt(h.r, h.c);
    if (!cell) return;
    setCellValue(h.r, h.c, replaceIn(cell.v, findQ, replaceQ, findOpts));
  };
  const replaceAll = () => {
    if (!hits.length) return;
    const n = hits.length;
    updateSheet((sh) => {
      const cells = { ...sh.cells };
      for (const h of hits) {
        const k = addr(h.r, h.c);
        if (cells[k]) cells[k] = { ...cells[k], v: replaceIn(cells[k].v, findQ, replaceQ, findOpts) };
      }
      return { ...sh, cells };
    });
    setFlash(t('s_replaced').replace('{n}', String(n)));
  };

  // ----- keyboard -----
  const onGridKey = (e: React.KeyboardEvent) => {
    if (editing) return;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (chartSel && (k === 'Delete' || k === 'Backspace')) { e.preventDefault(); deleteChart(chartSel); return; }
    if (chartSel && k === 'Escape') { setChartSel(null); return; }
    if (mod && k.toLowerCase() === 'z') { e.preventDefault(); if (e.shiftKey) redo(); else undo(); return; }
    if (mod && k.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
    if (mod && k.toLowerCase() === 'b') { e.preventDefault(); toggleStyle('b'); return; }
    if (mod && k.toLowerCase() === 'i') { e.preventDefault(); toggleStyle('i'); return; }
    if (mod && k.toLowerCase() === 'u') { e.preventDefault(); toggleStyle('u'); return; }
    if (mod && k.toLowerCase() === 'f') { e.preventDefault(); setFind({ replace: false }); return; }
    if (mod && k.toLowerCase() === 'h') { e.preventDefault(); setFind({ replace: true }); return; }
    if (mod && k.toLowerCase() === 'd') { e.preventDefault(); fillDir('down'); return; }
    if (mod && k.toLowerCase() === 'r') { e.preventDefault(); fillDir('right'); return; }
    if (mod && k === '`') { e.preventDefault(); setShowFormulas((s) => !s); return; }
    if (mod && e.shiftKey && k.toLowerCase() === 'l') { e.preventDefault(); toggleFilter(); return; }
    if (mod && k.toLowerCase() === 'a') { e.preventDefault(); setAnchor({ r: 0, c: 0 }); setFocus({ r: sheet.rows - 1, c: sheet.cols - 1 }); return; }
    if (e.altKey && k === 'ArrowDown' && activeCell?.list) { e.preventDefault(); setListOpen(true); return; }
    if (mod) return;
    switch (k) {
      case 'ArrowDown': e.preventDefault(); moveBy(1, 0, e.shiftKey); return;
      case 'ArrowUp': e.preventDefault(); moveBy(-1, 0, e.shiftKey); return;
      case 'ArrowLeft': e.preventDefault(); moveBy(0, -1, e.shiftKey); return;
      case 'ArrowRight': e.preventDefault(); moveBy(0, 1, e.shiftKey); return;
      case 'Tab': e.preventDefault(); moveBy(0, e.shiftKey ? -1 : 1, false); return;
      case 'Enter': e.preventDefault(); moveBy(e.shiftKey ? -1 : 1, 0, false); return;
      case 'F2': e.preventDefault(); startEdit(); return;
      case 'Escape': setCtx(null); setListOpen(false); return;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        clearContents();
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
  const copyPayload = () => {
    const rows: string[] = [];
    for (let r = sel.r1; r <= sel.r2; r++) {
      if (hiddenR.has(r)) continue;
      const vals: string[] = [];
      for (let c = sel.c1; c <= sel.c2; c++) vals.push(displayAt(sheet, valueAt, r, c));
      rows.push(vals.join('\t'));
    }
    const internal: (Cell | null)[][] = [];
    for (let r = sel.r1; r <= sel.r2; r++) {
      const row: (Cell | null)[] = [];
      for (let c = sel.c1; c <= sel.c2; c++) row.push(cellAt(r, c) ?? null);
      internal.push(row);
    }
    const text = rows.join('\n');
    const payload = JSON.stringify({ src: { r: sel.r1, c: sel.c1 }, cells: internal });
    lastCopy = { text, payload };
    return { text, payload };
  };

  const onCopy = (e: React.ClipboardEvent, cut = false) => {
    if (editing) return;
    e.preventDefault();
    const { text, payload } = copyPayload();
    e.clipboardData.setData('text/plain', text);
    e.clipboardData.setData('application/x-paperline-cells', payload);
    if (cut) forSel(() => undefined);
  };

  const pasteData = (internal: string | null, text: string) => {
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

  const onPaste = (e: React.ClipboardEvent) => {
    if (editing) return;
    e.preventDefault();
    pasteData(e.clipboardData.getData('application/x-paperline-cells') || null, e.clipboardData.getData('text/plain'));
  };

  const menuCopy = (cut: boolean) => {
    const { text } = copyPayload();
    navigator.clipboard?.writeText(text).catch(() => null);
    if (cut) forSel(() => undefined);
  };
  const menuPaste = async () => {
    let text = '';
    try { text = (await navigator.clipboard?.readText()) ?? ''; } catch { /* permission denied */ }
    if (lastCopy && (text === '' || text === lastCopy.text)) pasteData(lastCopy.payload, lastCopy.text);
    else if (text) pasteData(null, text);
  };

  // ----- mouse -----
  const onCellDown = (e: React.MouseEvent, r: number, c: number) => {
    setCtx(null);
    setListOpen(false);
    if (e.button === 2) {
      if (!inRange(sel, r, c)) { setAnchor({ r, c }); setFocus({ r, c }); }
      return;
    }
    if (e.button !== 0) return;
    setChartSel(null);
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
    if (filling.current) setFillTo({ r, c });
    else if (dragging.current) setFocus({ r, c });
    const has = !!sheet.cells[addr(r, c)]?.note;
    if (has || hoverNote) setHoverNote(has ? { r, c } : null);
  };
  useEffect(() => {
    const up = () => {
      dragging.current = false;
      if (filling.current) {
        filling.current = false;
        applyFillRef.current();
      }
    };
    window.addEventListener('mouseup', up);
    return () => window.removeEventListener('mouseup', up);
  }, []);
  useEffect(() => {
    if (!ctx) return;
    const h = (e: MouseEvent) => { if (!(e.target as HTMLElement).closest('.ctx-menu')) setCtx(null); };
    document.addEventListener('mousedown', h);
    return () => document.removeEventListener('mousedown', h);
  }, [ctx]);

  const selectCol = (c: number, extend: boolean) => {
    if (!extend) setAnchor({ r: 0, c });
    setFocus({ r: sheet.rows - 1, c });
    setChartSel(null);
    focusGrid();
  };
  const selectRow = (r: number, extend: boolean) => {
    if (!extend) setAnchor({ r, c: 0 });
    setFocus({ r, c: sheet.cols - 1 });
    setChartSel(null);
    focusGrid();
  };

  const startResize = (e: React.MouseEvent, axis: 'c' | 'r', i: number) => {
    e.preventDefault();
    e.stopPropagation();
    const start = axis === 'c' ? e.clientX : e.clientY;
    const startSize = axis === 'c' ? colW(i) : rowH(i);
    const rtl = axis === 'c' && getComputedStyle(gridRef.current!).direction === 'rtl';
    let size = startSize;
    const guide = document.createElement('div');
    guide.className = `resize-guide ${axis === 'r' ? 'row' : ''}`;
    document.body.appendChild(guide);
    const move = (ev: MouseEvent) => {
      const pos = axis === 'c' ? ev.clientX : ev.clientY;
      size = Math.max(axis === 'c' ? 24 : 14, startSize + ((rtl ? -1 : 1) * (pos - start)) / zoom);
      if (axis === 'c') guide.style.left = `${pos}px`;
      else guide.style.top = `${pos}px`;
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      guide.remove();
      if (size === startSize) return;
      // resizing one of several selected headers resizes them all, like Excel
      const whole = axis === 'c'
        ? sel.r1 === 0 && sel.r2 === sheet.rows - 1 && i >= sel.c1 && i <= sel.c2
        : sel.c1 === 0 && sel.c2 === sheet.cols - 1 && i >= sel.r1 && i <= sel.r2;
      const idx = whole ? (axis === 'c' ? range(sel.c1, sel.c2) : range(sel.r1, sel.r2)) : [i];
      updateSheet((sh) => {
        const key = axis === 'c' ? 'colW' : 'rowH';
        const m = { ...sh[key] };
        for (const j of idx) m[j] = Math.round(size);
        return { ...sh, [key]: m };
      });
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  /** Fit columns to their widest value. */
  const autofit = (cols: number[]) => {
    const g = document.createElement('canvas').getContext('2d')!;
    updateSheet((sh) => {
      const colWs = { ...sh.colW };
      for (const c of cols) {
        let w = 0;
        for (const k of Object.keys(sh.cells)) {
          const p = parseAddr(k)!;
          if (p.c !== c || mergeAt(sh, p.r, p.c)) continue;
          const st = sh.cells[k].st ?? {};
          g.font = `${st.b ? 700 : 400} ${st.fontSize ? (st.fontSize * 4) / 3 : 13}px ${st.font ? familyCss(st.font) : 'Inter, sans-serif'}`;
          const text = showFormulas ? sh.cells[k].v : formatValue(engine.cellValue(sh, p.r, p.c), st.fmt, st.dec);
          w = Math.max(w, g.measureText(text).width);
        }
        colWs[c] = w ? Math.min(600, Math.ceil(w + 18)) : DEFAULT_COL_W;
      }
      return { ...sh, colW: colWs };
    });
  };

  // ----- fill handle -----
  const fillRange = (target: { r: number; c: number } | null): { dir: 'down' | 'up' | 'right' | 'left'; n: number } | null => {
    if (!target) return null;
    const opts = [
      { dir: 'down' as const, n: target.r - sel.r2 },
      { dir: 'up' as const, n: sel.r1 - target.r },
      { dir: 'right' as const, n: target.c - sel.c2 },
      { dir: 'left' as const, n: sel.c1 - target.c },
    ].sort((a, b) => b.n - a.n);
    return opts[0].n > 0 ? opts[0] : null;
  };

  const doFill = (dir: 'down' | 'up' | 'right' | 'left', n: number, base: Range = sel) => {
    if (n <= 0) return;
    const vertical = dir === 'down' || dir === 'up';
    const sign = dir === 'down' || dir === 'right' ? 1 : -1;
    const edge = sign > 0 ? (vertical ? base.r2 : base.c2) : vertical ? base.r1 : base.c1;
    updateSheet((sh) => {
      const cells = { ...sh.cells };
      const lanes = vertical ? range(base.c1, base.c2) : range(base.r1, base.r2);
      for (const lane of lanes) {
        const along = vertical ? range(base.r1, base.r2) : range(base.c1, base.c2);
        if (sign < 0) along.reverse();
        const src = along.map((i) => (vertical ? sh.cells[addr(i, lane)] : sh.cells[addr(lane, i)]));
        const out = fillSeries(src, n, (v, k) => relocate(v, vertical ? sign * k : 0, vertical ? 0 : sign * k));
        out.forEach((cell, i) => {
          const pos = edge + sign * (i + 1);
          if (pos < 0) return;
          const k = vertical ? addr(pos, lane) : addr(lane, pos);
          if (cell) cells[k] = cell;
          else delete cells[k];
        });
      }
      const end = edge + sign * n;
      return { ...sh, cells, rows: Math.max(sh.rows, vertical ? end + 2 : sh.rows), cols: Math.max(sh.cols, vertical ? sh.cols : end + 1) };
    });
    const nr = { ...base };
    if (dir === 'down') nr.r2 += n;
    if (dir === 'up') nr.r1 = Math.max(0, nr.r1 - n);
    if (dir === 'right') nr.c2 += n;
    if (dir === 'left') nr.c1 = Math.max(0, nr.c1 - n);
    setAnchor({ r: nr.r1, c: nr.c1 });
    setFocus({ r: nr.r2, c: nr.c2 });
  };
  applyFillRef.current = () => {
    const f = fillRange(fillTo);
    setFillTo(null);
    if (f) doFill(f.dir, f.n);
  };
  /** Ctrl+D / Ctrl+R: copy the first row (or column) of the selection across it. */
  const fillDir = (dir: 'down' | 'right') => {
    if (dir === 'down' && sel.r2 > sel.r1) doFill('down', sel.r2 - sel.r1, { ...sel, r2: sel.r1 });
    else if (dir === 'right' && sel.c2 > sel.c1) doFill('right', sel.c2 - sel.c1, { ...sel, c2: sel.c1 });
  };
  /** Double-click on the fill handle: fill down as far as the neighbouring column has data. */
  const fillAuto = () => {
    const filled = (r: number, c: number) => c >= 0 && (sheet.cells[addr(r, c)]?.v ?? '') !== '';
    const side = filled(sel.r2 + 1, sel.c1 - 1) ? sel.c1 - 1 : filled(sel.r2 + 1, sel.c2 + 1) ? sel.c2 + 1 : -1;
    if (side < 0) return;
    let end = sel.r2 + 1;
    while (filled(end + 1, side)) end++;
    doFill('down', end - sel.r2);
  };

  // ----- commands -----
  const insertFn = (name: string) => {
    const r = sel.r1, c = sel.c1;
    let arg = '';
    if (!single) {
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

  /** The block a command works on: the selection, or the data around the active cell. */
  const workRange = () => {
    if (!single) return sel;
    const rg = currentRegion(sheet, anchor.r, anchor.c);
    // skip a title row above the table (one filled cell, e.g. a merged heading)
    const filled = (r: number) => range(rg.c1, rg.c2).filter((c) => (cellAt(r, c)?.v ?? '') !== '').length;
    while (rg.r2 - rg.r1 > 1 && anchor.r > rg.r1 && filled(rg.r1) === 1 && filled(rg.r1 + 1) > 1) rg.r1++;
    return rg;
  };
  const hasHeader = (rg: Range) => {
    for (let c = rg.c1; c <= rg.c2; c++) {
      const a = valueAt(rg.r1, c), b = valueAt(rg.r1 + 1, c);
      if (typeof a === 'string' && a !== '' && typeof b !== 'string') return true;
    }
    return false;
  };

  const sortRange = (rg: Range, key: number, dir: 1 | -1) => {
    const rows: { k: unknown; cells: (Cell | undefined)[] }[] = [];
    for (let r = rg.r1; r <= rg.r2; r++) {
      const cells: (Cell | undefined)[] = [];
      for (let c = rg.c1; c <= rg.c2; c++) cells.push(cellAt(r, c));
      rows.push({ k: valueAt(r, key), cells });
    }
    rows.sort((a, b) => {
      if ((a.k === null || a.k === '') && (b.k === null || b.k === '')) return 0;
      if (a.k === null || a.k === '') return 1;
      if (b.k === null || b.k === '') return -1;
      if (typeof a.k === 'number' && typeof b.k === 'number') return (a.k - b.k) * dir;
      return String(a.k).localeCompare(String(b.k), undefined, { numeric: true }) * dir;
    });
    updateSheet((sh) => {
      const cells = { ...sh.cells };
      rows.forEach((row, i) => row.cells.forEach((cell, j) => {
        const k = addr(rg.r1 + i, rg.c1 + j);
        if (cell) cells[k] = cell;
        else delete cells[k];
      }));
      return { ...sh, cells };
    });
  };
  const sortSel = (dir: 1 | -1) => {
    const f = sheet.filter;
    if (f && inRange({ ...f.range, r2: filterEnd(sheet) }, anchor.r, anchor.c)) {
      sortRange({ ...f.range, r1: f.range.r1 + 1, r2: filterEnd(sheet) }, anchor.c, dir);
      return;
    }
    const rg = { ...workRange() };
    if (single && hasHeader(rg)) rg.r1++;
    sortRange(rg, anchor.c, dir);
  };

  const insertRows = () => updateSheet((sh) => shiftSheet(sh, 'r', sel.r1, sel.r2 - sel.r1 + 1));
  const insertCols = () => updateSheet((sh) => shiftSheet(sh, 'c', sel.c1, sel.c2 - sel.c1 + 1));
  const deleteRows = () => updateSheet((sh) => shiftSheet(sh, 'r', sel.r1, -(sel.r2 - sel.r1 + 1)));
  const deleteCols = () => updateSheet((sh) => shiftSheet(sh, 'c', sel.c1, -(sel.c2 - sel.c1 + 1)));
  const setHidden = (axis: 'r' | 'c', hide: boolean) =>
    updateSheet((sh) => {
      const key = axis === 'r' ? 'hiddenR' : 'hiddenC';
      const [a, b] = axis === 'r' ? [sel.r1, sel.r2] : [sel.c1, sel.c2];
      const cur = new Set(sh[key] ?? []);
      if (hide) range(a, b).forEach((i) => cur.add(i));
      else {
        const inside = [...cur].filter((i) => i >= a - 1 && i <= b + 1);
        (inside.length ? inside : [...cur]).forEach((i) => cur.delete(i));
      }
      const list = [...cur].sort((x, y) => x - y);
      return { ...sh, [key]: list.length ? list : undefined };
    });

  const doMerge = (mode: 'center' | 'across' | 'merge' | 'unmerge') => {
    updateSheet((sh) => {
      if (mode === 'unmerge') return unmergeCells(sh, sel);
      if (mode === 'across') {
        let out = sh;
        for (let r = sel.r1; r <= sel.r2; r++) out = mergeCells(out, { r1: r, r2: r, c1: sel.c1, c2: sel.c2 }, false);
        return out;
      }
      return mergeCells(sh, sel, mode === 'center');
    });
  };
  const isMerged = !!sheet.merges?.some((m) => m.r1 >= sel.r1 && m.r2 <= sel.r2 && m.c1 >= sel.c1 && m.c2 <= sel.c2);

  // filter
  const toggleFilter = () => {
    if (sheet.filter) { updateSheet((sh) => ({ ...sh, filter: undefined })); return; }
    const rg = workRange();
    if (rg.r2 <= rg.r1) { setFlash(t('s_needData')); return; }
    updateSheet((sh) => ({ ...sh, filter: { range: rg, show: {} } }));
  };
  const setFilterShow = (c: number, show: string[] | null) =>
    updateSheet((sh) => {
      if (!sh.filter) return sh;
      const s = { ...sh.filter.show };
      if (show) s[c] = show;
      else delete s[c];
      return { ...sh, filter: { ...sh.filter, show: s } };
    });

  const removeDuplicates = () => {
    const rg = { ...workRange() };
    if (hasHeader(rg)) rg.r1++;
    const seen = new Set<string>();
    const keep: (Cell | undefined)[][] = [];
    for (let r = rg.r1; r <= rg.r2; r++) {
      const key = range(rg.c1, rg.c2).map((c) => displayAt(sheet, valueAt, r, c)).join('\u0001');
      if (seen.has(key)) continue;
      seen.add(key);
      keep.push(range(rg.c1, rg.c2).map((c) => cellAt(r, c)));
    }
    const removed = rg.r2 - rg.r1 + 1 - keep.length;
    if (removed) {
      updateSheet((sh) => {
        const cells = { ...sh.cells };
        for (let r = rg.r1; r <= rg.r2; r++)
          for (let c = rg.c1; c <= rg.c2; c++) {
            const cell = keep[r - rg.r1]?.[c - rg.c1];
            if (cell) cells[addr(r, c)] = cell;
            else delete cells[addr(r, c)];
          }
        return { ...sh, cells };
      });
    }
    setFlash(t('s_dupsRemoved').replace('{n}', String(removed)));
  };

  const formatAsTable = (color: string) => {
    const rg = workRange();
    const tint = mixHex(color, '#ffffff', 0.86);
    updateSheet((sh) => {
      const cells = { ...sh.cells };
      for (let r = rg.r1; r <= rg.r2; r++)
        for (let c = rg.c1; c <= rg.c2; c++) {
          const k = addr(r, c);
          const prev = cells[k];
          const head = r === rg.r1;
          const st: CellStyle = { ...prev?.st, bg: head ? color : (r - rg.r1) % 2 === 0 ? tint : undefined, color: head ? '#ffffff' : prev?.st?.color, b: head || prev?.st?.b };
          if (r === rg.r2) st.bd = { ...st.bd, b: 1 };
          cells[k] = { ...prev, v: prev?.v ?? '', st: clean(st) };
        }
      return { ...sh, cells, filter: { range: rg, show: {} } };
    });
  };

  // conditional formatting
  const addRule = (rule: Omit<CFRule, 'id' | 'range'>) => {
    const rg = { ...sel };
    updateSheet((sh) => ({ ...sh, cf: [...(sh.cf ?? []), { ...rule, id: Math.random().toString(36).slice(2, 9), range: rg }] }));
  };
  const askRule = (kind: CFKind) => {
    const styleField: AskField = {
      key: 'style', label: t('s_cfWith'), value: 'red', type: 'select',
      options: Object.keys(HL_STYLES).map((k) => ({ value: k, label: t(`s_hl_${k}` as Key) })),
    };
    const two = kind === 'between';
    const fields: AskField[] =
      kind === 'dup' || kind === 'blank' || kind === 'aboveAvg' ? [styleField]
      : kind === 'top' || kind === 'bottom' ? [{ key: 'a', label: t('s_cfCount'), value: '10', type: 'number' }, styleField]
      : [{ key: 'a', label: two ? t('s_cfFrom') : t('s_cfValue'), value: '' }, ...(two ? [{ key: 'b', label: t('s_cfTo'), value: '' }] : []), styleField];
    setAsk({
      title: t(`s_cf_${kind}` as Key),
      fields,
      onOk: (v) => addRule({ kind, a: v.a, b: v.b, ...HL_STYLES[v.style] }),
    });
  };
  const clearRules = (all: boolean) =>
    updateSheet((sh) => {
      const cf = all ? [] : (sh.cf ?? []).filter((x) => !(x.range.r1 <= sel.r2 && sel.r1 <= x.range.r2 && x.range.c1 <= sel.c2 && sel.c1 <= x.range.c2));
      return { ...sh, cf: cf.length ? cf : undefined };
    });

  // charts
  const insertChart = (type: ChartType) => {
    const rg = workRange();
    const ch: ChartSpec = {
      id: Math.random().toString(36).slice(2, 9),
      type,
      range: rg,
      title: hasHeader(rg) && rg.c2 - rg.c1 === 1 ? String(valueAt(rg.r1, rg.c2) ?? '') : '',
      x: colLefts[Math.min(Math.max(rg.c2, currentRegion(sheet, rg.r1, rg.c1).c2) + 1, sheet.cols)] + 16,
      y: rowTops[rg.r1],
      w: 460,
      h: 290,
      legend: true,
    };
    updateSheet((sh) => ({ ...sh, charts: [...(sh.charts ?? []), ch] }));
    setChartSel(ch.id);
    focusGrid();
  };
  const updateChart = (id: string, patch: Partial<ChartSpec>) =>
    updateSheet((sh) => ({ ...sh, charts: sh.charts?.map((x) => (x.id === id ? { ...x, ...patch } : x)) }));
  const deleteChart = (id: string) => {
    updateSheet((sh) => {
      const charts = sh.charts?.filter((x) => x.id !== id);
      return { ...sh, charts: charts?.length ? charts : undefined };
    });
    setChartSel(null);
    focusGrid();
  };
  const askChartTitle = (ch: ChartSpec) =>
    setAsk({ title: t('s_chartTitle'), fields: [{ key: 'title', label: t('title'), value: ch.title ?? '' }], onOk: (v) => updateChart(ch.id, { title: v.title }) });
  const startChartDrag = (e: React.MouseEvent, ch: ChartSpec, mode: 'move' | 'resize') => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    setChartSel(ch.id);
    focusGrid();
    const sx = e.clientX, sy = e.clientY;
    const el = (e.currentTarget as HTMLElement).closest('.sheet-chart') as HTMLElement;
    let next = { x: ch.x, y: ch.y, w: ch.w, h: ch.h };
    const move = (ev: MouseEvent) => {
      const dx = (ev.clientX - sx) / zoom, dy = (ev.clientY - sy) / zoom;
      next = mode === 'move'
        ? { ...next, x: Math.max(0, ch.x + dx), y: Math.max(0, ch.y + dy) }
        : { ...next, w: Math.max(180, ch.w + dx), h: Math.max(140, ch.h + dy) };
      el.style.left = `${HEADER_W + next.x}px`;
      el.style.top = `${HEADER_H + next.y}px`;
      if (mode === 'resize') { el.style.width = `${next.w}px`; el.style.height = `${next.h}px`; }
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      if (next.x !== ch.x || next.y !== ch.y || next.w !== ch.w || next.h !== ch.h)
        updateChart(ch.id, { x: Math.round(next.x), y: Math.round(next.y), w: Math.round(next.w), h: Math.round(next.h) });
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  // notes and lists
  const editNote = () =>
    setAsk({
      title: activeCell?.note ? t('s_editNote') : t('s_newNote'),
      fields: [{ key: 'note', label: t('s_note'), value: activeCell?.note ?? '', type: 'textarea' }],
      onOk: (v) => forRange({ r1: anchor.r, c1: anchor.c, r2: anchor.r, c2: anchor.c }, (cell) => ({ ...cell, v: cell?.v ?? '', note: v.note.trim() || undefined })),
    });
  const deleteNotes = () => forSel((cell) => (cell ? { ...cell, note: undefined } : cell));
  const editList = () =>
    setAsk({
      title: t('s_dropdown'),
      fields: [{ key: 'items', label: t('s_listItems'), value: activeCell?.list?.join(', ') ?? '', type: 'textarea' }],
      onOk: (v) => {
        const items = v.items.split(/[,،\n]/).map((x) => x.trim()).filter(Boolean);
        forSel((cell) => ({ ...cell, v: cell?.v ?? '', list: items.length ? [...new Set(items)] : undefined }));
      },
    });

  const askSize = (axis: 'r' | 'c') =>
    setAsk({
      title: axis === 'r' ? t('s_rowHeight') : t('s_colWidth'),
      fields: [{ key: 'v', label: 'px', value: String(axis === 'r' ? rowH(anchor.r) || DEFAULT_ROW_H : colW(anchor.c) || DEFAULT_COL_W), type: 'number' }],
      onOk: (v) => {
        const n = Math.max(axis === 'r' ? 10 : 16, Math.min(800, Number(v.v) || 0));
        updateSheet((sh) => {
          const key = axis === 'r' ? 'rowH' : 'colW';
          const m = { ...sh[key] };
          for (const i of axis === 'r' ? range(sel.r1, sel.r2) : range(sel.c1, sel.c2)) m[i] = n;
          return { ...sh, [key]: m };
        });
      },
    });

  const setActiveSheet = (i: number) => {
    if (editing) commitEdit();
    onChange({ ...data, active: i });
    setAnchor({ r: 0, c: 0 });
    setFocus({ r: 0, c: 0 });
    setChartSel(null);
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
    if (single) return null;
    let sum = 0, n = 0, cnt = 0, mn = Infinity, mx = -Infinity;
    const r2 = Math.min(sel.r2, sel.r1 + 5000);
    for (let r = sel.r1; r <= r2; r++) {
      if (hiddenR.has(r)) continue;
      for (let c = sel.c1; c <= sel.c2; c++) {
        const v = valueAt(r, c);
        if (v !== null && v !== '') cnt++;
        if (typeof v === 'number') { sum += v; n++; mn = Math.min(mn, v); mx = Math.max(mx, v); }
      }
    }
    return { sum, avg: n ? sum / n : 0, cnt, n, mn, mx };
  }, [single, sel.r1, sel.r2, sel.c1, sel.c2, valueAt, hiddenR]);

  // ----- ribbon -----
  const barValue = editing ? editing.value : activeCell?.v ?? '';
  const fmtOptions: { value: NumFmt; label: string }[] = [
    { value: 'general', label: t('formatGeneral') },
    { value: 'number', label: t('formatNumber') },
    { value: 'currency', label: t('formatCurrency') },
    { value: 'percent', label: t('formatPercent') },
    { value: 'date', label: t('formatDate') },
    { value: 'time', label: t('s_fmtTime') },
    { value: 'datetime', label: t('s_fmtDateTime') },
    { value: 'scientific', label: t('s_fmtScientific') },
    { value: 'text', label: t('s_fmtText') },
  ];
  const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48].map((s) => ({ value: s, label: String(s) }));
  const menuBtn = (icon: ReactNode, label: string) => (open: boolean, toggle: () => void) => (
    <Btn icon={<span className="with-caret">{icon}<ChevronDown size={11} /></span>} label={label} active={open} onClick={toggle} />
  );
  const bigMenuBtn = (icon: ReactNode, label: string) => (open: boolean, toggle: () => void) => <Btn big icon={icon} label={label} active={open} onClick={toggle} />;

  const borderItems: { mode: BorderMode; icon: ReactNode }[] = [
    { mode: 'bottom', icon: <PanelBottom size={16} /> },
    { mode: 'top', icon: <PanelTop size={16} /> },
    { mode: 'left', icon: <PanelLeft size={16} /> },
    { mode: 'right', icon: <PanelRight size={16} /> },
    { mode: 'all', icon: <Grid2x2 size={16} /> },
    { mode: 'outer', icon: <Square size={16} /> },
    { mode: 'inner', icon: <Grid3x3 size={16} /> },
    { mode: 'thickOuter', icon: <Square size={16} strokeWidth={3} /> },
    { mode: 'thickBottom', icon: <PanelBottom size={16} strokeWidth={3} /> },
    { mode: 'none', icon: <SquareDashed size={16} /> },
  ];

  const cfMenu = (
    <Menu trigger={bigMenuBtn(<Highlighter size={22} />, t('s_condFormat'))}>
      {(close) => (
        <>
          <div className="menu-label">{t('s_highlightRules')}</div>
          {(['gt', 'lt', 'between', 'eq', 'text', 'dup', 'blank'] as CFKind[]).map((k) => (
            <MenuItem key={k} label={t(`s_cf_${k}` as Key)} onClick={() => { close(); askRule(k); }} />
          ))}
          <div className="menu-label">{t('s_topBottom')}</div>
          {(['top', 'bottom', 'aboveAvg'] as CFKind[]).map((k) => (
            <MenuItem key={k} label={t(`s_cf_${k}` as Key)} onClick={() => { close(); askRule(k); }} />
          ))}
          <div className="menu-label">{t('s_dataBars')}</div>
          <div className="cf-swatches">
            {['#638ec6', '#63be7b', '#ffb628', '#f8696b', '#8b5cf6'].map((c) => (
              <button key={c} type="button" className="cf-swatch bar" title={t('s_dataBars')} onMouseDown={(e) => e.preventDefault()} onClick={() => { close(); addRule({ kind: 'bar', c1: c }); }}>
                <i style={{ background: `linear-gradient(90deg, ${c}, ${c}33)` }} />
              </button>
            ))}
          </div>
          <div className="menu-label">{t('s_colorScales')}</div>
          <div className="cf-swatches">
            {[
              ['#f8696b', '#ffeb84', '#63be7b'],
              ['#63be7b', '#ffeb84', '#f8696b'],
              ['#ffffff', '', '#5a8ac6'],
              ['#ffffff', '', '#63be7b'],
              ['#f8696b', '#ffffff', '#5a8ac6'],
            ].map(([lo, mid, hi]) => (
              <button key={lo + mid + hi} type="button" className="cf-swatch" title={t('s_colorScales')} onMouseDown={(e) => e.preventDefault()} onClick={() => { close(); addRule({ kind: 'scale', c2: lo, mid: mid || undefined, c1: hi }); }}>
                <i style={{ background: `linear-gradient(90deg, ${lo}, ${mid || lo}, ${hi})` }} />
              </button>
            ))}
          </div>
          <div className="menu-sep" />
          <MenuItem icon={<Eraser size={15} />} label={t('s_clearRulesSel')} onClick={() => { close(); clearRules(false); }} />
          <MenuItem icon={<Trash2 size={15} />} label={t('s_clearRulesAll')} onClick={() => { close(); clearRules(true); }} />
        </>
      )}
    </Menu>
  );

  const tableMenu = (
    <Menu trigger={bigMenuBtn(<Table2 size={22} />, t('s_formatTable'))}>
      {(close) => (
        <div className="table-styles">
          {TABLE_COLORS.map((c) => (
            <button key={c} type="button" className="table-style" title={t('s_formatTable')} onMouseDown={(e) => e.preventDefault()} onClick={() => { close(); formatAsTable(c); }}>
              <i style={{ background: c }} />
              <i style={{ background: mixHex(c, '#ffffff', 0.86) }} />
              <i />
              <i style={{ background: mixHex(c, '#ffffff', 0.86) }} />
            </button>
          ))}
        </div>
      )}
    </Menu>
  );

  const formatMenu = (
    <Menu trigger={menuBtn(<Rows3 size={18} />, t('format'))}>
      {(close) => (
        <>
          <MenuItem icon={<MoveVertical size={15} />} label={t('s_rowHeight')} onClick={() => { close(); askSize('r'); }} />
          <MenuItem icon={<MoveHorizontal size={15} />} label={t('s_colWidth')} onClick={() => { close(); askSize('c'); }} />
          <MenuItem icon={<ArrowLeftRight size={15} />} label={t('s_autofit')} onClick={() => { close(); autofit(range(sel.c1, sel.c2)); }} />
          <div className="menu-sep" />
          <MenuItem icon={<EyeOff size={15} />} label={t('s_hideRows')} onClick={() => { close(); setHidden('r', true); }} />
          <MenuItem icon={<EyeOff size={15} />} label={t('s_hideCols')} onClick={() => { close(); setHidden('c', true); }} />
          <MenuItem icon={<Eye size={15} />} label={t('s_unhideRows')} onClick={() => { close(); setHidden('r', false); }} />
          <MenuItem icon={<Eye size={15} />} label={t('s_unhideCols')} onClick={() => { close(); setHidden('c', false); }} />
        </>
      )}
    </Menu>
  );

  const homeTab = (
    <>
      <Group>
        <Btn icon={<Undo2 size={18} />} label={t('undo')} onClick={undo} />
        <Btn icon={<Redo2 size={18} />} label={t('redo')} onClick={redo} />
      </Group>
      <Group label={t('font')}>
        <FontPicker value={activeStyle.font ?? 'Calibri'} onChange={(f) => { applyStyle({ font: f === 'Calibri' ? undefined : f }); focusGrid(); }} width={140} />
        <Select label={t('size')} value={activeStyle.fontSize ?? 11} options={sizes} onChange={(v) => applyStyle({ fontSize: v === 11 ? undefined : v })} width={64} />
        <Btn icon={<Bold size={18} />} label={t('bold')} active={!!activeStyle.b} onClick={() => toggleStyle('b')} />
        <Btn icon={<Italic size={18} />} label={t('italic')} active={!!activeStyle.i} onClick={() => toggleStyle('i')} />
        <Btn icon={<Underline size={18} />} label={t('underline')} active={!!activeStyle.u} onClick={() => toggleStyle('u')} />
        <Btn icon={<Strikethrough size={18} />} label={t('strike')} active={!!activeStyle.s} onClick={() => toggleStyle('s')} />
        <Menu trigger={menuBtn(<Square size={18} />, t('borders'))}>
          {(close) => (
            <>
              {borderItems.map((b) => (
                <MenuItem key={b.mode} icon={b.icon} label={t(`s_bd_${b.mode}` as Key)} onClick={() => { close(); updateSheet((sh) => applyBorders(sh, sel, b.mode)); }} />
              ))}
            </>
          )}
        </Menu>
        <ColorBtn icon={<Type size={16} />} label={t('textColor')} color={activeStyle.color ?? '#000'} onPick={(c) => applyStyle({ color: c ?? undefined })} allowNone />
        <ColorBtn icon={<PaintBucket size={16} />} label={t('fillColor')} color={activeStyle.bg} onPick={(c) => applyStyle({ bg: c ?? undefined })} allowNone />
      </Group>
      <Group label={t('s_alignment')}>
        <Btn icon={<AlignVerticalJustifyStart size={18} />} label={t('s_alignTop')} active={activeStyle.va === 'top'} onClick={() => applyStyle({ va: activeStyle.va === 'top' ? undefined : 'top' })} />
        <Btn icon={<AlignVerticalJustifyCenter size={18} />} label={t('s_alignMiddle')} active={activeStyle.va === 'middle'} onClick={() => applyStyle({ va: activeStyle.va === 'middle' ? undefined : 'middle' })} />
        <Btn icon={<AlignVerticalJustifyEnd size={18} />} label={t('s_alignBottom')} active={activeStyle.va === 'bottom'} onClick={() => applyStyle({ va: activeStyle.va === 'bottom' ? undefined : 'bottom' })} />
        <Btn icon={<AlignLeft size={18} />} label={t('alignLeft')} active={activeStyle.align === 'left'} onClick={() => applyStyle({ align: activeStyle.align === 'left' ? undefined : 'left' })} />
        <Btn icon={<AlignCenter size={18} />} label={t('alignCenter')} active={activeStyle.align === 'center'} onClick={() => applyStyle({ align: activeStyle.align === 'center' ? undefined : 'center' })} />
        <Btn icon={<AlignRight size={18} />} label={t('alignRight')} active={activeStyle.align === 'right'} onClick={() => applyStyle({ align: activeStyle.align === 'right' ? undefined : 'right' })} />
        <Btn icon={<WrapText size={18} />} label={t('wrap')} active={!!activeStyle.wrap} onClick={() => toggleStyle('wrap')} />
        <Menu trigger={menuBtn(<Merge size={18} />, t('s_merge'))}>
          {(close) => (
            <>
              <MenuItem label={t('s_mergeCenter')} onClick={() => { close(); doMerge(isMerged ? 'unmerge' : 'center'); }} />
              <MenuItem label={t('s_mergeAcross')} onClick={() => { close(); doMerge('across'); }} />
              <MenuItem label={t('s_mergeCells')} onClick={() => { close(); doMerge('merge'); }} />
              <MenuItem label={t('s_unmerge')} onClick={() => { close(); doMerge('unmerge'); }} />
            </>
          )}
        </Menu>
      </Group>
      <Group label={t('formatNumber')}>
        <Select label={t('formatNumber')} value={activeStyle.fmt ?? 'general'} options={fmtOptions} onChange={(v) => applyStyle({ fmt: v === 'general' ? undefined : v, dec: undefined })} width={110} />
        <Btn icon={<DollarSign size={18} />} label={t('formatCurrency')} onClick={() => applyStyle({ fmt: 'currency', dec: undefined })} />
        <Btn icon={<Percent size={18} />} label={t('formatPercent')} onClick={() => applyStyle({ fmt: 'percent', dec: undefined })} />
        <Btn icon={<Plus size={14} />} label={t('decInc')} onClick={() => applyStyle({ dec: Math.min(10, (activeStyle.dec ?? defaultDec(activeStyle.fmt)) + 1), fmt: activeStyle.fmt ?? 'number' })} />
        <Btn icon={<Minus size={14} />} label={t('decDec')} onClick={() => applyStyle({ dec: Math.max(0, (activeStyle.dec ?? defaultDec(activeStyle.fmt)) - 1), fmt: activeStyle.fmt ?? 'number' })} />
      </Group>
      <Group label={t('style')}>
        {cfMenu}
        {tableMenu}
      </Group>
      <Group label={t('s_cells')}>
        <Btn icon={<Rows3 size={18} />} label={t('insertRow')} onClick={insertRows} />
        <Btn icon={<Columns3 size={18} />} label={t('insertCol')} onClick={insertCols} />
        <Btn icon={<Trash2 size={16} />} label={t('deleteRow')} onClick={deleteRows} />
        <Btn icon={<span className="rot90"><Trash2 size={16} /></span>} label={t('deleteCol')} onClick={deleteCols} />
        {formatMenu}
        <Btn icon={<Eraser size={18} />} label={t('clearFormat')} onClick={() => forSel((cell) => (cell?.v || cell?.note || cell?.list ? { v: cell.v, note: cell.note, list: cell.list } : undefined))} />
      </Group>
      <Group label={t('s_editing')}>
        <Btn big icon={<Sigma size={22} />} label={t('autoSum')} onClick={() => insertFn('SUM')} />
        <Btn icon={<Search size={18} />} label={t('find')} onClick={() => setFind({ replace: false })} />
        <Btn icon={<Replace size={18} />} label={t('replace')} onClick={() => setFind({ replace: true })} />
        <Btn icon={<ArrowDownAZ size={18} />} label={t('sortAsc')} onClick={() => sortSel(1)} />
        <Btn icon={<Filter size={18} />} label={t('s_filter')} active={!!sheet.filter} onClick={toggleFilter} />
      </Group>
    </>
  );

  const chartTypes: { type: ChartType; icon: ReactNode }[] = [
    { type: 'column', icon: <ChartColumn size={22} /> },
    { type: 'bar', icon: <ChartBar size={22} /> },
    { type: 'line', icon: <ChartLine size={22} /> },
    { type: 'area', icon: <ChartArea size={22} /> },
    { type: 'pie', icon: <ChartPie size={22} /> },
    { type: 'doughnut', icon: <Donut size={22} /> },
    { type: 'scatter', icon: <ChartScatter size={22} /> },
  ];

  const insertTab = (
    <>
      <Group label={t('s_tables')}>{tableMenu}</Group>
      <Group label={t('s_charts')}>
        {chartTypes.map((c) => <Btn key={c.type} big icon={c.icon} label={t(`s_chart_${c.type}` as Key)} onClick={() => insertChart(c.type)} />)}
      </Group>
      <Group label={t('s_notes')}>
        <Btn big icon={<StickyNote size={22} />} label={t('s_newNote')} onClick={editNote} />
      </Group>
      <Group label={t('functions')}>
        <Btn big icon={<Sigma size={22} />} label={t('autoSum')} onClick={() => insertFn('SUM')} />
      </Group>
    </>
  );

  const fnGroups: [Key, string[]][] = [
    ['s_fnMath', ['SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'ABS', 'INT', 'MOD', 'POWER', 'SQRT', 'CEILING', 'FLOOR', 'TRUNC', 'RANDBETWEEN']],
    ['s_fnStats', ['AVERAGE', 'AVERAGEIF', 'AVERAGEIFS', 'COUNT', 'COUNTA', 'COUNTIF', 'COUNTIFS', 'MAX', 'MAXIFS', 'MIN', 'MINIFS', 'MEDIAN', 'MODE', 'STDEV', 'VAR', 'LARGE', 'SMALL', 'RANK']],
    ['s_fnLogic', ['IF', 'IFS', 'IFERROR', 'AND', 'OR', 'NOT', 'SWITCH', 'CHOOSE']],
    ['s_fnLookup', ['XLOOKUP', 'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'ROWS', 'COLUMNS']],
    ['s_fnText', ['CONCAT', 'TEXTJOIN', 'LEFT', 'RIGHT', 'MID', 'LEN', 'UPPER', 'LOWER', 'PROPER', 'TRIM', 'SUBSTITUTE', 'REPLACE', 'FIND', 'SEARCH', 'TEXT', 'VALUE', 'REPT']],
    ['s_fnDate', ['TODAY', 'NOW', 'DATE', 'YEAR', 'MONTH', 'DAY', 'WEEKDAY', 'EDATE', 'EOMONTH', 'DATEDIF', 'NETWORKDAYS', 'TIME', 'HOUR', 'MINUTE']],
    ['s_fnFinance', ['PMT', 'FV', 'PV', 'NPV']],
  ];
  const formulasTab = (
    <>
      <Group label={t('functions')}>
        <Btn big icon={<Sigma size={22} />} label={t('sum')} onClick={() => insertFn('SUM')} />
        <Btn big icon={<Calculator size={22} />} label={t('average')} onClick={() => insertFn('AVERAGE')} />
        <Btn big icon={<Hash size={22} />} label={t('count')} onClick={() => insertFn('COUNT')} />
        <Btn big icon={<ArrowDownAZ size={22} />} label={t('min')} onClick={() => insertFn('MIN')} />
        <Btn big icon={<ArrowUpAZ size={22} />} label={t('max')} onClick={() => insertFn('MAX')} />
      </Group>
      <Group label={t('s_library')}>
        <div className="fn-chips">
          {fnGroups.map(([label, names]) => (
            <Menu key={label} trigger={(o, tg) => <button type="button" className={`r-chip ${o ? 'on' : ''}`} onMouseDown={(e) => e.preventDefault()} onClick={tg}>{t(label)} <ChevronDown size={12} /></button>}>
              {(close) => (
                <div className="fn-list">
                  {names.filter((n) => FUNCTION_NAMES.includes(n)).map((n) => <MenuItem key={n} label={n} onClick={() => { close(); insertFn(n); }} />)}
                </div>
              )}
            </Menu>
          ))}
          <Select label={t('functions')} value={'' as string} options={[{ value: '', label: `ƒx ${t('s_allFunctions')}` }, ...FUNCTION_NAMES.map((f) => ({ value: f, label: f }))]} onChange={(f) => f && insertFn(f)} width={150} />
        </div>
      </Group>
      <Group label={t('s_audit')}>
        <Btn big icon={<Braces size={22} />} label={t('s_showFormulas')} active={showFormulas} onClick={() => setShowFormulas((s) => !s)} />
      </Group>
    </>
  );

  const dataTab = (
    <>
      <Group label={t('s_sortFilter')}>
        <Btn big icon={<ArrowDownAZ size={22} />} label={t('sortAsc')} onClick={() => sortSel(1)} />
        <Btn big icon={<ArrowUpAZ size={22} />} label={t('sortDesc')} onClick={() => sortSel(-1)} />
        <Btn big icon={<Filter size={22} />} label={t('s_filter')} active={!!sheet.filter} onClick={toggleFilter} />
        <Btn big icon={<FilterX size={22} />} label={t('s_clearFilter')} disabled={!sheet.filter || !Object.keys(sheet.filter.show).length} onClick={() => updateSheet((sh) => (sh.filter ? { ...sh, filter: { ...sh.filter, show: {} } } : sh))} />
      </Group>
      <Group label={t('s_dataTools')}>
        <Btn big icon={<CopyX size={22} />} label={t('s_removeDups')} onClick={removeDuplicates} />
        <Btn big icon={<ListChecks size={22} />} label={t('s_dropdown')} onClick={editList} />
      </Group>
      <Group>
        <Btn big icon={<Pin size={22} />} label={t('freeze')} active={!!sheet.freezeTop} onClick={() => updateSheet((sh) => ({ ...sh, freezeTop: !sh.freezeTop }))} />
      </Group>
    </>
  );

  const reviewTab = (
    <Group label={t('s_notes')}>
      <Btn big icon={<StickyNote size={22} />} label={activeCell?.note ? t('s_editNote') : t('s_newNote')} onClick={editNote} />
      <Btn big icon={<StickyNoteX size={22} />} label={t('s_deleteNote')} onClick={deleteNotes} />
    </Group>
  );

  const viewTab = (
    <>
      <Group label={t('zoom')}>
        <Select label={t('zoom')} value={zoom} options={[0.5, 0.75, 0.9, 1, 1.25, 1.5, 2].map((z) => ({ value: z, label: `${Math.round(z * 100)}%` }))} onChange={setZoom} width={80} />
        <Btn icon={<ZoomIn size={18} />} label={t('zoom')} onClick={() => setZoom(1)} />
      </Group>
      <Group>
        <Btn big icon={<Grid3x3 size={22} />} label={t('gridlines')} active={gridlines} onClick={() => setGridlines((g) => !g)} />
        <Btn big icon={<Braces size={22} />} label={t('s_showFormulas')} active={showFormulas} onClick={() => setShowFormulas((s) => !s)} />
        <Btn big icon={<Pin size={22} />} label={t('freeze')} active={!!sheet.freezeTop} onClick={() => updateSheet((sh) => ({ ...sh, freezeTop: !sh.freezeTop }))} />
      </Group>
    </>
  );

  // ----- grid render -----
  const totalW = colLefts[sheet.cols];
  const totalH = rowTops[sheet.rows];
  const visibleRows: number[] = [];
  for (let r = firstRow; r <= lastRow; r++) if (!hiddenR.has(r)) visibleRows.push(r);
  const visibleCols: number[] = [];
  for (let c = 0; c < sheet.cols; c++) if (!hiddenC.has(c)) visibleCols.push(c);

  const covered = new Set<string>();
  const shownMerges = (sheet.merges ?? []).filter((m) => m.r2 >= firstRow && m.r1 <= lastRow);
  for (const m of shownMerges) for (let r = Math.max(m.r1, firstRow); r <= Math.min(m.r2, lastRow); r++) for (let c = m.c1; c <= m.c2; c++) covered.add(addr(r, c));

  const filt = sheet.filter;
  const filterHead = (r: number, c: number) => !!filt && r === filt.range.r1 && c >= filt.range.c1 && c <= filt.range.c2;

  /** A shared edge is drawn once: skip our left/top when the neighbour already draws its right/bottom. */
  const ownBorders = (r: number, c: number, st: CellStyle) => {
    const bd = bordersOf(st);
    if (!bd) return bd;
    const out = { ...bd };
    if (out.l && c > 0 && (bordersOf(sheet.cells[addr(r, c - 1)]?.st)?.r ?? 0) >= out.l) delete out.l;
    if (out.t && r > 0 && (bordersOf(sheet.cells[addr(r - 1, c)]?.st)?.b ?? 0) >= out.t) delete out.t;
    return out;
  };
  const renderCell = (r: number, c: number, cbox: CSSProperties, merged = false) => {
    const k = addr(r, c);
    const cell = sheet.cells[k];
    const cf = cfMap.get(k);
    const head = filterHead(r, c);
    const handlers = {
      onMouseDown: (e: React.MouseEvent) => onCellDown(e, r, c),
      onMouseEnter: () => onCellEnter(r, c),
      onDoubleClick: () => startEdit(),
    };
    const filterBtn = head && (
      <button
        type="button"
        className={`filter-btn ${filt!.show[c] ? 'on' : ''}`}
        aria-label={t('s_filter')}
        onMouseDown={(e) => {
          e.stopPropagation();
          e.preventDefault();
          const rc = e.currentTarget.getBoundingClientRect();
          setFilterPop(filterPop?.c === c ? null : { c, x: rc.left, y: rc.bottom + 4 });
        }}
      >
        {filt!.show[c] ? <Filter size={11} /> : <ChevronDown size={12} />}
      </button>
    );
    if (!cell && !cf && !head && !merged) return <div key={c} className="g-cell" style={cbox} {...handlers} />;
    const v = cell ? valueAt(r, c) : null;
    const st = cell?.st ?? {};
    const text = showFormulas && cell?.v.startsWith('=') ? cell.v : formatValue(v, st.fmt, st.dec);
    const isNum = typeof v === 'number' && !showFormulas;
    const height = Number(cbox.height);
    const va = st.va ?? (height > DEFAULT_ROW_H + 2 && !st.wrap ? 'bottom' : undefined);
    const bg = cf?.bg ?? st.bg ?? (merged ? '#fff' : undefined);
    const style: CSSProperties = {
      ...cbox,
      fontWeight: st.b ? 700 : undefined,
      fontStyle: st.i ? 'italic' : undefined,
      textDecoration: [st.u && 'underline', st.s && 'line-through'].filter(Boolean).join(' ') || undefined,
      color: typeof v === 'object' && v !== null ? '#dc2626' : cf?.color ?? st.color,
      backgroundColor: bg,
      backgroundImage: cf?.bar ? `linear-gradient(90deg, ${cf.bar.color}, ${cf.bar.color}55)` : undefined,
      backgroundSize: cf?.bar ? `${Math.max(2, cf.bar.pct * 100)}% 72%` : undefined,
      backgroundRepeat: cf?.bar ? 'no-repeat' : undefined,
      backgroundPosition: cf?.bar ? 'left center' : undefined,
      textAlign: st.align ?? (isNum ? 'right' : typeof v === 'boolean' || (typeof v === 'object' && v !== null) ? 'center' : 'start'),
      fontSize: st.fontSize ? `${st.fontSize}pt` : undefined,
      fontFamily: st.font ? familyCss(st.font) : undefined,
      whiteSpace: st.wrap ? 'normal' : 'nowrap',
      boxShadow: borderShadow(ownBorders(r, c, st)),
      zIndex: merged ? 2 : !st.wrap && !isNum && text && !bg ? 1 : undefined,
      ...(va && { display: 'flex', flexDirection: 'column', justifyContent: va === 'top' ? 'flex-start' : va === 'middle' ? 'center' : 'flex-end', lineHeight: 1.35, paddingBlock: 2 }),
      ...(head && { paddingInlineEnd: 22 }),
    };
    const overflow = !st.wrap && !isNum && !bg && !st.align && !merged && !head;
    return (
      <div key={c} className={`g-cell ${overflow ? 'overflow' : ''} ${cell?.note ? 'has-note' : ''}`} style={style} dir="auto" {...handlers}>
        {va ? <span className="cell-text">{text}</span> : text}
        {filterBtn}
      </div>
    );
  };

  const box = (r1: number, c1: number, r2: number, c2: number) => ({
    left: colLefts[c1],
    top: rowTops[r1],
    width: colLefts[c2 + 1] - colLefts[c1],
    height: rowTops[r2 + 1] - rowTops[r1],
  });
  const selBox = box(sel.r1, sel.c1, sel.r2, sel.c2);
  const activeBox = anchorMerge ? box(anchorMerge.r1, anchorMerge.c1, anchorMerge.r2, anchorMerge.c2) : box(anchor.r, anchor.c, anchor.r, anchor.c);
  const fillPreview = (() => {
    const f = fillRange(fillTo);
    if (!f) return null;
    const p = { ...sel };
    if (f.dir === 'down') p.r2 = Math.min(sheet.rows - 1, p.r2 + f.n);
    if (f.dir === 'up') p.r1 = Math.max(0, p.r1 - f.n);
    if (f.dir === 'right') p.c2 = Math.min(sheet.cols - 1, p.c2 + f.n);
    if (f.dir === 'left') p.c1 = Math.max(0, p.c1 - f.n);
    return box(p.r1, p.c1, p.r2, p.c2);
  })();

  const fmtAt = (r: number, c: number) => displayAt(sheet, valueAt, r, c);
  const noteCell = hoverNote && sheet.cells[addr(hoverNote.r, hoverNote.c)]?.note ? hoverNote : null;
  const at = (b: { left: number; top: number; width: number; height: number }) => ({ left: HEADER_W + b.left, top: HEADER_H + b.top, width: b.width, height: b.height });

  const ctxItems: ([ReactNode, string, () => void, string?] | null)[] = [
    [<Scissors size={15} />, t('s_cut'), () => menuCopy(true), 'Ctrl+X'],
    [<Copy size={15} />, t('copy'), () => menuCopy(false), 'Ctrl+C'],
    [<ClipboardPaste size={15} />, t('s_paste'), () => void menuPaste(), 'Ctrl+V'],
    null,
    [<Rows3 size={15} />, t('insertRow'), insertRows],
    [<Columns3 size={15} />, t('insertCol'), insertCols],
    [<Trash2 size={15} />, t('deleteRow'), deleteRows],
    [<Trash2 size={15} />, t('deleteCol'), deleteCols],
    [<Eraser size={15} />, t('s_clearContents'), clearContents, 'Del'],
    null,
    [<ArrowDownAZ size={15} />, t('sortAsc'), () => sortSel(1)],
    [<ArrowUpAZ size={15} />, t('sortDesc'), () => sortSel(-1)],
    [<Filter size={15} />, sheet.filter ? t('s_removeFilter') : t('s_filter'), toggleFilter],
    null,
    [<StickyNote size={15} />, activeCell?.note ? t('s_editNote') : t('s_newNote'), editNote],
    [<Merge size={15} />, isMerged ? t('s_unmerge') : t('s_mergeCenter'), () => doMerge(isMerged ? 'unmerge' : 'center')],
    [<EyeOff size={15} />, t('s_hideRows'), () => setHidden('r', true)],
    [<Eye size={15} />, t('s_unhideRows'), () => setHidden('r', false)],
  ];

  return (
    <div className="editor sheet-editor">
      <Ribbon
        tabs={[
          { id: 'home', label: t('tab_home'), content: homeTab },
          { id: 'insert', label: t('tab_insert'), content: insertTab },
          { id: 'formulas', label: t('tab_formulas'), content: formulasTab },
          { id: 'data', label: t('tab_data'), content: dataTab },
          { id: 'review', label: t('tab_review'), content: reviewTab },
          { id: 'view', label: t('tab_view'), content: viewTab },
        ]}
      />
      <div className="formula-bar" dir="ltr">
        <div className="name-box">{single ? addr(anchor.r, anchor.c) : rangeRef(sel)}</div>
        <span className="fx">ƒx</span>
        <input
          className="formula-input"
          value={barValue}
          aria-label="Formula"
          dir="auto"
          onFocus={() => !editing && setEditing({ r: anchor.r, c: anchor.c, value: activeCell?.v ?? '', fromBar: true })}
          onChange={(e) => setEditing({ r: editing?.r ?? anchor.r, c: editing?.c ?? anchor.c, value: e.target.value, fromBar: true })}
          onKeyDown={onEditKey}
          onBlur={() => editing?.fromBar && commitEdit()}
        />
      </div>
      <div className="sheet-body">
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
          onMouseLeave={() => hoverNote && setHoverNote(null)}
          onContextMenu={(e) => {
            if ((e.target as HTMLElement).closest('.g-cell, .row-h, .col-h')) {
              e.preventDefault();
              setCtx({ x: e.clientX, y: e.clientY });
            }
          }}
          onScroll={(e) => { setScroll({ top: e.currentTarget.scrollTop, height: e.currentTarget.clientHeight }); if (filterPop) setFilterPop(null); }}
        >
          <div className="grid-canvas" style={{ width: (totalW + HEADER_W) * zoom, height: (totalH + HEADER_H) * zoom }}>
            <div className="grid-zoom" style={{ transform: `scale(${zoom})`, width: totalW + HEADER_W, height: totalH + HEADER_H }}>
              {/* column headers */}
              <div className="col-headers" style={{ width: totalW + HEADER_W, height: HEADER_H }}>
                <div className="corner" style={{ width: HEADER_W, height: HEADER_H }} onMouseDown={() => { setAnchor({ r: 0, c: 0 }); setFocus({ r: sheet.rows - 1, c: sheet.cols - 1 }); }} />
                {visibleCols.map((c) => (
                  <div
                    key={c}
                    className={`col-h ${c >= sel.c1 && c <= sel.c2 ? 'hl' : ''} ${c > 0 && hiddenC.has(c - 1) ? 'after-hidden' : ''}`}
                    style={{ left: HEADER_W + colLefts[c], width: colW(c), height: HEADER_H }}
                    onMouseDown={(e) => { if (e.button === 0 || !inRange(sel, 0, c)) selectCol(c, e.shiftKey); }}
                  >
                    {colName(c)}
                    <span className="col-resize" onMouseDown={(e) => startResize(e, 'c', c)} onDoubleClick={(e) => { e.stopPropagation(); autofit(c >= sel.c1 && c <= sel.c2 ? range(sel.c1, sel.c2) : [c]); }} />
                  </div>
                ))}
              </div>
              {/* rows */}
              {visibleRows.map((r) => (
                <div key={r} className={`g-row ${sheet.freezeTop && r === 0 ? 'frozen' : ''}`} style={{ top: HEADER_H + rowTops[r], height: rowH(r), width: totalW + HEADER_W }}>
                  <div
                    className={`row-h ${r >= sel.r1 && r <= sel.r2 ? 'hl' : ''} ${r > 0 && hiddenR.has(r - 1) ? 'after-hidden' : ''} ${filt && r > filt.range.r1 && Object.keys(filt.show).length ? 'filtered' : ''}`}
                    style={{ width: HEADER_W, height: rowH(r) }}
                    onMouseDown={(e) => { if (e.button === 0 || !inRange(sel, r, 0)) selectRow(r, e.shiftKey); }}
                  >
                    {r + 1}
                    <span className="row-resize" onMouseDown={(e) => startResize(e, 'r', r)} />
                  </div>
                  {visibleCols.map((c) => (covered.has(addr(r, c)) ? null : renderCell(r, c, { left: HEADER_W + colLefts[c], width: colW(c), height: rowH(r) })))}
                </div>
              ))}
              {/* merged cells sit above the rows they span */}
              {shownMerges.map((m) => {
                const b = box(m.r1, m.c1, m.r2, m.c2);
                if (!b.width || !b.height) return null;
                return <div key={rangeRef(m)} className="merge-wrap">{renderCell(m.r1, m.c1, at(b), true)}</div>;
              })}
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
              {find && hits.map((h, i) => <div key={`h${i}`} className={`find-hit ${i === hitIdx ? 'on' : ''}`} style={at(box(h.r, h.c, h.r, h.c))} />)}
              {/* selection */}
              <div className="sel-box" style={at(selBox)} />
              <div className="sel-active" style={at(activeBox)} />
              {fillPreview && <div className="fill-preview" style={at(fillPreview)} />}
              {!editing && !chartSel && (
                <div
                  className="fill-handle"
                  title={t('s_fillHandle')}
                  style={{ left: HEADER_W + selBox.left + selBox.width - 4, top: HEADER_H + selBox.top + selBox.height - 4 }}
                  onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); filling.current = true; setFillTo(null); }}
                  onDoubleClick={fillAuto}
                />
              )}
              {activeCell?.list && !editing && (
                <button type="button" className="dv-btn" aria-label={t('s_dropdown')} style={{ left: HEADER_W + activeBox.left + activeBox.width + 2, top: HEADER_H + activeBox.top + activeBox.height - 20 }} onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); setListOpen((o) => !o); }}>
                  <ChevronDown size={13} />
                </button>
              )}
              {listOpen && activeCell?.list && (
                <div className="dv-pop" style={{ left: HEADER_W + activeBox.left, top: HEADER_H + activeBox.top + activeBox.height + 2, minWidth: activeBox.width }}>
                  {activeCell.list.map((item) => (
                    <button key={item} type="button" className={`dv-item ${item === activeCell.v ? 'on' : ''}`} dir="auto" onMouseDown={(e) => { e.preventDefault(); setCellValue(anchor.r, anchor.c, item); setListOpen(false); focusGrid(); }}>
                      {item}
                    </button>
                  ))}
                </div>
              )}
              {noteCell && (() => {
                const b = box(noteCell.r, noteCell.c, noteCell.r, noteCell.c);
                return <div className="note-pop" dir="auto" style={{ left: HEADER_W + b.left + b.width + 6, top: HEADER_H + b.top }}>{sheet.cells[addr(noteCell.r, noteCell.c)]!.note}</div>;
              })()}
              {/* charts float over the grid */}
              {sheet.charts?.map((ch) => {
                const on = ch.id === chartSel;
                return (
                  <div key={ch.id} className={`sheet-chart ${on ? 'on' : ''}`} style={{ left: HEADER_W + ch.x, top: HEADER_H + ch.y, width: ch.w, height: ch.h }} onMouseDown={(e) => startChartDrag(e, ch, 'move')} onDoubleClick={() => askChartTitle(ch)}>
                    <ChartSvg type={ch.type} data={chartData(ch.range, valueAt, fmtAt)} w={ch.w} h={ch.h} title={ch.title} legend={ch.legend !== false} />
                    {on && (
                      <>
                        <span className="chart-resize" onMouseDown={(e) => startChartDrag(e, ch, 'resize')} />
                        <div className={`chart-tools ${ch.y < 56 ? 'below' : ''}`} onMouseDown={(e) => e.stopPropagation()} onDoubleClick={(e) => e.stopPropagation()}>
                          {chartTypes.map((c) => (
                            <button key={c.type} type="button" className={`icon-btn ${ch.type === c.type ? 'on' : ''}`} title={t(`s_chart_${c.type}` as Key)} onClick={() => updateChart(ch.id, { type: c.type })}>
                              {c.icon}
                            </button>
                          ))}
                          <span className="sep" />
                          <button type="button" className="icon-btn" title={t('s_chartTitle')} onClick={() => askChartTitle(ch)}><Pencil size={15} /></button>
                          <button type="button" className={`icon-btn ${ch.legend !== false ? 'on' : ''}`} title={t('s_legend')} onClick={() => updateChart(ch.id, { legend: ch.legend === false })}><Tags size={15} /></button>
                          <button type="button" className="icon-btn danger" title={t('delete')} onClick={() => deleteChart(ch.id)}><Trash2 size={15} /></button>
                        </div>
                      </>
                    )}
                  </div>
                );
              })}
              {editing && !editing.fromBar && (
                <input
                  ref={editRef}
                  className="cell-editor"
                  dir="auto"
                  style={{
                    left: HEADER_W + colLefts[editing.c],
                    top: HEADER_H + rowTops[editing.r],
                    minWidth: mergeAt(sheet, editing.r, editing.c) ? activeBox.width : colW(editing.c),
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
        {find && (
          <div className="find-panel" role="search" onKeyDown={(e) => { if (e.key === 'Escape') { setFind(null); focusGrid(); } }}>
            <div className="find-row">
              <Search size={15} />
              <input autoFocus dir="auto" placeholder={t('s_findPlaceholder')} value={findQ} onChange={(e) => setFindQ(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); gotoHit(e.shiftKey ? -1 : 1); } }} />
              <span className="find-count">{findQ ? (hits.length ? `${Math.min(hitIdx + 1, hits.length)}/${hits.length}` : t('noResults')) : ''}</span>
              <button className="icon-btn" aria-label="Previous" onClick={() => gotoHit(-1)}><ChevronUp size={16} /></button>
              <button className="icon-btn" aria-label="Next" onClick={() => gotoHit(1)}><ChevronDown size={16} /></button>
              <button className="icon-btn" aria-label={t('close')} onClick={() => { setFind(null); focusGrid(); }}><X size={16} /></button>
            </div>
            {find.replace && (
              <div className="find-row">
                <Replace size={15} />
                <input dir="auto" placeholder={t('replaceWith')} value={replaceQ} onChange={(e) => setReplaceQ(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); replaceOne(); } }} />
                <button className="ghost-btn small" onClick={replaceOne}>{t('replaceOne')}</button>
                <button className="ghost-btn small" onClick={replaceAll}>{t('replaceAll')}</button>
              </div>
            )}
            <div className="find-opts">
              <label><input type="checkbox" checked={findOpts.matchCase} onChange={(e) => setFindOpts({ ...findOpts, matchCase: e.target.checked })} /> {t('matchCase')}</label>
              <label><input type="checkbox" checked={findOpts.whole} onChange={(e) => setFindOpts({ ...findOpts, whole: e.target.checked })} /> {t('s_wholeCell')}</label>
              <label><input type="checkbox" checked={findOpts.formulas} onChange={(e) => setFindOpts({ ...findOpts, formulas: e.target.checked })} /> {t('s_inFormulas')}</label>
              {!find.replace && <button className="link-btn" onClick={() => setFind({ replace: true })}>{t('replace')}</button>}
            </div>
          </div>
        )}
      </div>
      {filterPop && filt && (
        <FilterPopup
          key={filterPop.c}
          x={filterPop.x}
          y={filterPop.y}
          values={filterValues(sheet, valueAt, filterPop.c)}
          shown={filt.show[filterPop.c]}
          onSort={(dir) => { sortRange({ ...filt.range, r1: filt.range.r1 + 1, r2: filterEnd(sheet) }, filterPop.c, dir); setFilterPop(null); }}
          onApply={(show) => { setFilterShow(filterPop.c, show); setFilterPop(null); focusGrid(); }}
          onClose={() => { setFilterPop(null); focusGrid(); }}
        />
      )}
      {ctx && (
        <div className="menu-pop ctx-menu" style={{ position: 'fixed', left: Math.max(8, Math.min(ctx.x, window.innerWidth - 270)), top: Math.max(8, Math.min(ctx.y, window.innerHeight - 560)) }} onContextMenu={(e) => e.preventDefault()}>
          {ctxItems.map((it, i) =>
            it ? <MenuItem key={i} icon={it[0]} label={it[1]} hint={it[3]} onClick={() => { setCtx(null); it[2](); focusGrid(); }} /> : <div key={i} className="menu-sep" />,
          )}
        </div>
      )}
      {ask && <AskDialog ask={ask} onClose={() => { setAsk(null); focusGrid(); }} />}
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
          {flash && <span className="flash">{flash}</span>}
          {filt && filtered.size > 0 && <span>{t('s_filteredRows').replace('{n}', String(filtered.size))}</span>}
          {stats && (
            <>
              {stats.n > 0 && <span>{t('average')}: {formatValue(stats.avg, undefined)}</span>}
              <span>{t('count')}: {stats.cnt}</span>
              {stats.n > 1 && <span className="wide-only">{t('min')}: {formatValue(stats.mn, undefined)}</span>}
              {stats.n > 1 && <span className="wide-only">{t('max')}: {formatValue(stats.mx, undefined)}</span>}
              {stats.n > 0 && <span>{t('sum')}: {formatValue(stats.sum, undefined)}</span>}
            </>
          )}
          <span className="muted"><ArrowLeftRight size={12} /> {Math.round(zoom * 100)}%</span>
        </div>
      </div>
    </div>
  );
}

// ----- popups -----

function FilterPopup({ x, y, values, shown, onSort, onApply, onClose }: {
  x: number; y: number; values: string[]; shown?: string[];
  onSort: (dir: 1 | -1) => void; onApply: (show: string[] | null) => void; onClose: () => void;
}) {
  const { t } = useI18n();
  const [q, setQ] = useState('');
  const [on, setOn] = useState(() => new Set(shown ?? values));
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const h = (e: MouseEvent) => { if (!ref.current?.contains(e.target as Node)) onClose(); };
    const k = (e: KeyboardEvent) => e.key === 'Escape' && onClose();
    document.addEventListener('mousedown', h);
    document.addEventListener('keydown', k);
    return () => { document.removeEventListener('mousedown', h); document.removeEventListener('keydown', k); };
  }, [onClose]);
  const label = (v: string) => (v === BLANK ? t('s_blanks') : v);
  const list = values.filter((v) => label(v).toLowerCase().includes(q.toLowerCase()));
  const all = list.length > 0 && list.every((v) => on.has(v));
  return (
    <div ref={ref} className="filter-pop" style={{ left: Math.max(8, Math.min(x, window.innerWidth - 268)), top: Math.max(8, Math.min(y, window.innerHeight - 400)) }}>
      <button type="button" className="menu-item" onClick={() => onSort(1)}><span className="menu-icon"><ArrowDownAZ size={15} /></span><span className="menu-label">{t('sortAsc')}</span></button>
      <button type="button" className="menu-item" onClick={() => onSort(-1)}><span className="menu-icon"><ArrowUpAZ size={15} /></span><span className="menu-label">{t('sortDesc')}</span></button>
      <div className="menu-sep" />
      <div className="find-row"><Search size={14} /><input autoFocus dir="auto" placeholder={t('s_search')} value={q} onChange={(e) => setQ(e.target.value)} /></div>
      <div className="filter-list">
        <label className="filter-item all">
          <input type="checkbox" checked={all} onChange={() => setOn((s) => { const n = new Set(s); list.forEach((v) => (all ? n.delete(v) : n.add(v))); return n; })} />
          {t('s_selectAll')}
        </label>
        {list.map((v) => (
          <label key={v} className="filter-item" dir="auto">
            <input type="checkbox" checked={on.has(v)} onChange={() => setOn((s) => { const n = new Set(s); if (n.has(v)) n.delete(v); else n.add(v); return n; })} />
            {label(v)}
          </label>
        ))}
      </div>
      <div className="filter-actions">
        <button type="button" className="link-btn" onClick={() => onApply(null)}>{t('s_clearFilter')}</button>
        <span className="grow" />
        <button type="button" className="ghost-btn small" onClick={onClose}>{t('cancel')}</button>
        <button type="button" className="primary-btn small" disabled={on.size === 0} onClick={() => onApply(values.every((v) => on.has(v)) ? null : [...on])}>{t('ok')}</button>
      </div>
    </div>
  );
}

function AskDialog({ ask, onClose }: { ask: Ask; onClose: () => void }) {
  const { t } = useI18n();
  const [vals, setVals] = useState(() => Object.fromEntries(ask.fields.map((f) => [f.key, f.value])));
  const ok = () => { ask.onOk(vals); onClose(); };
  return (
    <div className="modal-back" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <form className="modal ask-modal" role="dialog" aria-label={ask.title} onSubmit={(e) => { e.preventDefault(); ok(); }} onKeyDown={(e) => e.key === 'Escape' && onClose()}>
        <header>
          <h3>{ask.title}</h3>
          <button type="button" className="icon-btn" aria-label={t('close')} onClick={onClose}><X size={18} /></button>
        </header>
        {ask.fields.map((f, i) => (
          <label key={f.key} className="ask-field">
            <span>{f.label}</span>
            {f.type === 'select' ? (
              <select value={vals[f.key]} onChange={(e) => setVals({ ...vals, [f.key]: e.target.value })}>
                {f.options!.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
              </select>
            ) : f.type === 'textarea' ? (
              <textarea autoFocus={i === 0} dir="auto" rows={4} value={vals[f.key]} onChange={(e) => setVals({ ...vals, [f.key]: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) ok(); }} />
            ) : (
              <input autoFocus={i === 0} dir="auto" type={f.type === 'number' ? 'number' : 'text'} value={vals[f.key]} onChange={(e) => setVals({ ...vals, [f.key]: e.target.value })} />
            )}
          </label>
        ))}
        <div className="modal-actions">
          <button type="button" className="ghost-btn" onClick={onClose}>{t('cancel')}</button>
          <button type="submit" className="primary-btn">{t('ok')}</button>
        </div>
      </form>
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
const range = (a: number, b: number) => Array.from({ length: Math.max(0, b - a + 1) }, (_, i) => a + i);
const clamp = (v: number, a: number, b: number) => Math.max(a, Math.min(b, v));
const defaultDec = (f?: NumFmt) => (f === 'percent' ? 0 : f === 'number' || f === 'currency' || f === 'scientific' ? 2 : 0);
function mixHex(a: string, b: string, t: number) {
  const pa = parseInt(a.slice(1), 16), pb = parseInt(b.slice(1), 16);
  const m = (s: number) => Math.round(((pa >> s) & 255) + (((pb >> s) & 255) - ((pa >> s) & 255)) * t);
  return `#${((m(16) << 16) | (m(8) << 8) | m(0)).toString(16).padStart(6, '0')}`;
}
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
