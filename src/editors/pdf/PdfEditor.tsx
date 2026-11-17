import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import {
  ArrowRight, Circle, Copy, Eraser, FilePlus2, FileStack, Hand, Highlighter, ImagePlus, Images, Italic, Bold,
  MessageSquare, Minus, MousePointer2, PenLine, PenTool, Redo2, RotateCcw, RotateCw, Scissors, Search, Signature, Square,
  Stamp, TextCursorInput, Trash2, Type, Undo2, ZoomIn, ZoomOut, Hash, Droplets, PanelTop, PanelBottom,
  ClipboardList, Info, ChevronUp, ChevronDown, ScanLine, Maximize2, FileOutput, X, Highlighter as Marker, EyeOff,
} from 'lucide-react';
import { useI18n } from '../../i18n';
import { Btn, ColorBtn, Group, Menu, MenuItem, Ribbon, Select } from '../../ui/Ribbon';
import { useHistory } from '../useHistory';
import {
  annotBounds, blankPage, FONT_CSS, isBoxLike, moveAnnot, pageLabel, pid, rotTransform, shownSize,
  type Annot, type BoxLike, type PdfDoc, type PdfPage, type Rot, type TextAnnot,
} from './model';
import { importPdf, pageText, sourcePage, type TextItem } from './pdfjs';
import { arrowHead, decorPos, inkPath } from './render';
import { SignaturePad } from './SignaturePad';
import { pagesToImages, pdfBlob, readFormFields, splitPdf, type FormField } from './save';
import { saveBlob } from '../../io/files';
import { imagesToPages } from './images';

type Tool =
  | 'select' | 'hand' | 'text' | 'edittext' | 'ink' | 'marker' | 'highlight' | 'rect' | 'ellipse' | 'line' | 'arrow'
  | 'note' | 'whiteout' | 'redact' | 'place';

interface Props {
  data: PdfDoc;
  name: string;
  onChange: (d: PdfDoc) => void;
}

interface ToolStyle {
  color: string;
  fill: string | null;
  width: number;
  size: number;
  font: TextAnnot['font'];
  bold: boolean;
  italic: boolean;
}

const PT2PX = 96 / 72;
const STAMPS = ['APPROVED', 'DRAFT', 'CONFIDENTIAL', 'REJECTED', 'PAID', 'COPY', 'معتمد', 'سري'];

export function PdfEditor({ data, name, onChange }: Props) {
  const { t } = useI18n();
  const { commit, undo, redo } = useHistory(data, onChange);
  const live = useRef(data);
  live.current = data;

  const [tool, setTool] = useState<Tool>('select');
  const [style, setStyle] = useState<ToolStyle>({ color: '#e11d48', fill: null, width: 2, size: 14, font: 'sans', bold: false, italic: false });
  const [zoom, setZoom] = useState(1);
  const [sel, setSel] = useState<{ page: string; id: string } | null>(null);
  const [editing, setEditing] = useState<string | null>(null);
  const [pending, setPending] = useState<{ kind: 'image' | 'stamp'; src?: string; label?: string; color?: string } | null>(null);
  const [selPages, setSelPages] = useState<Set<string>>(new Set());
  const [current, setCurrent] = useState(0);
  const [panel, setPanel] = useState<'none' | 'search' | 'form' | 'comments' | 'props'>('none');
  const [showSig, setShowSig] = useState(false);
  const [busy, setBusy] = useState(false);
  const scrollRef = useRef<HTMLDivElement>(null);
  const imgInput = useRef<HTMLInputElement>(null);
  const insertInput = useRef<HTMLInputElement>(null);
  const [search, setSearch] = useState<{ q: string; hits: { page: number; x: number; y: number; w: number; h: number }[]; at: number }>({ q: '', hits: [], at: 0 });
  const [fields, setFields] = useState<FormField[] | null>(null);

  const scale = zoom * PT2PX;

  // ----- mutations -----
  const setPages = (pages: PdfPage[], record = true) => (record ? commit({ ...live.current, pages }) : onChange({ ...live.current, pages }));
  const updatePage = (id: string, fn: (p: PdfPage) => PdfPage, record = true) => setPages(live.current.pages.map((p) => (p.id === id ? fn(p) : p)), record);
  const addAnnot = (pageId: string, a: Annot) => {
    updatePage(pageId, (p) => ({ ...p, annots: [...p.annots, a] }));
    setSel({ page: pageId, id: a.id });
  };
  const patchAnnot = (pageId: string, id: string, patch: Partial<Annot>, record = true) =>
    updatePage(pageId, (p) => ({ ...p, annots: p.annots.map((a) => (a.id === id ? ({ ...a, ...patch } as Annot) : a)) }), record);
  const deleteAnnot = (pageId: string, id: string) => {
    updatePage(pageId, (p) => ({ ...p, annots: p.annots.filter((a) => a.id !== id) }));
    setSel(null);
    setEditing(null);
  };
  /** Wrap a live gesture into one undo step. */
  const gesture = () => {
    const before = live.current;
    return () => {
      const after = live.current;
      if (after === before) return;
      onChange(before);
      commit(after);
    };
  };

  const selected = useMemo(() => {
    if (!sel) return null;
    return data.pages.find((p) => p.id === sel.page)?.annots.find((a) => a.id === sel.id) ?? null;
  }, [sel, data]);

  // Style controls edit the selected item, or the style for the next one.
  const applyStyle = (patch: Partial<ToolStyle>) => {
    setStyle((s) => ({ ...s, ...patch }));
    if (!sel || !selected) return;
    const a = selected;
    const p: Partial<Annot> & Record<string, unknown> = {};
    if (patch.color !== undefined) {
      if (a.kind === 'text' || a.kind === 'line' || a.kind === 'arrow' || a.kind === 'ink' || a.kind === 'note' || a.kind === 'stamp') p.color = patch.color;
      if (a.kind === 'rect' || a.kind === 'ellipse') p.stroke = patch.color;
      if (a.kind === 'highlight') p.fill = patch.color;
    }
    if (patch.fill !== undefined && (a.kind === 'rect' || a.kind === 'ellipse')) p.fill = patch.fill ?? 'none';
    if (patch.fill !== undefined && a.kind === 'text') p.bg = patch.fill ?? undefined;
    if (patch.width !== undefined && 'width' in a) p.width = patch.width;
    if (patch.size !== undefined && a.kind === 'text') p.size = patch.size;
    if (patch.font !== undefined && a.kind === 'text') p.font = patch.font;
    if (patch.bold !== undefined && a.kind === 'text') p.bold = patch.bold;
    if (patch.italic !== undefined && a.kind === 'text') p.italic = patch.italic;
    if (Object.keys(p).length) patchAnnot(sel.page, sel.id, p);
  };

  // ----- page operations -----
  const targetPages = () => (selPages.size ? data.pages.filter((p) => selPages.has(p.id)) : [data.pages[current]].filter(Boolean));
  const rotate = (dir: 1 | -1) => {
    const ids = new Set(targetPages().map((p) => p.id));
    setPages(data.pages.map((p) => (ids.has(p.id) ? { ...p, rot: (((p.rot + dir * 90) % 360) + 360) % 360 as Rot } : p)));
  };
  const deletePages = () => {
    const ids = new Set(targetPages().map((p) => p.id));
    if (ids.size >= data.pages.length) return;
    setPages(data.pages.filter((p) => !ids.has(p.id)));
    setSelPages(new Set());
  };
  const duplicatePages = () => {
    const ids = new Set(targetPages().map((p) => p.id));
    const out: PdfPage[] = [];
    for (const p of data.pages) {
      out.push(p);
      if (ids.has(p.id)) out.push({ ...p, id: pid(), annots: p.annots.map((a) => ({ ...a, id: pid() })) });
    }
    setPages(out);
  };
  const insertBlank = () => {
    const ref = data.pages[current];
    const pages = data.pages.slice();
    pages.splice(current + 1, 0, blankPage(ref?.w, ref?.h));
    setPages(pages);
  };
  const movePage = (from: number, to: number) => {
    if (to < 0 || to >= data.pages.length || from === to) return;
    const pages = data.pages.slice();
    const [p] = pages.splice(from, 1);
    pages.splice(to, 0, p);
    setPages(pages);
  };
  const insertFiles = async (files: FileList) => {
    setBusy(true);
    try {
      let sources = { ...live.current.sources };
      const added: PdfPage[] = [];
      const images: File[] = [];
      for (const f of Array.from(files)) {
        if (/\.pdf$/i.test(f.name) || f.type === 'application/pdf') {
          const imp = await importPdf(new Uint8Array(await f.arrayBuffer()));
          sources = { ...sources, [imp.id]: imp.b64 };
          added.push(...imp.pages);
        } else if (f.type.startsWith('image/')) images.push(f);
      }
      if (images.length) added.push(...(await imagesToPages(images)));
      const pages = live.current.pages.slice();
      pages.splice(current + 1, 0, ...added);
      commit({ ...live.current, sources, pages });
    } catch (e) {
      alert(String((e as Error)?.message ?? e));
    } finally {
      setBusy(false);
    }
  };
  const extract = async () => {
    setBusy(true);
    try {
      await saveBlob(await pdfBlob(data, targetPages()), `${name}-pages.pdf`);
    } finally {
      setBusy(false);
    }
  };
  const run = async (fn: () => Promise<Blob>, filename: string) => {
    setBusy(true);
    try {
      await saveBlob(await fn(), filename);
    } finally {
      setBusy(false);
    }
  };

  // ----- scrolling / current page -----
  const scrollToPage = (i: number) => {
    const el = scrollRef.current?.querySelector<HTMLElement>(`[data-page-index="${i}"]`);
    el?.scrollIntoView({ block: 'start', behavior: 'smooth' });
  };
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const onScroll = () => {
      const mid = el.scrollTop + el.clientHeight / 3;
      const pages = Array.from(el.querySelectorAll<HTMLElement>('[data-page-index]'));
      let idx = 0;
      for (const p of pages) if (p.offsetTop <= mid) idx = Number(p.dataset.pageIndex);
      setCurrent(idx);
    };
    el.addEventListener('scroll', onScroll, { passive: true });
    return () => el.removeEventListener('scroll', onScroll);
  }, []);

  const fit = (mode: 'width' | 'page') => {
    const el = scrollRef.current;
    const p = data.pages[current] ?? data.pages[0];
    if (!el || !p) return;
    const s = shownSize(p);
    const zw = (el.clientWidth - 64) / (s.w * PT2PX);
    const zh = (el.clientHeight - 48) / (s.h * PT2PX);
    setZoom(Math.max(0.25, Math.min(4, mode === 'width' ? zw : Math.min(zw, zh))));
  };

  // ----- keyboard -----
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const tgt = e.target as HTMLElement;
      if (tgt.closest('input, textarea, select, [contenteditable="true"]')) return;
      if (!tgt.closest('.pdf-editor') && tgt !== document.body) return;
      const mod = e.ctrlKey || e.metaKey;
      if (mod && e.key.toLowerCase() === 'z') { e.preventDefault(); if (e.shiftKey) redo(); else undo(); return; }
      if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
      if (mod && e.key.toLowerCase() === 'f') { e.preventDefault(); setPanel('search'); return; }
      if (mod && (e.key === '=' || e.key === '+')) { e.preventDefault(); setZoom((z) => Math.min(4, z * 1.2)); return; }
      if (mod && e.key === '-') { e.preventDefault(); setZoom((z) => Math.max(0.25, z / 1.2)); return; }
      if (e.key === 'Escape') { setSel(null); setEditing(null); setPending(null); setTool('select'); return; }
      if (sel && selected && (e.key === 'Delete' || e.key === 'Backspace')) { e.preventDefault(); deleteAnnot(sel.page, sel.id); return; }
      if (sel && selected) {
        const step = e.shiftKey ? 10 : 1;
        const d = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key];
        if (d) { e.preventDefault(); updatePage(sel.page, (p) => ({ ...p, annots: p.annots.map((a) => (a.id === sel.id ? moveAnnot(a, d[0], d[1]) : a)) })); }
        return;
      }
      if (!mod && !e.altKey) {
        const map: Record<string, Tool> = { v: 'select', h: 'hand', t: 'text', p: 'ink', r: 'rect', o: 'ellipse', l: 'line', a: 'arrow', n: 'note', m: 'marker' };
        if (map[e.key]) setTool(map[e.key]);
        if (e.key === 'PageDown') { e.preventDefault(); scrollToPage(Math.min(data.pages.length - 1, current + 1)); }
        if (e.key === 'PageUp') { e.preventDefault(); scrollToPage(Math.max(0, current - 1)); }
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  // ----- search -----
  const runSearch = async (q: string) => {
    const needle = q.trim().toLowerCase();
    if (!needle) { setSearch({ q, hits: [], at: 0 }); return; }
    const hits: typeof search.hits = [];
    for (let i = 0; i < data.pages.length; i++) {
      const items = await pageText(data, data.pages[i]);
      for (const it of items) {
        const s = it.str.toLowerCase();
        let from = 0, k: number;
        while ((k = s.indexOf(needle, from)) >= 0) {
          // Approximate the match by character share, along the text's own direction.
          const n = Math.max(1, it.str.length);
          const a = k / n, b = (k + needle.length) / n;
          if (Math.abs(it.ux) >= Math.abs(it.uy)) {
            const [f0, f1] = it.ux >= 0 ? [a, b] : [1 - b, 1 - a];
            hits.push({ page: i, x: it.x + it.w * f0, y: it.y, w: it.w * (f1 - f0), h: it.h });
          } else {
            const [f0, f1] = it.uy >= 0 ? [a, b] : [1 - b, 1 - a];
            hits.push({ page: i, x: it.x, y: it.y + it.h * f0, w: it.w, h: it.h * (f1 - f0) });
          }
          from = k + needle.length;
        }
      }
    }
    setSearch({ q, hits, at: 0 });
    if (hits.length) revealHit();
  };
  /** Bring the current search hit to the middle of the view (after it renders). */
  const revealHit = () =>
    requestAnimationFrame(() => requestAnimationFrame(() => scrollRef.current?.querySelector('.hit.on')?.scrollIntoView({ block: 'center', inline: 'center', behavior: 'smooth' })));
  const stepSearch = (d: 1 | -1) => {
    if (!search.hits.length) return;
    const at = (search.at + d + search.hits.length) % search.hits.length;
    setSearch({ ...search, at });
    revealHit();
  };

  // ----- forms -----
  useEffect(() => {
    if (panel !== 'form') return;
    let alive = true;
    readFormFields(live.current).then((f) => alive && setFields(f));
    return () => { alive = false; };
  }, [panel]);
  const setField = (f: FormField, v: string | boolean) => {
    setFields((fs) => fs?.map((x) => (x.src === f.src && x.name === f.name ? { ...x, value: v } : x)) ?? null);
    const form = { ...live.current.form, [f.src]: { ...live.current.form[f.src], [f.name]: v } };
    onChange({ ...live.current, form });
  };

  const pickImage = (f: File) => {
    const r = new FileReader();
    r.onload = () => { setPending({ kind: 'image', src: String(r.result) }); setTool('place'); };
    r.readAsDataURL(f);
  };

  const s = selected;
  const sw = (
    <Select label={t('lineWidth')} value={s && 'width' in s ? (s as { width?: number }).width ?? style.width : style.width} options={[1, 2, 3, 4, 6, 8, 12].map((v) => ({ value: v, label: `${v} pt` }))} onChange={(v) => applyStyle({ width: v })} width={70} />
  );
  const textSel = s?.kind === 'text' ? s : null;

  const toolBtn = (id: Tool, icon: ReactNode, label: string, big = false) => (
    <Btn big={big} icon={icon} label={label} active={tool === id} onClick={() => { setTool(id); setPending(null); if (id !== 'select') setSel(null); }} />
  );

  const editTab = (
    <>
      <Group>
        <Btn icon={<Undo2 size={18} />} label={t('undo')} onClick={undo} />
        <Btn icon={<Redo2 size={18} />} label={t('redo')} onClick={redo} />
      </Group>
      <Group label={t('tools')}>
        {toolBtn('select', <MousePointer2 size={22} />, t('select'), true)}
        {toolBtn('hand', <Hand size={22} />, t('pan'), true)}
        {toolBtn('edittext', <TextCursorInput size={22} />, t('editText'), true)}
        {toolBtn('text', <Type size={22} />, t('addText'), true)}
      </Group>
      <Group label={t('annotate')}>
        <div className="r-col">
          <div className="r-row">
            {toolBtn('marker', <Marker size={18} />, t('highlight'))}
            {toolBtn('highlight', <Highlighter size={18} />, t('highlightArea'))}
            {toolBtn('ink', <PenLine size={18} />, t('draw'))}
            {toolBtn('note', <MessageSquare size={18} />, t('comment'))}
          </div>
          <div className="r-row">
            {toolBtn('rect', <Square size={18} />, t('rectangle'))}
            {toolBtn('ellipse', <Circle size={18} />, t('ellipse'))}
            {toolBtn('line', <Minus size={18} />, t('line'))}
            {toolBtn('arrow', <ArrowRight size={18} />, t('arrow'))}
          </div>
        </div>
      </Group>
      <Group label={t('insert')}>
        <Btn big icon={<Signature size={22} />} label={t('signature')} onClick={() => setShowSig(true)} />
        <Btn big icon={<ImagePlus size={22} />} label={t('image')} onClick={() => imgInput.current?.click()} />
        <Menu trigger={(_o, toggle) => <Btn big icon={<Stamp size={22} />} label={t('stamp')} onClick={toggle} />}>
          {(close) => (
            <>
              {STAMPS.map((l) => (
                <MenuItem key={l} icon={<Stamp size={14} />} label={l} onClick={() => { setPending({ kind: 'stamp', label: l, color: /REJECTED|CONFIDENTIAL|سري/.test(l) ? '#dc2626' : /DRAFT|COPY/.test(l) ? '#2563eb' : '#16a34a' }); setTool('place'); close(); }} />
              ))}
              <MenuItem icon={<Stamp size={14} />} label={new Date().toLocaleDateString()} onClick={() => { setPending({ kind: 'stamp', label: new Date().toLocaleDateString(), color: '#475569' }); setTool('place'); close(); }} />
            </>
          )}
        </Menu>
      </Group>
      <Group label={t('protect')}>
        {toolBtn('whiteout', <Eraser size={22} />, t('whiteout'), true)}
        {toolBtn('redact', <EyeOff size={22} />, t('redact'), true)}
      </Group>
      <Group label={t('style')}>
        <div className="r-col">
          <div className="r-row">
            <ColorBtn icon={<PenTool size={16} />} label={t('color')} color={style.color} onPick={(c) => c && applyStyle({ color: c })} />
            <ColorBtn icon={<Droplets size={16} />} label={t('fillColor')} color={style.fill ?? undefined} allowNone onPick={(c) => applyStyle({ fill: c })} />
            {sw}
          </div>
          <div className="r-row">
            <Select label={t('font')} value={textSel?.font ?? style.font} options={[{ value: 'sans', label: 'Sans' }, { value: 'serif', label: 'Serif' }, { value: 'mono', label: 'Mono' }] as { value: TextAnnot['font']; label: string }[]} onChange={(v) => applyStyle({ font: v })} width={78} />
            <Select label={t('size')} value={textSel?.size ?? style.size} options={[8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72].map((v) => ({ value: v, label: String(v) }))} onChange={(v) => applyStyle({ size: v })} width={58} />
            <Btn icon={<Bold size={16} />} label={t('bold')} active={textSel?.bold ?? style.bold} onClick={() => applyStyle({ bold: !(textSel?.bold ?? style.bold) })} />
            <Btn icon={<Italic size={16} />} label={t('italic')} active={textSel?.italic ?? style.italic} onClick={() => applyStyle({ italic: !(textSel?.italic ?? style.italic) })} />
          </div>
        </div>
      </Group>
      {sel && <Group><Btn big icon={<Trash2 size={22} />} label={t('delete')} onClick={() => deleteAnnot(sel.page, sel.id)} /></Group>}
    </>
  );

  const pagesTab = (
    <>
      <Group label={t('pages')}>
        <Btn big icon={<FilePlus2 size={22} />} label={t('blankPage')} onClick={insertBlank} />
        <Btn big icon={<FileStack size={22} />} label={t('insertFile')} onClick={() => insertInput.current?.click()} />
        <Btn big icon={<Copy size={22} />} label={t('duplicate')} onClick={duplicatePages} />
        <Btn big icon={<Trash2 size={22} />} label={t('delete')} disabled={data.pages.length <= 1} onClick={deletePages} />
      </Group>
      <Group label={t('rotate')}>
        <Btn big icon={<RotateCcw size={22} />} label={t('rotateLeft')} onClick={() => rotate(-1)} />
        <Btn big icon={<RotateCw size={22} />} label={t('rotateRight')} onClick={() => rotate(1)} />
      </Group>
      <Group label={t('order')}>
        <Btn icon={<ChevronUp size={18} />} label={t('moveUp')} onClick={() => { movePage(current, current - 1); setCurrent(current - 1); }} />
        <Btn icon={<ChevronDown size={18} />} label={t('moveDown')} onClick={() => { movePage(current, current + 1); setCurrent(current + 1); }} />
      </Group>
      <Group label={t('exportGroup')}>
        <Btn big icon={<FileOutput size={22} />} label={t('extractPages')} onClick={extract} />
        <Btn big icon={<Scissors size={22} />} label={t('splitPages')} onClick={() => run(() => splitPdf(data), `${name}-split.zip`)} />
        <Btn big icon={<Images size={22} />} label={t('toImages')} onClick={() => run(() => pagesToImages(data), `${name}-images.zip`)} />
      </Group>
    </>
  );

  const decor = data.decor;
  const setDecor = (patch: Partial<PdfDoc['decor']>) => commit({ ...data, decor: { ...data.decor, ...patch } });
  const toolsTab = (
    <>
      <Group label={t('pageNumbers')}>
        <Btn big icon={<Hash size={22} />} label={t('pageNumbers')} active={!!decor.pageNumbers} onClick={() => setDecor({ pageNumbers: decor.pageNumbers ? undefined : { pos: 'bottom-center', fmt: 'n', size: 10 } })} />
        {decor.pageNumbers && (
          <div className="r-col">
            <Select label={t('position')} value={decor.pageNumbers.pos} options={[{ value: 'bottom-center', label: '⎵ center' }, { value: 'bottom-right', label: '⎵ right' }, { value: 'bottom-left', label: '⎵ left' }, { value: 'top-right', label: '⎴ right' }] as { value: NonNullable<PdfDoc['decor']['pageNumbers']>['pos']; label: string }[]} onChange={(v) => setDecor({ pageNumbers: { ...decor.pageNumbers!, pos: v } })} width={110} />
            <Select label={t('format')} value={decor.pageNumbers.fmt} options={[{ value: 'n', label: '1' }, { value: 'n-of-N', label: '1 / 9' }, { value: 'page-n', label: 'Page 1' }] as { value: 'n' | 'n-of-N' | 'page-n'; label: string }[]} onChange={(v) => setDecor({ pageNumbers: { ...decor.pageNumbers!, fmt: v } })} width={110} />
          </div>
        )}
      </Group>
      <Group label={t('watermark')}>
        <Btn big icon={<Droplets size={22} />} label={t('watermark')} active={!!decor.watermark} onClick={() => setDecor({ watermark: decor.watermark ? undefined : { text: 'CONFIDENTIAL', color: '#94a3b8', opacity: 0.25, size: 64 } })} />
        {decor.watermark && (
          <div className="r-col">
            <input className="r-input" dir="auto" value={decor.watermark.text} aria-label={t('watermark')} onChange={(e) => onChange({ ...data, decor: { ...decor, watermark: { ...decor.watermark!, text: e.target.value } } })} />
            <div className="r-row">
              <ColorBtn icon={<Droplets size={14} />} label={t('color')} color={decor.watermark.color} onPick={(c) => c && setDecor({ watermark: { ...decor.watermark!, color: c } })} />
              <Select label={t('opacity')} value={decor.watermark.opacity} options={[0.1, 0.2, 0.3, 0.5, 0.8].map((v) => ({ value: v, label: `${v * 100}%` }))} onChange={(v) => setDecor({ watermark: { ...decor.watermark!, opacity: v } })} width={72} />
            </div>
          </div>
        )}
      </Group>
      <Group label={t('headerFooter')}>
        <div className="r-col">
          <label className="r-field"><PanelTop size={14} /><input className="r-input" dir="auto" placeholder={t('header')} value={decor.header ?? ''} onChange={(e) => onChange({ ...data, decor: { ...decor, header: e.target.value || undefined } })} /></label>
          <label className="r-field"><PanelBottom size={14} /><input className="r-input" dir="auto" placeholder={t('footer')} value={decor.footer ?? ''} onChange={(e) => onChange({ ...data, decor: { ...decor, footer: e.target.value || undefined } })} /></label>
        </div>
      </Group>
      <Group label={t('panels')}>
        <Btn big icon={<Search size={22} />} label={t('find')} active={panel === 'search'} onClick={() => setPanel(panel === 'search' ? 'none' : 'search')} />
        <Btn big icon={<ClipboardList size={22} />} label={t('fillForm')} active={panel === 'form'} onClick={() => setPanel(panel === 'form' ? 'none' : 'form')} />
        <Btn big icon={<MessageSquare size={22} />} label={t('comments')} active={panel === 'comments'} onClick={() => setPanel(panel === 'comments' ? 'none' : 'comments')} />
        <Btn big icon={<Info size={22} />} label={t('properties')} active={panel === 'props'} onClick={() => setPanel(panel === 'props' ? 'none' : 'props')} />
      </Group>
    </>
  );

  const viewTab = (
    <>
      <Group label={t('zoom')}>
        <Btn big icon={<ZoomOut size={22} />} label={t('zoomOut')} onClick={() => setZoom((z) => Math.max(0.25, z / 1.2))} />
        <Select label={t('zoom')} value={Math.round(zoom * 100) / 100} options={[...new Set([0.5, 0.75, 1, 1.25, 1.5, 2, 3, Math.round(zoom * 100) / 100])].sort((a, b) => a - b).map((z) => ({ value: z, label: `${Math.round(z * 100)}%` }))} onChange={setZoom} width={84} />
        <Btn big icon={<ZoomIn size={22} />} label={t('zoomIn')} onClick={() => setZoom((z) => Math.min(4, z * 1.2))} />
      </Group>
      <Group label={t('fit')}>
        <Btn big icon={<ScanLine size={22} />} label={t('fitWidth')} onClick={() => fit('width')} />
        <Btn big icon={<Maximize2 size={22} />} label={t('fitPage')} onClick={() => fit('page')} />
      </Group>
    </>
  );

  const comments = data.pages.flatMap((p, i) => p.annots.filter((a) => a.kind === 'note' || a.kind === 'text').map((a) => ({ page: i, pageId: p.id, a })));

  return (
    <div className={`editor pdf-editor tool-${tool}`}>
      <Ribbon
        tabs={[
          { id: 'edit', label: t('tab_edit'), content: editTab },
          { id: 'pages', label: t('tab_pages'), content: pagesTab },
          { id: 'tools', label: t('tab_tools'), content: toolsTab },
          { id: 'view', label: t('tab_view'), content: viewTab },
        ]}
        extra={busy ? <span className="busy-pill">{t('working')}</span> : undefined}
      />
      <input ref={imgInput} type="file" accept="image/*" hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) pickImage(f); e.target.value = ''; }} />
      <input ref={insertInput} type="file" accept="application/pdf,.pdf,image/*" multiple hidden onChange={(e) => { if (e.target.files?.length) insertFiles(e.target.files); e.target.value = ''; }} />
      {tool === 'place' && pending && <div className="place-hint">{t('clickToPlace')} <button onClick={() => { setPending(null); setTool('select'); }}><X size={14} /></button></div>}
      <div className="pdf-body">
        <aside className="pdf-thumbs" aria-label={t('pages')}>
          {data.pages.map((p, i) => (
            <Thumb
              key={p.id}
              doc={data}
              page={p}
              index={i}
              active={i === current}
              selected={selPages.has(p.id)}
              onClick={(e) => {
                if (e.ctrlKey || e.metaKey) {
                  const n = new Set(selPages);
                  if (n.has(p.id)) n.delete(p.id); else n.add(p.id);
                  setSelPages(n);
                } else if (e.shiftKey) {
                  const a = Math.min(i, current), b = Math.max(i, current);
                  setSelPages(new Set(data.pages.slice(a, b + 1).map((x) => x.id)));
                } else {
                  setSelPages(new Set());
                  scrollToPage(i);
                }
              }}
              onDrop={(from) => movePage(from, i)}
            />
          ))}
          <div className="thumb-hint">{t('thumbHint')}</div>
        </aside>
        <div className="pdf-scroll" ref={scrollRef} onPointerDown={(e) => { if (e.target === e.currentTarget) { setSel(null); setEditing(null); } }}>
          {data.pages.map((p, i) => (
            <PageView
              key={p.id}
              doc={data}
              page={p}
              index={i}
              scale={scale}
              tool={tool}
              style={style}
              pending={pending}
              sel={sel?.page === p.id ? sel.id : null}
              editing={editing}
              hits={search.hits.filter((h) => h.page === i)}
              activeHit={search.hits[search.at]?.page === i ? search.hits[search.at] : null}
              scrollRef={scrollRef}
              onSelect={(id) => { setSel(id ? { page: p.id, id } : null); if (!id) setEditing(null); }}
              onEdit={(id) => setEditing(id)}
              onAdd={(a) => { addAnnot(p.id, a); if (a.kind === 'text') setEditing(a.id); if (tool === 'place') { setPending(null); setTool('select'); } else if (!['ink', 'marker', 'highlight'].includes(tool)) setTool('select'); }}
              onAddMany={(as, editId) => { updatePage(p.id, (pg) => ({ ...pg, annots: [...pg.annots, ...as] })); if (editId) { setSel({ page: p.id, id: editId }); setEditing(editId); } setTool('select'); }}
              onPatch={(id, patch, record) => patchAnnot(p.id, id, patch, record)}
              gesture={gesture}
              onDelete={(id) => deleteAnnot(p.id, id)}
            />
          ))}
        </div>
        {panel !== 'none' && (
          <aside className="pdf-panel">
            <header>
              <h3>{panel === 'search' ? t('find') : panel === 'form' ? t('fillForm') : panel === 'comments' ? t('comments') : t('properties')}</h3>
              <button className="icon-btn" onClick={() => setPanel('none')} aria-label={t('close')}><X size={16} /></button>
            </header>
            {panel === 'search' && (
              <div className="panel-body">
                <input className="panel-input" dir="auto" autoFocus placeholder={t('findPlaceholder')} defaultValue={search.q} onKeyDown={(e) => { if (e.key === 'Enter') { if (e.currentTarget.value === search.q && search.hits.length) stepSearch(e.shiftKey ? -1 : 1); else runSearch(e.currentTarget.value); } }} />
                <div className="panel-row">
                  <span>{search.hits.length ? `${search.at + 1} / ${search.hits.length}` : search.q ? t('noResults') : ''}</span>
                  <span className="grow" />
                  <button className="icon-btn" onClick={() => stepSearch(-1)} aria-label="Previous"><ChevronUp size={16} /></button>
                  <button className="icon-btn" onClick={() => stepSearch(1)} aria-label="Next"><ChevronDown size={16} /></button>
                </div>
                <p className="panel-note">{t('findNote')}</p>
              </div>
            )}
            {panel === 'form' && (
              <div className="panel-body">
                {fields === null && <p className="panel-note">{t('working')}</p>}
                {fields?.length === 0 && <p className="panel-note">{t('noForm')}</p>}
                {fields?.map((f) => (
                  <label key={f.src + f.name} className="form-field">
                    <span>{f.name}</span>
                    {f.type === 'checkbox' ? (
                      <input type="checkbox" checked={Boolean(f.value)} onChange={(e) => setField(f, e.target.checked)} />
                    ) : f.type === 'choice' || f.type === 'radio' ? (
                      <select value={String(f.value)} onChange={(e) => setField(f, e.target.value)}>
                        <option value="" />
                        {f.options?.map((o) => <option key={o} value={o}>{o}</option>)}
                      </select>
                    ) : f.multiline ? (
                      <textarea dir="auto" value={String(f.value)} onChange={(e) => setField(f, e.target.value)} />
                    ) : (
                      <input dir="auto" value={String(f.value)} onChange={(e) => setField(f, e.target.value)} />
                    )}
                  </label>
                ))}
                {!!fields?.length && <p className="panel-note">{t('formNote')}</p>}
              </div>
            )}
            {panel === 'comments' && (
              <div className="panel-body">
                {comments.length === 0 && <p className="panel-note">{t('noComments')}</p>}
                {comments.map(({ page, pageId, a }) => (
                  <button key={a.id} className="comment-item" onClick={() => { scrollToPage(page); setSel({ page: pageId, id: a.id }); }}>
                    <span className="comment-page">{t('page')} {page + 1}</span>
                    <span dir="auto">{(a as { text: string }).text || '…'}</span>
                  </button>
                ))}
              </div>
            )}
            {panel === 'props' && (
              <div className="panel-body">
                {(['title', 'author', 'subject'] as const).map((k) => (
                  <label key={k} className="form-field">
                    <span>{t(k === 'title' ? 'docTitle' : k === 'author' ? 'author' : 'subject')}</span>
                    <input dir="auto" value={data.meta[k] ?? ''} onChange={(e) => onChange({ ...data, meta: { ...data.meta, [k]: e.target.value || undefined } })} />
                  </label>
                ))}
                <p className="panel-note">{data.pages.length} {t('pagesCount')}</p>
              </div>
            )}
          </aside>
        )}
      </div>
      <div className="status-bar">
        <span>{t('page')} {Math.min(current + 1, data.pages.length)} / {data.pages.length}</span>
        {selPages.size > 0 && <span>{selPages.size} {t('selectedPages')}</span>}
        <span className="grow" />
        <button className="sb-btn" onClick={() => setZoom((z) => Math.max(0.25, z / 1.2))} aria-label={t('zoomOut')}><ZoomOut size={14} /></button>
        <span>{Math.round(zoom * 100)}%</span>
        <button className="sb-btn" onClick={() => setZoom((z) => Math.min(4, z * 1.2))} aria-label={t('zoomIn')}><ZoomIn size={14} /></button>
      </div>
      {showSig && <SignaturePad onClose={() => setShowSig(false)} onDone={(src) => { setShowSig(false); setPending({ kind: 'image', src }); setTool('place'); }} />}
    </div>
  );
}

// ---------------------------------------------------------------------------

function useVisible(ref: React.RefObject<HTMLElement | null>, margin = '600px') {
  const [v, setV] = useState(false);
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const io = new IntersectionObserver(([e]) => setV(e.isIntersecting), { rootMargin: margin });
    io.observe(el);
    return () => io.disconnect();
  }, [ref, margin]);
  return v;
}

/** pdf.js canvas for a page, rendered when visible. */
function PdfCanvas({ doc, page, scale }: { doc: PdfDoc; page: PdfPage; scale: number }) {
  const holder = useRef<HTMLDivElement>(null);
  const visible = useVisible(holder);
  const shown = shownSize(page);
  useEffect(() => {
    if (!visible || !page.src) return;
    let cancelled = false;
    let task: { cancel: () => void } | null = null;
    const timer = setTimeout(async () => {
      const pg = await sourcePage(doc, page);
      if (!pg || cancelled) return;
      const dpr = Math.min(2.5, window.devicePixelRatio || 1);
      const vp = pg.getViewport({ scale: scale * dpr, rotation: (page.srcRot + page.rot) % 360 });
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(vp.width);
      canvas.height = Math.floor(vp.height);
      const rt = pg.render({ canvas, canvasContext: canvas.getContext('2d')!, viewport: vp });
      task = rt;
      try {
        await rt.promise;
      } catch {
        return;
      }
      if (cancelled || !holder.current) return;
      holder.current.replaceChildren(canvas);
    }, 30);
    return () => {
      cancelled = true;
      clearTimeout(timer);
      task?.cancel();
    };
  }, [visible, doc, page.src, page.index, page.rot, page.srcRot, scale, page]);
  return <div ref={holder} className="pdf-canvas" style={{ width: shown.w * scale, height: shown.h * scale }} />;
}

function Thumb({ doc, page, index, active, selected, onClick, onDrop }: { doc: PdfDoc; page: PdfPage; index: number; active: boolean; selected: boolean; onClick: (e: React.MouseEvent) => void; onDrop: (from: number) => void }) {
  const shown = shownSize(page);
  const scale = 132 / Math.max(shown.w, shown.h * 0.78);
  const [over, setOver] = useState(false);
  return (
    <div
      className={`thumb ${active ? 'active' : ''} ${selected ? 'selected' : ''} ${over ? 'over' : ''}`}
      draggable
      onDragStart={(e) => { e.dataTransfer.setData('text/x-page', String(index)); e.dataTransfer.effectAllowed = 'move'; }}
      onDragOver={(e) => { if (e.dataTransfer.types.includes('text/x-page')) { e.preventDefault(); setOver(true); } }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => { setOver(false); const from = Number(e.dataTransfer.getData('text/x-page')); if (!Number.isNaN(from)) onDrop(from); }}
      onClick={onClick}
    >
      <div className="thumb-page" style={{ width: shown.w * scale, height: shown.h * scale }}>
        <PdfCanvas doc={doc} page={page} scale={scale} />
        <AnnotLayer page={page} scale={scale} decor={doc.decor} index={index} total={doc.pages.length} />
      </div>
      <span>{index + 1}</span>
    </div>
  );
}

/** Non-interactive drawing of annotations and decorations. */
function AnnotLayer({ page, scale, decor, index, total }: { page: PdfPage; scale: number; decor: PdfDoc['decor']; index: number; total: number }) {
  const shown = shownSize(page);
  return (
    <svg className="annot-layer" width={shown.w * scale} height={shown.h * scale} viewBox={`0 0 ${shown.w} ${shown.h}`}>
      <g transform={`matrix(${rotTransform(page).join(' ')})`}>
        {page.annots.map((a) => <AnnotShape key={a.id} a={a} />)}
      </g>
      <DecorLayer decor={decor} w={shown.w} h={shown.h} n={index + 1} total={total} />
    </svg>
  );
}

function DecorLayer({ decor, w, h, n, total }: { decor: PdfDoc['decor']; w: number; h: number; n: number; total: number }) {
  return (
    <g className="decor" pointerEvents="none">
      {decor.watermark?.text && (
        <text x={w / 2} y={h / 2} textAnchor="middle" dominantBaseline="middle" fontSize={decor.watermark.size} fontWeight={700} fill={decor.watermark.color} opacity={decor.watermark.opacity} transform={`rotate(${(-Math.atan2(h, w) * 180) / Math.PI} ${w / 2} ${h / 2})`} fontFamily="Helvetica, Arial, 'IBM Plex Sans Arabic', sans-serif">
          {decor.watermark.text}
        </text>
      )}
      {decor.header && <text x={w / 2} y={24} textAnchor="middle" dominantBaseline="middle" fontSize={10} fill="#475569">{decor.header}</text>}
      {decor.footer && <text x={w / 2} y={h - 24} textAnchor="middle" dominantBaseline="middle" fontSize={10} fill="#475569">{decor.footer}</text>}
      {decor.pageNumbers && (() => {
        const p = decorPos(decor.pageNumbers.pos, w, h);
        return <text x={p.x} y={p.y} textAnchor={p.align === 'center' ? 'middle' : p.align === 'right' ? 'end' : 'start'} dominantBaseline="middle" fontSize={decor.pageNumbers.size} fill="#475569">{pageLabel(n, total, decor.pageNumbers.fmt)}</text>;
      })()}
    </g>
  );
}

function AnnotShape({ a, editing, children }: { a: Annot; editing?: boolean; children?: ReactNode }) {
  switch (a.kind) {
    case 'rect':
      return <rect x={a.x} y={a.y} width={a.w} height={a.h} fill={a.fill ?? 'none'} stroke={a.stroke ?? 'none'} strokeWidth={a.width ?? 2} opacity={a.opacity} />;
    case 'ellipse':
      return <ellipse cx={a.x + a.w / 2} cy={a.y + a.h / 2} rx={Math.abs(a.w / 2)} ry={Math.abs(a.h / 2)} fill={a.fill ?? 'none'} stroke={a.stroke ?? 'none'} strokeWidth={a.width ?? 2} opacity={a.opacity} />;
    case 'highlight':
      return <rect x={a.x} y={a.y} width={a.w} height={a.h} fill={a.fill ?? '#fde047'} opacity={a.opacity ?? 0.4} style={{ mixBlendMode: 'multiply' }} />;
    case 'whiteout':
      return <rect x={a.x} y={a.y} width={a.w} height={a.h} fill={a.fill ?? '#ffffff'} />;
    case 'redact':
      return <rect x={a.x} y={a.y} width={a.w} height={a.h} fill="#000" />;
    case 'line':
    case 'arrow': {
      const head = a.kind === 'arrow' ? arrowHead(a.x1, a.y1, a.x2, a.y2, a.width) : null;
      return (
        <g>
          <line x1={a.x1} y1={a.y1} x2={a.x2} y2={a.y2} stroke={a.color} strokeWidth={a.width} strokeLinecap="round" />
          {head && <polygon points={head.map((p) => p.join(',')).join(' ')} fill={a.color} />}
        </g>
      );
    }
    case 'ink':
      return <path d={inkPath(a.points)} fill="none" stroke={a.color} strokeWidth={a.width} strokeLinecap="round" strokeLinejoin="round" opacity={a.opacity} style={a.opacity && a.opacity < 1 ? { mixBlendMode: 'multiply' } : undefined} />;
    case 'image':
      return <image href={a.src} x={a.x} y={a.y} width={a.w} height={a.h} preserveAspectRatio="none" />;
    case 'stamp': {
      const fs = Math.min(a.h * 0.5, (a.w / Math.max(4, a.label.length)) * 1.5);
      return (
        <g opacity={0.9}>
          <rect x={a.x} y={a.y} width={a.w} height={a.h} rx={6} fill="none" stroke={a.color} strokeWidth={2.5} />
          <text x={a.x + a.w / 2} y={a.y + a.h / 2 + 1} textAnchor="middle" dominantBaseline="middle" fontWeight={700} fontSize={fs} fill={a.color} fontFamily="Helvetica, Arial, 'IBM Plex Sans Arabic', sans-serif">{a.label}</text>
        </g>
      );
    }
    case 'note':
      return (
        <g>
          <rect x={a.x} y={a.y} width={20} height={20} rx={3} fill={a.color} stroke="rgba(0,0,0,.25)" strokeWidth={0.6} />
          <path d={`M${a.x + 5} ${a.y + 6.5}h10M${a.x + 5} ${a.y + 10}h10M${a.x + 5} ${a.y + 13.5}h7`} stroke="rgba(0,0,0,.55)" strokeWidth={1.4} />
        </g>
      );
    case 'text':
      return (
        <foreignObject x={a.x} y={a.y} width={Math.max(4, a.w)} height={Math.max(4, a.h)} style={{ overflow: 'visible' }}>
          {children ?? (
            <div
              className={`pdf-textbox ${editing ? 'editing' : ''}`}
              dir="auto"
              style={{ fontSize: a.size, color: a.color, fontFamily: FONT_CSS[a.font], fontWeight: a.bold ? 700 : 400, fontStyle: a.italic ? 'italic' : undefined, textAlign: a.align, background: a.bg }}
            >
              {a.text}
            </div>
          )}
        </foreignObject>
      );
  }
}

type Draft = { kind: 'box'; tool: Tool; x0: number; y0: number; x1: number; y1: number } | { kind: 'line'; tool: Tool; x0: number; y0: number; x1: number; y1: number } | { kind: 'ink'; tool: Tool; pts: number[] };

interface PageViewProps {
  doc: PdfDoc;
  page: PdfPage;
  index: number;
  scale: number;
  tool: Tool;
  style: ToolStyle;
  pending: { kind: 'image' | 'stamp'; src?: string; label?: string; color?: string } | null;
  sel: string | null;
  editing: string | null;
  hits: { x: number; y: number; w: number; h: number }[];
  activeHit: { x: number; y: number; w: number; h: number } | null;
  scrollRef: React.RefObject<HTMLDivElement | null>;
  onSelect: (id: string | null) => void;
  onEdit: (id: string | null) => void;
  onAdd: (a: Annot) => void;
  onAddMany: (a: Annot[], editId?: string) => void;
  onPatch: (id: string, patch: Partial<Annot>, record?: boolean) => void;
  gesture: () => () => void;
  onDelete: (id: string) => void;
}

function PageView(props: PageViewProps) {
  const { doc, page, index, scale, tool, style, pending, sel, editing, hits, activeHit, scrollRef } = props;
  const { t } = useI18n();
  const shown = shownSize(page);
  const gRef = useRef<SVGGElement>(null);
  const [draft, setDraft] = useState<Draft | null>(null);
  const [items, setItems] = useState<TextItem[] | null>(null);
  const [hover, setHover] = useState<TextItem | null>(null);
  const wrapRef = useRef<HTMLDivElement>(null);
  const visible = useVisible(wrapRef);

  const wantsText = tool === 'edittext' || tool === 'marker';
  useEffect(() => {
    if (!wantsText || !visible || items) return;
    pageText(doc, page).then(setItems);
  }, [wantsText, visible, items, doc, page]);

  const toLocal = useCallback((e: { clientX: number; clientY: number }) => {
    const g = gRef.current!;
    const m = g.getScreenCTM()!.inverse();
    const pt = new DOMPoint(e.clientX, e.clientY).matrixTransform(m);
    return { x: pt.x, y: pt.y };
  }, []);

  const selected = sel ? page.annots.find((a) => a.id === sel) ?? null : null;

  const onDown = (e: React.PointerEvent) => {
    if (e.button !== 0) return;
    const p = toLocal(e);
    if (tool === 'hand') {
      const el = scrollRef.current!;
      const sx = e.clientX, sy = e.clientY, st = el.scrollTop, sl = el.scrollLeft;
      const move = (ev: PointerEvent) => { el.scrollTop = st - (ev.clientY - sy); el.scrollLeft = sl - (ev.clientX - sx); };
      const up = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); };
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', up);
      return;
    }
    if (tool === 'select') { props.onSelect(null); return; }
    if (tool === 'place' && pending) {
      e.preventDefault();
      if (pending.kind === 'stamp') {
        const w = Math.max(110, (pending.label?.length ?? 6) * 13), h = 40;
        props.onAdd({ id: pid(), kind: 'stamp', x: p.x - w / 2, y: p.y - h / 2, w, h, label: pending.label!, color: pending.color! });
      } else if (pending.src) {
        const img = new Image();
        img.onload = () => {
          let w = Math.min(180, img.naturalWidth), h = (w * img.naturalHeight) / img.naturalWidth;
          if (h > 200) { h = 200; w = (h * img.naturalWidth) / img.naturalHeight; }
          props.onAdd({ id: pid(), kind: 'image', x: p.x - w / 2, y: p.y - h / 2, w, h, src: pending.src! });
        };
        img.src = pending.src;
      }
      return;
    }
    if (tool === 'note') {
      props.onAdd({ id: pid(), kind: 'note', x: p.x - 10, y: p.y - 10, text: '', color: '#fde047' });
      return;
    }
    if (tool === 'text') {
      e.preventDefault();
      props.onAdd({ id: pid(), kind: 'text', x: p.x, y: p.y - style.size * 0.6, w: 220, h: style.size * 1.6, text: '', size: style.size, color: style.color, font: style.font, bold: style.bold, italic: style.italic });
      return;
    }
    if (tool === 'edittext' || (tool === 'marker' && hover)) return; // handled by text item rects
    e.preventDefault();
    (e.target as Element).setPointerCapture?.(e.pointerId);
    if (tool === 'ink' || tool === 'marker') setDraft({ kind: 'ink', tool, pts: [p.x, p.y] });
    else if (tool === 'line' || tool === 'arrow') setDraft({ kind: 'line', tool, x0: p.x, y0: p.y, x1: p.x, y1: p.y });
    else setDraft({ kind: 'box', tool, x0: p.x, y0: p.y, x1: p.x, y1: p.y });
  };

  const onMove = (e: React.PointerEvent) => {
    if (!draft) return;
    const p = toLocal(e);
    if (draft.kind === 'ink') setDraft({ ...draft, pts: [...draft.pts, p.x, p.y] });
    else setDraft({ ...draft, x1: p.x, y1: p.y });
  };

  const onUp = () => {
    if (!draft) return;
    const d = draft;
    setDraft(null);
    if (d.kind === 'ink') {
      if (d.pts.length < 4) return;
      const marker = d.tool === 'marker';
      props.onAdd({ id: pid(), kind: 'ink', points: simplify(d.pts), color: marker ? '#facc15' : style.color, width: marker ? 12 : style.width, opacity: marker ? 0.45 : undefined });
      return;
    }
    if (d.kind === 'line') {
      if (Math.hypot(d.x1 - d.x0, d.y1 - d.y0) < 3) return;
      props.onAdd({ id: pid(), kind: d.tool as 'line' | 'arrow', x1: d.x0, y1: d.y0, x2: d.x1, y2: d.y1, color: style.color, width: style.width });
      return;
    }
    const x = Math.min(d.x0, d.x1), y = Math.min(d.y0, d.y1), w = Math.abs(d.x1 - d.x0), h = Math.abs(d.y1 - d.y0);
    if (w < 3 || h < 3) return;
    const kind = d.tool as 'rect' | 'ellipse' | 'highlight' | 'whiteout' | 'redact';
    if (kind === 'rect' || kind === 'ellipse') props.onAdd({ id: pid(), kind, x, y, w, h, stroke: style.color, fill: style.fill ?? 'none', width: style.width });
    else if (kind === 'highlight') props.onAdd({ id: pid(), kind, x, y, w, h, fill: '#fde047', opacity: 0.4 });
    else props.onAdd({ id: pid(), kind, x, y, w, h });
  };

  // Turn existing PDF text into editable text: cover the original and type over it.
  const editItem = (it: TextItem) => {
    const pad = 1.5;
    const cover: Annot = { id: pid(), kind: 'whiteout', x: it.x - pad, y: it.y - pad, w: it.w + pad * 2, h: it.h + pad * 2 };
    const text: TextAnnot = { id: pid(), kind: 'text', x: it.x - 2, y: it.y - 2, w: it.w + 40, h: it.h + 4, text: it.str, size: Math.round(it.size * 10) / 10, color: '#111111', font: it.font as TextAnnot['font'], bold: it.bold, italic: it.italic };
    props.onAddMany([cover, text], text.id);
  };
  const highlightItem = (it: TextItem) => props.onAdd({ id: pid(), kind: 'highlight', x: it.x - 1, y: it.y, w: it.w + 2, h: it.h, fill: '#fde047', opacity: 0.4 });

  // ----- select / move / resize -----
  const startMove = (e: React.PointerEvent, a: Annot) => {
    if (tool !== 'select' || e.button !== 0) return;
    e.stopPropagation();
    props.onSelect(a.id);
    if (editing === a.id) return;
    const start = toLocal(e);
    const done = props.gesture();
    let last = start;
    const move = (ev: PointerEvent) => {
      const p = toLocal(ev);
      const dx = p.x - last.x, dy = p.y - last.y;
      last = p;
      props.onPatch(a.id, {}, false);
      const cur = liveAnnot(a.id);
      if (cur) props.onPatch(a.id, moveAnnot(cur, dx, dy), false);
    };
    const up = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); done(); };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };
  const annotsRef = useRef(page.annots);
  annotsRef.current = page.annots;
  const liveAnnot = (id: string) => annotsRef.current.find((a) => a.id === id);

  const startResize = (e: React.PointerEvent, a: BoxLike, handle: string) => {
    e.stopPropagation();
    e.preventDefault();
    const done = props.gesture();
    const o = { x: a.x, y: a.y, w: a.w, h: a.h };
    const s0 = toLocal(e);
    const keep = a.kind === 'image' && handle.length === 2;
    const move = (ev: PointerEvent) => {
      const p = toLocal(ev);
      const dx = p.x - s0.x, dy = p.y - s0.y;
      let { x, y, w, h } = o;
      if (handle.includes('e')) w = Math.max(6, o.w + dx);
      if (handle.includes('s')) h = Math.max(6, o.h + dy);
      if (handle.includes('w')) { w = Math.max(6, o.w - dx); x = o.x + o.w - w; }
      if (handle.includes('n')) { h = Math.max(6, o.h - dy); y = o.y + o.h - h; }
      if (keep) {
        const r = o.w / o.h;
        if (w / h > r) w = h * r; else h = w / r;
        if (handle.includes('w')) x = o.x + o.w - w;
        if (handle.includes('n')) y = o.y + o.h - h;
      }
      props.onPatch(a.id, { x, y, w, h }, false);
    };
    const up = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); done(); };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };
  const startEndpoint = (e: React.PointerEvent, id: string, which: 1 | 2) => {
    e.stopPropagation();
    e.preventDefault();
    const done = props.gesture();
    const move = (ev: PointerEvent) => {
      const p = toLocal(ev);
      props.onPatch(id, which === 1 ? { x1: p.x, y1: p.y } : { x2: p.x, y2: p.y }, false);
    };
    const up = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); done(); };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const inv = 1 / scale;
  const handleR = 5 * inv;

  return (
    <div className="pdf-page-wrap" data-page-index={index} ref={wrapRef}>
      <div className="pdf-page" style={{ width: shown.w * scale, height: shown.h * scale }}>
        <PdfCanvas doc={doc} page={page} scale={scale} />
        <svg
          className="pdf-overlay"
          width={shown.w * scale}
          height={shown.h * scale}
          viewBox={`0 0 ${shown.w} ${shown.h}`}
          onPointerDown={onDown}
          onPointerMove={onMove}
          onPointerUp={onUp}
        >
          <g ref={gRef} transform={`matrix(${rotTransform(page).join(' ')})`}>
            <rect x={0} y={0} width={page.w} height={page.h} fill="transparent" />
            {page.annots.map((a) => {
              const isEditing = editing === a.id && (a.kind === 'text' || a.kind === 'note');
              return (
                <g key={a.id} className={`annot k-${a.kind} ${sel === a.id ? 'sel' : ''}`} onPointerDown={(e) => startMove(e, a)} onDoubleClick={(e) => { if (a.kind === 'text' || a.kind === 'note') { e.stopPropagation(); props.onSelect(a.id); props.onEdit(a.id); } }}>
                  {a.kind === 'text' && isEditing ? (
                    <AnnotShape a={a}>
                      <TextEditor a={a} onChange={(text, h) => props.onPatch(a.id, { text, h: Math.max(a.h, h) } as Partial<Annot>, false)} gesture={props.gesture} onDone={() => { props.onEdit(null); if (!a.text.trim() && !liveAnnot(a.id)?.['text' as keyof Annot]) props.onDelete(a.id); }} />
                    </AnnotShape>
                  ) : (
                    <AnnotShape a={a} />
                  )}
                  {/* fat invisible hit area for thin strokes */}
                  {(a.kind === 'line' || a.kind === 'arrow') && <line x1={a.x1} y1={a.y1} x2={a.x2} y2={a.y2} stroke="transparent" strokeWidth={10 * inv + a.width} />}
                  {a.kind === 'ink' && <path d={inkPath(a.points)} fill="none" stroke="transparent" strokeWidth={10 * inv + a.width} />}
                </g>
              );
            })}
            {/* search hits */}
            {hits.map((h, i) => <rect key={i} x={h.x} y={h.y} width={h.w} height={h.h} className={`hit ${activeHit && activeHit.x === h.x && activeHit.y === h.y ? 'on' : ''}`} />)}
            {/* text items for edit / highlight */}
            {wantsText && items?.filter((it) => tool === 'marker' || it.ux > 0.9).map((it, i) => (
              <rect
                key={i}
                x={it.x}
                y={it.y}
                width={it.w}
                height={it.h}
                className={`text-item ${hover === it ? 'hover' : ''} ${tool}`}
                onPointerEnter={() => setHover(it)}
                onPointerLeave={() => setHover(null)}
                onPointerDown={(e) => { e.stopPropagation(); e.preventDefault(); if (tool === 'edittext') editItem(it); else highlightItem(it); }}
              />
            ))}
            {/* draft while drawing */}
            {draft?.kind === 'ink' && <path d={inkPath(draft.pts)} fill="none" stroke={draft.tool === 'marker' ? '#facc15' : style.color} strokeWidth={draft.tool === 'marker' ? 12 : style.width} strokeLinecap="round" strokeLinejoin="round" opacity={draft.tool === 'marker' ? 0.45 : 1} />}
            {draft?.kind === 'line' && <line x1={draft.x0} y1={draft.y0} x2={draft.x1} y2={draft.y1} stroke={style.color} strokeWidth={style.width} strokeLinecap="round" />}
            {draft?.kind === 'box' && (() => {
              const x = Math.min(draft.x0, draft.x1), y = Math.min(draft.y0, draft.y1), w = Math.abs(draft.x1 - draft.x0), h = Math.abs(draft.y1 - draft.y0);
              if (draft.tool === 'ellipse') return <ellipse cx={x + w / 2} cy={y + h / 2} rx={w / 2} ry={h / 2} fill={style.fill ?? 'none'} stroke={style.color} strokeWidth={style.width} />;
              const fill = draft.tool === 'highlight' ? '#fde047' : draft.tool === 'whiteout' ? '#fff' : draft.tool === 'redact' ? '#000' : style.fill ?? 'none';
              return <rect x={x} y={y} width={w} height={h} fill={fill} opacity={draft.tool === 'highlight' ? 0.4 : 1} stroke={draft.tool === 'rect' ? style.color : draft.tool === 'whiteout' ? '#94a3b8' : 'none'} strokeWidth={draft.tool === 'rect' ? style.width : inv} strokeDasharray={draft.tool === 'whiteout' ? `${3 * inv}` : undefined} />;
            })()}
            {/* selection */}
            {selected && tool === 'select' && editing !== selected.id && (() => {
              const b = annotBounds(selected);
              const pad = 3 * inv;
              return (
                <g className="sel-ui">
                  <rect x={b.x - pad} y={b.y - pad} width={b.w + pad * 2} height={b.h + pad * 2} fill="none" stroke="var(--accent)" strokeWidth={1.5 * inv} strokeDasharray={`${4 * inv} ${3 * inv}`} pointerEvents="none" />
                  {isBoxLike(selected) && ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map((h) => {
                    const hx = h.includes('w') ? b.x : h.includes('e') ? b.x + b.w : b.x + b.w / 2;
                    const hy = h.includes('n') ? b.y : h.includes('s') ? b.y + b.h : b.y + b.h / 2;
                    return <circle key={h} cx={hx} cy={hy} r={handleR} className={`h-${h}`} fill="#fff" stroke="var(--accent)" strokeWidth={1.5 * inv} onPointerDown={(e) => startResize(e, selected, h)} />;
                  })}
                  {(selected.kind === 'line' || selected.kind === 'arrow') && (
                    <>
                      <circle cx={selected.x1} cy={selected.y1} r={handleR} fill="#fff" stroke="var(--accent)" strokeWidth={1.5 * inv} onPointerDown={(e) => startEndpoint(e, selected.id, 1)} />
                      <circle cx={selected.x2} cy={selected.y2} r={handleR} fill="#fff" stroke="var(--accent)" strokeWidth={1.5 * inv} onPointerDown={(e) => startEndpoint(e, selected.id, 2)} />
                    </>
                  )}
                </g>
              );
            })()}
            {/* sticky note popup */}
            {selected?.kind === 'note' && (
              <foreignObject x={selected.x + 24} y={selected.y} width={200} height={130} style={{ overflow: 'visible' }}>
                <div className="note-pop" onPointerDown={(e) => e.stopPropagation()}>
                  <NoteEditor
                    value={selected.text}
                    placeholder={t('commentPlaceholder')}
                    gesture={props.gesture}
                    onChange={(v) => props.onPatch(selected.id, { text: v } as Partial<Annot>, false)}
                  />
                </div>
              </foreignObject>
            )}
          </g>
          <DecorLayer decor={doc.decor} w={shown.w} h={shown.h} n={index + 1} total={doc.pages.length} />
        </svg>
      </div>
      <div className="page-num">{index + 1}</div>
    </div>
  );
}

function TextEditor({ a, onChange, onDone, gesture }: { a: TextAnnot; onChange: (text: string, h: number) => void; onDone: () => void; gesture: () => () => void }) {
  const ref = useRef<HTMLTextAreaElement>(null);
  const done = useRef<(() => void) | null>(null);
  useLayoutEffect(() => {
    const el = ref.current!;
    el.focus();
    el.setSelectionRange(el.value.length, el.value.length);
    done.current = gesture();
    return () => done.current?.();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  return (
    <textarea
      ref={ref}
      className="pdf-textbox editing"
      dir="auto"
      value={a.text}
      style={{ fontSize: a.size, color: a.color, fontFamily: FONT_CSS[a.font], fontWeight: a.bold ? 700 : 400, fontStyle: a.italic ? 'italic' : undefined, textAlign: a.align, background: a.bg ?? 'rgba(255,255,255,.6)' }}
      onPointerDown={(e) => e.stopPropagation()}
      onChange={(e) => {
        const el = e.currentTarget;
        el.style.height = 'auto';
        onChange(el.value, el.scrollHeight);
        el.style.height = '';
      }}
      onBlur={onDone}
      onKeyDown={(e) => { if (e.key === 'Escape') { e.preventDefault(); e.currentTarget.blur(); } e.stopPropagation(); }}
    />
  );
}

function NoteEditor({ value, placeholder, onChange, gesture }: { value: string; placeholder: string; onChange: (v: string) => void; gesture: () => () => void }) {
  const done = useRef<(() => void) | null>(null);
  return (
    <textarea
      dir="auto"
      autoFocus={!value}
      placeholder={placeholder}
      value={value}
      onFocus={() => (done.current = gesture())}
      onBlur={() => { done.current?.(); done.current = null; }}
      onChange={(e) => onChange(e.target.value)}
      onKeyDown={(e) => e.stopPropagation()}
    />
  );
}

/** Drop points closer than 1pt to keep ink light. */
function simplify(pts: number[]): number[] {
  const out = [pts[0], pts[1]];
  for (let i = 2; i < pts.length; i += 2) {
    const dx = pts[i] - out[out.length - 2], dy = pts[i + 1] - out[out.length - 1];
    if (dx * dx + dy * dy >= 1) out.push(Math.round(pts[i] * 10) / 10, Math.round(pts[i + 1] * 10) / 10);
  }
  return out;
}

