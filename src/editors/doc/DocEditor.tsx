import { useEffect, useRef, useState } from 'react';
import { useEditor, EditorContent, useEditorState, type Editor } from '@tiptap/react';
import Placeholder from '@tiptap/extension-placeholder';
import {
  AlignCenter, AlignJustify, AlignLeft, AlignRight, Bold, Code2, Eraser, Highlighter, ImagePlus, Italic, Link2, List,
  ListOrdered, Minus, Pilcrow, Quote, Redo2, Strikethrough, Table2, Type, Underline, Undo2, FileX2, Rows3, Columns3,
  Trash2, SeparatorHorizontal, PilcrowLeft, PilcrowRight, RectangleVertical, RectangleHorizontal, Subscript as SubIcon,
  Superscript as SupIcon, CaseSensitive, Paintbrush, IndentIncrease, IndentDecrease, ListChecks, Search, Replace, ListTree,
  Omega, CalendarClock, Volume2, VolumeX, SpellCheck, AArrowUp, AArrowDown, ChevronDown, ChevronUp, X, BookOpenText,
  PaintBucket, Hash, PanelTop, PanelBottom, WholeWord, ALargeSmall,
} from 'lucide-react';
import { useI18n } from '../../i18n';
import { AskDialog, type Ask } from '../../ui/AskDialog';
import { Btn, ColorBtn, Group, Menu, MenuItem, Ribbon, Select } from '../../ui/Ribbon';
import { FontPicker } from '../../ui/FontPicker';
import { familyCss } from '../../fonts';
import type { DocData, PageNumberPos, PaperSize } from '../../types';
import { docExtensions, headingsOf, searchKey } from './extensions';
import type { Mark } from '@tiptap/pm/model';
import { MARGIN_PX, pageSize } from './page';

interface Props {
  data: DocData;
  onChange: (d: DocData) => void;
  onEditor?: (e: Editor | null) => void;
}

const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
const SPACINGS = ['1', '1.15', '1.5', '2', '2.5', '3'];
const SYMBOLS = [
  '©', '®', '™', '§', '¶', '†', '•', '…', '–', '—', '«', '»', '“', '”', '‘', '’', '€', '£', '¥', '¢', '₹', '﷼', '±', '×',
  '÷', '≠', '≈', '≤', '≥', '∞', '√', '∑', 'π', 'µ', 'Ω', '°', '‰', '½', '¼', '¾', '←', '→', '↑', '↓', '↔', '⇒', '✓', '✗',
  '★', '☆', '♥', '☎', '✉', '☐', '☑', '،', '؛', '؟', '٪', '٫', 'ﷺ', '۞', '٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩',
];

type CaseMode = 'upper' | 'lower' | 'title' | 'sentence';
function recase(text: string, mode: CaseMode, atStart: boolean) {
  if (mode === 'upper') return text.toLocaleUpperCase();
  if (mode === 'lower') return text.toLocaleLowerCase();
  if (mode === 'title') return text.toLocaleLowerCase().replace(/(^|[\s\-(“"'])(\p{L})/gu, (_, a, b) => a + b.toLocaleUpperCase());
  let start = atStart;
  return text.toLocaleLowerCase().replace(/\p{L}|[.!?؟]/gu, (ch) => {
    if (/[.!?؟]/.test(ch)) { start = true; return ch; }
    if (start) { start = false; return ch.toLocaleUpperCase(); }
    return ch;
  });
}


export function DocEditor({ data, onChange, onEditor }: Props) {
  const { t } = useI18n();
  const dataRef = useRef(data);
  dataRef.current = data;
  const [zoom, setZoom] = useState(1);
  const [ask, setAsk] = useState<Ask | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const [find, setFind] = useState<null | { replace: boolean }>(null);
  const [findQ, setFindQ] = useState('');
  const [replaceQ, setReplaceQ] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [wholeWord, setWholeWord] = useState(false);
  const [nav, setNav] = useState(false);
  const [speaking, setSpeaking] = useState(false);
  const [spell, setSpell] = useState(true);
  const painter = useRef<{ marks: readonly Mark[]; block: Record<string, unknown> } | null>(null);
  const [painting, setPainting] = useState(false);

  const editor = useEditor({
    extensions: [...docExtensions, Placeholder.configure({ placeholder: () => t('clickText') })],
    content: data.html,
    textDirection: 'auto',
    editorProps: { attributes: { class: 'doc-content', spellcheck: 'true' } },
    onUpdate: ({ editor }) => onChange({ ...dataRef.current, html: editor.getHTML() }),
  });

  useEffect(() => {
    onEditor?.(editor);
    return () => onEditor?.(null);
  }, [editor, onEditor]);

  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => ({
      bold: e.isActive('bold'),
      italic: e.isActive('italic'),
      underline: e.isActive('underline'),
      strike: e.isActive('strike'),
      bullet: e.isActive('bulletList'),
      ordered: e.isActive('orderedList'),
      quote: e.isActive('blockquote'),
      code: e.isActive('codeBlock'),
      table: e.isActive('table'),
      link: e.isActive('link'),
      align: (['left', 'center', 'right', 'justify'] as const).find((a) => e.isActive({ textAlign: a })) ?? 'left',
      block: e.isActive('heading', { level: 1 }) ? 'h1' : e.isActive('heading', { level: 2 }) ? 'h2' : e.isActive('heading', { level: 3 }) ? 'h3' : e.isActive('heading', { level: 4 }) ? 'h4' : 'p',
      font: (e.getAttributes('textStyle').fontFamily as string | undefined)?.replace(/["']/g, '').split(',')[0] ?? 'Inter',
      size: parseFloat(e.getAttributes('textStyle').fontSize ?? '') || 12,
      color: e.getAttributes('textStyle').color as string | undefined,
      words: e.state.doc.textContent.split(/\s+/).filter(Boolean).length,
      chars: e.state.doc.textContent.length,
      canUndo: e.can().undo(),
      canRedo: e.can().redo(),
      sub: e.isActive('subscript'),
      sup: e.isActive('superscript'),
      task: e.isActive('taskList'),
      lineHeight: (e.getAttributes('paragraph').lineHeight ?? e.getAttributes('heading').lineHeight ?? '') as string,
      search: searchKey.getState(e.state),
      headings: nav ? headingsOf(e.state.doc) : [],
    }),
  });

  // Find & replace follows the inputs.
  useEffect(() => {
    if (!editor) return;
    editor.commands.setSearch(find ? findQ : '', { caseSensitive: matchCase, wholeWord });
  }, [editor, find, findQ, matchCase, wholeWord]);

  // Ctrl+F find, Ctrl+H replace (only while this editor is on screen).
  useEffect(() => {
    if (!editor) return;
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || !editor.view.dom.isConnected || !(editor.view.dom as HTMLElement).offsetParent) return;
      const k = e.key.toLowerCase();
      if (k === 'f' || k === 'h') {
        e.preventDefault();
        const sel = editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to, ' ');
        if (sel && sel.length < 80) setFindQ(sel);
        setFind({ replace: k === 'h' });
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [editor]);

  // Format painter: the next selection takes the copied formatting.
  useEffect(() => {
    if (!editor) return;
    const onSel = () => {
      const p = painter.current;
      const { from, to, empty } = editor.state.selection;
      if (!p || empty) return;
      const tr = editor.state.tr;
      Object.values(editor.state.schema.marks).forEach((mt) => tr.removeMark(from, to, mt));
      p.marks.forEach((m) => tr.addMark(from, to, m));
      editor.state.doc.nodesBetween(from, to, (n, pos) => {
        if (n.type.name === 'paragraph' || n.type.name === 'heading') tr.setNodeMarkup(pos, undefined, { ...n.attrs, ...p.block });
      });
      editor.view.dispatch(tr);
      painter.current = null;
      setPainting(false);
    };
    editor.on('selectionUpdate', onSel);
    return () => { editor.off('selectionUpdate', onSel); };
  }, [editor]);

  useEffect(() => () => window.speechSynthesis?.cancel(), []);

  if (!editor || !s) return null;
  const chain = () => editor.chain().focus();

  const setBlock = (v: string) => {
    if (v === 'p') chain().setParagraph().run();
    else chain().setHeading({ level: Number(v.slice(1)) as 1 | 2 | 3 | 4 }).run();
  };

  const insertImage = (file: File) => {
    const reader = new FileReader();
    reader.onload = () => chain().setImage({ src: String(reader.result) }).run();
    reader.readAsDataURL(file);
  };

  const setLink = () => {
    const prev = editor.getAttributes('link').href as string | undefined;
    setAsk({
      title: t('link'),
      fields: [{ key: 'url', label: t('linkPrompt'), value: prev ?? 'https://' }],
      onOk: ({ url }) => {
        const u = url.trim();
        if (!u || u === 'https://') chain().extendMarkRange('link').unsetLink().run();
        else chain().extendMarkRange('link').setLink({ href: u }).run();
      },
    });
  };

  const changeCase = (mode: CaseMode) => {
    const { from, to } = editor.state.selection;
    if (from === to) return;
    const tr = editor.state.tr;
    let first = true;
    editor.state.doc.nodesBetween(from, to, (n, pos) => {
      if (!n.isText) return;
      const a = Math.max(from, pos), b = Math.min(to, pos + n.nodeSize);
      const text = n.text!.slice(a - pos, b - pos);
      const before = editor.state.doc.textBetween(Math.max(0, a - 3), a, '\n');
      const atStart = first && (a === 1 || /[.!?؟]\s*$|^\s*$/.test(before));
      first = false;
      tr.replaceWith(tr.mapping.map(a), tr.mapping.map(b), editor.state.schema.text(recase(text, mode, atStart), n.marks));
    });
    editor.view.dispatch(tr.setSelection(editor.state.selection.map(tr.doc, tr.mapping)));
    editor.commands.focus();
  };

  const copyFormat = () => {
    if (painting) { painter.current = null; setPainting(false); return; }
    const { $from } = editor.state.selection;
    const block = $from.parent.attrs as Record<string, unknown>;
    painter.current = { marks: editor.state.storedMarks ?? $from.marks(), block: { textAlign: block.textAlign, lineHeight: block.lineHeight, indent: block.indent, spaceBefore: block.spaceBefore, spaceAfter: block.spaceAfter } };
    setPainting(true);
  };

  const growFont = (dir: 1 | -1) => {
    const i = SIZES.findIndex((x) => x >= s.size);
    const next = SIZES[Math.max(0, Math.min(SIZES.length - 1, (i < 0 ? SIZES.length - 1 : i) + dir))];
    chain().setFontSize(`${next}pt`).run();
  };

  const insertDate = () => {
    const ar = document.documentElement.lang === 'ar';
    chain().insertContent(new Date().toLocaleString(ar ? 'ar' : undefined, { dateStyle: 'long', timeStyle: 'short' })).run();
  };

  const readAloud = () => {
    const synth = window.speechSynthesis;
    if (!synth) return;
    if (speaking) { synth.cancel(); setSpeaking(false); return; }
    const { from, to } = editor.state.selection;
    const text = from !== to ? editor.state.doc.textBetween(from, to, '\n') : editor.getText();
    if (!text.trim()) return;
    const u = new SpeechSynthesisUtterance(text);
    u.lang = /[\u0600-\u06FF]/.test(text) ? 'ar' : 'en';
    u.onend = u.onerror = () => setSpeaking(false);
    synth.speak(u);
    setSpeaking(true);
  };

  const setPage = (patch: Partial<DocData['page']>) => onChange({ ...data, page: { ...data.page, ...patch } });

  const page = data.page;
  const { w: pageW, h: pageH } = pageSize(page);
  const pad = MARGIN_PX[page.margin];

  const homeTab = (
    <>
      <Group>
        <Btn icon={<Undo2 size={18} />} label={t('undo')} disabled={!s.canUndo} onClick={() => chain().undo().run()} />
        <Btn icon={<Redo2 size={18} />} label={t('redo')} disabled={!s.canRedo} onClick={() => chain().redo().run()} />
      </Group>
      <Group label={t('font')}>
        <div className="r-col">
          <div className="r-row">
            <FontPicker value={s.font} onChange={(f) => chain().setFontFamily(familyCss(f)).run()} width={150} />
            <Select label={t('size')} value={s.size} options={SIZES.map((x) => ({ value: x, label: String(x) }))} onChange={(v) => chain().setFontSize(`${v}pt`).run()} width={60} />
            <Btn icon={<AArrowUp size={18} />} label={t('growFont')} onClick={() => growFont(1)} />
            <Btn icon={<AArrowDown size={18} />} label={t('shrinkFont')} onClick={() => growFont(-1)} />
            <Menu trigger={(o, tg) => <Btn icon={<CaseSensitive size={18} />} label={t('changeCase')} active={o} onClick={tg} />}>
              {(close) => (
                <>
                  {(['sentence', 'lower', 'upper', 'title'] as const).map((m) => (
                    <MenuItem key={m} label={t(`case_${m}`)} onClick={() => { close(); changeCase(m); }} />
                  ))}
                </>
              )}
            </Menu>
          </div>
          <div className="r-row">
            <Btn icon={<Bold size={18} />} label={t('bold')} active={s.bold} onClick={() => chain().toggleBold().run()} />
            <Btn icon={<Italic size={18} />} label={t('italic')} active={s.italic} onClick={() => chain().toggleItalic().run()} />
            <Btn icon={<Underline size={18} />} label={t('underline')} active={s.underline} onClick={() => chain().toggleUnderline().run()} />
            <Btn icon={<Strikethrough size={18} />} label={t('strike')} active={s.strike} onClick={() => chain().toggleStrike().run()} />
            <ColorBtn icon={<Type size={16} />} label={t('textColor')} color={s.color ?? '#111827'} onPick={(c) => (c ? chain().setColor(c).run() : chain().unsetColor().run())} allowNone />
            <ColorBtn icon={<Highlighter size={16} />} label={t('highlight')} color="#fef08a" onPick={(c) => (c ? chain().setHighlight({ color: c }).run() : chain().unsetHighlight().run())} allowNone />
            <Btn icon={<SubIcon size={18} />} label={t('subscript')} active={s.sub} onClick={() => chain().toggleSubscript().run()} />
            <Btn icon={<SupIcon size={18} />} label={t('superscript')} active={s.sup} onClick={() => chain().toggleSuperscript().run()} />
            <Btn icon={<Paintbrush size={18} />} label={t('formatPainter')} active={painting} onClick={copyFormat} />
            <Btn icon={<Eraser size={18} />} label={t('clearFormat')} onClick={() => chain().unsetAllMarks().clearNodes().run()} />
          </div>
        </div>
      </Group>
      <Group label={t('style')}>
        <div className="r-col">
          <div className="r-row">
            <Btn icon={<List size={18} />} label={t('bullets')} active={s.bullet} onClick={() => chain().toggleBulletList().run()} />
            <Btn icon={<ListOrdered size={18} />} label={t('numbering')} active={s.ordered} onClick={() => chain().toggleOrderedList().run()} />
            <Btn icon={<Quote size={18} />} label={t('quote')} active={s.quote} onClick={() => chain().toggleBlockquote().run()} />
            <Btn icon={<ListChecks size={18} />} label={t('checklist')} active={s.task} onClick={() => chain().toggleTaskList().run()} />
            <Btn icon={<IndentDecrease size={18} />} label={t('outdent')} onClick={() => chain().outdent().run()} />
            <Btn icon={<IndentIncrease size={18} />} label={t('indent')} onClick={() => chain().indent().run()} />
            <Menu trigger={(o, tg) => <Btn icon={<ALargeSmall size={18} />} label={t('lineSpacing')} active={o} onClick={tg} />}>
              {(close) => (
                <>
                  {SPACINGS.map((v) => (
                    <MenuItem key={v} label={v} hint={s.lineHeight === v ? '✓' : undefined} onClick={() => { close(); chain().setLineHeight(v).run(); }} />
                  ))}
                  <div className="menu-sep" />
                  <MenuItem label={t('addSpaceBefore')} onClick={() => { close(); chain().setSpacing(12, null).run(); }} />
                  <MenuItem label={t('addSpaceAfter')} onClick={() => { close(); chain().setSpacing(null, 12).run(); }} />
                  <MenuItem label={t('removeSpacing')} onClick={() => { close(); chain().setSpacing(0, 0).setLineHeight(null).run(); }} />
                </>
              )}
            </Menu>
            <Btn icon={<Code2 size={18} />} label={t('code')} active={s.code} onClick={() => chain().toggleCodeBlock().run()} />
          </div>
          <div className="r-row">
            <Btn icon={<AlignLeft size={18} />} label={t('alignLeft')} active={s.align === 'left'} onClick={() => chain().setTextAlign('left').run()} />
            <Btn icon={<AlignCenter size={18} />} label={t('alignCenter')} active={s.align === 'center'} onClick={() => chain().setTextAlign('center').run()} />
            <Btn icon={<AlignRight size={18} />} label={t('alignRight')} active={s.align === 'right'} onClick={() => chain().setTextAlign('right').run()} />
            <Btn icon={<AlignJustify size={18} />} label={t('justify')} active={s.align === 'justify'} onClick={() => chain().setTextAlign('justify').run()} />
            <Btn icon={<PilcrowRight size={18} />} label={t('ltr')} onClick={() => chain().setTextDirection('ltr').run()} />
            <Btn icon={<PilcrowLeft size={18} />} label={t('rtl')} onClick={() => chain().setTextDirection('rtl').run()} />
          </div>
        </div>
      </Group>
      <Group label={t('style')}>
        <div className="style-gallery">
          {(
            [
              ['p', t('normalText')],
              ['h1', t('heading1')],
              ['h2', t('heading2')],
              ['h3', t('heading3')],
              ['h4', t('heading4')],
            ] as const
          ).map(([v, label]) => (
            <button key={v} type="button" className={`style-chip ${v} ${s.block === v ? 'on' : ''}`} onMouseDown={(e) => e.preventDefault()} onClick={() => setBlock(v)}>
              {label}
            </button>
          ))}
        </div>
      </Group>
      <Group label={t('editing')}>
        <div className="r-col">
          <Btn icon={<Search size={18} />} label={t('find')} showLabel onClick={() => setFind({ replace: false })} />
          <Btn icon={<Replace size={18} />} label={t('replace')} showLabel onClick={() => setFind({ replace: true })} />
        </div>
      </Group>
    </>
  );

  const insertTab = (
    <>
      <Group label={t('table')}>
        <Btn big icon={<Table2 size={22} />} label={t('table')} onClick={() => chain().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()} />
        {s.table && (
          <>
            <Btn icon={<Rows3 size={18} />} label={t('addRowAfter')} showLabel onClick={() => chain().addRowAfter().run()} />
            <Btn icon={<Columns3 size={18} />} label={t('addColAfter')} showLabel onClick={() => chain().addColumnAfter().run()} />
            <Btn icon={<Trash2 size={16} />} label={t('deleteRow')} onClick={() => chain().deleteRow().run()} />
            <Btn icon={<span className="rot90"><Trash2 size={16} /></span>} label={t('deleteCol')} onClick={() => chain().deleteColumn().run()} />
            <Btn icon={<FileX2 size={16} />} label={t('deleteTable')} onClick={() => chain().deleteTable().run()} />
          </>
        )}
      </Group>
      <Group>
        <Btn big icon={<ImagePlus size={22} />} label={t('image')} onClick={() => fileRef.current?.click()} />
        <Btn big icon={<Link2 size={22} />} label={t('link')} active={s.link} onClick={setLink} />
        <Btn big icon={<Minus size={22} />} label={t('divider')} onClick={() => chain().setHorizontalRule().run()} />
        <Btn big icon={<SeparatorHorizontal size={22} />} label={t('pageBreak')} onClick={() => chain().setPageBreak().run()} />
      </Group>
      <Group>
        <Btn big icon={<BookOpenText size={22} />} label={t('toc')} onClick={() => chain().insertToc().run()} />
        <Menu trigger={(o, tg) => <Btn big icon={<Omega size={22} />} label={t('symbol')} active={o} onClick={tg} />}>
          {(close) => (
            <div className="symbol-grid">
              {SYMBOLS.map((c) => (
                <button key={c} type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => { close(); chain().insertContent(c).run(); }}>{c}</button>
              ))}
            </div>
          )}
        </Menu>
        <Btn big icon={<CalendarClock size={22} />} label={t('dateTime')} onClick={insertDate} />
      </Group>
    </>
  );

  const layoutTab = (
    <>
      <Group label={t('orientation')}>
        <Btn big icon={<RectangleVertical size={22} />} label={t('portrait')} active={page.orientation === 'portrait'} onClick={() => onChange({ ...data, page: { ...page, orientation: 'portrait' } })} />
        <Btn big icon={<RectangleHorizontal size={22} />} label={t('landscape')} active={page.orientation === 'landscape'} onClick={() => onChange({ ...data, page: { ...page, orientation: 'landscape' } })} />
      </Group>
      <Group label={t('margins')}>
        {(['narrow', 'normal', 'wide'] as const).map((m) => (
          <Btn key={m} big icon={<Pilcrow size={22} />} label={t(m)} active={page.margin === m} onClick={() => onChange({ ...data, page: { ...page, margin: m } })} />
        ))}
      </Group>
      <Group label={t('paperSize')}>
        <div className="r-col">
          <Select label={t('paperSize')} value={page.size ?? 'A4'} options={(['A4', 'Letter', 'Legal', 'A5'] as PaperSize[]).map((v) => ({ value: v, label: v }))} onChange={(v) => setPage({ size: v === 'A4' ? undefined : v })} width={96} />
          <ColorBtn icon={<PaintBucket size={16} />} label={t('pageColor')} color={page.color ?? '#ffffff'} onPick={(c) => setPage({ color: c && c.toLowerCase() !== '#ffffff' ? c : undefined })} allowNone />
        </div>
      </Group>
      <Group label={t('headerFooter')}>
        <div className="r-col">
          <label className="r-field"><PanelTop size={15} /><input className="r-input" dir="auto" placeholder={t('header')} value={page.header ?? ''} onChange={(e) => setPage({ header: e.target.value || undefined })} /></label>
          <label className="r-field"><PanelBottom size={15} /><input className="r-input" dir="auto" placeholder={t('footer')} value={page.footer ?? ''} onChange={(e) => setPage({ footer: e.target.value || undefined })} /></label>
        </div>
      </Group>
      <Group label={t('pageNumbers')}>
        <div className="r-col">
          <Hash size={16} />
          <Select label={t('pageNumbers')} value={page.pageNumbers ?? 'none'} options={[{ value: 'none', label: t('none') }, { value: 'bottom-center', label: t('pos_bottom_center') }, { value: 'bottom-right', label: t('pos_bottom_right') }, { value: 'top-right', label: t('pos_top_right') }]} onChange={(v) => setPage({ pageNumbers: v === 'none' ? undefined : (v as PageNumberPos) })} width={130} />
        </div>
      </Group>
    </>
  );

  const reviewTab = (
    <>
      <Group label={t('proofing')}>
        <Btn big icon={<SpellCheck size={22} />} label={t('spelling')} active={spell} onClick={() => setSpell((v) => !v)} />
        <Btn big icon={<WholeWord size={22} />} label={`${s.words} ${t('words')}`} onClick={() => undefined} />
      </Group>
      <Group label={t('speech')}>
        <Btn big icon={speaking ? <VolumeX size={22} /> : <Volume2 size={22} />} label={speaking ? t('stopReading') : t('readAloud')} active={speaking} onClick={readAloud} />
      </Group>
    </>
  );

  const viewTab = (
    <>
      <Group label={t('zoom')}>
        <Select label={t('zoom')} value={zoom} options={[0.5, 0.75, 1, 1.25, 1.5, 2].map((z) => ({ value: z, label: `${Math.round(z * 100)}%` }))} onChange={setZoom} width={90} />
      </Group>
      <Group label={t('panels')}>
        <Btn big icon={<ListTree size={22} />} label={t('navigation')} active={nav} onClick={() => setNav((v) => !v)} />
      </Group>
    </>
  );

  return (
    <div className="editor doc-editor">
      <Ribbon
        tabs={[
          { id: 'home', label: t('tab_home'), content: homeTab },
          { id: 'insert', label: t('tab_insert'), content: insertTab },
          { id: 'layout', label: t('tab_layout'), content: layoutTab },
          { id: 'review', label: t('tab_review'), content: reviewTab },
          { id: 'view', label: t('tab_view'), content: viewTab },
        ]}
      />
      <input ref={fileRef} type="file" accept="image/*" hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) insertImage(f); e.target.value = ''; }} />
      <div className="doc-body">
        {nav && (
          <aside className="doc-nav" aria-label={t('navigation')}>
            <header><b>{t('navigation')}</b><button className="icon-btn" onClick={() => setNav(false)} aria-label={t('close')}><X size={15} /></button></header>
            {s.headings.length === 0 && <p className="panel-note">{t('noHeadings')}</p>}
            {s.headings.map((h, i) => (
              <button key={i} className={`nav-item l${h.level}`} dir="auto" onClick={() => editor.chain().focus().setTextSelection(h.pos + 1).scrollIntoView().run()}>{h.text}</button>
            ))}
          </aside>
        )}
        <div className="doc-scroll" onMouseDown={(e) => { if (e.target === e.currentTarget) editor.commands.focus('end'); }}>
          <div className="doc-zoom" style={{ zoom }}>
            <div
              className="doc-page"
              style={{ width: pageW, minHeight: pageH, padding: pad, backgroundColor: page.color, ['--page-h' as string]: `${pageH}px`, ['--band-h' as string]: `${pad}px`, ['--band-pad' as string]: `${pad}px` }}
              onMouseDown={(e) => { if (e.target === e.currentTarget) editor.commands.focus('end'); }}
            >
              {page.header && <div className="page-band top" style={{ height: pad }} dir="auto">{page.header}</div>}
              <EditorContent editor={editor} spellCheck={spell} />
              {(page.footer || page.pageNumbers) && (
                <div className="page-band bottom" style={{ height: pad }}>
                  {page.pageNumbers === 'top-right' ? null : <span />}
                  <span dir="auto">{page.footer}</span>
                  <span>{page.pageNumbers && page.pageNumbers !== 'top-right' ? '1' : ''}</span>
                </div>
              )}
              {page.pageNumbers === 'top-right' && <div className="page-band top num" style={{ height: pad }}>1</div>}
            </div>
          </div>
        </div>
        {find && (
          <div className="find-panel" role="search" onKeyDown={(e) => { if (e.key === 'Escape') { setFind(null); editor.commands.focus(); } }}>
            <div className="find-row">
              <Search size={15} />
              <input autoFocus dir="auto" placeholder={t('findPlaceholder')} value={findQ} onChange={(e) => setFindQ(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); editor.commands.gotoMatch(e.shiftKey ? -1 : 1); } }} />
              <span className="find-count">{findQ ? (s.search?.matches.length ? `${(s.search.at ?? 0) + 1}/${s.search.matches.length}` : t('noResults')) : ''}</span>
              <button className="icon-btn" aria-label="Previous" onClick={() => editor.commands.gotoMatch(-1)}><ChevronUp size={16} /></button>
              <button className="icon-btn" aria-label="Next" onClick={() => editor.commands.gotoMatch(1)}><ChevronDown size={16} /></button>
              <button className="icon-btn" aria-label={t('close')} onClick={() => setFind(null)}><X size={16} /></button>
            </div>
            {find.replace && (
              <div className="find-row">
                <Replace size={15} />
                <input dir="auto" placeholder={t('replaceWith')} value={replaceQ} onChange={(e) => setReplaceQ(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); editor.commands.replaceMatch(replaceQ); } }} />
                <button className="ghost-btn small" onClick={() => editor.commands.replaceMatch(replaceQ)}>{t('replaceOne')}</button>
                <button className="ghost-btn small" onClick={() => editor.commands.replaceAll(replaceQ)}>{t('replaceAll')}</button>
              </div>
            )}
            <div className="find-opts">
              <label><input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} /> {t('matchCase')}</label>
              <label><input type="checkbox" checked={wholeWord} onChange={(e) => setWholeWord(e.target.checked)} /> {t('wholeWord')}</label>
              {!find.replace && <button className="link-btn" onClick={() => setFind({ replace: true })}>{t('replace')}</button>}
            </div>
          </div>
        )}
      </div>
      <div className="status-bar">
        <span>{s.words} {t('words')}</span>
        <span>{s.chars} {t('characters')}</span>
        <span className="grow" />
        <span>{Math.round(zoom * 100)}%</span>
      </div>
      {ask && <AskDialog ask={ask} onClose={() => { setAsk(null); editor.commands.focus(); }} />}
    </div>
  );
}
