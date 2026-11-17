import { useEffect, useRef, useState } from 'react';
import { useEditor, EditorContent, useEditorState, type Editor } from '@tiptap/react';
import Placeholder from '@tiptap/extension-placeholder';
import {
  AlignCenter, AlignJustify, AlignLeft, AlignRight, Bold, Code2, Eraser, Highlighter, ImagePlus, Italic, Link2, List,
  ListOrdered, Minus, Pilcrow, Quote, Redo2, Strikethrough, Table2, Type, Underline, Undo2, FileX2, Rows3, Columns3,
  Trash2, SeparatorHorizontal, PilcrowLeft, PilcrowRight, RectangleVertical, RectangleHorizontal,
} from 'lucide-react';
import { useI18n } from '../../i18n';
import { Btn, ColorBtn, Group, Ribbon, Select } from '../../ui/Ribbon';
import type { DocData } from '../../types';
import { docExtensions } from './extensions';
import { MARGIN_PX, PAGE } from './page';

interface Props {
  data: DocData;
  onChange: (d: DocData) => void;
  onEditor?: (e: Editor | null) => void;
}

export const FONTS = [
  { value: 'Inter', label: 'Inter' },
  { value: 'Calibri', label: 'Calibri' },
  { value: 'Arial', label: 'Arial' },
  { value: 'Georgia', label: 'Georgia' },
  { value: 'Times New Roman', label: 'Times New Roman' },
  { value: 'Courier New', label: 'Courier New' },
  { value: 'IBM Plex Sans Arabic', label: 'IBM Plex Sans Arabic' },
  { value: 'Tahoma', label: 'Tahoma' },
];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];


export function DocEditor({ data, onChange, onEditor }: Props) {
  const { t } = useI18n();
  const dataRef = useRef(data);
  dataRef.current = data;
  const [zoom, setZoom] = useState(1);
  const fileRef = useRef<HTMLInputElement>(null);

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
    }),
  });

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
    const url = window.prompt(t('linkPrompt'), prev ?? 'https://');
    if (url === null) return;
    if (url === '') chain().extendMarkRange('link').unsetLink().run();
    else chain().extendMarkRange('link').setLink({ href: url }).run();
  };

  const page = data.page;
  const pageW = page.orientation === 'landscape' ? PAGE.h : PAGE.w;
  const pageH = page.orientation === 'landscape' ? PAGE.w : PAGE.h;
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
            <Select label={t('font')} value={s.font} options={FONTS} onChange={(f) => chain().setFontFamily(f).run()} width={150} />
            <Select label={t('size')} value={s.size} options={SIZES.map((x) => ({ value: x, label: String(x) }))} onChange={(v) => chain().setFontSize(`${v}pt`).run()} width={60} />
          </div>
          <div className="r-row">
            <Btn icon={<Bold size={18} />} label={t('bold')} active={s.bold} onClick={() => chain().toggleBold().run()} />
            <Btn icon={<Italic size={18} />} label={t('italic')} active={s.italic} onClick={() => chain().toggleItalic().run()} />
            <Btn icon={<Underline size={18} />} label={t('underline')} active={s.underline} onClick={() => chain().toggleUnderline().run()} />
            <Btn icon={<Strikethrough size={18} />} label={t('strike')} active={s.strike} onClick={() => chain().toggleStrike().run()} />
            <ColorBtn icon={<Type size={16} />} label={t('textColor')} color={s.color ?? '#111827'} onPick={(c) => (c ? chain().setColor(c).run() : chain().unsetColor().run())} allowNone />
            <ColorBtn icon={<Highlighter size={16} />} label={t('highlight')} color="#fef08a" onPick={(c) => (c ? chain().setHighlight({ color: c }).run() : chain().unsetHighlight().run())} allowNone />
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
            ] as const
          ).map(([v, label]) => (
            <button key={v} type="button" className={`style-chip ${v} ${s.block === v ? 'on' : ''}`} onMouseDown={(e) => e.preventDefault()} onClick={() => setBlock(v)}>
              {label}
            </button>
          ))}
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
    </>
  );

  const viewTab = (
    <Group label={t('zoom')}>
      <Select label={t('zoom')} value={zoom} options={[0.5, 0.75, 1, 1.25, 1.5, 2].map((z) => ({ value: z, label: `${Math.round(z * 100)}%` }))} onChange={setZoom} width={90} />
    </Group>
  );

  return (
    <div className="editor doc-editor">
      <Ribbon
        tabs={[
          { id: 'home', label: t('tab_home'), content: homeTab },
          { id: 'insert', label: t('tab_insert'), content: insertTab },
          { id: 'layout', label: t('tab_layout'), content: layoutTab },
          { id: 'view', label: t('tab_view'), content: viewTab },
        ]}
      />
      <input ref={fileRef} type="file" accept="image/*" hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) insertImage(f); e.target.value = ''; }} />
      <div className="doc-scroll" onMouseDown={(e) => { if (e.target === e.currentTarget) editor.commands.focus('end'); }}>
        <div className="doc-zoom" style={{ zoom }}>
          <div
            className="doc-page"
            style={{ width: pageW, minHeight: pageH, padding: pad, ['--page-h' as string]: `${pageH}px` }}
            onMouseDown={(e) => { if (e.target === e.currentTarget) editor.commands.focus('end'); }}
          >
            <EditorContent editor={editor} />
          </div>
        </div>
      </div>
      <div className="status-bar">
        <span>{s.words} {t('words')}</span>
        <span>{s.chars} {t('characters')}</span>
        <span className="grow" />
        <span>{Math.round(zoom * 100)}%</span>
      </div>
    </div>
  );
}
