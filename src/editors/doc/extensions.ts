import { Extension, Node, mergeAttributes, type Extensions } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import type { Node as PMNode } from '@tiptap/pm/model';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';
import { TaskItem, TaskList } from '@tiptap/extension-list';
import StarterKit from '@tiptap/starter-kit';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyleKit } from '@tiptap/extension-text-style';
import Highlight from '@tiptap/extension-highlight';
import { TableKit } from '@tiptap/extension-table';
import Image from '@tiptap/extension-image';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    pageBreak: { setPageBreak: () => ReturnType };
    paragraphFormat: {
      setLineHeight: (v: string | null) => ReturnType;
      setSpacing: (before: number | null, after: number | null) => ReturnType;
      indent: () => ReturnType;
      outdent: () => ReturnType;
    };
    toc: { insertToc: () => ReturnType };
    search: {
      setSearch: (q: string, opts?: { caseSensitive?: boolean; wholeWord?: boolean }) => ReturnType;
      gotoMatch: (dir: 1 | -1) => ReturnType;
      replaceMatch: (text: string) => ReturnType;
      replaceAll: (text: string) => ReturnType;
    };
  }
}

export const PageBreak = Node.create({
  name: 'pageBreak',
  group: 'block',
  atom: true,
  selectable: true,
  parseHTML: () => [{ tag: 'div[data-page-break]' }, { tag: 'hr[data-page-break]' }],
  renderHTML: ({ HTMLAttributes }) => ['div', mergeAttributes(HTMLAttributes, { 'data-page-break': 'true', class: 'page-break' })],
  addCommands() {
    return {
      setPageBreak:
        () =>
        ({ chain }) =>
          chain().insertContent({ type: this.name }).insertContent({ type: 'paragraph' }).run(),
    };
  },
});

export const ResizableImage = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      width: {
        default: null,
        parseHTML: (el) => el.getAttribute('width') || (el as HTMLElement).style.width?.replace('px', '') || null,
        renderHTML: (a) => (a.width ? { width: a.width } : {}),
      },
    };
  },
}).configure({ allowBase64: true, inline: false });

const BLOCKS = ['paragraph', 'heading'];
const INDENT_STEP = 36; // px, half an inch

/** Line spacing, space before/after and indentation for paragraphs and headings (Word's Paragraph dialog). */
export const ParagraphFormat = Extension.create({
  name: 'paragraphFormat',
  addGlobalAttributes() {
    return [
      {
        types: BLOCKS,
        attributes: {
          lineHeight: {
            default: null,
            parseHTML: (el) => el.style.lineHeight || null,
            renderHTML: (a) => (a.lineHeight ? { style: `line-height: ${a.lineHeight}` } : {}),
          },
          spaceBefore: {
            default: null,
            parseHTML: (el) => (el.style.marginTop ? parseFloat(el.style.marginTop) : null),
            renderHTML: (a) => (a.spaceBefore != null ? { style: `margin-top: ${a.spaceBefore}pt` } : {}),
          },
          spaceAfter: {
            default: null,
            parseHTML: (el) => (el.style.marginBottom ? parseFloat(el.style.marginBottom) : null),
            renderHTML: (a) => (a.spaceAfter != null ? { style: `margin-bottom: ${a.spaceAfter}pt` } : {}),
          },
          indent: {
            default: 0,
            parseHTML: (el) => Math.round((parseFloat(el.style.marginInlineStart || el.style.marginLeft || '0') || 0) / INDENT_STEP),
            renderHTML: (a) => (a.indent ? { style: `margin-inline-start: ${a.indent * INDENT_STEP}px` } : {}),
          },
        },
      },
    ];
  },
  addCommands() {
    const each =
      (fn: (attrs: Record<string, unknown>) => Record<string, unknown> | null) =>
      ({ tr, state, dispatch }: { tr: import('@tiptap/pm/state').Transaction; state: import('@tiptap/pm/state').EditorState; dispatch?: unknown }) => {
        const { from, to } = state.selection;
        let changed = false;
        state.doc.nodesBetween(from, to, (node, pos) => {
          if (!BLOCKS.includes(node.type.name)) return;
          const next = fn(node.attrs);
          if (next) {
            tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...next });
            changed = true;
          }
        });
        if (changed && dispatch) (dispatch as (t: typeof tr) => void)(tr);
        return changed;
      };
    return {
      setLineHeight: (v) => each(() => ({ lineHeight: v })),
      setSpacing: (before, after) => each(() => ({ spaceBefore: before, spaceAfter: after })),
      indent:
        () =>
        ({ editor, commands, ...rest }) =>
          editor.isActive('listItem') ? commands.sinkListItem('listItem') : editor.isActive('taskItem') ? commands.sinkListItem('taskItem') : each((a) => ({ indent: Math.min(10, ((a.indent as number) || 0) + 1) }))(rest as never),
      outdent:
        () =>
        ({ editor, commands, ...rest }) =>
          editor.isActive('listItem') ? commands.liftListItem('listItem') : editor.isActive('taskItem') ? commands.liftListItem('taskItem') : each((a) => ((a.indent as number) ? { indent: (a.indent as number) - 1 } : null))(rest as never),
    };
  },
  addKeyboardShortcuts() {
    return {
      'Mod-]': () => this.editor.commands.indent(),
      'Mod-[': () => this.editor.commands.outdent(),
    };
  },
});

/** Headings of a document, for the table of contents and the navigation pane. */
export function headingsOf(doc: PMNode) {
  const out: { level: number; text: string; pos: number }[] = [];
  doc.descendants((n, pos) => {
    if (n.type.name === 'heading' && n.textContent.trim()) out.push({ level: n.attrs.level as number, text: n.textContent, pos });
    return n.type.name !== 'table';
  });
  return out;
}

/** A live table of contents built from the document's headings. */
export const TableOfContents = Node.create({
  name: 'toc',
  group: 'block',
  atom: true,
  selectable: true,
  draggable: true,
  parseHTML: () => [{ tag: 'div[data-toc]' }],
  renderHTML: ({ HTMLAttributes }) => ['div', mergeAttributes(HTMLAttributes, { 'data-toc': 'true', class: 'toc' })],
  addCommands() {
    return {
      insertToc:
        () =>
        ({ chain }) =>
          chain().insertContent([{ type: this.name }, { type: 'paragraph' }]).run(),
    };
  },
  addNodeView() {
    return ({ editor }) => {
      const dom = document.createElement('div');
      dom.className = 'toc';
      dom.setAttribute('data-toc', 'true');
      dom.contentEditable = 'false';
      const render = () => {
        const hs = headingsOf(editor.state.doc);
        const title = document.createElement('div');
        title.className = 'toc-title';
        title.textContent = document.documentElement.lang === 'ar' ? 'المحتويات' : 'Contents';
        const list = document.createElement('div');
        for (const h of hs) {
          const row = document.createElement('a');
          row.className = `toc-row l${h.level}`;
          row.textContent = h.text;
          row.onmousedown = (e) => {
            e.preventDefault();
            editor.chain().focus().setTextSelection(h.pos + 1).scrollIntoView().run();
          };
          list.appendChild(row);
        }
        if (!hs.length) {
          const empty = document.createElement('div');
          empty.className = 'toc-empty';
          empty.textContent = document.documentElement.lang === 'ar' ? 'أضف عناوين لتظهر هنا.' : 'Add headings and they will appear here.';
          list.appendChild(empty);
        }
        dom.replaceChildren(title, list);
      };
      render();
      const onUpdate = () => render();
      editor.on('update', onUpdate);
      return { dom, destroy: () => editor.off('update', onUpdate), ignoreMutation: () => true, stopEvent: () => false };
    };
  },
});

// ---------- find & replace ----------
interface SearchState {
  q: string;
  caseSensitive: boolean;
  wholeWord: boolean;
  matches: { from: number; to: number }[];
  at: number;
}
export const searchKey = new PluginKey<SearchState>('search');

function findAll(doc: PMNode, q: string, caseSensitive: boolean, wholeWord: boolean) {
  const out: { from: number; to: number }[] = [];
  if (!q) return out;
  const esc = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(wholeWord ? `(?<![\\p{L}\\p{N}_])${esc}(?![\\p{L}\\p{N}_])` : esc, caseSensitive ? 'gu' : 'giu');
  doc.descendants((n, pos) => {
    if (!n.isTextblock) return true;
    // map text offsets back to document positions, child by child
    let text = '';
    const map: number[] = [];
    n.forEach((child, offset) => {
      if (child.isText) {
        for (let i = 0; i < child.text!.length; i++) map.push(pos + 1 + offset + i);
        text += child.text;
      } else {
        map.push(-1);
        text += '￼';
      }
    });
    let m: RegExpExecArray | null;
    while ((m = re.exec(text))) {
      if (!m[0].length) { re.lastIndex++; continue; }
      const a = map[m.index], b = map[m.index + m[0].length - 1];
      if (a >= 0 && b >= 0) out.push({ from: a, to: b + 1 });
    }
    return false;
  });
  return out;
}

export const Search = Extension.create({
  name: 'search',
  addProseMirrorPlugins() {
    return [
      new Plugin<SearchState>({
        key: searchKey,
        state: {
          init: () => ({ q: '', caseSensitive: false, wholeWord: false, matches: [], at: 0 }),
          apply(tr, prev) {
            const meta = tr.getMeta(searchKey) as Partial<SearchState> | undefined;
            const next = meta ? { ...prev, ...meta } : prev;
            if (meta || tr.docChanged) {
              const matches = findAll(tr.doc, next.q, next.caseSensitive, next.wholeWord);
              return { ...next, matches, at: Math.min(next.at, Math.max(0, matches.length - 1)) };
            }
            return next;
          },
        },
        props: {
          decorations(state) {
            const s = searchKey.getState(state);
            if (!s?.matches.length) return null;
            return DecorationSet.create(state.doc, s.matches.map((m, i) => Decoration.inline(m.from, m.to, { class: i === s.at ? 'search-hit current' : 'search-hit' })));
          },
        },
      }),
    ];
  },
  addCommands() {
    const select = (s: SearchState) => s.matches[s.at];
    return {
      setSearch:
        (q, opts = {}) =>
        ({ tr, dispatch }) => {
          if (dispatch) dispatch(tr.setMeta(searchKey, { q, caseSensitive: !!opts.caseSensitive, wholeWord: !!opts.wholeWord, at: 0 }));
          return true;
        },
      gotoMatch:
        (dir) =>
        ({ state, tr, dispatch, editor }) => {
          const s = searchKey.getState(state);
          if (!s?.matches.length) return false;
          const at = (s.at + dir + s.matches.length) % s.matches.length;
          if (dispatch) {
            dispatch(tr.setMeta(searchKey, { at }));
            const m = s.matches[at];
            requestAnimationFrame(() => editor.chain().setTextSelection(m).scrollIntoView().run());
          }
          return true;
        },
      replaceMatch:
        (text) =>
        ({ state, tr, dispatch }) => {
          const s = searchKey.getState(state);
          const m = s && select(s);
          if (!m) return false;
          if (dispatch) dispatch(text ? tr.insertText(text, m.from, m.to) : tr.delete(m.from, m.to));
          return true;
        },
      replaceAll:
        (text) =>
        ({ state, tr, dispatch }) => {
          const s = searchKey.getState(state);
          if (!s?.matches.length) return false;
          if (dispatch) {
            for (const m of [...s.matches].reverse()) {
              if (text) tr.insertText(text, m.from, m.to);
              else tr.delete(m.from, m.to);
            }
            dispatch(tr);
          }
          return true;
        },
    };
  },
});

/** Shared by the editor, HTML→JSON conversion and export. */
export const docExtensions: Extensions = [
  StarterKit.configure({ heading: { levels: [1, 2, 3, 4] }, link: { openOnClick: false, autolink: true } }),
  TextAlign.configure({ types: ['heading', 'paragraph'] }),
  TextStyleKit,
  Highlight.configure({ multicolor: true }),
  TableKit.configure({ table: { resizable: true } }),
  ResizableImage,
  PageBreak,
  Subscript,
  Superscript,
  TaskList,
  TaskItem.configure({ nested: true }),
  ParagraphFormat,
  TableOfContents,
  Search,
];
