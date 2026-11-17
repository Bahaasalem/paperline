import { Node, mergeAttributes, type Extensions } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyleKit } from '@tiptap/extension-text-style';
import Highlight from '@tiptap/extension-highlight';
import { TableKit } from '@tiptap/extension-table';
import Image from '@tiptap/extension-image';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    pageBreak: { setPageBreak: () => ReturnType };
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

/** Shared by the editor, HTML→JSON conversion and export. */
export const docExtensions: Extensions = [
  StarterKit.configure({ heading: { levels: [1, 2, 3, 4] }, link: { openOnClick: false, autolink: true } }),
  TextAlign.configure({ types: ['heading', 'paragraph'] }),
  TextStyleKit,
  Highlight.configure({ multicolor: true }),
  TableKit.configure({ table: { resizable: true } }),
  ResizableImage,
  PageBreak,
];
