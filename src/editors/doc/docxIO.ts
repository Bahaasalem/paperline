import type { JSONContent } from '@tiptap/react';
import {
  AlignmentType, BorderStyle, Document, ExternalHyperlink, Footer, Header, HeadingLevel, LineRuleType, PageNumber, ImageRun, LevelFormat, Packer, PageBreak,
  PageOrientation, Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, WidthType, type IRunOptions, type ParagraphChild,
} from 'docx';
import type { DocPageSetup } from '../../types';
import { PAPER } from './page';

export async function readDocx(buf: ArrayBuffer): Promise<string> {
  const mammoth = await import('mammoth');
  const res = await mammoth.convertToHtml(
    { arrayBuffer: buf },
    {
      styleMap: ["p[style-name='Title'] => h1:fresh", "p[style-name='Subtitle'] => h2:fresh", 'u => u', 'strike => s'],
    },
  );
  return res.value || '<p></p>';
}

const MARGINS: Record<DocPageSetup['margin'], number> = { narrow: 720, normal: 1440, wide: 2160 }; // twips

type Mark = { type: string; attrs?: Record<string, unknown> };

function runOpts(text: string, marks: Mark[] = []): IRunOptions {
  const o: Record<string, unknown> = { text };
  for (const m of marks) {
    switch (m.type) {
      case 'bold': o.bold = true; break;
      case 'italic': o.italics = true; break;
      case 'underline': o.underline = {}; break;
      case 'strike': o.strike = true; break;
      case 'code': o.font = 'Consolas'; break;
      case 'subscript': o.subScript = true; break;
      case 'superscript': o.superScript = true; break;
      case 'highlight': {
        const c = (m.attrs?.color as string) || '#fef08a';
        o.shading = { type: ShadingType.CLEAR, fill: hex(c), color: 'auto' };
        break;
      }
      case 'textStyle': {
        const a = m.attrs ?? {};
        if (a.color) o.color = hex(a.color as string);
        if (a.fontFamily) o.font = String(a.fontFamily).split(',')[0].replace(/["']/g, '').trim();
        if (a.fontSize) {
          const px = parseFloat(String(a.fontSize));
          const pt = String(a.fontSize).endsWith('pt') ? px : px * 0.75;
          o.size = Math.round(pt * 2);
        }
        if (a.backgroundColor) o.shading = { type: ShadingType.CLEAR, fill: hex(a.backgroundColor as string), color: 'auto' };
        break;
      }
      case 'link':
        o.style = 'Hyperlink';
        break;
    }
  }
  return o as IRunOptions;
}

function hex(c: string): string {
  if (c.startsWith('#')) {
    const h = c.slice(1);
    return (h.length === 3 ? h.split('').map((x) => x + x).join('') : h).toUpperCase();
  }
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
  if (m) return [m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, '0')).join('').toUpperCase();
  return '000000';
}

function dataUrlToBytes(src: string): { data: Uint8Array; type: 'png' | 'jpg' | 'gif' | 'bmp' } | null {
  const m = /^data:image\/(png|jpe?g|gif|bmp);base64,(.*)$/.exec(src);
  if (!m) return null;
  const bin = atob(m[2]);
  const data = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) data[i] = bin.charCodeAt(i);
  const t = m[1] === 'jpeg' ? 'jpg' : (m[1] as 'png' | 'jpg' | 'gif' | 'bmp');
  return { data, type: t };
}

async function imageSize(src: string): Promise<{ w: number; h: number }> {
  return new Promise((res) => {
    const img = new Image();
    img.onload = () => res({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => res({ w: 300, h: 200 });
    img.src = src;
  });
}

const ALIGN: Record<string, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
};

interface Ctx {
  maxImgW: number;
  numberingRefs: number;
  headings: { level: number; text: string }[];
  tocTitle: string;
}

function collectHeadings(nodes: JSONContent[] | undefined, out: { level: number; text: string }[] = []) {
  for (const n of nodes ?? []) {
    if (n.type === 'heading') {
      const text = (n.content ?? []).map((c) => c.text ?? '').join('').trim();
      if (text) out.push({ level: Number(n.attrs?.level ?? 1), text });
    } else if (n.content && n.type !== 'table') collectHeadings(n.content, out);
  }
  return out;
}

async function inline(nodes: JSONContent[] | undefined, ctx: Ctx): Promise<ParagraphChild[]> {
  const out: ParagraphChild[] = [];
  for (const n of nodes ?? []) {
    if (n.type === 'text') {
      const link = n.marks?.find((m) => m.type === 'link');
      const run = new TextRun(runOpts(n.text ?? '', n.marks as Mark[]));
      if (link) out.push(new ExternalHyperlink({ link: String(link.attrs?.href ?? ''), children: [run] }));
      else out.push(run);
    } else if (n.type === 'hardBreak') {
      out.push(new TextRun({ break: 1 }));
    } else if (n.type === 'image') {
      const img = await imageRun(n, ctx);
      if (img) out.push(img);
    }
  }
  return out;
}

async function imageRun(n: JSONContent, ctx: Ctx) {
  const src = String(n.attrs?.src ?? '');
  const bytes = dataUrlToBytes(src);
  if (!bytes) return null;
  let { w, h } = await imageSize(src);
  const wantW = n.attrs?.width ? parseFloat(String(n.attrs.width)) : w;
  if (wantW && wantW !== w) { h = (h * wantW) / w; w = wantW; }
  if (w > ctx.maxImgW) { h = (h * ctx.maxImgW) / w; w = ctx.maxImgW; }
  return new ImageRun({ data: bytes.data, type: bytes.type, transformation: { width: Math.round(w), height: Math.round(h) } } as ConstructorParameters<typeof ImageRun>[0]);
}

function paraOpts(n: JSONContent) {
  const a = n.attrs ?? {};
  const o: Record<string, unknown> = {};
  if (a.textAlign && ALIGN[a.textAlign as string]) o.alignment = ALIGN[a.textAlign as string];
  if (a.dir === 'rtl') o.bidirectional = true;
  const spacing: Record<string, unknown> = {};
  if (a.lineHeight) { spacing.line = Math.round(parseFloat(String(a.lineHeight)) * 240); spacing.lineRule = LineRuleType.AUTO; }
  if (a.spaceBefore != null) spacing.before = Math.round(Number(a.spaceBefore) * 20);
  if (a.spaceAfter != null) spacing.after = Math.round(Number(a.spaceAfter) * 20);
  if (Object.keys(spacing).length) o.spacing = spacing;
  if (a.indent) o.indent = { start: Number(a.indent) * 540 };
  return o;
}

async function blocks(nodes: JSONContent[] | undefined, ctx: Ctx, list?: { ref: string; level: number }): Promise<(Paragraph | Table)[]> {
  const out: (Paragraph | Table)[] = [];
  for (const n of nodes ?? []) {
    switch (n.type) {
      case 'paragraph':
        out.push(new Paragraph({ ...paraOpts(n), children: await inline(n.content, ctx), numbering: list ? { reference: list.ref, level: list.level } : undefined }));
        break;
      case 'heading': {
        const lvl = Number(n.attrs?.level ?? 1);
        const heading = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6][lvl - 1];
        out.push(new Paragraph({ ...paraOpts(n), heading, children: await inline(n.content, ctx) }));
        break;
      }
      case 'blockquote':
        for (const p of n.content ?? []) {
          out.push(new Paragraph({ ...paraOpts(p), indent: { left: 720 }, border: { left: { style: BorderStyle.SINGLE, size: 12, color: 'A5B4FC', space: 8 } }, children: await inline(p.content, ctx) }));
        }
        break;
      case 'codeBlock':
        for (const line of (n.content?.map((c) => c.text ?? '').join('') ?? '').split('\n')) {
          out.push(new Paragraph({ shading: { type: ShadingType.CLEAR, fill: 'F1F5F9', color: 'auto' }, children: [new TextRun({ text: line, font: 'Consolas', size: 20 })] }));
        }
        break;
      case 'bulletList':
      case 'orderedList': {
        const ref = n.type === 'bulletList' ? 'bullets' : `numbers-${ctx.numberingRefs++}`;
        const level = list ? list.level + 1 : 0;
        for (const li of n.content ?? []) {
          for (const child of li.content ?? []) {
            if (child.type === 'bulletList' || child.type === 'orderedList') out.push(...(await blocks([child], ctx, { ref, level })));
            else out.push(...(await blocks([child], ctx, { ref, level })));
          }
        }
        break;
      }
      case 'taskList':
        for (const li of n.content ?? []) {
          const mark = li.attrs?.checked ? '☑ ' : '☐ ';
          const [first, ...rest] = li.content ?? [];
          if (first) out.push(new Paragraph({ ...paraOpts(first), indent: { start: 360 * ((list?.level ?? -1) + 1) }, children: [new TextRun(mark), ...(await inline(first.content, ctx))] }));
          if (rest.length) out.push(...(await blocks(rest, ctx, { ref: 'bullets', level: (list?.level ?? -1) + 1 })));
        }
        break;
      case 'toc':
        out.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(ctx.tocTitle)] }));
        for (const h of ctx.headings) out.push(new Paragraph({ indent: { start: 360 * (h.level - 1) }, children: [new TextRun(h.text)] }));
        break;
      case 'horizontalRule':
        out.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CBD5E1', space: 1 } }, children: [] }));
        break;
      case 'pageBreak':
        out.push(new Paragraph({ children: [new PageBreak()] }));
        break;
      case 'image': {
        const img = await imageRun(n, ctx);
        if (img) out.push(new Paragraph({ children: [img] }));
        break;
      }
      case 'table': {
        const rows: TableRow[] = [];
        for (const row of n.content ?? []) {
          const cells: TableCell[] = [];
          for (const cell of row.content ?? []) {
            const inner = (await blocks(cell.content, ctx)).filter((b): b is Paragraph => b instanceof Paragraph);
            cells.push(
              new TableCell({
                children: inner.length ? inner : [new Paragraph('')],
                columnSpan: Number(cell.attrs?.colspan ?? 1),
                rowSpan: Number(cell.attrs?.rowspan ?? 1),
                shading: cell.type === 'tableHeader' ? { type: ShadingType.CLEAR, fill: 'EEF2FF', color: 'auto' } : undefined,
              }),
            );
          }
          rows.push(new TableRow({ children: cells }));
        }
        out.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }));
        break;
      }
      default:
        if (n.content) out.push(...(await blocks(n.content, ctx, list)));
    }
  }
  return out;
}

export async function writeDocx(json: JSONContent, page: DocPageSetup): Promise<Blob> {
  const landscape = page.orientation === 'landscape';
  const margin = MARGINS[page.margin];
  const paper = PAPER[page.size ?? 'A4'];
  const pageWTwips = landscape ? paper.th : paper.tw;
  const maxImgW = ((pageWTwips - 2 * margin) / 1440) * 96;
  const ar = document.documentElement.lang === 'ar';
  const ctx: Ctx = { maxImgW, numberingRefs: 0, headings: collectHeadings(json.content), tocTitle: ar ? 'المحتويات' : 'Contents' };
  const children = await blocks(json.content, ctx);
  const numberConfigs = Array.from({ length: Math.max(1, ctx.numberingRefs) }, (_, i) => ({
    reference: `numbers-${i}`,
    levels: Array.from({ length: 6 }, (_, l) => ({
      level: l,
      format: [LevelFormat.DECIMAL, LevelFormat.LOWER_LETTER, LevelFormat.LOWER_ROMAN][l % 3],
      text: `%${l + 1}.`,
      alignment: AlignmentType.START,
      style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } },
    })),
  }));
  const doc = new Document({
    creator: 'Paperline',
    background: page.color ? { color: hex(page.color) } : undefined,
    styles: {
      default: { document: { run: { font: 'Calibri', size: 24 } } },
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: Array.from({ length: 6 }, (_, l) => ({
            level: l,
            format: LevelFormat.BULLET,
            text: ['•', '◦', '▪'][l % 3],
            alignment: AlignmentType.START,
            style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } },
          })),
        },
        ...numberConfigs,
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT, width: paper.tw, height: paper.th },
            margin: { top: margin, bottom: margin, left: margin, right: margin },
          },
        },
        ...headerFooter(page),
        children: children.length ? children : [new Paragraph('')],
      },
    ],
  });
  return Packer.toBlob(doc);
}

/** Word header and footer, with a live page number field. */
function headerFooter(page: DocPageSetup) {
  const pos = page.pageNumbers === 'bottom-center' && page.footer ? 'bottom-right' : page.pageNumbers;
  const small = { size: 18, color: '475569' };
  const num = () => new TextRun({ ...small, children: [PageNumber.CURRENT] });
  const top: Paragraph[] = [];
  const bottom: Paragraph[] = [];
  if (page.header) top.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ ...small, text: page.header })] }));
  if (pos === 'top-right') top.push(new Paragraph({ alignment: AlignmentType.RIGHT, children: [num()] }));
  if (page.footer) bottom.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ ...small, text: page.footer })] }));
  if (pos === 'bottom-center' || pos === 'bottom-right') bottom.push(new Paragraph({ alignment: pos === 'bottom-center' ? AlignmentType.CENTER : AlignmentType.RIGHT, children: [num()] }));
  return {
    headers: top.length ? { default: new Header({ children: top }) } : undefined,
    footers: bottom.length ? { default: new Footer({ children: bottom }) } : undefined,
  };
}
