import type { JSONContent } from '@tiptap/react';
import {
  AlignmentType, BorderStyle, Document, ExternalHyperlink, HeadingLevel, ImageRun, LevelFormat, Packer, PageBreak,
  PageOrientation, Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, WidthType, type IRunOptions, type ParagraphChild,
} from 'docx';
import type { DocPageSetup } from '../../types';

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
  const pageWTwips = landscape ? 16838 : 11906;
  const maxImgW = ((pageWTwips - 2 * margin) / 1440) * 96;
  const ctx: Ctx = { maxImgW, numberingRefs: 0 };
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
            size: { orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT, width: 11906, height: 16838 },
            margin: { top: margin, bottom: margin, left: margin, right: margin },
          },
        },
        children: children.length ? children : [new Paragraph('')],
      },
    ],
  });
  return Packer.toBlob(doc);
}
