export type Rot = 0 | 90 | 180 | 270;

interface AnnotBase {
  id: string;
}

export interface TextAnnot extends AnnotBase {
  kind: 'text';
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  size: number; // pt
  color: string;
  /** 'sans' | 'serif' | 'mono' are the PDF standard fonts (real text); any other family is a device font. */
  font: string;
  bold?: boolean;
  italic?: boolean;
  align?: 'left' | 'center' | 'right';
  bg?: string;
}

export interface BoxAnnot extends AnnotBase {
  kind: 'rect' | 'ellipse' | 'highlight' | 'whiteout' | 'redact';
  x: number;
  y: number;
  w: number;
  h: number;
  stroke?: string;
  fill?: string;
  width?: number;
  opacity?: number;
}

export interface LineAnnot extends AnnotBase {
  kind: 'line' | 'arrow';
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
  width: number;
}

export interface InkAnnot extends AnnotBase {
  kind: 'ink';
  points: number[]; // flat x,y pairs
  color: string;
  width: number;
  opacity?: number;
}

export interface ImageAnnot extends AnnotBase {
  kind: 'image';
  x: number;
  y: number;
  w: number;
  h: number;
  src: string; // data URL (png/jpeg)
}

export interface NoteAnnot extends AnnotBase {
  kind: 'note';
  x: number;
  y: number;
  text: string;
  color: string;
}

export interface StampAnnot extends AnnotBase {
  kind: 'stamp';
  x: number;
  y: number;
  w: number;
  h: number;
  label: string;
  color: string;
}

export type Annot = TextAnnot | BoxAnnot | LineAnnot | InkAnnot | ImageAnnot | NoteAnnot | StampAnnot;
export type BoxLike = TextAnnot | BoxAnnot | ImageAnnot | StampAnnot;

export interface PdfPage {
  id: string;
  /** Source PDF id, or null for a blank page. */
  src: string | null;
  /** Page index inside the source. */
  index: number;
  /** Size in points as the source shows it (its own /Rotate applied). */
  w: number;
  h: number;
  /** The source page's own rotation. */
  srcRot: number;
  /** Extra rotation added by the user. */
  rot: Rot;
  annots: Annot[];
}

export interface PageDecor {
  pageNumbers?: { pos: 'bottom-center' | 'bottom-right' | 'bottom-left' | 'top-right'; fmt: 'n' | 'n-of-N' | 'page-n'; size: number };
  watermark?: { text: string; color: string; opacity: number; size: number };
  header?: string;
  footer?: string;
}

export interface PdfDoc {
  /** Original PDF files, base64, keyed by id. */
  sources: Record<string, string>;
  pages: PdfPage[];
  /** Form field values keyed by source id, then field name. */
  form: Record<string, Record<string, string | boolean>>;
  decor: PageDecor;
  meta: { title?: string; author?: string; subject?: string };
}

export const A4 = { w: 595.28, h: 841.89 };

let seq = 0;
export const pid = () => `p${Date.now().toString(36)}${(seq++).toString(36)}${Math.random().toString(36).slice(2, 5)}`;

export function newPdfDoc(): PdfDoc {
  return { sources: {}, pages: [blankPage()], form: {}, decor: {}, meta: {} };
}

export function blankPage(w = A4.w, h = A4.h): PdfPage {
  return { id: pid(), src: null, index: 0, w, h, srcRot: 0, rot: 0, annots: [] };
}

export function bytesToB64(bytes: Uint8Array): string {
  let s = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) s += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(s);
}

export function b64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/** Size of the page as displayed, after the user's rotation. */
export const shownSize = (p: PdfPage) => (p.rot % 180 ? { w: p.h, h: p.w } : { w: p.w, h: p.h });

/** SVG/canvas transform that maps page coordinates into the user-rotated display. */
export function rotTransform(p: PdfPage): [number, number, number, number, number, number] {
  switch (p.rot) {
    case 90:
      return [0, 1, -1, 0, p.h, 0];
    case 180:
      return [-1, 0, 0, -1, p.w, p.h];
    case 270:
      return [0, -1, 1, 0, 0, p.w];
    default:
      return [1, 0, 0, 1, 0, 0];
  }
}

export const isBoxLike = (a: Annot): a is BoxLike => 'w' in a;

export function annotBounds(a: Annot): { x: number; y: number; w: number; h: number } {
  switch (a.kind) {
    case 'line':
    case 'arrow':
      return { x: Math.min(a.x1, a.x2), y: Math.min(a.y1, a.y2), w: Math.abs(a.x2 - a.x1), h: Math.abs(a.y2 - a.y1) };
    case 'ink': {
      let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
      for (let i = 0; i < a.points.length; i += 2) {
        x1 = Math.min(x1, a.points[i]); x2 = Math.max(x2, a.points[i]);
        y1 = Math.min(y1, a.points[i + 1]); y2 = Math.max(y2, a.points[i + 1]);
      }
      return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
    }
    case 'note':
      return { x: a.x, y: a.y, w: 22, h: 22 };
    default:
      return { x: a.x, y: a.y, w: a.w, h: a.h };
  }
}

export function moveAnnot(a: Annot, dx: number, dy: number): Annot {
  switch (a.kind) {
    case 'line':
    case 'arrow':
      return { ...a, x1: a.x1 + dx, y1: a.y1 + dy, x2: a.x2 + dx, y2: a.y2 + dy };
    case 'ink':
      return { ...a, points: a.points.map((v, i) => v + (i % 2 ? dy : dx)) };
    default:
      return { ...a, x: a.x + dx, y: a.y + dy } as Annot;
  }
}

export const FONT_CSS: Record<string, string> = {
  sans: 'Helvetica, Arial, "IBM Plex Sans Arabic", sans-serif',
  serif: '"Times New Roman", Times, serif',
  mono: '"Courier New", Courier, monospace',
};

export function pageLabel(n: number, total: number, fmt: NonNullable<PageDecor['pageNumbers']>['fmt']) {
  return fmt === 'n-of-N' ? `${n} / ${total}` : fmt === 'page-n' ? `Page ${n}` : String(n);
}

export const isStdFont = (f: string) => f === 'sans' || f === 'serif' || f === 'mono';
/** CSS font-family for a text annotation. */
export const annotFont = (f: string) => FONT_CSS[f] ?? `"${f}", "IBM Plex Sans Arabic", sans-serif`;
