import { FONT_CSS, pageLabel, rotTransform, shownSize, type Annot, type PageDecor, type PdfDoc, type PdfPage, type TextAnnot } from './model';
import { sourcePage } from './pdfjs';

/** Draw the annotations of a page in page coordinates (ctx already scaled). */
export async function drawAnnots(ctx: CanvasRenderingContext2D, annots: Annot[]) {
  for (const a of annots) {
    ctx.save();
    switch (a.kind) {
      case 'rect':
      case 'ellipse':
      case 'highlight':
      case 'whiteout':
      case 'redact': {
        const fill = a.kind === 'highlight' ? a.fill ?? '#fde047' : a.kind === 'whiteout' ? a.fill ?? '#ffffff' : a.kind === 'redact' ? '#000000' : a.fill;
        ctx.globalAlpha = a.kind === 'highlight' ? a.opacity ?? 0.4 : a.opacity ?? 1;
        if (a.kind === 'highlight') ctx.globalCompositeOperation = 'multiply';
        ctx.beginPath();
        if (a.kind === 'ellipse') ctx.ellipse(a.x + a.w / 2, a.y + a.h / 2, Math.abs(a.w / 2), Math.abs(a.h / 2), 0, 0, Math.PI * 2);
        else ctx.rect(a.x, a.y, a.w, a.h);
        if (fill && fill !== 'none') {
          ctx.fillStyle = fill;
          ctx.fill();
        }
        if (a.stroke && (a.kind === 'rect' || a.kind === 'ellipse')) {
          ctx.globalAlpha = 1;
          ctx.strokeStyle = a.stroke;
          ctx.lineWidth = a.width ?? 2;
          ctx.stroke();
        }
        break;
      }
      case 'line':
      case 'arrow': {
        ctx.strokeStyle = a.color;
        ctx.fillStyle = a.color;
        ctx.lineWidth = a.width;
        ctx.lineCap = 'round';
        ctx.beginPath();
        ctx.moveTo(a.x1, a.y1);
        ctx.lineTo(a.x2, a.y2);
        ctx.stroke();
        if (a.kind === 'arrow') {
          const [p1, p2, p3] = arrowHead(a.x1, a.y1, a.x2, a.y2, a.width);
          ctx.beginPath();
          ctx.moveTo(...p1);
          ctx.lineTo(...p2);
          ctx.lineTo(...p3);
          ctx.closePath();
          ctx.fill();
        }
        break;
      }
      case 'ink': {
        ctx.strokeStyle = a.color;
        ctx.lineWidth = a.width;
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
        ctx.globalAlpha = a.opacity ?? 1;
        if (a.opacity && a.opacity < 1) ctx.globalCompositeOperation = 'multiply';
        ctx.stroke(new Path2D(inkPath(a.points)));
        break;
      }
      case 'image': {
        const img = await loadImage(a.src);
        ctx.drawImage(img, a.x, a.y, a.w, a.h);
        break;
      }
      case 'text':
        drawTextBox(ctx, a);
        break;
      case 'note': {
        ctx.fillStyle = a.color;
        roundRect(ctx, a.x, a.y, 20, 20, 3);
        ctx.fill();
        ctx.fillStyle = 'rgba(0,0,0,0.55)';
        ctx.fillRect(a.x + 5, a.y + 6, 10, 1.5);
        ctx.fillRect(a.x + 5, a.y + 9.5, 10, 1.5);
        ctx.fillRect(a.x + 5, a.y + 13, 7, 1.5);
        break;
      }
      case 'stamp': {
        ctx.strokeStyle = a.color;
        ctx.fillStyle = a.color;
        ctx.lineWidth = 2.5;
        ctx.globalAlpha = 0.9;
        roundRect(ctx, a.x, a.y, a.w, a.h, 6);
        ctx.stroke();
        ctx.font = `700 ${Math.min(a.h * 0.5, (a.w / Math.max(4, a.label.length)) * 1.5)}px Helvetica, Arial, "IBM Plex Sans Arabic", sans-serif`;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(a.label, a.x + a.w / 2, a.y + a.h / 2 + 1);
        break;
      }
    }
    ctx.restore();
  }
}

export function drawDecor(ctx: CanvasRenderingContext2D, decor: PageDecor, w: number, h: number, n: number, total: number) {
  ctx.save();
  if (decor.watermark?.text) {
    const wm = decor.watermark;
    ctx.globalAlpha = wm.opacity;
    ctx.fillStyle = wm.color;
    ctx.font = `700 ${wm.size}px Helvetica, Arial, "IBM Plex Sans Arabic", sans-serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.translate(w / 2, h / 2);
    ctx.rotate(-Math.atan2(h, w));
    ctx.fillText(wm.text, 0, 0);
  }
  ctx.restore();
  ctx.save();
  ctx.fillStyle = '#475569';
  ctx.font = '10px Helvetica, Arial, "IBM Plex Sans Arabic", sans-serif';
  ctx.textBaseline = 'middle';
  if (decor.header) {
    ctx.textAlign = 'center';
    ctx.fillText(decor.header, w / 2, 24);
  }
  if (decor.footer) {
    ctx.textAlign = 'center';
    ctx.fillText(decor.footer, w / 2, h - 24);
  }
  if (decor.pageNumbers) {
    const pn = decor.pageNumbers;
    const label = pageLabel(n, total, pn.fmt);
    ctx.font = `${pn.size}px Helvetica, Arial, sans-serif`;
    const pos = decorPos(pn.pos, w, h);
    ctx.textAlign = pos.align;
    ctx.fillText(label, pos.x, pos.y);
  }
  ctx.restore();
}

export function decorPos(pos: NonNullable<PageDecor['pageNumbers']>['pos'], w: number, h: number): { x: number; y: number; align: CanvasTextAlign } {
  switch (pos) {
    case 'bottom-right':
      return { x: w - 36, y: h - 24, align: 'right' };
    case 'bottom-left':
      return { x: 36, y: h - 24, align: 'left' };
    case 'top-right':
      return { x: w - 36, y: 24, align: 'right' };
    default:
      return { x: w / 2, y: h - 24, align: 'center' };
  }
}

export function inkPath(pts: number[]): string {
  if (pts.length < 2) return '';
  let d = `M${pts[0].toFixed(1)} ${pts[1].toFixed(1)}`;
  if (pts.length === 2) return d + ` l0.01 0`;
  for (let i = 2; i < pts.length - 2; i += 2) {
    const mx = (pts[i] + pts[i + 2]) / 2;
    const my = (pts[i + 1] + pts[i + 3]) / 2;
    d += ` Q${pts[i].toFixed(1)} ${pts[i + 1].toFixed(1)} ${mx.toFixed(1)} ${my.toFixed(1)}`;
  }
  d += ` L${pts[pts.length - 2].toFixed(1)} ${pts[pts.length - 1].toFixed(1)}`;
  return d;
}

export function arrowHead(x1: number, y1: number, x2: number, y2: number, w: number): [[number, number], [number, number], [number, number]] {
  const ang = Math.atan2(y2 - y1, x2 - x1);
  const len = Math.max(8, w * 4);
  const spread = 0.45;
  return [
    [x2, y2],
    [x2 - len * Math.cos(ang - spread), y2 - len * Math.sin(ang - spread)],
    [x2 - len * Math.cos(ang + spread), y2 - len * Math.sin(ang + spread)],
  ];
}

function roundRect(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, r: number) {
  ctx.beginPath();
  ctx.roundRect(x, y, w, h, r);
}

const imgCache = new Map<string, Promise<HTMLImageElement>>();
export function loadImage(src: string): Promise<HTMLImageElement> {
  let p = imgCache.get(src);
  if (!p) {
    p = new Promise((res, rej) => {
      const i = new Image();
      i.onload = () => res(i);
      i.onerror = rej;
      i.src = src;
    });
    imgCache.set(src, p);
  }
  return p;
}

const RTL = /[֐-ࣿיִ-﷿ﹰ-﻿]/;
const STRONG = /[A-Za-zÀ-ɏͰ-ӿ֐-ࣿיִ-﷿ﹰ-﻿]/;
/** Direction from the first strong character, like dir="auto" and unicode-bidi: plaintext. */
export const isRtl = (s: string) => {
  const m = STRONG.exec(s);
  return !!m && RTL.test(m[0]);
};

/** Wrap and draw a text box. Used for previews, flattening and non-Latin text in saved PDFs. */
export function drawTextBox(ctx: CanvasRenderingContext2D, a: TextAnnot) {
  ctx.save();
  if (a.bg) {
    ctx.fillStyle = a.bg;
    ctx.fillRect(a.x, a.y, a.w, a.h);
  }
  ctx.fillStyle = a.color;
  ctx.font = `${a.italic ? 'italic ' : ''}${a.bold ? '700 ' : ''}${a.size}px ${FONT_CSS[a.font]}`;
  ctx.textBaseline = 'top';
  const lines = wrapLines(ctx, a.text, a.w - 4);
  const lh = a.size * 1.25;
  lines.forEach((line, i) => {
    const rtl = isRtl(line);
    ctx.direction = rtl ? 'rtl' : 'ltr';
    const align = a.align ?? (rtl ? 'right' : 'left');
    ctx.textAlign = align;
    const x = align === 'center' ? a.x + a.w / 2 : align === 'right' ? a.x + a.w - 2 : a.x + 2;
    ctx.fillText(line, x, a.y + 2 + i * lh);
  });
  ctx.restore();
}

export function wrapLines(ctx: CanvasRenderingContext2D, text: string, maxW: number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    const words = para.split(/(\s+)/);
    let line = '';
    for (const w of words) {
      const test = line + w;
      if (line && ctx.measureText(test).width > maxW && w.trim()) {
        out.push(line.trimEnd());
        line = w.trimStart();
      } else line = test;
    }
    out.push(line);
  }
  return out;
}

/** Render a whole page (PDF content + annotations + decorations) into a canvas. */
export async function renderPageCanvas(doc: PdfDoc, page: PdfPage, scale: number, opts: { annots?: boolean; decor?: { n: number; total: number } } = {}): Promise<HTMLCanvasElement> {
  const shown = shownSize(page);
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(shown.w * scale);
  canvas.height = Math.round(shown.h * scale);
  const ctx = canvas.getContext('2d')!;
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  const pg = await sourcePage(doc, page);
  if (pg) {
    const vp = pg.getViewport({ scale, rotation: (page.srcRot + page.rot) % 360 });
    await pg.render({ canvas, canvasContext: ctx, viewport: vp }).promise;
  }
  if (opts.annots !== false) {
    ctx.save();
    ctx.scale(scale, scale);
    ctx.transform(...rotTransform(page));
    await drawAnnots(ctx, page.annots);
    ctx.restore();
  }
  if (opts.decor) {
    ctx.save();
    ctx.scale(scale, scale);
    drawDecor(ctx, doc.decor, shown.w, shown.h, opts.decor.n, opts.decor.total);
    ctx.restore();
  }
  return canvas;
}
