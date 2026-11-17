import * as pdfjs from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import type { PDFDocumentProxy, PDFPageProxy } from 'pdfjs-dist';
import { b64ToBytes, pid, type PdfDoc, type PdfPage } from './model';

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

export { pdfjs };

const docs = new Map<string, Promise<PDFDocumentProxy>>();

/** pdf.js document for a source, loaded once. */
export function sourceDoc(doc: PdfDoc, src: string): Promise<PDFDocumentProxy> {
  let p = docs.get(src);
  if (!p) {
    p = pdfjs.getDocument({ data: b64ToBytes(doc.sources[src]) }).promise;
    docs.set(src, p);
  }
  return p;
}

export async function sourcePage(doc: PdfDoc, page: PdfPage): Promise<PDFPageProxy | null> {
  if (!page.src) return null;
  const d = await sourceDoc(doc, page.src);
  return d.getPage(page.index + 1);
}

/** Read a PDF file into pages. Throws "password" for encrypted files. */
export async function importPdf(bytes: Uint8Array): Promise<{ id: string; b64: string; pages: PdfPage[]; title?: string }> {
  const id = pid();
  const { bytesToB64 } = await import('./model');
  const b64 = bytesToB64(bytes);
  const task = pdfjs.getDocument({ data: bytes.slice() });
  let needsPassword = false;
  task.onPassword = () => {
    needsPassword = true;
    task.destroy();
  };
  let d: PDFDocumentProxy;
  try {
    d = await task.promise;
  } catch (e) {
    if (needsPassword || (e as { name?: string })?.name === 'PasswordException') throw new Error('password');
    throw e;
  }
  docs.set(id, Promise.resolve(d));
  const pages: PdfPage[] = [];
  for (let i = 0; i < d.numPages; i++) {
    const pg = await d.getPage(i + 1);
    const vp = pg.getViewport({ scale: 1 });
    pages.push({ id: pid(), src: id, index: i, w: vp.width, h: vp.height, srcRot: pg.rotate, rot: 0, annots: [] });
  }
  let title: string | undefined;
  try {
    const meta = await d.getMetadata();
    title = (meta.info as { Title?: string })?.Title || undefined;
  } catch {
    /* no metadata */
  }
  return { id, b64, pages, title };
}

export interface TextItem {
  str: string;
  x: number;
  y: number; // top
  w: number;
  h: number;
  size: number;
  font: string;
  bold: boolean;
  italic: boolean;
  /** Unit vector of the reading direction on the page (1,0 for normal horizontal text). */
  ux: number;
  uy: number;
}

const textCache = new Map<string, Promise<TextItem[]>>();

/** Text runs of a page in page coordinates (top-left origin, points). */
export function pageText(doc: PdfDoc, page: PdfPage): Promise<TextItem[]> {
  if (!page.src) return Promise.resolve([]);
  const key = `${page.src}:${page.index}`;
  let p = textCache.get(key);
  if (!p) {
    p = (async () => {
      const pg = await sourcePage(doc, page);
      if (!pg) return [];
      const vp = pg.getViewport({ scale: 1 });
      const tc = await pg.getTextContent();
      const styles = tc.styles as Record<string, { fontFamily?: string }>;
      const out: TextItem[] = [];
      for (const it of tc.items) {
        if (!('str' in it) || !it.str.trim()) continue;
        const tx = pdfjs.Util.transform(vp.transform, it.transform);
        const size = Math.hypot(tx[2], tx[3]);
        const fontName = String(it.fontName ?? '');
        const family = styles[it.fontName]?.fontFamily ?? '';
        // Text may run in any direction (rotated pages, vertical labels): box the four corners.
        const sx = Math.hypot(tx[0], tx[1]) || 1;
        const ux = tx[0] / sx, uy = tx[1] / sx;
        const len = it.width * vp.scale;
        const vx = tx[2] / size, vy = tx[3] / size; // "up" for the glyphs
        const ox = tx[4], oy = tx[5];
        const xs = [ox, ox + ux * len, ox + vx * size, ox + ux * len + vx * size];
        const ys = [oy, oy + uy * len, oy + vy * size, oy + uy * len + vy * size];
        const x1 = Math.min(...xs), x2 = Math.max(...xs), y1 = Math.min(...ys), y2 = Math.max(...ys);
        const pad = size * 0.15;
        const horiz = Math.abs(ux) >= Math.abs(uy);
        out.push({
          str: it.str,
          x: horiz ? x1 : x1 - pad / 2,
          y: horiz ? y1 - pad / 2 : y1,
          w: horiz ? x2 - x1 : x2 - x1 + pad,
          h: horiz ? y2 - y1 + pad : y2 - y1,
          size,
          ux,
          uy,
          font: /serif/i.test(family) && !/sans/i.test(family) ? 'serif' : /mono|courier/i.test(family + fontName) ? 'mono' : 'sans',
          bold: /bold|black|heavy/i.test(fontName),
          italic: /italic|oblique/i.test(fontName),
        });
      }
      return out;
    })();
    textCache.set(key, p);
  }
  return p;
}
