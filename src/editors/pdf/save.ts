import {
  BlendMode, PDFArray, PDFDocument, PDFFont, PDFHexString, PDFName, PDFPage, PDFRef, StandardFonts, concatTransformationMatrix, degrees,
  popGraphicsState, pushGraphicsState, rgb, type PDFImage,
} from 'pdf-lib';
import JSZip from 'jszip';
import { b64ToBytes, isStdFont, pageLabel, rotTransform, shownSize, type Annot, type PdfDoc, type PdfPage, type TextAnnot } from './model';
import { pdfjs, sourcePage } from './pdfjs';
import { arrowHead, decorPos, drawTextBox, inkPath, isRtl, loadImage, renderPageCanvas } from './render';

type M = [number, number, number, number, number, number];

const mul = (m1: M, m2: M): M => pdfjs.Util.transform(m1, m2) as M;
const inv = (m: M): M => pdfjs.Util.inverseTransform(m) as M;
const apply = (p: [number, number], m: M): [number, number] => [m[0] * p[0] + m[2] * p[1] + m[4], m[1] * p[0] + m[3] * p[1] + m[5]];

function color(hex?: string) {
  const h = (hex ?? '#000000').replace('#', '');
  const v = h.length === 3 ? h.split('').map((c) => c + c).join('') : h.padEnd(6, '0');
  return rgb(parseInt(v.slice(0, 2), 16) / 255, parseInt(v.slice(2, 4), 16) / 255, parseInt(v.slice(4, 6), 16) / 255);
}

/** Text pdf-lib's standard fonts can encode (WinAnsi). */
const latinOnly = (s: string) => /^[\x20-\x7E\xA0-\xFF\n\r\t’‘“”–—•…€]*$/.test(s);

interface Fonts {
  get(font: TextAnnot['font'], bold?: boolean, italic?: boolean): Promise<PDFFont>;
}

function fontCache(pdf: PDFDocument): Fonts {
  const cache = new Map<string, Promise<PDFFont>>();
  const pick = (f: TextAnnot['font'], b?: boolean, i?: boolean) => {
    if (f === 'serif') return b && i ? StandardFonts.TimesRomanBoldItalic : b ? StandardFonts.TimesRomanBold : i ? StandardFonts.TimesRomanItalic : StandardFonts.TimesRoman;
    if (f === 'mono') return b && i ? StandardFonts.CourierBoldOblique : b ? StandardFonts.CourierBold : i ? StandardFonts.CourierOblique : StandardFonts.Courier;
    return b && i ? StandardFonts.HelveticaBoldOblique : b ? StandardFonts.HelveticaBold : i ? StandardFonts.HelveticaOblique : StandardFonts.Helvetica;
  };
  return {
    get(f, b, i) {
      const name = pick(f, b, i);
      if (!cache.has(name)) cache.set(name, pdf.embedFont(name));
      return cache.get(name)!;
    },
  };
}

async function embedDataUrl(pdf: PDFDocument, src: string): Promise<PDFImage> {
  if (src.startsWith('data:image/jpeg') || src.startsWith('data:image/jpg')) return pdf.embedJpg(b64ToBytes(src.split(',')[1]));
  if (src.startsWith('data:image/png')) return pdf.embedPng(b64ToBytes(src.split(',')[1]));
  // other formats: convert through a canvas
  const img = await loadImage(src);
  const c = document.createElement('canvas');
  c.width = img.naturalWidth;
  c.height = img.naturalHeight;
  c.getContext('2d')!.drawImage(img, 0, 0);
  return pdf.embedPng(b64ToBytes(c.toDataURL('image/png').split(',')[1]));
}

/** Rasterize a text box (for scripts the standard fonts cannot encode, such as Arabic). */
function textBoxPng(a: TextAnnot, scale = 4): string {
  const c = document.createElement('canvas');
  c.width = Math.max(1, Math.ceil(a.w * scale));
  c.height = Math.max(1, Math.ceil(a.h * scale));
  const ctx = c.getContext('2d')!;
  ctx.scale(scale, scale);
  drawTextBox(ctx, { ...a, x: 0, y: 0 });
  return c.toDataURL('image/png');
}

function wrap(font: PDFFont, size: number, text: string, maxW: number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    let line = '';
    for (const word of para.split(/(\s+)/)) {
      const test = line + word;
      if (line && font.widthOfTextAtSize(test, size) > maxW && word.trim()) {
        out.push(line.trimEnd());
        line = word.trimStart();
      } else line = test;
    }
    out.push(line);
  }
  return out;
}

/** Draws in a y-up space of height H that maps onto the page through `m`. */
async function drawAnnotsPdf(pdf: PDFDocument, page: PDFPage, fonts: Fonts, annots: Annot[], H: number, toPdf: M) {
  const Y = (y: number) => H - y;
  page.pushOperators(pushGraphicsState(), concatTransformationMatrix(...toPdf));
  for (const a of annots) {
    switch (a.kind) {
      case 'rect':
      case 'highlight':
      case 'whiteout':
      case 'redact': {
        const fill = a.kind === 'highlight' ? a.fill ?? '#fde047' : a.kind === 'whiteout' ? a.fill ?? '#ffffff' : a.kind === 'redact' ? '#000000' : a.fill;
        page.drawRectangle({
          x: a.x, y: Y(a.y + a.h), width: a.w, height: a.h,
          color: fill && fill !== 'none' ? color(fill) : undefined,
          opacity: a.kind === 'highlight' ? a.opacity ?? 0.4 : a.opacity ?? 1,
          borderColor: a.kind === 'rect' && a.stroke ? color(a.stroke) : undefined,
          borderWidth: a.kind === 'rect' && a.stroke ? a.width ?? 2 : 0,
          blendMode: a.kind === 'highlight' ? BlendMode.Multiply : undefined,
        });
        break;
      }
      case 'ellipse':
        page.drawEllipse({
          x: a.x + a.w / 2, y: Y(a.y + a.h / 2), xScale: Math.abs(a.w / 2), yScale: Math.abs(a.h / 2),
          color: a.fill && a.fill !== 'none' ? color(a.fill) : undefined, opacity: a.opacity ?? 1,
          borderColor: a.stroke ? color(a.stroke) : undefined, borderWidth: a.stroke ? a.width ?? 2 : 0,
        });
        break;
      case 'line':
      case 'arrow':
        page.drawLine({ start: { x: a.x1, y: Y(a.y1) }, end: { x: a.x2, y: Y(a.y2) }, thickness: a.width, color: color(a.color), lineCap: 1 as never });
        if (a.kind === 'arrow') {
          const [p1, p2, p3] = arrowHead(a.x1, a.y1, a.x2, a.y2, a.width);
          page.drawSvgPath(`M${p1[0]} ${p1[1]} L${p2[0]} ${p2[1]} L${p3[0]} ${p3[1]} Z`, { x: 0, y: H, color: color(a.color) });
        }
        break;
      case 'ink':
        page.drawSvgPath(inkPath(a.points), { x: 0, y: H, borderColor: color(a.color), borderWidth: a.width, borderLineCap: 1 as never, opacity: a.opacity, borderOpacity: a.opacity ?? 1, blendMode: a.opacity && a.opacity < 1 ? BlendMode.Multiply : undefined });
        break;
      case 'image': {
        const img = await embedDataUrl(pdf, a.src);
        page.drawImage(img, { x: a.x, y: Y(a.y + a.h), width: a.w, height: a.h });
        break;
      }
      case 'text': {
        if (a.bg) page.drawRectangle({ x: a.x, y: Y(a.y + a.h), width: a.w, height: a.h, color: color(a.bg) });
        // Device fonts and non-Latin scripts cannot use the standard PDF fonts: embed the text as an image.
        if (!latinOnly(a.text) || !isStdFont(a.font)) {
          const img = await embedDataUrl(pdf, textBoxPng({ ...a, bg: undefined }));
          page.drawImage(img, { x: a.x, y: Y(a.y + a.h), width: a.w, height: a.h });
          break;
        }
        const font = await fonts.get(a.font, a.bold, a.italic);
        const lh = a.size * 1.25;
        wrap(font, a.size, a.text, a.w - 4).forEach((line, i) => {
          const tw = font.widthOfTextAtSize(line, a.size);
          const x = a.align === 'center' ? a.x + (a.w - tw) / 2 : a.align === 'right' ? a.x + a.w - 2 - tw : a.x + 2;
          // canvas uses a top baseline; place the alphabetic baseline about 0.8em below the line top
          page.drawText(line, { x, y: Y(a.y + 2 + i * lh + a.size * 0.8), size: a.size, font, color: color(a.color) });
        });
        break;
      }
      case 'stamp': {
        const label = a.label;
        page.drawRectangle({ x: a.x, y: Y(a.y + a.h), width: a.w, height: a.h, borderColor: color(a.color), borderWidth: 2.5, opacity: 0, borderOpacity: 0.9 });
        if (latinOnly(label)) {
          const font = await fonts.get('sans', true);
          const size = Math.min(a.h * 0.5, (a.w / Math.max(4, label.length)) * 1.5);
          const tw = font.widthOfTextAtSize(label, size);
          page.drawText(label, { x: a.x + (a.w - tw) / 2, y: Y(a.y + a.h / 2) - size * 0.35, size, font, color: color(a.color), opacity: 0.9 });
        } else {
          const img = await embedDataUrl(pdf, textBoxPng({ id: '', kind: 'text', x: 0, y: 0, w: a.w, h: a.h, text: label, size: a.h * 0.45, color: a.color, font: 'sans', bold: true, align: 'center' }));
          page.drawImage(img, { x: a.x, y: Y(a.y + a.h), width: a.w, height: a.h });
        }
        break;
      }
      case 'note':
        break; // written as a real PDF comment below
    }
  }
  page.pushOperators(popGraphicsState());
}

function addNoteAnnots(pdf: PDFDocument, page: PDFPage, annots: Annot[], H: number, toPdf: M) {
  for (const a of annots) {
    if (a.kind !== 'note') continue;
    const p1 = apply([a.x, H - a.y], toPdf);
    const p2 = apply([a.x + 20, H - a.y - 20], toPdf);
    const c = color(a.color);
    const dict = pdf.context.obj({
      Type: 'Annot',
      Subtype: 'Text',
      Rect: [Math.min(p1[0], p2[0]), Math.min(p1[1], p2[1]), Math.max(p1[0], p2[0]), Math.max(p1[1], p2[1])],
      Contents: PDFHexString.fromText(a.text),
      Name: 'Comment',
      C: [c.red, c.green, c.blue],
      F: 4,
      Open: false,
    });
    const ref = pdf.context.register(dict);
    const existing = page.node.lookup(PDFName.of('Annots'));
    if (existing && 'push' in existing) (existing as unknown as { push: (r: unknown) => void }).push(ref);
    else page.node.set(PDFName.of('Annots'), pdf.context.obj([ref]));
  }
}

/** A line of text rendered to PNG (for scripts the standard PDF fonts cannot show), sized in points. */
function linePng(text: string, size: number, fill: string, bold: boolean, scale = 4) {
  const font = `${bold ? '700 ' : ''}${size}px Helvetica, Arial, "IBM Plex Sans Arabic", sans-serif`;
  const m = document.createElement('canvas').getContext('2d')!;
  m.font = font;
  const w = Math.ceil(m.measureText(text).width + 2);
  const h = Math.ceil(size * 1.3);
  const c = document.createElement('canvas');
  c.width = w * scale;
  c.height = h * scale;
  const ctx = c.getContext('2d')!;
  ctx.scale(scale, scale);
  ctx.font = font;
  ctx.fillStyle = fill;
  ctx.textBaseline = 'middle';
  ctx.direction = isRtl(text) ? 'rtl' : 'ltr';
  ctx.textAlign = 'center';
  ctx.fillText(text, w / 2, h / 2);
  return { src: c.toDataURL('image/png'), w, h };
}

async function drawDecorPdf(pdf: PDFDocument, page: PDFPage, fonts: Fonts, doc: PdfDoc, n: number, total: number, shown: { w: number; h: number }, toPdf: M) {
  const d = doc.decor;
  if (!d.watermark?.text && !d.pageNumbers && !d.header && !d.footer) return;
  const H = shown.h;
  page.pushOperators(pushGraphicsState(), concatTransformationMatrix(...toPdf));
  const font = await fonts.get('sans');
  const bold = await fonts.get('sans', true);
  const grey = '#475569';
  if (d.watermark?.text) {
    const wm = d.watermark;
    const ang = Math.atan2(shown.h, shown.w);
    const cx = shown.w / 2, cy = H - shown.h / 2;
    const cos = Math.cos(ang), sin = Math.sin(ang);
    if (latinOnly(wm.text)) {
      const tw = bold.widthOfTextAtSize(wm.text, wm.size);
      // start point so the text is centered after rotation
      const x = cx - (cos * tw) / 2 + sin * wm.size * 0.35;
      const y = cy - (sin * tw) / 2 - cos * wm.size * 0.35;
      page.drawText(wm.text, { x, y, size: wm.size, font: bold, color: color(wm.color), opacity: wm.opacity, rotate: degrees((ang * 180) / Math.PI) });
    } else {
      const im = linePng(wm.text, wm.size, wm.color, true);
      const img = await embedDataUrl(pdf, im.src);
      const x = cx - (cos * im.w) / 2 + (sin * im.h) / 2;
      const y = cy - (sin * im.w) / 2 - (cos * im.h) / 2;
      page.drawImage(img, { x, y, width: im.w, height: im.h, opacity: wm.opacity, rotate: degrees((ang * 180) / Math.PI) });
    }
  }
  const put = async (text: string, pos: { x: number; y: number; align: CanvasTextAlign }, size: number) => {
    if (latinOnly(text)) {
      const tw = font.widthOfTextAtSize(text, size);
      const x = pos.align === 'center' ? pos.x - tw / 2 : pos.align === 'right' ? pos.x - tw : pos.x;
      page.drawText(text, { x, y: H - pos.y - size * 0.35, size, font, color: color(grey) });
      return;
    }
    const im = linePng(text, size, grey, false);
    const x = pos.align === 'center' ? pos.x - im.w / 2 : pos.align === 'right' ? pos.x - im.w : pos.x;
    page.drawImage(await embedDataUrl(pdf, im.src), { x, y: H - pos.y - im.h / 2, width: im.w, height: im.h });
  };
  if (d.header) await put(d.header, { x: shown.w / 2, y: 24, align: 'center' }, 10);
  if (d.footer) await put(d.footer, { x: shown.w / 2, y: shown.h - 24, align: 'center' }, 10);
  if (d.pageNumbers) await put(pageLabel(n, total, d.pageNumbers.fmt), decorPos(d.pageNumbers.pos, shown.w, shown.h), d.pageNumbers.size);
  page.pushOperators(popGraphicsState());
}

/** pdf-lib's flatten deletes widget objects but leaves their refs in /Annots; drop those so viewers don't see broken refs. */
function dropDanglingAnnots(src: PDFDocument) {
  for (const page of src.getPages()) {
    const annots = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray);
    if (!annots) continue;
    for (let i = annots.size() - 1; i >= 0; i--) {
      const r = annots.get(i);
      if (r instanceof PDFRef && !src.context.lookup(r)) annots.remove(i);
    }
  }
}

async function loadSources(doc: PdfDoc, pages: PdfPage[]) {
  const out = new Map<string, PDFDocument | null>();
  for (const p of pages) {
    if (!p.src || out.has(p.src)) continue;
    try {
      const src = await PDFDocument.load(b64ToBytes(doc.sources[p.src]), { ignoreEncryption: true, updateMetadata: false });
      const values = doc.form[p.src];
      if (values && Object.keys(values).length) {
        const form = src.getForm();
        for (const [name, v] of Object.entries(values)) {
          try {
            const f = form.getField(name);
            const type = f.constructor.name;
            if (type === 'PDFTextField') (f as unknown as { setText: (s: string) => void }).setText(String(v));
            else if (type === 'PDFCheckBox') (v ? (f as unknown as { check: () => void }).check() : (f as unknown as { uncheck: () => void }).uncheck());
            else if (type === 'PDFDropdown' || type === 'PDFOptionList' || type === 'PDFRadioGroup') (f as unknown as { select: (s: string) => void }).select(String(v));
          } catch {
            /* skip fields that cannot be set */
          }
        }
        try {
          form.flatten();
        } catch {
          /* keep the form as is */
        }
        dropDanglingAnnots(src);
      }
      out.set(p.src, src);
    } catch {
      out.set(p.src, null); // unreadable for pdf-lib: such pages are rasterized
    }
  }
  return out;
}

/** Build the final PDF: original content, edits, comments and page decorations. */
export async function writePdf(doc: PdfDoc, pages = doc.pages): Promise<Uint8Array> {
  const out = await PDFDocument.create();
  out.setProducer('Paperline');
  out.setCreator('Paperline');
  if (doc.meta.title) out.setTitle(doc.meta.title);
  if (doc.meta.author) out.setAuthor(doc.meta.author);
  if (doc.meta.subject) out.setSubject(doc.meta.subject);
  const fonts = fontCache(out);
  const sources = await loadSources(doc, pages);
  const total = pages.length;

  for (let i = 0; i < pages.length; i++) {
    const p = pages[i];
    const shown = shownSize(p);
    const src = p.src ? sources.get(p.src) : null;
    const flatten = p.annots.some((a) => a.kind === 'redact') || (p.src && !src);
    let page: PDFPage;
    let toPdf: M;

    if (flatten) {
      // Redactions must remove what is underneath, so the page becomes a picture.
      const canvas = await renderPageCanvas(doc, p, 2.5);
      const img = await out.embedJpg(b64ToBytes(canvas.toDataURL('image/jpeg', 0.92).split(',')[1]));
      page = out.addPage([shown.w, shown.h]);
      page.drawImage(img, { x: 0, y: 0, width: shown.w, height: shown.h });
      // keep comments as real PDF notes, moved into the rotated page's space
      const notes = p.annots
        .filter((a) => a.kind === 'note')
        .map((n) => {
          const [x, y] = apply([n.x, n.y], rotTransform(p) as M);
          return { ...n, x, y };
        });
      addNoteAnnots(out, page, notes, shown.h, [1, 0, 0, 1, 0, 0]);
      await drawDecorPdf(out, page, fonts, doc, i + 1, total, shown, [1, 0, 0, 1, 0, 0]);
      continue;
    }

    if (p.src && src) {
      const [copied] = await out.copyPages(src, [p.index]);
      page = out.addPage(copied);
      const pg = await sourcePage(doc, p);
      const T = pg!.getViewport({ scale: 1 }).transform as M; // PDF → display (y-down)
      toPdf = mul(inv(T), [1, 0, 0, -1, 0, p.h]);
      page.setRotation(degrees((p.srcRot + p.rot) % 360));
    } else {
      page = out.addPage([p.w, p.h]);
      toPdf = [1, 0, 0, 1, 0, 0];
      page.setRotation(degrees(p.rot));
    }
    await drawAnnotsPdf(out, page, fonts, p.annots, p.h, toPdf);
    addNoteAnnots(out, page, p.annots, p.h, toPdf);
    // decorations live in the user-rotated (shown) space
    const shownToPage = inv(rotTransform(p) as M); // shown y-down → page y-down
    const decorM = mul(mul(mul(toPdf, [1, 0, 0, -1, 0, p.h]), shownToPage), [1, 0, 0, -1, 0, shown.h]);
    await drawDecorPdf(out, page, fonts, doc, i + 1, total, shown, decorM);
  }
  return out.save();
}

export async function pdfBlob(doc: PdfDoc, pages?: PdfPage[]) {
  const bytes = await writePdf(doc, pages);
  return new Blob([bytes as BlobPart], { type: 'application/pdf' });
}

/** One PDF per page, zipped. */
export async function splitPdf(doc: PdfDoc): Promise<Blob> {
  const zip = new JSZip();
  for (let i = 0; i < doc.pages.length; i++) zip.file(`page-${String(i + 1).padStart(3, '0')}.pdf`, await writePdf(doc, [doc.pages[i]]));
  return zip.generateAsync({ type: 'blob' });
}

/** Every page as a PNG picture, zipped. */
export async function pagesToImages(doc: PdfDoc, scale = 2): Promise<Blob> {
  const zip = new JSZip();
  for (let i = 0; i < doc.pages.length; i++) {
    const c = await renderPageCanvas(doc, doc.pages[i], scale, { decor: { n: i + 1, total: doc.pages.length } });
    zip.file(`page-${String(i + 1).padStart(3, '0')}.png`, c.toDataURL('image/png').split(',')[1], { base64: true });
  }
  return zip.generateAsync({ type: 'blob' });
}

export interface FormField {
  src: string;
  name: string;
  type: 'text' | 'checkbox' | 'choice' | 'radio' | 'other';
  value: string | boolean;
  options?: string[];
  multiline?: boolean;
}

/** Fillable form fields of every source PDF. */
export async function readFormFields(doc: PdfDoc): Promise<FormField[]> {
  const out: FormField[] = [];
  for (const [id, b64] of Object.entries(doc.sources)) {
    try {
      const pdf = await PDFDocument.load(b64ToBytes(b64), { ignoreEncryption: true, updateMetadata: false });
      for (const f of pdf.getForm().getFields()) {
        const name = f.getName();
        const t = f.constructor.name;
        const saved = doc.form[id]?.[name];
        const any = f as unknown as Record<string, () => unknown>;
        if (t === 'PDFTextField') out.push({ src: id, name, type: 'text', value: saved ?? String(any.getText() ?? ''), multiline: Boolean(any.isMultiline?.()) });
        else if (t === 'PDFCheckBox') out.push({ src: id, name, type: 'checkbox', value: saved ?? Boolean(any.isChecked()) });
        else if (t === 'PDFDropdown' || t === 'PDFOptionList') out.push({ src: id, name, type: 'choice', value: saved ?? String((any.getSelected() as string[])[0] ?? ''), options: any.getOptions() as string[] });
        else if (t === 'PDFRadioGroup') out.push({ src: id, name, type: 'radio', value: saved ?? String(any.getSelected() ?? ''), options: any.getOptions() as string[] });
      }
    } catch {
      /* no form */
    }
  }
  return out;
}

export function isLatin(s: string) {
  return latinOnly(s) && !isRtl(s);
}
