import { A4, blankPage, pid, type PdfPage } from './model';

function readAsDataUrl(f: File): Promise<string> {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result));
    r.onerror = () => rej(r.error);
    r.readAsDataURL(f);
  });
}

function size(src: string): Promise<{ w: number; h: number }> {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => res({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = rej;
    img.src = src;
  });
}

/** Re-encode anything that is not PNG/JPEG (webp, gif, bmp…) as PNG so pdf-lib can embed it. */
async function normalize(src: string): Promise<string> {
  if (/^data:image\/(png|jpe?g)/i.test(src)) return src;
  const img = new Image();
  img.src = src;
  await img.decode();
  const c = document.createElement('canvas');
  c.width = img.naturalWidth;
  c.height = img.naturalHeight;
  c.getContext('2d')!.drawImage(img, 0, 0);
  return c.toDataURL('image/png');
}

/** One page per image, sized to the image's shape and fitted inside A4. */
export async function imagesToPages(files: File[]): Promise<PdfPage[]> {
  const out: PdfPage[] = [];
  for (const f of files) {
    const src = await normalize(await readAsDataUrl(f));
    const { w, h } = await size(src);
    const landscape = w > h;
    const box = landscape ? { w: A4.h, h: A4.w } : A4;
    const k = Math.min(box.w / w, box.h / h);
    const pw = Math.round(w * k * 100) / 100;
    const ph = Math.round(h * k * 100) / 100;
    const page = blankPage(pw, ph);
    page.annots.push({ id: pid(), kind: 'image', x: 0, y: 0, w: pw, h: ph, src });
    out.push(page);
  }
  return out;
}
