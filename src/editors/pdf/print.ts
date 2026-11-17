import { shownSize, type PdfDoc } from './model';
import { renderPageCanvas } from './render';

/** Print a PDF document: each page rendered to an image, one image per printed page. */
export async function printPdfDoc(doc: PdfDoc) {
  const root = document.createElement('div');
  root.id = 'print-root';
  root.className = 'print-pdf';
  document.body.appendChild(root);
  const first = shownSize(doc.pages[0]);
  try {
    for (let i = 0; i < doc.pages.length; i++) {
      const p = doc.pages[i];
      const c = await renderPageCanvas(doc, p, 2, { annots: true, decor: { n: i + 1, total: doc.pages.length } });
      const img = new Image();
      img.src = c.toDataURL('image/jpeg', 0.92);
      img.className = 'print-pdf-page';
      const s = shownSize(p);
      img.style.aspectRatio = `${s.w} / ${s.h}`;
      root.appendChild(img);
      await img.decode().catch(() => null);
    }
    const style = document.createElement('style');
    style.textContent = `@page { size: ${first.w}pt ${first.h}pt; margin: 0; }`;
    document.head.appendChild(style);
    document.body.classList.add('printing');
    await new Promise<void>((resolve) => {
      let finished = false;
      const done = () => {
        if (finished) return;
        finished = true;
        style.remove();
        document.body.classList.remove('printing');
        window.removeEventListener('afterprint', done);
        resolve();
      };
      window.addEventListener('afterprint', done);
      window.print();
      setTimeout(done, 1000);
    });
  } finally {
    root.remove();
  }
}
