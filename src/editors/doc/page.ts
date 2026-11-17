import type { DocPageSetup, PaperSize } from '../../types';

/** Paper sizes in CSS px (96 dpi) and in twips (Word's unit). */
export const PAPER: Record<PaperSize, { w: number; h: number; tw: number; th: number }> = {
  A4: { w: 794, h: 1123, tw: 11906, th: 16838 },
  Letter: { w: 816, h: 1056, tw: 12240, th: 15840 },
  Legal: { w: 816, h: 1344, tw: 12240, th: 20160 },
  A5: { w: 559, h: 794, tw: 8391, th: 11906 },
};
export const PAGE = PAPER.A4;
export const MARGIN_PX: Record<DocPageSetup['margin'], number> = { narrow: 48, normal: 96, wide: 144 };

/** Page size in px after orientation. */
export function pageSize(p: DocPageSetup) {
  const s = PAPER[p.size ?? 'A4'];
  return p.orientation === 'landscape' ? { w: s.h, h: s.w } : { w: s.w, h: s.h };
}

/** Fill in table-of-contents blocks for output that has no live editor (print, PDF). */
export function expandToc(html: string): string {
  if (!html.includes('data-toc')) return html;
  const dom = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  const heads = Array.from(dom.body.querySelectorAll('h1, h2, h3, h4')).filter((h) => h.textContent?.trim());
  const ar = document.documentElement.lang === 'ar';
  dom.body.querySelectorAll('[data-toc]').forEach((toc) => {
    toc.className = 'toc';
    toc.innerHTML = `<div class="toc-title">${ar ? 'المحتويات' : 'Contents'}</div>`;
    for (const h of heads) {
      const row = dom.createElement('div');
      row.className = `toc-row l${h.tagName.slice(1)}`;
      row.textContent = h.textContent;
      toc.appendChild(row);
    }
  });
  return dom.body.innerHTML;
}
