import { useEffect, useRef, type CSSProperties } from 'react';
import { createPortal } from 'react-dom';
import type { OpenFile } from '../types';
import { expandToc, MARGIN_PX, PAGE, pageSize } from '../editors/doc/page';
import { Engine, formatValue } from '../editors/sheet/formula';
import { addr, colName, DEFAULT_COL_W, parseAddr, type Sheet, type Workbook } from '../editors/sheet/model';
import { SlideView } from '../editors/slides/SlideView';

export type PrintMode = 'print' | 'pdf';
/** Files printed through HTML. PDFs print their own rendered pages (editors/pdf/print). */
export type Printable = Exclude<OpenFile, { kind: 'pdf' }>;

interface Geometry {
  pageW: number; // px
  pageH: number;
  margin: number;
  landscape: boolean;
}

function geometry(f: Printable): Geometry {
  if (f.kind === 'doc') {
    const landscape = f.data.page.orientation === 'landscape';
    const s = pageSize(f.data.page);
    return { pageW: s.w, pageH: s.h, margin: MARGIN_PX[f.data.page.margin], landscape };
  }
  if (f.kind === 'slides') return { pageW: f.data.w, pageH: f.data.h, margin: 0, landscape: true };
  // sheets: landscape when wide
  const sh = f.data.sheets[f.data.active];
  const used = usedCols(sh);
  let w = 0;
  for (let c = 0; c <= used; c++) w += sh.colW[c] ?? DEFAULT_COL_W;
  const landscape = w > PAGE.w - 96;
  return { pageW: landscape ? PAGE.h : PAGE.w, pageH: landscape ? PAGE.w : PAGE.h, margin: 48, landscape };
}

function usedCols(sh: Sheet) {
  let c2 = 0;
  for (const k of Object.keys(sh.cells)) if (sh.cells[k].v !== '') c2 = Math.max(c2, parseAddr(k)!.c);
  return c2;
}
function usedRows(sh: Sheet) {
  let r2 = 0;
  for (const k of Object.keys(sh.cells)) if (sh.cells[k].v !== '') r2 = Math.max(r2, parseAddr(k)!.r);
  return r2;
}

function SheetTable({ sheet, wb }: { sheet: Sheet; wb: Workbook }) {
  const engine = new Engine(wb);
  const r2 = usedRows(sheet);
  const c2 = usedCols(sheet);
  return (
    <table className="print-sheet" dir="ltr">
      <colgroup>
        {Array.from({ length: c2 + 1 }, (_, c) => (
          <col key={c} style={{ width: sheet.colW[c] ?? DEFAULT_COL_W }} />
        ))}
      </colgroup>
      <tbody>
        {Array.from({ length: r2 + 1 }, (_, r) => (
          <tr key={r} data-row={r}>
            {Array.from({ length: c2 + 1 }, (_, c) => {
              const cell = sheet.cells[addr(r, c)];
              const v = cell ? engine.cellValue(sheet, r, c) : null;
              const st = cell?.st ?? {};
              const style: CSSProperties = {
                fontWeight: st.b ? 700 : undefined,
                fontStyle: st.i ? 'italic' : undefined,
                textDecoration: [st.u && 'underline', st.s && 'line-through'].filter(Boolean).join(' ') || undefined,
                color: st.color,
                background: st.bg,
                textAlign: st.align ?? (typeof v === 'number' ? 'right' : 'left'),
                fontSize: st.fontSize ? `${st.fontSize}pt` : undefined,
                fontFamily: st.font ? `"${st.font}", "IBM Plex Sans Arabic", sans-serif` : undefined,
                whiteSpace: st.wrap ? 'normal' : 'nowrap',
                border: st.border ? '1px solid #94a3b8' : undefined,
              };
              return (
                <td key={c} style={style} title={`${colName(c)}${r + 1}`} dir="auto">
                  {formatValue(v, st.fmt, st.dec)}
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function PrintContent({ file, g }: { file: Printable; g: Geometry }) {
  if (file.kind === 'doc')
    return (
      <div className="print-doc doc-content ProseMirror" style={{ width: g.pageW - 2 * g.margin }} dangerouslySetInnerHTML={{ __html: expandToc(file.data.html) }} />
    );
  if (file.kind === 'sheet')
    return (
      <div className="print-sheet-wrap" style={{ width: g.pageW - 2 * g.margin }}>
        <SheetTable sheet={file.data.sheets[file.data.active]} wb={file.data} />
      </div>
    );
  return (
    <div className="print-slides">
      {file.data.slides.map((s) => (
        <div key={s.id} className="print-slide" style={{ width: file.data.w, height: file.data.h }}>
          <SlideView slide={s} w={file.data.w} h={file.data.h} />
        </div>
      ))}
    </div>
  );
}

async function settle(root: HTMLElement) {
  await document.fonts?.ready;
  await Promise.all(Array.from(root.querySelectorAll('img')).map((img) => (img.complete ? null : img.decode().catch(() => null))));
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
}

/** Break offsets (px from container top) at block boundaries, so pages never cut a line in half. */
function pageBreaks(container: HTMLElement, contentH: number, kind: Printable['kind']): number[] {
  const top = container.getBoundingClientRect().top;
  const blocks: { y: number; bottom: number; force: boolean }[] = [];
  const els = kind === 'sheet' ? Array.from(container.querySelectorAll('tr')) : Array.from(container.children);
  for (const el of els) {
    const r = el.getBoundingClientRect();
    blocks.push({ y: r.top - top, bottom: r.bottom - top, force: el.classList.contains('page-break') });
  }
  const total = container.scrollHeight;
  const breaks = [0];
  let pageStart = 0;
  for (const b of blocks) {
    if (b.force) {
      breaks.push(b.bottom);
      pageStart = b.bottom;
      continue;
    }
    if (b.bottom - pageStart > contentH) {
      // block doesn't fit: break before it, or slice it if it alone is taller than a page
      let start = b.y > pageStart ? b.y : pageStart + contentH;
      breaks.push(start);
      pageStart = start;
      while (b.bottom - pageStart > contentH) {
        start = pageStart + contentH;
        breaks.push(start);
        pageStart = start;
      }
    }
  }
  if (total > pageStart + 1) breaks.push(total);
  else if (breaks.length === 1) breaks.push(total);
  return breaks;
}

export async function runPdf(root: HTMLElement, file: Printable): Promise<Blob> {
  const [{ jsPDF }, { default: html2canvas }] = await Promise.all([import('jspdf'), import('html2canvas-pro')]);
  await settle(root);
  const g = geometry(file);
  const pdf = new jsPDF({ orientation: g.landscape ? 'landscape' : 'portrait', unit: 'px', format: [g.pageW, g.pageH], hotfixes: ['px_scaling'], compress: true });
  const scale = 2;
  if (file.kind === 'slides') {
    const slides = Array.from(root.querySelectorAll<HTMLElement>('.print-slide'));
    for (let i = 0; i < slides.length; i++) {
      const canvas = await html2canvas(slides[i], { scale, backgroundColor: null, useCORS: true, logging: false });
      if (i > 0) pdf.addPage([g.pageW, g.pageH], 'landscape');
      pdf.addImage(canvas, 'JPEG', 0, 0, g.pageW, g.pageH, undefined, 'FAST');
    }
    return pdf.output('blob');
  }
  const content = root.firstElementChild as HTMLElement;
  const contentH = g.pageH - 2 * g.margin;
  const inner = file.kind === 'sheet' ? (content.querySelector('table') as HTMLElement) : content;
  const breaks = pageBreaks(inner, contentH, file.kind);
  const setup = file.kind === 'doc' ? file.data.page : null;
  const paper = setup?.color || '#ffffff';
  const canvas = await html2canvas(content, { scale, backgroundColor: paper, useCORS: true, logging: false });
  const total = breaks.length - 1;
  let n = 0;
  for (let i = 0; i < total; i++) {
    const y0 = breaks[i];
    const h = Math.min(contentH, breaks[i + 1] - y0);
    if (h <= 0) continue;
    n++;
    const slice = document.createElement('canvas');
    slice.width = canvas.width;
    slice.height = Math.ceil(h * scale);
    const ctx = slice.getContext('2d')!;
    ctx.fillStyle = paper;
    ctx.fillRect(0, 0, slice.width, slice.height);
    ctx.drawImage(canvas, 0, Math.floor(y0 * scale), canvas.width, slice.height, 0, 0, canvas.width, slice.height);
    if (i > 0) pdf.addPage([g.pageW, g.pageH], g.landscape ? 'landscape' : 'portrait');
    if (paper !== '#ffffff') {
      pdf.setFillColor(paper);
      pdf.rect(0, 0, g.pageW, g.pageH, 'F');
    }
    pdf.addImage(slice, 'JPEG', g.margin, g.margin, canvas.width / scale, h, undefined, 'FAST');
    if (setup) {
      // Header, footer and page numbers sit in the margins, drawn as images so any script works.
      const band = (text: string, y: number, align: 'center' | 'right') => {
        const img = textLine(text, 12, '#475569');
        const w = img.width / 4, hh = img.height / 4;
        const x = align === 'center' ? (g.pageW - w) / 2 : g.pageW - g.margin - w;
        pdf.addImage(img, 'PNG', x, y - hh / 2, w, hh);
      };
      const top = g.margin / 2, bottom = g.pageH - g.margin / 2;
      if (setup.header) band(setup.header, top, 'center');
      if (setup.footer) band(setup.footer, bottom, 'center');
      const pos = setup.pageNumbers === 'bottom-center' && setup.footer ? 'bottom-right' : setup.pageNumbers;
      if (pos) band(String(n), pos === 'top-right' ? top : bottom, pos === 'bottom-center' ? 'center' : 'right');
    }
  }
  return pdf.output('blob');
}

/** One line of text drawn at 4x for crisp header/footer images. */
function textLine(text: string, px: number, color: string) {
  const font = `${px}px Inter, "IBM Plex Sans Arabic", sans-serif`;
  const m = document.createElement('canvas').getContext('2d')!;
  m.font = font;
  const w = Math.ceil(m.measureText(text).width) + 4;
  const c = document.createElement('canvas');
  c.width = w * 4;
  c.height = Math.ceil(px * 1.5) * 4;
  const ctx = c.getContext('2d')!;
  ctx.scale(4, 4);
  ctx.font = font;
  ctx.fillStyle = color;
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';
  ctx.fillText(text, w / 2, (px * 1.5) / 2);
  return c;
}

const cssString = (s: string) => `"${s.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, ' ')}"`;

export function runPrint(file: Printable) {
  const g = geometry(file);
  const style = document.createElement('style');
  const size = file.kind === 'sheet' ? `A4 ${g.landscape ? 'landscape' : 'portrait'}` : `${g.pageW}px ${g.pageH}px`;
  let boxes = '';
  if (file.kind === 'doc') {
    const p = file.data.page;
    const box = (where: string, content: string) => `@${where} { content: ${content}; font: 9pt Inter, "IBM Plex Sans Arabic", sans-serif; color: #475569; }`;
    if (p.header) boxes += box('top-center', cssString(p.header));
    if (p.footer) boxes += box('bottom-center', cssString(p.footer));
    const pos = p.pageNumbers === 'bottom-center' && p.footer ? 'bottom-right' : p.pageNumbers;
    if (pos) boxes += box(pos, 'counter(page)');
    if (p.color) boxes += `background: ${p.color};`;
  }
  style.textContent = `@page { size: ${size}; margin: ${file.kind === 'slides' ? 0 : g.margin + 'px'}; ${boxes} }`;
  document.head.appendChild(style);
  document.body.classList.add('printing');
  const done = () => {
    style.remove();
    document.body.classList.remove('printing');
    window.removeEventListener('afterprint', done);
  };
  window.addEventListener('afterprint', done);
  window.print();
  setTimeout(done, 1000);
}

/** Mounts the printable rendering offscreen, runs the job, then calls onDone. */
export function PrintJob({ file, mode, onDone }: { file: Printable; mode: PrintMode; onDone: (pdf?: Blob, err?: unknown) => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const g = geometry(file);
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        if (!ref.current) return;
        if (mode === 'pdf') {
          const blob = await runPdf(ref.current, file);
          if (!cancelled) onDone(blob);
        } else {
          await settle(ref.current);
          runPrint(file);
          if (!cancelled) onDone();
        }
      } catch (e) {
        if (!cancelled) onDone(undefined, e);
      }
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  return createPortal(
    <div id="print-root" ref={ref} className={`print-${file.kind}`} aria-hidden>
      <PrintContent file={file} g={g} />
    </div>,
    document.body,
  );
}
