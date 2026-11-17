import JSZip from 'jszip';
import { textStyle, uid, type ChartData, type Deck, type ShapeEl, type ShapeKind, type Slide, type SlideEl, type TextEl, type TextStyle, type Transition } from './model';
import type { ChartType } from '../sheet/model';
import { CHART_COLORS } from '../sheet/features';

const EMU_PER_PX = 9525;
const PX_PER_IN = 96;

// ---------------- export ----------------
/** Our shape kinds → DrawingML preset geometry names (pptxgenjs passes these through). */
const PPTX_SHAPES: Record<ShapeKind, string> = {
  rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rtTriangle: 'rtTriangle', diamond: 'diamond',
  pentagon: 'pentagon', hexagon: 'hexagon', octagon: 'octagon', parallelogram: 'parallelogram', trapezoid: 'trapezoid',
  star4: 'star4', star5: 'star5', star8: 'star8', heart: 'heart', plus: 'plus', callout: 'wedgeRoundRectCallout', cloud: 'cloud',
  rightArrow: 'rightArrow', leftArrow: 'leftArrow', upArrow: 'upArrow', downArrow: 'downArrow', leftRightArrow: 'leftRightArrow',
  chevron: 'chevron', homePlate: 'homePlate', line: 'line', arrowLine: 'line',
};
/** Preset names we read back, including common aliases PowerPoint writes. */
const FROM_PRST: Record<string, ShapeKind> = {
  ...Object.fromEntries(Object.entries(PPTX_SHAPES).filter(([k]) => k !== 'arrowLine').map(([k, v]) => [v, k as ShapeKind])),
  mathPlus: 'plus', wedgeRectCallout: 'callout', wedgeEllipseCallout: 'callout', cloudCallout: 'cloud', flowChartProcess: 'rect',
  flowChartAlternateProcess: 'roundRect', flowChartDecision: 'diamond', flowChartConnector: 'ellipse', snip1Rect: 'rect', round2SameRect: 'roundRect',
  notchedRightArrow: 'rightArrow', stripedRightArrow: 'rightArrow', star6: 'star8', star7: 'star8', star10: 'star8', star12: 'star8', decagon: 'octagon', heptagon: 'hexagon',
};
const TRANSITION_XML: Record<Exclude<Transition, 'none'>, string> = {
  fade: '<p:fade/>', push: '<p:push dir="u"/>', wipe: '<p:wipe dir="r"/>', zoom: '<p:zoom/>', cover: '<p:cover dir="l"/>',
};

const hex = (c?: string) => (c && c.startsWith('#') ? c.slice(1).toUpperCase().padEnd(6, '0').slice(0, 6) : undefined);
const transparency = (el: SlideEl) => (el.opacity !== undefined && el.opacity < 1 ? Math.round((1 - el.opacity) * 100) : undefined);
const SHADOW = { type: 'outer', blur: 8, offset: 4, angle: 90, color: '0F172A', opacity: 0.3 };

function runsOf(el: TextEl | ShapeEl) {
  const lines = el.text.split('\n');
  const list = el.type === 'text' && (el.bullets || el.numbered);
  return lines.map((l, i) => {
    const sub = l.startsWith('  ');
    return {
      text: l.trimStart(),
      options: {
        breakLine: i < lines.length - 1,
        bullet: list ? (el.type === 'text' && el.numbered ? { type: 'number', ...(sub ? { indent: 18 } : {}) } : sub ? { indent: 18 } : true) : undefined,
        indentLevel: list && sub ? 1 : undefined,
      },
    };
  });
}

export async function writePptx(deck: Deck): Promise<Blob> {
  const { default: PptxGenJS } = await import('pptxgenjs');
  const pptx = new PptxGenJS();
  const wIn = deck.w / PX_PER_IN;
  const hIn = deck.h / PX_PER_IN;
  pptx.defineLayout({ name: 'PAPERLINE', width: wIn, height: hIn });
  pptx.layout = 'PAPERLINE';
  pptx.author = 'Paperline';
  const inch = (px: number) => px / PX_PER_IN;
  for (const s of deck.slides) {
    const slide = pptx.addSlide();
    slide.background = { color: hex(s.bg) ?? 'FFFFFF' };
    if (s.hidden) slide.hidden = true;
    if (s.notes) slide.addNotes(s.notes);
    const deco = isDarkHex(s.bg) ? 'E2E8F0' : '64748B';
    for (const el of s.els) {
      const pos = { x: inch(el.x), y: inch(el.y), w: inch(el.w), h: inch(el.h), rotate: el.rot, flipH: el.flipH, flipV: el.flipV };
      const shadow = el.shadow ? SHADOW : undefined;
      if (el.type === 'image') {
        slide.addImage({ data: el.src, ...pos, transparency: transparency(el), shadow } as never);
        continue;
      }
      if (el.type === 'table') {
        const band = tintHex(el.color, 0.86);
        const border = { type: 'solid', pt: 0.75, color: tintHex(el.color, 0.55) };
        const rows = el.rows.map((row, r) => {
          const head = el.header && r === 0;
          const striped = el.banded && !head && (r - (el.header ? 1 : 0)) % 2 === 0;
          return row.map((text) => ({
            text,
            options: { fill: { color: head ? hex(el.color) : striped ? band : 'FFFFFF' }, color: head ? 'FFFFFF' : hex(el.style.color), bold: head || el.style.bold, border },
          }));
        });
        slide.addTable(rows as never, {
          x: pos.x, y: pos.y, w: pos.w, h: pos.h, rowH: pos.h / el.rows.length, fontSize: el.style.fontSize, fontFace: el.style.fontFace,
          align: el.style.align, valign: 'middle', margin: 4,
        } as never);
        continue;
      }
      if (el.type === 'chart') {
        const scatter = el.chart === 'scatter';
        const type = { column: 'bar', bar: 'bar', line: 'line', area: 'area', pie: 'pie', doughnut: 'doughnut', scatter: 'scatter' }[el.chart];
        const series = el.data.series.map((x) => ({ name: x.name, labels: el.data.cats, values: x.values.map((v) => v ?? 0) }));
        const xs = el.data.cats.map((c, i) => (Number.isFinite(parseFloat(c)) ? parseFloat(c) : i + 1));
        const chartData = scatter ? [{ name: 'X', values: xs }, ...series.map((x) => ({ name: x.name, values: x.values }))] : series;
        slide.addChart(type as never, chartData as never, {
          ...pos, barDir: el.chart === 'bar' ? 'bar' : 'col', chartColors: CHART_COLORS.map((c) => c.slice(1)),
          showLegend: el.legend !== false && (el.chart === 'pie' || el.chart === 'doughnut' || el.data.series.length > 1), legendPos: 'b',
          showTitle: !!el.title, title: el.title || undefined, titleFontSize: 14, fontFace: 'Arial',
          showPercent: el.chart === 'pie' || el.chart === 'doughnut', showValue: false, lineSize: scatter ? 0 : 2, lineDataSymbolSize: 6,
        } as never);
        continue;
      }
      const textOpts = textProps(el.style);
      const runs = runsOf(el);
      if (el.type === 'shape') {
        const isLine = el.shape === 'line' || el.shape === 'arrowLine';
        const fill = !isLine && el.fill && el.fill !== 'transparent' ? { color: hex(el.fill), transparency: transparency(el) } : undefined;
        const dashType = el.dash === 'dash' ? 'dash' : el.dash === 'dot' ? 'sysDot' : undefined;
        const line = isLine
          ? { color: hex(el.stroke ?? el.fill), width: el.strokeW || 2, dashType, endArrowType: el.shape === 'arrowLine' ? 'triangle' : undefined }
          : el.stroke && el.strokeW ? { color: hex(el.stroke), width: el.strokeW, dashType } : undefined;
        // a line is drawn through the middle of its box
        const lpos = isLine ? { ...pos, y: inch(el.y + el.h / 2), h: 0 } : pos;
        const opts = { ...lpos, fill, line, shadow };
        if (el.text && !isLine) slide.addText(runs as never, { ...opts, ...textOpts, shape: PPTX_SHAPES[el.shape] } as never);
        else slide.addShape(PPTX_SHAPES[el.shape] as never, opts as never);
      } else {
        slide.addText((el.text ? runs : '') as never, { ...pos, ...textOpts, fill: el.fill ? { color: hex(el.fill), transparency: transparency(el) } : undefined, shadow, margin: 4 } as never);
      }
    }
    if (deck.footer) slide.addText(deck.footer, { x: inch(deck.w * 0.25), y: inch(deck.h - 34), w: inch(deck.w * 0.5), h: inch(24), fontSize: 10, color: deco, align: 'center' });
    if (deck.slideNumbers) slide.slideNumber = { x: inch(deck.w - 70), y: inch(deck.h - 34), w: inch(50), h: inch(24), fontSize: 10, color: deco, align: 'right' } as never;
  }
  const out = (await pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  const bytes = await addTransitions(out, deck);
  return new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
}

/** pptxgenjs has no transitions: write <p:transition> into each slide part ourselves. */
async function addTransitions(buf: ArrayBuffer, deck: Deck): Promise<ArrayBuffer> {
  if (!deck.slides.some((s) => s.transition && s.transition !== 'none')) return buf;
  const zip = await JSZip.loadAsync(buf);
  for (const [i, s] of deck.slides.entries()) {
    if (!s.transition || s.transition === 'none') continue;
    const path = `ppt/slides/slide${i + 1}.xml`;
    const xml = await zip.file(path)?.async('string');
    if (!xml) continue;
    const tr = `<p:transition spd="med">${TRANSITION_XML[s.transition]}</p:transition>`;
    const at = xml.indexOf('</p:clrMapOvr>');
    const next = at >= 0 ? xml.slice(0, at + 14) + tr + xml.slice(at + 14) : xml.replace('</p:cSld>', `</p:cSld>${tr}`);
    zip.file(path, next);
  }
  return zip.generateAsync({ type: 'arraybuffer' });
}

function textProps(st: TextStyle) {
  return {
    fontSize: st.fontSize,
    fontFace: st.fontFace,
    color: hex(st.color),
    bold: st.bold,
    italic: st.italic,
    underline: st.underline ? { style: 'sng' } : undefined,
    strike: st.strike ? 'sngStrike' : undefined,
    lineSpacingMultiple: st.lineSpacing,
    align: st.align,
    valign: st.valign,
  } as Record<string, unknown>;
}

function tintHex(c: string, t: number) {
  const n = parseInt((hex(c) ?? '6366F1'), 16);
  const m = (s: number) => Math.round(((n >> s) & 255) + (255 - ((n >> s) & 255)) * t);
  return ((m(16) << 16) | (m(8) << 8) | m(0)).toString(16).padStart(6, '0').toUpperCase();
}
const isDarkHex = (c: string) => bgIsDark(c);

// ---------------- import ----------------
const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
};

const parseXml = (s: string) => new DOMParser().parseFromString(s, 'application/xml');
const kids = (el: Element | null | undefined, ns: string, name: string) => (el ? Array.from(el.children).filter((c) => c.namespaceURI === ns && c.localName === name) : []);
const kid = (el: Element | null | undefined, ns: string, name: string) => kids(el, ns, name)[0] ?? null;
const desc = (el: Element | null | undefined, ns: string, name: string) => (el ? Array.from(el.getElementsByTagNameNS(ns, name)) : []);

function resolvePath(base: string, target: string) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/');
  parts.pop();
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

async function readRels(zip: JSZip, part: string): Promise<Record<string, { target: string; type: string }>> {
  const dir = part.split('/').slice(0, -1).join('/');
  const file = part.split('/').pop();
  const relsPath = `${dir}/_rels/${file}.rels`;
  const xml = await zip.file(relsPath)?.async('string');
  const out: Record<string, { target: string; type: string }> = {};
  if (!xml) return out;
  for (const r of Array.from(parseXml(xml).getElementsByTagName('Relationship'))) {
    out[r.getAttribute('Id')!] = { target: resolvePath(part, r.getAttribute('Target')!), type: r.getAttribute('Type') ?? '' };
  }
  return out;
}

interface Theme {
  colors: Record<string, string>;
  majorFont?: string;
  minorFont?: string;
}

function colorOf(el: Element | null, theme: Theme): string | undefined {
  if (!el) return undefined;
  const srgb = kid(el, NS.a, 'srgbClr');
  if (srgb) return '#' + srgb.getAttribute('val')!.toLowerCase();
  const scheme = kid(el, NS.a, 'schemeClr');
  if (scheme) {
    const v = scheme.getAttribute('val')!;
    const alias: Record<string, string> = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
    return theme.colors[alias[v] ?? v];
  }
  const sys = kid(el, NS.a, 'sysClr');
  if (sys) return '#' + (sys.getAttribute('lastClr') ?? '000000').toLowerCase();
  return undefined;
}

interface Placeholder {
  type: string;
  idx?: string;
  xfrm?: { x: number; y: number; w: number; h: number };
  size?: number;
  algn?: string;
  anchor?: string;
  bold?: boolean;
}

function xfrmOf(spPr: Element | null) {
  return readXfrm(kid(spPr, NS.a, 'xfrm'));
}
function readXfrm(x: Element | null) {
  if (!x) return undefined;
  const off = kid(x, NS.a, 'off'), ext = kid(x, NS.a, 'ext');
  if (!off || !ext) return undefined;
  return {
    x: Number(off.getAttribute('x')) / EMU_PER_PX,
    y: Number(off.getAttribute('y')) / EMU_PER_PX,
    w: Number(ext.getAttribute('cx')) / EMU_PER_PX,
    h: Number(ext.getAttribute('cy')) / EMU_PER_PX,
    rot: x.getAttribute('rot') ? Number(x.getAttribute('rot')) / 60000 : undefined,
    flipH: x.getAttribute('flipH') === '1' || undefined,
    flipV: x.getAttribute('flipV') === '1' || undefined,
  };
}

function placeholdersOf(doc: Document): Placeholder[] {
  const out: Placeholder[] = [];
  for (const sp of desc(doc.documentElement, NS.p, 'sp')) {
    const ph = desc(sp, NS.p, 'ph')[0];
    if (!ph) continue;
    const lvl1 = desc(sp, NS.a, 'lvl1pPr')[0];
    const rPr = desc(lvl1 ?? sp, NS.a, 'defRPr')[0];
    out.push({
      type: ph.getAttribute('type') ?? 'body',
      idx: ph.getAttribute('idx') ?? undefined,
      xfrm: xfrmOf(kid(sp, NS.p, 'spPr')),
      size: rPr?.getAttribute('sz') ? Number(rPr.getAttribute('sz')) / 100 : undefined,
      algn: lvl1?.getAttribute('algn') ?? undefined,
      anchor: desc(sp, NS.a, 'bodyPr')[0]?.getAttribute('anchor') ?? undefined,
      bold: rPr?.getAttribute('b') ? rPr.getAttribute('b') === '1' : undefined,
    });
  }
  return out;
}

function findPh(list: Placeholder[], type: string, idx?: string, needXfrm = true) {
  const norm = (t: string) => (t === 'ctrTitle' ? 'title' : t === 'subTitle' ? 'body' : t);
  const ok = (p: Placeholder) => !needXfrm || !!p.xfrm;
  return (idx ? list.find((p) => p.idx === idx && ok(p)) : undefined) ?? list.find((p) => p.type === type && ok(p)) ?? list.find((p) => norm(p.type) === norm(type) && ok(p));
}

/** Master text styles: default size/bold for title and body text. */
function masterStyles(master: Document | null) {
  const read = (name: string) => {
    const st = master ? desc(master.documentElement, NS.p, name)[0] : null;
    const lvl1 = desc(st, NS.a, 'lvl1pPr')[0];
    const rPr = desc(lvl1, NS.a, 'defRPr')[0];
    return { size: rPr?.getAttribute('sz') ? Number(rPr.getAttribute('sz')) / 100 : undefined, bold: rPr?.getAttribute('b') === '1', algn: lvl1?.getAttribute('algn') ?? undefined };
  };
  return { title: read('titleStyle'), body: read('bodyStyle') };
}

const toDataUrl = async (zip: JSZip, path: string) => {
  const f = zip.file(path);
  if (!f) return null;
  const b64 = await f.async('base64');
  const ext = path.split('.').pop()!.toLowerCase();
  const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'svg' ? 'image/svg+xml' : ext === 'gif' ? 'image/gif' : `image/${ext}`;
  return `data:${mime};base64,${b64}`;
};

export async function readPptx(buf: ArrayBuffer): Promise<Deck> {
  const zip = await JSZip.loadAsync(buf);
  const presPath = 'ppt/presentation.xml';
  const pres = parseXml(await zip.file(presPath)!.async('string'));
  const sz = desc(pres.documentElement, NS.p, 'sldSz')[0];
  const cx = Number(sz?.getAttribute('cx') ?? 12192000);
  const cy = Number(sz?.getAttribute('cy') ?? 6858000);
  // normalize width to 960px
  const W = 960;
  const scale = W / (cx / EMU_PER_PX);
  const H = Math.round((cy / EMU_PER_PX) * scale);
  const presRels = await readRels(zip, presPath);

  // theme
  const theme: Theme = { colors: {} };
  const themeRel = Object.values(presRels).find((r) => r.type.endsWith('/theme'));
  let themePath = themeRel?.target;
  if (!themePath) themePath = Object.keys(zip.files).find((f) => /^ppt\/theme\/theme\d+\.xml$/.test(f));
  if (themePath && zip.file(themePath)) {
    const th = parseXml(await zip.file(themePath)!.async('string'));
    const scheme = desc(th.documentElement, NS.a, 'clrScheme')[0];
    for (const c of Array.from(scheme?.children ?? [])) {
      const col = colorOf(c, theme);
      if (col) theme.colors[c.localName] = col;
    }
    theme.majorFont = desc(desc(th.documentElement, NS.a, 'majorFont')[0], NS.a, 'latin')[0]?.getAttribute('typeface') ?? undefined;
    theme.minorFont = desc(desc(th.documentElement, NS.a, 'minorFont')[0], NS.a, 'latin')[0]?.getAttribute('typeface') ?? undefined;
  }

  const slideIds = desc(pres.documentElement, NS.p, 'sldId');
  let sawNum = false;
  let footer: string | undefined;
  const slides: Slide[] = [];
  const cache = new Map<string, Document>();
  const load = async (p: string) => {
    if (!cache.has(p)) cache.set(p, parseXml((await zip.file(p)?.async('string')) ?? '<x/>'));
    return cache.get(p)!;
  };

  for (const sid of slideIds) {
    const rid = sid.getAttributeNS(NS.r, 'id') ?? sid.getAttribute('r:id');
    const slidePath = rid ? presRels[rid]?.target : undefined;
    if (!slidePath || !zip.file(slidePath)) continue;
    const doc = await load(slidePath);
    const rels = await readRels(zip, slidePath);
    const layoutPath = Object.values(rels).find((r) => r.type.endsWith('/slideLayout'))?.target;
    const layoutDoc = layoutPath ? await load(layoutPath) : null;
    const layoutRels = layoutPath ? await readRels(zip, layoutPath) : {};
    const masterPath = Object.values(layoutRels).find((r) => r.type.endsWith('/slideMaster'))?.target;
    const masterDoc = masterPath ? await load(masterPath) : null;
    const layoutPh = layoutDoc ? placeholdersOf(layoutDoc) : [];
    const masterPh = masterDoc ? placeholdersOf(masterDoc) : [];
    const mStyles = masterStyles(masterDoc);

    // background: slide → layout → master
    let bg = '#ffffff';
    for (const d of [doc, layoutDoc, masterDoc]) {
      const bgPr = d ? desc(d.documentElement, NS.p, 'bgPr')[0] : null;
      const fill = bgPr ? kid(bgPr, NS.a, 'solidFill') : null;
      const c = colorOf(fill, theme);
      if (c) { bg = c; break; }
      const bgRef = d ? desc(d.documentElement, NS.p, 'bgRef')[0] : null;
      const c2 = colorOf(bgRef, theme);
      if (c2) { bg = c2; break; }
    }

    const els: SlideEl[] = [];
    const spTree = desc(doc.documentElement, NS.p, 'spTree')[0];
    const walk = async (parent: Element | null) => {
      for (const node of Array.from(parent?.children ?? [])) {
        if (node.namespaceURI !== NS.p) continue;
        if (node.localName === 'grpSp') { await walk(node); continue; }
        if (node.localName === 'pic') {
          const blip = desc(node, NS.a, 'blip')[0];
          const embed = blip?.getAttributeNS(NS.r, 'embed') ?? blip?.getAttribute('r:embed');
          const target = embed ? rels[embed]?.target : undefined;
          const src = target ? await toDataUrl(zip, target) : null;
          const xf = xfrmOf(kid(node, NS.p, 'spPr'));
          if (src && xf) els.push({ id: uid(), type: 'image', src, x: xf.x * scale, y: xf.y * scale, w: xf.w * scale, h: xf.h * scale, rot: xf.rot, flipH: xf.flipH, flipV: xf.flipV });
          continue;
        }
        if (node.localName === 'graphicFrame') {
          const xf = readXfrm(kid(node, NS.p, 'xfrm'));
          if (!xf) continue;
          const box = { id: uid(), x: xf.x * scale, y: xf.y * scale, w: xf.w * scale, h: xf.h * scale };
          const tbl = desc(node, NS.a, 'tbl')[0];
          if (tbl) { const t = readTable(tbl, theme, scale, box); if (t) els.push(t); continue; }
          const chartRef = desc(node, NS.c, 'chart')[0];
          const crid = chartRef?.getAttributeNS(NS.r, 'id') ?? chartRef?.getAttribute('r:id');
          const chartPath = crid ? rels[crid]?.target : undefined;
          if (chartPath && zip.file(chartPath)) {
            const ch = readChart(await load(chartPath));
            if (ch) els.push({ ...box, type: 'chart', ...ch });
          }
          continue;
        }
        if (node.localName !== 'sp' && node.localName !== 'cxnSp') continue;
        const spPr = kid(node, NS.p, 'spPr');
        const ph = desc(node, NS.p, 'ph')[0];
        const phType = ph ? ph.getAttribute('type') ?? 'body' : undefined;
        const phIdx = ph?.getAttribute('idx') ?? undefined;
        let xf = xfrmOf(spPr);
        const inherited = ph ? findPh(layoutPh, phType!, phIdx) ?? findPh(masterPh, phType!, phIdx) : undefined;
        const layoutStyle = ph ? findPh(layoutPh, phType!, phIdx, false) : undefined;
        const masterStyle = ph ? findPh(masterPh, phType!, phIdx, false) : undefined;
        if (!xf && inherited?.xfrm) xf = { ...inherited.xfrm, rot: undefined, flipH: undefined, flipV: undefined };
        if (!xf) continue;
        if (phType === 'sldNum') { sawNum = true; continue; }
        if (phType === 'ftr') { const ft = desc(node, NS.a, 't').map((x) => x.textContent).join('').trim(); if (ft) footer ??= ft; continue; }
        if (phType === 'dt') continue;

        // text
        const txBody = kid(node, NS.p, 'txBody');
        const paras = kids(txBody, NS.a, 'p');
        const lines: string[] = [];
        let first: Element | null = null;
        const isTitleType = phType === 'title' || phType === 'ctrTitle';
        const inhAlgn = layoutStyle?.algn ?? masterStyle?.algn ?? (ph ? (isTitleType ? mStyles.title.algn : mStyles.body.algn) : undefined);
        let align: TextStyle['align'] = inhAlgn === 'ctr' ? 'center' : inhAlgn === 'r' ? 'right' : phType === 'ctrTitle' ? 'center' : 'left';
        let hasBullets = false;
        let numbered = false;
        let lineSpacing: number | undefined;
        let explicitAlign = false;
        for (const p of paras) {
          const pPr = kid(p, NS.a, 'pPr');
          const lvl = Number(pPr?.getAttribute('lvl') ?? 0);
          const algn = pPr?.getAttribute('algn');
          if (algn) explicitAlign = true;
          if (algn === 'ctr') align = 'center';
          else if (algn === 'r') align = 'right';
          else if (algn === 'l') align = 'left';
          if (pPr && kid(pPr, NS.a, 'buChar')) hasBullets = true;
          if (pPr && kid(pPr, NS.a, 'buAutoNum')) { hasBullets = true; numbered = true; }
          const spc = desc(kid(pPr, NS.a, 'lnSpc'), NS.a, 'spcPct')[0]?.getAttribute('val');
          if (spc && lineSpacing === undefined) lineSpacing = Number(spc) / 100000;
          let text = '';
          for (const r of Array.from(p.children)) {
            if (r.localName === 'r' || r.localName === 'fld') {
              text += kid(r, NS.a, 't')?.textContent ?? '';
              if (!first) first = kid(r, NS.a, 'rPr');
            } else if (r.localName === 'br') text += '\n';
          }
          lines.push((lvl > 0 ? '  ' : '') + text);
        }
        while (lines.length && lines[lines.length - 1] === '') lines.pop();
        const text = lines.join('\n');

        const isTitle = phType === 'title' || phType === 'ctrTitle';
        const sizeAttr = first?.getAttribute('sz') ?? desc(txBody, NS.a, 'defRPr')[0]?.getAttribute('sz');
        const fontSize = sizeAttr
          ? Number(sizeAttr) / 100
          : layoutStyle?.size ?? masterStyle?.size ?? (ph ? (isTitle ? mStyles.title.size : mStyles.body.size) : undefined) ?? (isTitle ? (phType === 'ctrTitle' ? 40 : 32) : 18);
        const spFill = colorOf(kid(spPr, NS.a, 'solidFill'), theme);
        const fontRefColor = colorOf(desc(kid(node, NS.p, 'style'), NS.a, 'fontRef')[0] ?? null, theme);
        const color = colorOf(kid(first, NS.a, 'solidFill'), theme) ?? (!ph && spFill ? fontRefColor ?? (bgIsDark(spFill) ? '#ffffff' : '#111827') : undefined) ?? (bgIsDark(bg) ? '#f8fafc' : theme.colors.dk1 ?? '#1f2937');
        const latin = kid(first, NS.a, 'latin')?.getAttribute('typeface');
        const fontFace = latin && !latin.startsWith('+') ? latin : isTitle ? theme.majorFont : theme.minorFont;
        const bodyPr = kid(txBody, NS.a, 'bodyPr');
        const anchor = bodyPr?.getAttribute('anchor') ?? layoutStyle?.anchor ?? masterStyle?.anchor;
        const style = textStyle({
          fontSize: Math.round(fontSize * scale * 10) / 10,
          color,
          bold: first?.getAttribute('b') ? first.getAttribute('b') === '1' : layoutStyle?.bold ?? masterStyle?.bold ?? (isTitle ? mStyles.title.bold : false),
          italic: first?.getAttribute('i') === '1',
          underline: !!first?.getAttribute('u') && first.getAttribute('u') !== 'none',
          strike: !!first?.getAttribute('strike') && first.getAttribute('strike') !== 'noStrike' ? true : undefined,
          lineSpacing: lineSpacing && Math.abs(lineSpacing - 1) > 0.01 ? Math.round(lineSpacing * 100) / 100 : undefined,
          align,
          valign: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : isTitle ? 'middle' : 'top',
          fontFace: fontFace ?? undefined,
        });
        const geom = kid(spPr, NS.a, 'prstGeom')?.getAttribute('prst');
        const fillColor = spFill ?? (!ph ? colorOf(desc(kid(node, NS.p, 'style'), NS.a, 'fillRef')[0] ?? null, theme) : undefined);
        const ln = kid(spPr, NS.a, 'ln');
        const lnColor = ln && !kid(ln, NS.a, 'noFill') ? colorOf(kid(ln, NS.a, 'solidFill'), theme) : undefined;
        const isLine = node.localName === 'cxnSp' || geom === 'line' || !!geom?.includes('Connector');
        const lnW = ln?.getAttribute('w') ? Number(ln.getAttribute('w')) / 12700 : undefined;
        const prstDash = desc(ln, NS.a, 'prstDash')[0]?.getAttribute('val');
        const dash: ShapeEl['dash'] = !prstDash || prstDash === 'solid' ? undefined : /dot/i.test(prstDash) && !/dash/i.test(prstDash) ? 'dot' : 'dash';
        const arrowEnd = [kid(ln, NS.a, 'tailEnd'), kid(ln, NS.a, 'headEnd')].some((x) => x && (x.getAttribute('type') ?? 'none') !== 'none');
        // a line's box may be flat; give it room and keep it centred on the same path
        const lineH = 20;
        const base = isLine
          ? lineBox(xf, scale, lineH)
          : { id: uid(), x: xf.x * scale, y: xf.y * scale, w: Math.max(4, xf.w * scale), h: Math.max(4, xf.h * scale), rot: xf.rot, flipH: xf.flipH, flipV: xf.flipV };
        const shapeKind: ShapeKind | undefined = isLine ? (arrowEnd ? 'arrowLine' : 'line') : geom && geom !== 'rect' ? FROM_PRST[geom] : geom === 'rect' && (fillColor || lnColor) ? 'rect' : undefined;
        const alpha = kid(kid(spPr, NS.a, 'solidFill')?.firstElementChild ?? null, NS.a, 'alpha')?.getAttribute('val');
        const extra = { opacity: alpha ? Number(alpha) / 100000 : undefined, shadow: desc(spPr, NS.a, 'outerShdw').length > 0 || undefined };
        if (!ph && shapeKind) {
          els.push({
            ...base, ...extra, type: 'shape', shape: shapeKind, fill: fillColor ?? 'transparent', stroke: isLine ? lnColor ?? fillColor ?? '#1f2937' : lnColor,
            strokeW: isLine ? lnW ?? 2 : lnColor ? lnW ?? 2 : 0, dash, text: isLine ? '' : text,
            style: { ...style, align: explicitAlign ? style.align : 'center', valign: anchor === 't' ? 'top' : anchor === 'b' ? 'bottom' : 'middle' },
          });
        } else if (text || ph) {
          els.push({ ...base, ...extra, type: 'text', text, style, fill: fillColor, bullets: hasBullets || (phType === 'body' && !!phIdx && lines.length > 1), numbered: numbered || undefined, placeholder: isTitle ? 'title' : phType === 'subTitle' ? 'subtitle' : ph ? 'body' : undefined });
        }
      }
    };
    await walk(spTree);

    // notes
    let notes: string | undefined;
    const notesPath = Object.values(rels).find((r) => r.type.endsWith('/notesSlide'))?.target;
    if (notesPath && zip.file(notesPath)) {
      const nd = await load(notesPath);
      const bodySp = desc(nd.documentElement, NS.p, 'sp').find((sp) => desc(sp, NS.p, 'ph')[0]?.getAttribute('type') === 'body');
      if (bodySp) notes = desc(bodySp, NS.a, 'p').map((p) => desc(p, NS.a, 't').map((x) => x.textContent).join('')).join('\n').trim() || undefined;
    }
    const trEl = kid(doc.documentElement, NS.p, 'transition');
    const trKind = trEl ? Array.from(trEl.children).find((c) => c.namespaceURI === NS.p)?.localName : undefined;
    const transition: Transition | undefined = !trEl ? undefined
      : trKind === 'push' || trKind === 'pull' ? 'push' : trKind === 'wipe' || trKind === 'split' || trKind === 'strips' ? 'wipe'
      : trKind === 'zoom' || trKind === 'zoomIn' ? 'zoom' : trKind === 'cover' ? 'cover' : 'fade';
    slides.push({ id: uid(), bg, els, notes, hidden: doc.documentElement.getAttribute('show') === '0' || undefined, transition });
  }
  if (!slides.length) slides.push({ id: uid(), bg: '#ffffff', els: [] });
  return { w: W, h: H, theme: 'paper', slides, slideNumbers: sawNum || undefined, footer };
}

function bgIsDark(c: string) {
  const h = c.replace('#', '');
  if (h.length !== 6) return false;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
  return 0.299 * r + 0.587 * g + 0.114 * b < 110;
}

/** A line drawn across its box: re-centre it in a box tall enough to grab, rotated to the same angle. */
function lineBox(xf: NonNullable<ReturnType<typeof readXfrm>>, scale: number, h: number) {
  const w0 = xf.w * scale, h0 = xf.h * scale;
  const cx = xf.x * scale + w0 / 2, cy = xf.y * scale + h0 / 2;
  // the path goes corner to corner (flips choose which diagonal)
  let ang = (Math.atan2(h0, w0) * 180) / Math.PI;
  if (xf.flipH) ang = 180 - ang;
  if (xf.flipV) ang = -ang;
  ang += xf.rot ?? 0;
  const len = Math.max(4, Math.hypot(w0, h0));
  const rot = Math.round((((ang % 360) + 360) % 360) * 10) / 10;
  return { id: uid(), x: cx - len / 2, y: cy - h / 2, w: len, h, rot: rot || undefined };
}

function readTable(tbl: Element, theme: Theme, scale: number, box: { id: string; x: number; y: number; w: number; h: number }): SlideEl | null {
  const trs = kids(tbl, NS.a, 'tr');
  if (!trs.length) return null;
  const rows = trs.map((tr) => kids(tr, NS.a, 'tc').map((tc) => kids(kid(tc, NS.a, 'txBody'), NS.a, 'p').map((p) => desc(p, NS.a, 't').map((x) => x.textContent).join('')).join('\n')));
  const cols = Math.max(...rows.map((r) => r.length));
  for (const r of rows) while (r.length < cols) r.push('');
  const tblPr = kid(tbl, NS.a, 'tblPr');
  const firstFill = colorOf(kid(kid(kids(trs[0], NS.a, 'tc')[0], NS.a, 'tcPr'), NS.a, 'solidFill'), theme);
  const rPr = desc(tbl, NS.a, 'rPr')[0] ?? desc(tbl, NS.a, 'endParaRPr')[0];
  const sz = rPr?.getAttribute('sz');
  return {
    ...box, type: 'table', rows,
    color: firstFill ?? theme.colors.accent1 ?? '#6366f1',
    header: tblPr ? tblPr.getAttribute('firstRow') === '1' : true,
    banded: tblPr ? tblPr.getAttribute('bandRow') === '1' : true,
    style: textStyle({ fontSize: Math.round((sz ? Number(sz) / 100 : 14) * scale * 10) / 10, color: theme.colors.dk1 ?? '#1f2937', fontFace: theme.minorFont }),
  };
}

function readChart(doc: Document): { chart: ChartType; data: ChartData; title?: string; legend: boolean } | null {
  const plot = desc(doc.documentElement, NS.c, 'plotArea')[0];
  if (!plot) return null;
  const kinds: [string, ChartType][] = [['barChart', 'column'], ['bar3DChart', 'column'], ['lineChart', 'line'], ['line3DChart', 'line'], ['areaChart', 'area'], ['pieChart', 'pie'], ['pie3DChart', 'pie'], ['doughnutChart', 'doughnut'], ['scatterChart', 'scatter']];
  let group: Element | null = null;
  let chart: ChartType = 'column';
  for (const [name, type] of kinds) { const g = kid(plot, NS.c, name); if (g) { group = g; chart = type; break; } }
  if (!group) return null;
  if (chart === 'column' && kid(group, NS.c, 'barDir')?.getAttribute('val') === 'bar') chart = 'bar';
  const pts = (el: Element | null) => {
    const out: string[] = [];
    for (const pt of desc(el, NS.c, 'pt')) out[Number(pt.getAttribute('idx') ?? out.length)] = desc(pt, NS.c, 'v')[0]?.textContent ?? '';
    return out;
  };
  const sers = kids(group, NS.c, 'ser');
  let cats: string[] = [];
  const series = sers.map((ser, i) => {
    const catEl = kid(ser, NS.c, chart === 'scatter' ? 'xVal' : 'cat');
    if (!cats.length) cats = pts(catEl);
    const vals = pts(kid(ser, NS.c, chart === 'scatter' ? 'yVal' : 'val'));
    const name = desc(kid(ser, NS.c, 'tx'), NS.c, 'v')[0]?.textContent ?? desc(kid(ser, NS.c, 'tx'), NS.a, 't')[0]?.textContent ?? `Series ${i + 1}`;
    return { name, values: vals.map((v) => (v === undefined || v === '' || !Number.isFinite(Number(v)) ? null : Number(v))) };
  });
  const n = Math.max(cats.length, ...series.map((x) => x.values.length));
  cats = Array.from({ length: n }, (_, i) => cats[i] ?? String(i + 1));
  for (const x of series) while (x.values.length < n) x.values.push(null);
  const titleEl = desc(doc.documentElement, NS.c, 'title')[0];
  const title = titleEl ? desc(titleEl, NS.a, 't').map((x) => x.textContent).join('') || undefined : undefined;
  return { chart, data: { cats, series }, title, legend: desc(doc.documentElement, NS.c, 'legend').length > 0 };
}
