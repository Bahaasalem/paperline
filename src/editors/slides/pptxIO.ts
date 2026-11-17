import JSZip from 'jszip';
import { textStyle, uid, type Deck, type ShapeKind, type Slide, type SlideEl, type TextStyle } from './model';

const EMU_PER_PX = 9525;
const PX_PER_IN = 96;

// ---------------- export ----------------
const PPTX_SHAPES: Record<ShapeKind, string> = {
  rect: 'rect',
  roundRect: 'roundRect',
  ellipse: 'ellipse',
  triangle: 'triangle',
  rightArrow: 'rightArrow',
  line: 'line',
};

const hex = (c?: string) => (c && c.startsWith('#') ? c.slice(1).toUpperCase().padEnd(6, '0').slice(0, 6) : undefined);

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
    if (s.notes) slide.addNotes(s.notes);
    for (const el of s.els) {
      const pos = { x: inch(el.x), y: inch(el.y), w: inch(el.w), h: inch(el.h), rotate: el.rot };
      if (el.type === 'image') {
        slide.addImage({ data: el.src, ...pos });
        continue;
      }
      const textOpts = textProps(el.style);
      const lines = el.text.split('\n');
      const runs = lines.map((l, i) => ({
        text: l.trimStart(),
        options: {
          breakLine: i < lines.length - 1,
          bullet: el.type === 'text' && el.bullets ? (l.startsWith('  ') ? { indent: 18 } : true) : undefined,
          indentLevel: el.type === 'text' && el.bullets && l.startsWith('  ') ? 1 : undefined,
        },
      }));
      if (el.type === 'shape') {
        const fill = el.fill && el.fill !== 'transparent' ? { color: hex(el.fill) } : undefined;
        const line = el.shape === 'line' ? { color: hex(el.stroke ?? el.fill), width: el.strokeW ?? 2 } : el.stroke ? { color: hex(el.stroke), width: el.strokeW ?? 2 } : undefined;
        if (el.text) slide.addText(runs, { ...pos, ...textOpts, shape: PPTX_SHAPES[el.shape] as never, fill: fill as never, line: line as never });
        else slide.addShape(PPTX_SHAPES[el.shape] as never, { ...pos, fill: fill as never, line: line as never });
      } else {
        slide.addText(el.text ? runs : '', { ...pos, ...textOpts, fill: el.fill ? { color: hex(el.fill) } : undefined, margin: 4 } as never);
      }
    }
  }
  const out = (await pptx.write({ outputType: 'blob' })) as Blob;
  return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
}

function textProps(st: TextStyle) {
  return {
    fontSize: st.fontSize,
    fontFace: st.fontFace,
    color: hex(st.color),
    bold: st.bold,
    italic: st.italic,
    underline: st.underline ? { style: 'sng' } : undefined,
    align: st.align,
    valign: st.valign,
  } as Record<string, unknown>;
}

// ---------------- import ----------------
const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
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
  const x = kid(spPr, NS.a, 'xfrm');
  if (!x) return undefined;
  const off = kid(x, NS.a, 'off'), ext = kid(x, NS.a, 'ext');
  if (!off || !ext) return undefined;
  return {
    x: Number(off.getAttribute('x')) / EMU_PER_PX,
    y: Number(off.getAttribute('y')) / EMU_PER_PX,
    w: Number(ext.getAttribute('cx')) / EMU_PER_PX,
    h: Number(ext.getAttribute('cy')) / EMU_PER_PX,
    rot: x.getAttribute('rot') ? Number(x.getAttribute('rot')) / 60000 : undefined,
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
          if (src && xf) els.push({ id: uid(), type: 'image', src, x: xf.x * scale, y: xf.y * scale, w: xf.w * scale, h: xf.h * scale, rot: xf.rot });
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
        if (!xf && inherited?.xfrm) xf = { ...inherited.xfrm, rot: undefined };
        if (!xf) continue;
        if (['sldNum', 'dt', 'ftr'].includes(phType ?? '')) continue;

        // text
        const txBody = kid(node, NS.p, 'txBody');
        const paras = kids(txBody, NS.a, 'p');
        const lines: string[] = [];
        let first: Element | null = null;
        const isTitleType = phType === 'title' || phType === 'ctrTitle';
        const inhAlgn = layoutStyle?.algn ?? masterStyle?.algn ?? (ph ? (isTitleType ? mStyles.title.algn : mStyles.body.algn) : undefined);
        let align: TextStyle['align'] = inhAlgn === 'ctr' ? 'center' : inhAlgn === 'r' ? 'right' : phType === 'ctrTitle' ? 'center' : 'left';
        let hasBullets = false;
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
          if (pPr && kid(pPr, NS.a, 'buAutoNum')) hasBullets = true;
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
          align,
          valign: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : isTitle ? 'middle' : 'top',
          fontFace: fontFace ?? undefined,
        });
        const geom = kid(spPr, NS.a, 'prstGeom')?.getAttribute('prst');
        const fillColor = spFill ?? (!ph ? colorOf(desc(kid(node, NS.p, 'style'), NS.a, 'fillRef')[0] ?? null, theme) : undefined);
        const ln = kid(spPr, NS.a, 'ln');
        const lnColor = ln && !kid(ln, NS.a, 'noFill') ? colorOf(kid(ln, NS.a, 'solidFill'), theme) : undefined;
        const base = { id: uid(), x: xf.x * scale, y: xf.y * scale, w: Math.max(4, xf.w * scale), h: Math.max(4, xf.h * scale), rot: xf.rot };
        const shapeKind: ShapeKind | undefined =
          node.localName === 'cxnSp' || geom === 'line' || geom?.includes('Connector') ? 'line'
            : geom === 'ellipse' ? 'ellipse' : geom === 'roundRect' ? 'roundRect' : geom === 'triangle' ? 'triangle' : geom === 'rightArrow' ? 'rightArrow' : geom === 'rect' && (fillColor || lnColor) ? 'rect' : undefined;
        if (!ph && shapeKind) {
          els.push({ ...base, type: 'shape', shape: shapeKind, fill: fillColor ?? 'transparent', stroke: lnColor, strokeW: lnColor ? 2 : 0, text, style: { ...style, align: explicitAlign ? style.align : 'center', valign: anchor === 't' ? 'top' : anchor === 'b' ? 'bottom' : 'middle' } });
        } else if (text || ph) {
          els.push({ ...base, type: 'text', text, style, fill: fillColor, bullets: hasBullets || (phType === 'body' && !!phIdx && lines.length > 1), placeholder: isTitle ? 'title' : phType === 'subTitle' ? 'subtitle' : ph ? 'body' : undefined });
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
    slides.push({ id: uid(), bg, els, notes });
  }
  if (!slides.length) slides.push({ id: uid(), bg: '#ffffff', els: [] });
  return { w: W, h: H, theme: 'paper', slides };
}

function bgIsDark(c: string) {
  const h = c.replace('#', '');
  if (h.length !== 6) return false;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
  return 0.299 * r + 0.587 * g + 0.114 * b < 110;
}
