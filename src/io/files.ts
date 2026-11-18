import JSZip from 'jszip';
import { nativeSave } from '../platform';
import { generateJSON } from '@tiptap/core';
import type { DocData, DocKind, OpenFile } from '../types';
import { docExtensions } from '../editors/doc/extensions';
import { readCsv, writeCsv } from '../editors/sheet/csv';

// Office converters are loaded on demand to keep startup fast.
const docxIO = () => import('../editors/doc/docxIO');
const xlsxIO = () => import('../editors/sheet/xlsxIO');
const pptxIO = () => import('../editors/slides/pptxIO');
import type { Workbook } from '../editors/sheet/model';
import type { Deck } from '../editors/slides/model';
import type { PdfDoc } from '../editors/pdf/model';
// PDF engine (pdf.js + pdf-lib) is loaded on demand too.
const pdfjsMod = () => import('../editors/pdf/pdfjs');
const pdfSave = () => import('../editors/pdf/save');
const pdfImages = () => import('../editors/pdf/images');

export const NATIVE_EXT = 'pline';
export const ACCEPT = '.docx,.xlsx,.pptx,.csv,.tsv,.txt,.md,.html,.htm,.pline,.pdf,.png,.jpg,.jpeg,.webp,.gif,.bmp';

export type SaveFormat = 'native' | 'docx' | 'xlsx' | 'pptx' | 'csv' | 'pdf';

type Loaded = { kind: 'doc'; data: DocData } | { kind: 'sheet'; data: Workbook } | { kind: 'slides'; data: Deck } | { kind: 'pdf'; data: PdfDoc };

const baseName = (n: string) => n.replace(/\.[^.]+$/, '');
const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export const defaultPage = (): DocData['page'] => ({ orientation: 'portrait', margin: 'normal' });

export async function openFile(file: File): Promise<Loaded & { name: string }> {
  const ext = file.name.split('.').pop()?.toLowerCase() ?? '';
  const name = baseName(file.name);
  switch (ext) {
    case 'docx':
      return { kind: 'doc', name, data: { html: await (await docxIO()).readDocx(await file.arrayBuffer()), page: defaultPage() } };
    case 'xlsx':
    case 'xlsm':
      return { kind: 'sheet', name, data: await (await xlsxIO()).readXlsx(await file.arrayBuffer()) };
    case 'csv':
    case 'tsv':
      return { kind: 'sheet', name, data: readCsv(await file.text(), name.slice(0, 31) || 'Sheet1') };
    case 'pptx':
      return { kind: 'slides', name, data: await (await pptxIO()).readPptx(await file.arrayBuffer()) };
    case 'txt':
    case 'md': {
      const text = await file.text();
      const html = ext === 'md' ? markdownToHtml(text) : text.split(/\r?\n/).map((l) => `<p>${esc(l)}</p>`).join('');
      return { kind: 'doc', name, data: { html, page: defaultPage() } };
    }
    case 'html':
    case 'htm':
      return { kind: 'doc', name, data: { html: new DOMParser().parseFromString(await file.text(), 'text/html').body.innerHTML, page: defaultPage() } };
    case 'pdf': {
      const { importPdf } = await pdfjsMod();
      const r = await importPdf(new Uint8Array(await file.arrayBuffer()));
      return { kind: 'pdf', name, data: { sources: { [r.id]: r.b64 }, pages: r.pages, form: {}, decor: {}, meta: { title: r.title } } };
    }
    case 'png':
    case 'jpg':
    case 'jpeg':
    case 'webp':
    case 'gif':
    case 'bmp': {
      const pages = await (await pdfImages()).imagesToPages([file]);
      return { kind: 'pdf', name, data: { sources: {}, pages, form: {}, decor: {}, meta: {} } };
    }
    case NATIVE_EXT:
      return { ...(await readNative(await file.arrayBuffer())), name };
    default:
      throw new Error('unsupported');
  }
}

// -------- native .pline: a zip with a manifest and the JSON model --------
interface Manifest {
  format: 'paperline';
  version: 1;
  kind: DocKind;
  app: string;
  modified: string;
}

export async function writeNative(f: OpenFile): Promise<Blob> {
  const zip = new JSZip();
  const manifest: Manifest = { format: 'paperline', version: 1, kind: f.kind, app: 'Paperline 0.1', modified: new Date().toISOString() };
  zip.file('manifest.json', JSON.stringify(manifest, null, 2));
  zip.file('content.json', JSON.stringify(f.data));
  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.paperline+zip' });
}

async function readNative(buf: ArrayBuffer): Promise<Loaded> {
  const zip = await JSZip.loadAsync(buf);
  const manifest = JSON.parse(await zip.file('manifest.json')!.async('string')) as Manifest;
  if (manifest.format !== 'paperline') throw new Error('not a paperline file');
  const data = JSON.parse(await zip.file('content.json')!.async('string'));
  return { kind: manifest.kind, data } as Loaded;
}

export async function exportAs(f: OpenFile, fmt: SaveFormat): Promise<{ blob: Blob; filename: string }> {
  const name = f.name || 'Untitled';
  switch (fmt) {
    case 'native':
      return { blob: await writeNative(f), filename: `${name}.${NATIVE_EXT}` };
    case 'docx': {
      if (f.kind !== 'doc') throw new Error('wrong kind');
      const json = generateJSON(f.data.html, docExtensions);
      return { blob: await (await docxIO()).writeDocx(json, f.data.page), filename: `${name}.docx` };
    }
    case 'xlsx':
      if (f.kind !== 'sheet') throw new Error('wrong kind');
      return { blob: await (await xlsxIO()).writeXlsx(f.data), filename: `${name}.xlsx` };
    case 'csv':
      if (f.kind !== 'sheet') throw new Error('wrong kind');
      return { blob: new Blob(['﻿' + writeCsv(f.data.sheets[f.data.active], f.data)], { type: 'text/csv' }), filename: `${name}.csv` };
    case 'pptx':
      if (f.kind !== 'slides') throw new Error('wrong kind');
      return { blob: await (await pptxIO()).writePptx(f.data), filename: `${name}.pptx` };
    case 'pdf':
      if (f.kind !== 'pdf') throw new Error('wrong kind');
      return { blob: await (await pdfSave()).pdfBlob(f.data), filename: `${name}.pdf` };
  }
}

/** Save with the system dialog when the browser has one, else download. */
export async function saveBlob(blob: Blob, filename: string) {
  const native = await nativeSave(blob, filename);
  if (native !== null) return native;
  const w = window as unknown as { showSaveFilePicker?: (o: unknown) => Promise<FileSystemFileHandle> };
  if (w.showSaveFilePicker && window.self === window.top) {
    try {
      const ext = filename.split('.').pop()!;
      const handle = await w.showSaveFilePicker({ suggestedName: filename, types: [{ description: ext.toUpperCase(), accept: { [blob.type || 'application/octet-stream']: [`.${ext}`] } }] });
      const ws = await handle.createWritable();
      await ws.write(blob);
      await ws.close();
      return;
    } catch (e) {
      if ((e as DOMException)?.name === 'AbortError') return false;
      // fall through to download
    }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
  return true;
}

/** Small Markdown → HTML for opening .md files (headings, lists, emphasis, code, links). */
function markdownToHtml(md: string): string {
  const inline = (s: string) =>
    esc(s)
      .replace(/`([^`]+)`/g, '<code>$1</code>')
      .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
      .replace(/\*([^*]+)\*/g, '<em>$1</em>')
      .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');
  const out: string[] = [];
  let list: 'ul' | 'ol' | null = null;
  let code = false;
  let codeBuf: string[] = [];
  const closeList = () => { if (list) { out.push(`</${list}>`); list = null; } };
  for (const line of md.split(/\r?\n/)) {
    if (line.startsWith('```')) {
      if (code) { out.push(`<pre><code>${esc(codeBuf.join('\n'))}</code></pre>`); codeBuf = []; }
      else closeList();
      code = !code;
      continue;
    }
    if (code) { codeBuf.push(line); continue; }
    let m: RegExpExecArray | null;
    if ((m = /^(#{1,4})\s+(.*)$/.exec(line))) { closeList(); out.push(`<h${m[1].length}>${inline(m[2])}</h${m[1].length}>`); }
    else if ((m = /^\s*[-*+]\s+(.*)$/.exec(line))) { if (list !== 'ul') { closeList(); out.push('<ul>'); list = 'ul'; } out.push(`<li><p>${inline(m[1])}</p></li>`); }
    else if ((m = /^\s*\d+[.)]\s+(.*)$/.exec(line))) { if (list !== 'ol') { closeList(); out.push('<ol>'); list = 'ol'; } out.push(`<li><p>${inline(m[1])}</p></li>`); }
    else if ((m = /^>\s?(.*)$/.exec(line))) { closeList(); out.push(`<blockquote><p>${inline(m[1])}</p></blockquote>`); }
    else if (/^(-{3,}|\*{3,})$/.test(line.trim())) { closeList(); out.push('<hr>'); }
    else if (line.trim() === '') closeList();
    else { closeList(); out.push(`<p>${inline(line)}</p>`); }
  }
  closeList();
  return out.join('');
}
