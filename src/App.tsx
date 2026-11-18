import { lazy, Suspense, useCallback, useEffect, useRef, useState } from 'react';
import { ChevronDown, Download, FileDown, FilePlus2, FolderOpen, Home, Languages, Loader2, Printer, Save, X } from 'lucide-react';
import { useI18n } from './i18n';
import type { DocKind, OpenFile } from './types';
import { StartScreen } from './StartScreen';
import { DocEditor } from './editors/doc/DocEditor';
import { SheetEditor } from './editors/sheet/SheetEditor';
import { SlidesEditor } from './editors/slides/SlidesEditor';
import { newPdfDoc } from './editors/pdf/model';
import { newWorkbook } from './editors/sheet/model';
import { newDeck } from './editors/slides/model';
import { ask, listenForOpenedFiles, printPdfBlob, printsViaPdf, shell } from './platform';
import { ACCEPT, defaultPage, exportAs, openFile, saveBlob, type SaveFormat } from './io/files';
import { PrintJob, type Printable, type PrintMode } from './io/print';
import { Menu, MenuItem } from './ui/Ribbon';
import { AppIcon, Logo } from './ui/Logo';

// The PDF editor brings pdf.js and pdf-lib, so it loads only when a PDF is opened.
const PdfEditor = lazy(() => import('./editors/pdf/PdfEditor').then((m) => ({ default: m.PdfEditor })));

let nextId = 1;

export default function App() {
  const { t, toggle } = useI18n();
  const [files, setFiles] = useState<OpenFile[]>([]);
  const [activeId, setActiveId] = useState<number | null>(null);
  const [dirty, setDirty] = useState<Record<number, boolean>>({});
  const [busy, setBusy] = useState(false);
  const [toast, setToast] = useState<string | null>(null);
  const [job, setJob] = useState<{ mode: PrintMode; file: Printable; print?: boolean } | null>(null);
  const openInput = useRef<HTMLInputElement>(null);

  const active = files.find((f) => f.id === activeId) ?? null;

  const flash = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast((m) => (m === msg ? null : m)), 2600);
  };

  const add = (f: Omit<OpenFile, 'id'>) => {
    const file = { ...f, id: nextId++ } as OpenFile;
    setFiles((fs) => [...fs, file]);
    setActiveId(file.id);
  };

  const createNew = (kind: DocKind) => {
    const name = `${t('untitled')} ${t(kind === 'doc' ? 'kind_doc' : kind === 'sheet' ? 'kind_sheet' : kind === 'pdf' ? 'kind_pdf' : 'kind_slides')}`;
    if (kind === 'pdf') add({ kind, name, data: newPdfDoc() });
    else if (kind === 'doc') add({ kind, name, data: { html: '', page: defaultPage() } });
    else if (kind === 'sheet') add({ kind, name, data: newWorkbook(t('sheet')) });
    else add({ kind, name, data: newDeck() });
  };

  const open = async (file: File) => {
    setBusy(true);
    try {
      const loaded = await openFile(file);
      add(loaded as Omit<OpenFile, 'id'>);
    } catch (e) {
      console.error(e);
      const m = (e as Error)?.message;
      flash(m === 'unsupported' ? t('unsupported') : m === 'password' ? t('pdfPassword') : t('errorOpen'));
    } finally {
      setBusy(false);
    }
  };

  const update = useCallback(
    (id: number) => (data: OpenFile['data']) => {
      setFiles((fs) => fs.map((f) => (f.id === id ? ({ ...f, data } as OpenFile) : f)));
      setDirty((d) => (d[id] ? d : { ...d, [id]: true }));
    },
    [],
  );
  const updaters = useRef(new Map<number, (d: OpenFile['data']) => void>());
  const updaterFor = (id: number) => {
    if (!updaters.current.has(id)) updaters.current.set(id, update(id));
    return updaters.current.get(id)!;
  };

  const save = async (fmt: SaveFormat) => {
    if (!active) return;
    setBusy(true);
    try {
      const { blob, filename } = await exportAs(active, fmt);
      if ((await saveBlob(blob, filename)) === false) return;
      if (fmt === 'native' || fmt === 'docx' || fmt === 'xlsx' || fmt === 'pptx' || fmt === 'pdf') setDirty((d) => ({ ...d, [active.id]: false }));
      flash(t('saved'));
    } catch (e) {
      console.error(e);
      flash(String((e as Error)?.message ?? e));
    } finally {
      setBusy(false);
    }
  };

  const officeFormat: SaveFormat | null = active
    ? active.kind === 'doc' ? 'docx' : active.kind === 'sheet' ? 'xlsx' : active.kind === 'pdf' ? 'pdf' : 'pptx'
    : null;

  /** Print / export PDF. A PDF file already is one: save it directly and print its rendered pages. */
  const startJob = (mode: PrintMode, file: OpenFile) => {
    if (file.kind === 'pdf') {
      if (mode === 'pdf') return save('pdf');
      setBusy(true);
      if (printsViaPdf) {
        import('./editors/pdf/save')
          .then(({ pdfBlob }) => pdfBlob(file.data))
          .then((b) => printPdfBlob(b, `${file.name}.pdf`))
          .catch((e) => flash(String((e as Error)?.message ?? e)))
          .finally(() => setBusy(false));
        return;
      }
      import('./editors/pdf/print')
        .then(({ printPdfDoc }) => printPdfDoc(file.data))
        .catch((e) => flash(String((e as Error)?.message ?? e)))
        .finally(() => setBusy(false));
      return;
    }
    if (mode === 'print' && printsViaPdf) { setBusy(true); setJob({ mode: 'pdf', file, print: true }); return; }
    if (mode === 'pdf') setBusy(true);
    setJob({ mode, file });
  };

  const close = async (id: number) => {
    if (dirty[id] && !(await ask(t('unsavedConfirm')))) return;
    setFiles((fs) => {
      const rest = fs.filter((f) => f.id !== id);
      if (activeId === id) setActiveId(rest.length ? rest[rest.length - 1].id : null);
      return rest;
    });
    updaters.current.delete(id);
  };

  const rename = (name: string) => {
    if (!active) return;
    const n = name.trim() || t('untitled');
    setFiles((fs) => fs.map((f) => (f.id === active.id ? { ...f, name: n } : f)));
  };

  // keyboard shortcuts
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;
      const k = e.key.toLowerCase();
      if (k === 's') { e.preventDefault(); if (active && officeFormat) save(e.shiftKey ? 'native' : officeFormat); }
      else if (k === 'o') { e.preventDefault(); openInput.current?.click(); }
      else if (k === 'p') { e.preventDefault(); if (active) startJob('print', active); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  // files the system asks us to open (double-click, "Open with Paperline")
  const openRef = useRef(open);
  openRef.current = open;
  useEffect(() => {
    let off: (() => void) | undefined;
    let gone = false;
    listenForOpenedFiles((f) => openRef.current(f)).then((u) => { if (gone) u(); else off = u; });
    return () => { gone = true; off?.(); };
  }, []);

  useEffect(() => {
    const onUnload = (e: BeforeUnloadEvent) => {
      if (Object.values(dirty).some(Boolean)) e.preventDefault();
    };
    window.addEventListener('beforeunload', onUnload);
    return () => window.removeEventListener('beforeunload', onUnload);
  }, [dirty]);

  const kindClass = active ? `kind-${active.kind}` : 'kind-home';

  return (
    <div className={`app ${kindClass}`}>
      {(active || files.length > 0) && (
      <header className="topbar">
        <button className={`home-btn ${!active ? 'on' : ''}`} onClick={() => setActiveId(null)} title={t('home')} aria-label={t('home')}>
          <Logo size={26} />
        </button>
        {active && (
          <Menu
            trigger={(o, toggleMenu) => (
              <button className={`file-btn ${o ? 'on' : ''}`} onClick={toggleMenu}>
                {t('file')} <ChevronDown size={14} />
              </button>
            )}
          >
            {(closeMenu) => (
              <>
                <MenuItem icon={<FilePlus2 size={16} />} label={t('new')} onClick={() => { closeMenu(); setActiveId(null); }} />
                <MenuItem icon={<FolderOpen size={16} />} label={t('open')} hint="Ctrl+O" onClick={() => { closeMenu(); openInput.current?.click(); }} />
                <div className="menu-sep" />
                <MenuItem icon={<Save size={16} />} label={`${t('save')} · ${officeFormat?.toUpperCase()}`} hint="Ctrl+S" onClick={() => { closeMenu(); save(officeFormat!); }} />
                <div className="menu-label">{t('saveAs')}</div>
                <MenuItem icon={<Download size={16} />} label={t('saveNative')} hint="Ctrl+Shift+S" onClick={() => { closeMenu(); save('native'); }} />
                {active.kind === 'doc' && <MenuItem icon={<Download size={16} />} label={t('saveDocx')} onClick={() => { closeMenu(); save('docx'); }} />}
                {active.kind === 'sheet' && <MenuItem icon={<Download size={16} />} label={t('saveXlsx')} onClick={() => { closeMenu(); save('xlsx'); }} />}
                {active.kind === 'sheet' && <MenuItem icon={<Download size={16} />} label={t('saveCsv')} onClick={() => { closeMenu(); save('csv'); }} />}
                {active.kind === 'slides' && <MenuItem icon={<Download size={16} />} label={t('savePptx')} onClick={() => { closeMenu(); save('pptx'); }} />}
                {active.kind === 'pdf' && <MenuItem icon={<Download size={16} />} label={t('savePdfFile')} onClick={() => { closeMenu(); save('pdf'); }} />}
                <MenuItem icon={<FileDown size={16} />} label={t('exportPdf')} onClick={() => { closeMenu(); startJob('pdf', active); }} />
                <div className="menu-sep" />
                <MenuItem icon={<Printer size={16} />} label={t('print')} hint="Ctrl+P" onClick={() => { closeMenu(); startJob('print', active); }} />
                <MenuItem icon={<X size={16} />} label={t('close')} onClick={() => { closeMenu(); close(active.id); }} />
              </>
            )}
          </Menu>
        )}
        <nav className="doc-tabs" aria-label="Open files">
          {files.map((f) => (
            <div key={f.id} className={`doc-tab k-${f.kind} ${f.id === activeId ? 'active' : ''}`} onClick={() => setActiveId(f.id)}>
              <AppIcon kind={f.kind} size={18} />
              {f.id === activeId ? (
                <input className="doc-name" value={f.name} onChange={(e) => rename(e.target.value)} onBlur={(e) => rename(e.target.value)} aria-label="File name" size={Math.max(6, f.name.length)} />
              ) : (
                <span className="doc-name">{f.name}</span>
              )}
              {dirty[f.id] && <span className="dirty" title="●">●</span>}
              <button className="tab-x" aria-label={t('close')} onClick={(e) => { e.stopPropagation(); close(f.id); }}>
                <X size={13} />
              </button>
            </div>
          ))}
        </nav>
        <div className="top-actions">
          {busy && <Loader2 className="spin" size={18} aria-label={t('working')} />}
          {active && (
            <>
              <button className="icon-btn" title={t('print')} aria-label={t('print')} onClick={() => startJob('print', active)}>
                <Printer size={18} />
              </button>
              <button className="icon-btn" title={t('exportPdf')} aria-label={t('exportPdf')} onClick={() => startJob('pdf', active)}>
                <FileDown size={18} />
              </button>
              <button className="save-btn" onClick={() => save(officeFormat!)}>
                <Save size={16} /> {t('save')}
              </button>
            </>
          )}
          {active && (
            <button className="lang-btn small" onClick={toggle} title="Language">
              <Languages size={16} /> {t('language')}
            </button>
          )}
          {!active && files.length > 0 && (
            <button className="icon-btn" onClick={() => setActiveId(files[files.length - 1].id)} aria-label="Back">
              <Home size={18} />
            </button>
          )}
        </div>
      </header>
      )}

      <main className="workspace">
        {!active && <StartScreen onNew={createNew} onOpen={open} onTemplate={add} />}
        {files.map((f) => (
          <div key={f.id} className="editor-host" hidden={f.id !== activeId}>
            {f.kind === 'doc' && <DocEditor data={f.data} onChange={updaterFor(f.id) as never} />}
            {f.kind === 'sheet' && f.id === activeId && <SheetEditor data={f.data} onChange={updaterFor(f.id) as never} />}
            {f.kind === 'slides' && f.id === activeId && <SlidesEditor data={f.data} onChange={updaterFor(f.id) as never} />}
            {f.kind === 'pdf' && f.id === activeId && (
              <Suspense fallback={<div className="editor-loading"><Loader2 className="spin" size={22} /></div>}>
                <PdfEditor data={f.data} name={f.name} onChange={updaterFor(f.id) as never} />
              </Suspense>
            )}
          </div>
        ))}
      </main>

      <input ref={openInput} type="file" accept={shell === 'mobile' ? undefined : ACCEPT} hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) open(f); e.target.value = ''; }} />
      {toast && <div className="toast" role="status">{toast}</div>}
      {job && (
        <PrintJob
          file={job.file}
          mode={job.mode}
          onDone={async (pdf, err) => {
            setJob(null);
            setBusy(false);
            if (err) { console.error(err); flash(String((err as Error)?.message ?? err)); return; }
            if (pdf && job.print) await printPdfBlob(pdf, `${job.file.name}.pdf`).catch((e) => flash(String((e as Error)?.message ?? e)));
            else if (pdf) await saveBlob(pdf, `${job.file.name}.pdf`);
          }}
        />
      )}
    </div>
  );
}
