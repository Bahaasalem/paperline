import { useRef, useState } from 'react';
import { FileText, FileType2, Sheet, Presentation, NotebookPen, Database, Newspaper, Workflow, GanttChart, Upload, Languages } from 'lucide-react';
import { useI18n, type Key } from './i18n';
import type { DocKind } from './types';
import { ACCEPT } from './io/files';
import { Logo } from './ui/Logo';

interface Props {
  onNew: (k: DocKind) => void;
  onOpen: (f: File) => void;
}

const MAIN: { kind: DocKind; icon: typeof FileText; label: Key; desc: Key; cls: string }[] = [
  { kind: 'doc', icon: FileText, label: 'kind_doc', desc: 'kind_doc_desc', cls: 'doc' },
  { kind: 'sheet', icon: Sheet, label: 'kind_sheet', desc: 'kind_sheet_desc', cls: 'sheet' },
  { kind: 'slides', icon: Presentation, label: 'kind_slides', desc: 'kind_slides_desc', cls: 'slides' },
  { kind: 'pdf', icon: FileType2, label: 'kind_pdf', desc: 'kind_pdf_desc', cls: 'pdf' },
];

const SOON: { icon: typeof FileText; label: Key; desc: Key }[] = [
  { icon: NotebookPen, label: 'kind_notes', desc: 'kind_notes_desc' },
  { icon: Database, label: 'kind_db', desc: 'kind_db_desc' },
  { icon: Newspaper, label: 'kind_pub', desc: 'kind_pub_desc' },
  { icon: Workflow, label: 'kind_diagram', desc: 'kind_diagram_desc' },
  { icon: GanttChart, label: 'kind_project', desc: 'kind_project_desc' },
];

export function StartScreen({ onNew, onOpen }: Props) {
  const { t, toggle } = useI18n();
  const input = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);
  return (
    <div className="start">
      <header className="start-head">
        <div className="brand">
          <Logo size={40} />
          <div>
            <h1>{t('appName')}</h1>
            <p>{t('tagline')}</p>
          </div>
        </div>
        <button className="lang-btn" onClick={toggle}>
          <Languages size={16} /> {t('language')}
        </button>
      </header>

      <section>
        <h2>{t('createTitle')}</h2>
        <div className="kind-grid">
          {MAIN.map(({ kind, icon: Icon, label, desc, cls }) => (
            <button key={kind} className={`kind-card ${cls}`} onClick={() => onNew(kind)}>
              <span className="kind-art">
                <Icon size={34} strokeWidth={1.75} />
              </span>
              <span className="kind-name">{t(label)}</span>
              <span className="kind-desc">{t(desc)}</span>
            </button>
          ))}
        </div>
        <div className="soon-grid">
          {SOON.map(({ icon: Icon, label, desc }) => (
            <div key={label} className="soon-card" aria-disabled>
              <Icon size={20} strokeWidth={1.75} />
              <div>
                <div className="soon-name">
                  {t(label)} <em>{t('comingSoon')}</em>
                </div>
                <div className="soon-desc">{t(desc)}</div>
              </div>
            </div>
          ))}
        </div>
      </section>

      <section>
        <h2>{t('openTitle')}</h2>
        <div
          className={`drop ${over ? 'over' : ''}`}
          onClick={() => input.current?.click()}
          onDragOver={(e) => { e.preventDefault(); setOver(true); }}
          onDragLeave={() => setOver(false)}
          onDrop={(e) => {
            e.preventDefault();
            setOver(false);
            const f = e.dataTransfer.files[0];
            if (f) onOpen(f);
          }}
          role="button"
          tabIndex={0}
          onKeyDown={(e) => (e.key === 'Enter' || e.key === ' ') && input.current?.click()}
        >
          <Upload size={28} />
          <p>{t('openHint')}</p>
          <span className="drop-btn">{t('browse')}</span>
        </div>
        <input ref={input} type="file" accept={ACCEPT} hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) onOpen(f); e.target.value = ''; }} />
      </section>
    </div>
  );
}
