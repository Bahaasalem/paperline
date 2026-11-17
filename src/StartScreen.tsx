import { useRef, useState } from 'react';
import { FileText, NotebookPen, Database, Newspaper, Workflow, GanttChart, Upload, Languages, LockKeyhole } from 'lucide-react';
import { useI18n, type Key } from './i18n';
import type { DocKind, OpenFile } from './types';
import { TEMPLATES } from './templates';
import { ACCEPT } from './io/files';
import { AppIcon, Wordmark } from './ui/Logo';

interface Props {
  onNew: (k: DocKind) => void;
  onOpen: (f: File) => void;
  onTemplate: (f: Omit<OpenFile, 'id'>) => void;
}

const MAIN: { kind: DocKind; label: Key; desc: Key }[] = [
  { kind: 'doc', label: 'kind_doc', desc: 'kind_doc_desc' },
  { kind: 'sheet', label: 'kind_sheet', desc: 'kind_sheet_desc' },
  { kind: 'slides', label: 'kind_slides', desc: 'kind_slides_desc' },
  { kind: 'pdf', label: 'kind_pdf', desc: 'kind_pdf_desc' },
];

const SOON: { icon: typeof FileText; label: Key; desc: Key }[] = [
  { icon: NotebookPen, label: 'kind_notes', desc: 'kind_notes_desc' },
  { icon: Database, label: 'kind_db', desc: 'kind_db_desc' },
  { icon: Newspaper, label: 'kind_pub', desc: 'kind_pub_desc' },
  { icon: Workflow, label: 'kind_diagram', desc: 'kind_diagram_desc' },
  { icon: GanttChart, label: 'kind_project', desc: 'kind_project_desc' },
];

function greeting(): Key {
  const h = new Date().getHours();
  return h < 5 ? 'helloEvening' : h < 12 ? 'helloMorning' : h < 18 ? 'helloAfternoon' : 'helloEvening';
}

export function StartScreen({ onNew, onOpen, onTemplate }: Props) {
  const { t, toggle, lang } = useI18n();
  const input = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);
  return (
    <div className="start">
      <header className="start-head">
        <Wordmark size={34} />
        <button className="lang-btn" onClick={toggle}>
          <Languages size={16} /> {t('language')}
        </button>
      </header>

      <section className="hero">
        <p className="hero-hi">{t(greeting())}</p>
        <h1>{t('createTitle')}</h1>
        <div className="kind-grid">
          {MAIN.map(({ kind, label, desc }) => (
            <button key={kind} className={`kind-card ${kind}`} onClick={() => onNew(kind)}>
              <AppIcon kind={kind} size={58} />
              <span className="kind-name">{t(label)}</span>
              <span className="kind-desc">{t(desc)}</span>
            </button>
          ))}
        </div>
      </section>

      <section>
        <h2>{t('templatesTitle')}</h2>
        <div className="tpl-row">
          {TEMPLATES.map((tp) => (
            <button key={tp.id} className={`tpl-card ${tp.kind}`} onClick={() => onTemplate(tp.build(lang))}>
              <span className={`tpl-art ${tp.kind === 'slides' ? 'wide' : ''}`} aria-hidden>
                <i /><i /><i /><i />
              </span>
              <span className="tpl-name">{tp.name[lang]}</span>
              <span className="tpl-kind">{t(tp.kind === 'doc' ? 'kind_doc' : tp.kind === 'sheet' ? 'kind_sheet' : 'kind_slides')}</span>
            </button>
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
          <Upload size={26} />
          <p>{t('openHint')}</p>
          <span className="drop-btn">{t('browse')}</span>
        </div>
        <input ref={input} type="file" accept={ACCEPT} hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) onOpen(f); e.target.value = ''; }} />
      </section>

      <section>
        <h2>{t('comingSoon')}</h2>
        <div className="soon-grid">
          {SOON.map(({ icon: Icon, label, desc }) => (
            <div key={label} className="soon-card" aria-disabled>
              <span className="soon-icon"><Icon size={18} strokeWidth={1.75} /></span>
              <div>
                <div className="soon-name">{t(label)}</div>
                <div className="soon-desc">{t(desc)}</div>
              </div>
            </div>
          ))}
        </div>
      </section>

      <footer className="start-foot">
        <LockKeyhole size={14} /> {t('privacyNote')}
      </footer>
    </div>
  );
}
