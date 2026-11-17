import { useState } from 'react';
import { X } from 'lucide-react';
import { useI18n } from '../i18n';

export interface AskField {
  key: string;
  label: string;
  value: string;
  type?: 'text' | 'number' | 'textarea' | 'select' | 'color';
  options?: { value: string; label: string }[];
  /** Extra hint under the field. */
  hint?: string;
  rows?: number;
}
export interface Ask {
  title: string;
  fields: AskField[];
  onOk: (v: Record<string, string>) => void;
  wide?: boolean;
}

/** A small form dialog: a title, a few fields, Cancel and OK. */
export function AskDialog({ ask, onClose }: { ask: Ask; onClose: () => void }) {
  const { t } = useI18n();
  const [vals, setVals] = useState(() => Object.fromEntries(ask.fields.map((f) => [f.key, f.value])));
  const ok = () => { ask.onOk(vals); onClose(); };
  return (
    <div className="modal-back" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <form className={`modal ask-modal ${ask.wide ? 'wide' : ''}`} role="dialog" aria-label={ask.title} onSubmit={(e) => { e.preventDefault(); ok(); }} onKeyDown={(e) => e.key === 'Escape' && onClose()}>
        <header>
          <h3>{ask.title}</h3>
          <button type="button" className="icon-btn" aria-label={t('close')} onClick={onClose}><X size={18} /></button>
        </header>
        {ask.fields.map((f, i) => (
          <label key={f.key} className="ask-field">
            <span>{f.label}</span>
            {f.type === 'select' ? (
              <select value={vals[f.key]} onChange={(e) => setVals({ ...vals, [f.key]: e.target.value })}>
                {f.options!.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
              </select>
            ) : f.type === 'textarea' ? (
              <textarea autoFocus={i === 0} dir="auto" rows={f.rows ?? 4} value={vals[f.key]} onChange={(e) => setVals({ ...vals, [f.key]: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) ok(); }} />
            ) : f.type === 'color' ? (
              <input type="color" value={vals[f.key]} onChange={(e) => setVals({ ...vals, [f.key]: e.target.value })} />
            ) : (
              <input autoFocus={i === 0} dir="auto" type={f.type === 'number' ? 'number' : 'text'} value={vals[f.key]} onChange={(e) => setVals({ ...vals, [f.key]: e.target.value })} />
            )}
            {f.hint && <small>{f.hint}</small>}
          </label>
        ))}
        <div className="modal-actions">
          <button type="button" className="ghost-btn" onClick={onClose}>{t('cancel')}</button>
          <button type="submit" className="primary-btn">{t('ok')}</button>
        </div>
      </form>
    </div>
  );
}

