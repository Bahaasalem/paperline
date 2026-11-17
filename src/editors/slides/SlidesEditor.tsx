import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import {
  AlignCenter, AlignLeft, AlignRight, ArrowRight, BringToFront, Circle, Copy, ImagePlus, Italic, Bold, Underline, LayoutTemplate,
  Minus, PaintBucket, Play, Plus, Presentation, Redo2, RectangleHorizontal, SendToBack, Square, SquareDashed, Trash2, Triangle,
  Type, Undo2, ChevronUp, ChevronDown, Palette, List, Columns2, AlignVerticalJustifyCenter, AlignVerticalJustifyStart,
  AlignVerticalJustifyEnd, PenLine,
} from 'lucide-react';
import { useI18n } from '../../i18n';
import { Btn, ColorBtn, Group, Menu, MenuItem, Ribbon, Select } from '../../ui/Ribbon';
import { useHistory } from '../useHistory';
import { FONTS } from '../doc/DocEditor';
import {
  applyTheme, makeSlide, textStyle, THEMES, themeOf, uid, type Deck, type Layout, type ShapeKind, type Slide, type SlideEl, type TextEl, type TextStyle,
} from './model';
import { ElView, ScaledSlide, SlideView, TextBody } from './SlideView';

interface Props {
  data: Deck;
  onChange: (d: Deck) => void;
}

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';
const HANDLES: Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];

export function SlidesEditor({ data, onChange }: Props) {
  const { t } = useI18n();
  const { commit, undo, redo } = useHistory(data, onChange);
  const [cur, setCur] = useState(0);
  const [selId, setSelId] = useState<string | null>(null);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [presenting, setPresenting] = useState<number | null>(null);
  const [fit, setFit] = useState(1);
  const stageRef = useRef<HTMLDivElement>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const live = useRef(data);
  live.current = data;

  const idx = Math.min(cur, data.slides.length - 1);
  const slide = data.slides[idx];
  const sel = slide?.els.find((e) => e.id === selId) ?? null;
  const theme = themeOf(data);

  const placeholderFor = useCallback(
    (el: TextEl) => (el.placeholder === 'title' ? t('clickTitle') : el.placeholder === 'subtitle' ? t('clickSubtitle') : el.placeholder === 'body' ? t('clickText') : undefined),
    [t],
  );

  // fit the slide into the stage
  useLayoutEffect(() => {
    const el = stageRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => {
      const pad = 48;
      setFit(Math.max(0.1, Math.min((el.clientWidth - pad) / data.w, (el.clientHeight - pad) / data.h)));
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, [data.w, data.h]);

  // ----- mutations -----
  const setSlides = (slides: Slide[]) => commit({ ...live.current, slides });
  const updateSlide = (fn: (s: Slide) => Slide, i = idx) => {
    const slides = live.current.slides.slice();
    slides[i] = fn(slides[i]);
    setSlides(slides);
  };
  const updateEl = (id: string, patch: Partial<SlideEl>, record = true) => {
    const slides = live.current.slides.slice();
    slides[idx] = { ...slides[idx], els: slides[idx].els.map((e) => (e.id === id ? ({ ...e, ...patch } as SlideEl) : e)) };
    if (record) commit({ ...live.current, slides });
    else onChange({ ...live.current, slides });
  };
  const updateStyle = (patch: Partial<TextStyle>) => {
    if (!sel || sel.type === 'image') return;
    updateEl(sel.id, { style: { ...sel.style, ...patch } } as Partial<SlideEl>);
  };
  const addEl = (el: SlideEl) => {
    updateSlide((s) => ({ ...s, els: [...s.els, el] }));
    setSelId(el.id);
  };
  const deleteSel = () => {
    if (!sel) return;
    updateSlide((s) => ({ ...s, els: s.els.filter((e) => e.id !== sel.id) }));
    setSelId(null);
  };

  const addSlide = (layout: Layout) => {
    const slides = data.slides.slice();
    slides.splice(idx + 1, 0, makeSlide(layout, data));
    setSlides(slides);
    setCur(idx + 1);
    setSelId(null);
  };
  const dupSlide = (i = idx) => {
    const slides = data.slides.slice();
    const s = data.slides[i];
    slides.splice(i + 1, 0, { ...s, id: uid(), els: s.els.map((e) => ({ ...e, id: uid() })) });
    setSlides(slides);
    setCur(i + 1);
  };
  const delSlide = (i = idx) => {
    if (data.slides.length <= 1) return;
    setSlides(data.slides.filter((_, j) => j !== i));
    setCur(Math.max(0, Math.min(i, data.slides.length - 2)));
    setSelId(null);
  };
  const moveSlide = (dir: -1 | 1) => {
    const j = idx + dir;
    if (j < 0 || j >= data.slides.length) return;
    const slides = data.slides.slice();
    [slides[idx], slides[j]] = [slides[j], slides[idx]];
    setSlides(slides);
    setCur(j);
  };

  const addText = () =>
    addEl({ id: uid(), type: 'text', text: '', x: data.w * 0.3, y: data.h * 0.4, w: data.w * 0.4, h: 60, style: textStyle({ color: theme.text, fontFace: theme.font }) });
  const addShape = (shape: ShapeKind) =>
    addEl({
      id: uid(), type: 'shape', shape, text: '',
      x: data.w / 2 - 90, y: data.h / 2 - (shape === 'line' ? 10 : 60), w: 180, h: shape === 'line' ? 20 : 120,
      fill: theme.accent, stroke: shape === 'line' ? theme.accent : undefined,
      style: textStyle({ color: '#ffffff', align: 'center', valign: 'middle', fontFace: theme.font }),
    });
  const addImage = (file: File) => {
    const r = new FileReader();
    r.onload = () => {
      const src = String(r.result);
      const img = new Image();
      img.onload = () => {
        let w = img.naturalWidth, h = img.naturalHeight;
        const max = Math.min(data.w * 0.6, data.h * 0.7 * (w / h));
        if (w > max) { h = (h * max) / w; w = max; }
        addEl({ id: uid(), type: 'image', src, x: (data.w - w) / 2, y: (data.h - h) / 2, w, h });
      };
      img.src = src;
    };
    r.readAsDataURL(file);
  };
  const reorder = (front: boolean) => {
    if (!sel) return;
    updateSlide((s) => {
      const rest = s.els.filter((e) => e.id !== sel.id);
      return { ...s, els: front ? [...rest, sel] : [sel, ...rest] };
    });
  };

  // ----- drag / resize -----
  const startDrag = (e: React.PointerEvent, el: SlideEl, handle?: Handle) => {
    if (e.button !== 0 || editingId === el.id) return;
    e.stopPropagation();
    e.preventDefault();
    setSelId(el.id);
    const sx = e.clientX, sy = e.clientY;
    const o = { x: el.x, y: el.y, w: el.w, h: el.h };
    const before = live.current;
    let moved = false;
    const keepRatio = el.type === 'image';
    const move = (ev: PointerEvent) => {
      const dx = (ev.clientX - sx) / fit, dy = (ev.clientY - sy) / fit;
      if (!moved && Math.abs(dx) + Math.abs(dy) < 2) return;
      moved = true;
      let { x, y, w, h } = o;
      if (!handle) { x += dx; y += dy; }
      else {
        if (handle.includes('e')) w = Math.max(10, o.w + dx);
        if (handle.includes('s')) h = Math.max(10, o.h + dy);
        if (handle.includes('w')) { w = Math.max(10, o.w - dx); x = o.x + o.w - w; }
        if (handle.includes('n')) { h = Math.max(10, o.h - dy); y = o.y + o.h - h; }
        if (keepRatio && handle.length === 2) {
          const r = o.w / o.h;
          if (w / h > r) w = h * r; else h = w / r;
          if (handle.includes('w')) x = o.x + o.w - w;
          if (handle.includes('n')) y = o.y + o.h - h;
        }
      }
      // snap to center lines
      const cx = x + w / 2, cy = y + h / 2;
      if (!handle && Math.abs(cx - data.w / 2) < 6) x = data.w / 2 - w / 2;
      if (!handle && Math.abs(cy - data.h / 2) < 6) y = data.h / 2 - h / 2;
      updateEl(el.id, { x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(h) }, false);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      if (moved) {
        // record a single undo step for the whole drag
        const after = live.current;
        onChange(before);
        commit(after);
      }
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  // ----- keyboard -----
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (presenting !== null || editingId) return;
      const tgt = e.target as HTMLElement;
      if (tgt.closest('input, textarea, select, [contenteditable="true"]')) return;
      if (!tgt.closest('.slides-editor')) return;
      const mod = e.ctrlKey || e.metaKey;
      if (mod && e.key.toLowerCase() === 'z') { e.preventDefault(); if (e.shiftKey) redo(); else undo(); return; }
      if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
      if (mod && e.key.toLowerCase() === 'd') { e.preventDefault(); if (sel) { const c = { ...sel, id: uid(), x: sel.x + 20, y: sel.y + 20 }; addEl(c); } else dupSlide(); return; }
      if (sel) {
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteSel(); return; }
        const step = e.shiftKey ? 10 : 1;
        const d = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key];
        if (d) { e.preventDefault(); updateEl(sel.id, { x: sel.x + d[0], y: sel.y + d[1] }); return; }
        if (e.key === 'Enter' && sel.type !== 'image') { e.preventDefault(); setEditingId(sel.id); return; }
        if (e.key === 'Escape') setSelId(null);
      } else {
        if (e.key === 'PageDown' || e.key === 'ArrowDown') { e.preventDefault(); setCur((c) => Math.min(data.slides.length - 1, c + 1)); }
        if (e.key === 'PageUp' || e.key === 'ArrowUp') { e.preventDefault(); setCur((c) => Math.max(0, c - 1)); }
        if (e.key === 'Delete') { e.preventDefault(); delSlide(); }
      }
      if (e.key === 'F5') { e.preventDefault(); setPresenting(e.shiftKey ? idx : 0); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  if (!slide) return null;
  const textSel = sel && sel.type !== 'image' ? sel : null;
  const st = textSel?.style;

  const layoutMenu = (onPick: (l: Layout) => void) => (
    <Menu trigger={(_o, toggle) => <Btn big icon={<Plus size={22} />} label={t('newSlide')} onClick={toggle} />}>
      {(close) =>
        (
          [
            ['title', t('layoutTitle'), <Presentation size={16} />],
            ['content', t('layoutContent'), <List size={16} />],
            ['two', t('layoutTwo'), <Columns2 size={16} />],
            ['blank', t('layoutBlank'), <SquareDashed size={16} />],
          ] as const
        ).map(([l, label, icon]) => <MenuItem key={l} icon={icon} label={label} onClick={() => { onPick(l); close(); }} />)
      }
    </Menu>
  );

  const textTools = (
    <Group label={t('font')}>
      <div className="r-col">
        <div className="r-row">
          <Select label={t('font')} value={st?.fontFace ?? theme.font} options={FONTS} onChange={(f) => updateStyle({ fontFace: f })} width={140} />
          <Select label={t('size')} value={st?.fontSize ?? 18} options={[10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 48, 60, 72, 96].map((x) => ({ value: x, label: String(x) }))} onChange={(v) => updateStyle({ fontSize: v })} width={60} />
        </div>
        <div className="r-row">
          <Btn icon={<Bold size={18} />} label={t('bold')} active={!!st?.bold} disabled={!textSel} onClick={() => updateStyle({ bold: !st?.bold })} />
          <Btn icon={<Italic size={18} />} label={t('italic')} active={!!st?.italic} disabled={!textSel} onClick={() => updateStyle({ italic: !st?.italic })} />
          <Btn icon={<Underline size={18} />} label={t('underline')} active={!!st?.underline} disabled={!textSel} onClick={() => updateStyle({ underline: !st?.underline })} />
          <ColorBtn icon={<Type size={16} />} label={t('textColor')} color={st?.color} onPick={(c) => c && updateStyle({ color: c })} />
          <Btn icon={<AlignLeft size={18} />} label={t('alignLeft')} active={st?.align === 'left'} disabled={!textSel} onClick={() => updateStyle({ align: 'left' })} />
          <Btn icon={<AlignCenter size={18} />} label={t('alignCenter')} active={st?.align === 'center'} disabled={!textSel} onClick={() => updateStyle({ align: 'center' })} />
          <Btn icon={<AlignRight size={18} />} label={t('alignRight')} active={st?.align === 'right'} disabled={!textSel} onClick={() => updateStyle({ align: 'right' })} />
          {textSel?.type === 'text' && <Btn icon={<List size={18} />} label={t('bullets')} active={!!textSel.bullets} onClick={() => updateEl(textSel.id, { bullets: !textSel.bullets } as Partial<SlideEl>)} />}
        </div>
      </div>
    </Group>
  );

  const shapeTools = sel && sel.type !== 'image' && (
    <Group label={t('shape')}>
      <div className="r-col">
        <div className="r-row">
          <ColorBtn icon={<PaintBucket size={16} />} label={t('fillColor')} color={sel.type === 'shape' ? sel.fill : sel.fill} allowNone onPick={(c) => updateEl(sel.id, { fill: c ?? (sel.type === 'shape' ? 'transparent' : undefined) } as Partial<SlideEl>)} />
          {sel.type === 'shape' && <ColorBtn icon={<PenLine size={16} />} label={t('outline')} color={sel.stroke} allowNone onPick={(c) => updateEl(sel.id, { stroke: c ?? undefined, strokeW: c ? sel.strokeW ?? 2 : 0 } as Partial<SlideEl>)} />}
        </div>
        <div className="r-row">
          <Btn icon={<AlignVerticalJustifyStart size={18} />} label="Top" active={st?.valign === 'top'} onClick={() => updateStyle({ valign: 'top' })} />
          <Btn icon={<AlignVerticalJustifyCenter size={18} />} label="Middle" active={st?.valign === 'middle'} onClick={() => updateStyle({ valign: 'middle' })} />
          <Btn icon={<AlignVerticalJustifyEnd size={18} />} label="Bottom" active={st?.valign === 'bottom'} onClick={() => updateStyle({ valign: 'bottom' })} />
        </div>
      </div>
    </Group>
  );

  const arrangeTools = sel && (
    <Group>
      <Btn icon={<BringToFront size={18} />} label={t('bringFront')} onClick={() => reorder(true)} />
      <Btn icon={<SendToBack size={18} />} label={t('sendBack')} onClick={() => reorder(false)} />
      <Btn icon={<Trash2 size={18} />} label={t('delete')} onClick={deleteSel} />
    </Group>
  );

  const homeTab = (
    <>
      <Group>
        <Btn icon={<Undo2 size={18} />} label={t('undo')} onClick={undo} />
        <Btn icon={<Redo2 size={18} />} label={t('redo')} onClick={redo} />
      </Group>
      <Group label={t('slide')}>
        {layoutMenu(addSlide)}
        <Btn icon={<Copy size={18} />} label={t('duplicate')} onClick={() => dupSlide()} />
        <Btn icon={<Trash2 size={18} />} label={t('delete')} disabled={data.slides.length <= 1} onClick={() => delSlide()} />
      </Group>
      {textTools}
      {shapeTools}
      {arrangeTools}
    </>
  );

  const insertTab = (
    <>
      <Group>
        <Btn big icon={<Type size={22} />} label={t('textBox')} onClick={addText} />
        <Btn big icon={<ImagePlus size={22} />} label={t('image')} onClick={() => fileRef.current?.click()} />
      </Group>
      <Group label={t('shape')}>
        <Btn icon={<Square size={18} />} label={t('rectangle')} onClick={() => addShape('rect')} />
        <Btn icon={<RectangleHorizontal size={18} />} label={t('rounded')} onClick={() => addShape('roundRect')} />
        <Btn icon={<Circle size={18} />} label={t('ellipse')} onClick={() => addShape('ellipse')} />
        <Btn icon={<Triangle size={18} />} label={t('triangle')} onClick={() => addShape('triangle')} />
        <Btn icon={<ArrowRight size={18} />} label={t('arrow')} onClick={() => addShape('rightArrow')} />
        <Btn icon={<Minus size={18} />} label={t('line')} onClick={() => addShape('line')} />
      </Group>
      {arrangeTools}
    </>
  );

  const designTab = (
    <>
      <Group label={t('theme')}>
        <div className="theme-row">
          {THEMES.map((th) => (
            <button key={th.id} type="button" className={`theme-chip ${data.theme === th.id ? 'on' : ''}`} style={{ background: th.bg, color: th.title, fontFamily: th.font }} title={th.id} onClick={() => commit(applyTheme(data, th.id))}>
              Aa
              <i style={{ background: th.accent }} />
            </button>
          ))}
        </div>
      </Group>
      <Group label={t('background')}>
        <ColorBtn icon={<Palette size={16} />} label={t('background')} color={slide.bg} onPick={(c) => c && updateSlide((s) => ({ ...s, bg: c }))} />
        <Btn icon={<LayoutTemplate size={18} />} label={`${t('background')} → *`} showLabel onClick={() => setSlides(data.slides.map((s) => ({ ...s, bg: slide.bg })))} />
      </Group>
    </>
  );

  const showTab = (
    <Group label={t('present')}>
      <Btn big icon={<Play size={22} />} label={t('fromStart')} onClick={() => setPresenting(0)} />
      <Btn big icon={<Presentation size={22} />} label={t('fromCurrent')} onClick={() => setPresenting(idx)} />
    </Group>
  );

  return (
    <div className="editor slides-editor">
      <Ribbon
        tabs={[
          { id: 'home', label: t('tab_home'), content: homeTab },
          { id: 'insert', label: t('tab_insert'), content: insertTab },
          { id: 'design', label: t('tab_design'), content: designTab },
          { id: 'show', label: t('tab_slideshow'), content: showTab },
        ]}
        extra={
          <button className="present-btn" onClick={() => setPresenting(idx)}>
            <Play size={14} /> {t('present')}
          </button>
        }
      />
      <input ref={fileRef} type="file" accept="image/*" hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) addImage(f); e.target.value = ''; }} />
      <div className="slides-body">
        <aside className="slide-rail" aria-label="Slides">
          {data.slides.map((s, i) => (
            <div key={s.id} className={`rail-item ${i === idx ? 'active' : ''}`} onClick={() => { setCur(i); setSelId(null); setEditingId(null); }} tabIndex={0}>
              <span className="rail-num">{i + 1}</span>
              <div className="rail-thumb">
                <ScaledSlide slide={s} w={data.w} h={data.h} width={168} />
              </div>
              {i === idx && (
                <div className="rail-actions">
                  <button title={t('moveUp')} aria-label={t('moveUp')} onClick={(e) => { e.stopPropagation(); moveSlide(-1); }}><ChevronUp size={14} /></button>
                  <button title={t('moveDown')} aria-label={t('moveDown')} onClick={(e) => { e.stopPropagation(); moveSlide(1); }}><ChevronDown size={14} /></button>
                  <button title={t('duplicate')} aria-label={t('duplicate')} onClick={(e) => { e.stopPropagation(); dupSlide(i); }}><Copy size={13} /></button>
                  <button title={t('delete')} aria-label={t('delete')} onClick={(e) => { e.stopPropagation(); delSlide(i); }}><Trash2 size={13} /></button>
                </div>
              )}
            </div>
          ))}
          <button className="rail-add" onClick={() => addSlide('content')}>
            <Plus size={16} /> {t('newSlide')}
          </button>
        </aside>
        <div className="slide-main">
          <div className="slide-stage" ref={stageRef} onPointerDown={() => { (document.activeElement as HTMLElement | null)?.blur?.(); setSelId(null); setEditingId(null); }}>
            <div className="slide-frame" style={{ width: data.w * fit, height: data.h * fit }}>
              <div className="slide-scale" style={{ transform: `scale(${fit})`, width: data.w, height: data.h }}>
                <div className="slide-surface editing" style={{ width: data.w, height: data.h, background: slide.bg }}>
                  {slide.els.map((el) => {
                    const isSel = el.id === selId;
                    const isEditing = el.id === editingId && el.type !== 'image';
                    return (
                      <div
                        key={el.id}
                        className={`el-wrap ${isSel ? 'selected' : ''} ${isEditing ? 'is-editing' : ''}`}
                        onPointerDown={(e) => startDrag(e, el)}
                        onDoubleClick={(e) => { e.stopPropagation(); if (el.type !== 'image') { setSelId(el.id); setEditingId(el.id); } }}
                      >
                        <ElView el={el} placeholder={el.type === 'text' ? placeholderFor(el) : undefined}>
                          {isEditing ? (
                            <InlineText
                              value={el.text}
                              onDone={(text) => { setEditingId(null); if (text !== el.text) updateEl(el.id, { text } as Partial<SlideEl>); }}
                            />
                          ) : el.type !== 'image' && el.type === 'shape' ? <TextBody el={el} /> : undefined}
                        </ElView>
                        {isSel && !isEditing && (
                          <div className="sel-frame" style={{ left: el.x, top: el.y, width: el.w, height: el.h, ['--inv' as string]: 1 / fit }}>
                            {HANDLES.map((h) => (
                              <span key={h} className={`handle ${h}`} onPointerDown={(e) => startDrag(e, el, h)} />
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>
          <textarea
            className="notes"
            dir="auto"
            placeholder={t('notesPlaceholder')}
            aria-label={t('notes')}
            value={slide.notes ?? ''}
            onChange={(e) => {
              const slides = data.slides.slice();
              slides[idx] = { ...slide, notes: e.target.value };
              onChange({ ...data, slides });
            }}
          />
        </div>
      </div>
      <div className="status-bar">
        <span>{t('slide')} {idx + 1} / {data.slides.length}</span>
        <span className="grow" />
        <span>{Math.round(fit * 100)}%</span>
      </div>
      {presenting !== null && <Presenter deck={data} start={presenting} onExit={(i) => { setPresenting(null); setCur(i); }} hint={t('exitShow')} />}
    </div>
  );
}

function InlineText({ value, onDone }: { value: string; onDone: (v: string) => void }) {
  const ref = useRef<HTMLTextAreaElement>(null);
  const [v, setV] = useState(value);
  useEffect(() => {
    const el = ref.current;
    if (el) { el.focus(); el.select(); }
  }, []);
  return (
    <textarea
      ref={ref}
      className="inline-text"
      dir="auto"
      value={v}
      onChange={(e) => setV(e.target.value)}
      onPointerDown={(e) => e.stopPropagation()}
      onBlur={() => onDone(v)}
      onKeyDown={(e) => {
        if (e.key === 'Escape') { e.preventDefault(); onDone(v); }
        if (e.key === 'Tab') {
          // indent / outdent bullet line
          e.preventDefault();
          const el = e.currentTarget;
          const start = v.lastIndexOf('\n', el.selectionStart - 1) + 1;
          const next = e.shiftKey ? (v.slice(start).startsWith('  ') ? v.slice(0, start) + v.slice(start + 2) : v) : v.slice(0, start) + '  ' + v.slice(start);
          setV(next);
        }
      }}
    />
  );
}

function Presenter({ deck, start, onExit, hint }: { deck: Deck; start: number; onExit: (i: number) => void; hint: string }) {
  const [i, setI] = useState(start);
  const [size, setSize] = useState({ w: window.innerWidth, h: window.innerHeight });
  const ref = useRef<HTMLDivElement>(null);
  const iRef = useRef(i);
  iRef.current = i;
  useEffect(() => {
    ref.current?.requestFullscreen?.().catch(() => {});
    const onResize = () => setSize({ w: window.innerWidth, h: window.innerHeight });
    const onKey = (e: KeyboardEvent) => {
      if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter'].includes(e.key)) { e.preventDefault(); setI((x) => Math.min(deck.slides.length - 1, x + 1)); }
      else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(e.key)) { e.preventDefault(); setI((x) => Math.max(0, x - 1)); }
      else if (e.key === 'Home') setI(0);
      else if (e.key === 'End') setI(deck.slides.length - 1);
      else if (e.key === 'Escape') exit();
    };
    const onFs = () => { if (!document.fullscreenElement) onExit(iRef.current); };
    const exit = () => { if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); else onExit(iRef.current); };
    window.addEventListener('resize', onResize);
    window.addEventListener('keydown', onKey);
    document.addEventListener('fullscreenchange', onFs);
    return () => {
      window.removeEventListener('resize', onResize);
      window.removeEventListener('keydown', onKey);
      document.removeEventListener('fullscreenchange', onFs);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  const s = Math.min(size.w / deck.w, size.h / deck.h);
  return (
    <div className="presenter" ref={ref} onClick={(e) => setI((x) => (e.clientX < size.w / 4 ? Math.max(0, x - 1) : Math.min(deck.slides.length - 1, x + 1)))}>
      <div style={{ width: deck.w * s, height: deck.h * s, overflow: 'hidden' }}>
        <div style={{ transform: `scale(${s})`, transformOrigin: '0 0', width: deck.w, height: deck.h }}>
          <SlideView slide={deck.slides[i]} w={deck.w} h={deck.h} />
        </div>
      </div>
      <div className="presenter-hint">{i + 1} / {deck.slides.length} · {hint}</div>
    </div>
  );
}
