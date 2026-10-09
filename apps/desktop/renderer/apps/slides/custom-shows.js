// Slide Show → Custom Slide Show → Custom Shows: PowerPoint's dialog of the
// deck's named shows — New, Edit, Copy, Remove and Show — and its Define
// Custom Show page: a name, the deck's slides, and the show's own order.
// Every change is written to the deck at once, as PowerPoint's is.

import React, { useState } from 'react';
import { Button, Dialog, Field, t, tn } from '@rutba/office-ui';

const label = (outline, i) => `${i + 1}. ${outline[i]?.title || t('Slide {number}', { number: i + 1 })}`;

export function CustomShowsDialog({ shows = [], outline = [], onSave, onPlay, onClose }) {
  const [list, setList] = useState(() => shows.map((s) => ({ ...s, slides: [...s.slides] })));
  const [sel, setSel] = useState(list.length ? 0 : -1);
  const [editing, setEditing] = useState(null);
  const [picked, setPicked] = useState([]);
  const [inSel, setInSel] = useState(-1);
  const [error, setError] = useState(null);

  const save = async (next) => {
    try {
      setError(null);
      await onSave(next);
      setList(next);
      return true;
    } catch (err) {
      setError(err.message || String(err));
      return false;
    }
  };
  const uniqueName = (base) => {
    let n = 1;
    let name = base;
    while (list.some((s) => s.name.toLowerCase() === name.toLowerCase())) name = `${base} ${++n}`;
    return name;
  };

  if (editing) {
    const move = (d) => {
      const to = inSel + d;
      if (inSel < 0 || to < 0 || to >= editing.slides.length) return;
      const slides = [...editing.slides];
      [slides[inSel], slides[to]] = [slides[to], slides[inSel]];
      setEditing({ ...editing, slides });
      setInSel(to);
    };
    const ok = async () => {
      const show = { id: editing.id, name: editing.name.trim(), slides: editing.slides };
      const next = editing.index == null ? [...list, show] : list.map((s, i) => (i === editing.index ? show : s));
      if (await save(next)) { setSel(editing.index ?? next.length - 1); setEditing(null); }
    };
    return (
      <Dialog
        title={t('Define Custom Show')}
        width={620}
        onClose={() => setEditing(null)}
        actions={<><Button label={t('Cancel')} onClick={() => setEditing(null)} /><Button primary label={t('OK')} className="sl-cs-ok" disabled={!editing.name.trim()} onClick={ok} /></>}
      >
        <div className="sl-cs">
          <Field label={t('Slide show name')}>
            <input className="rw-input sl-cs-name" value={editing.name} autoFocus onChange={(e) => setEditing({ ...editing, name: e.target.value })} />
          </Field>
          <div className="sl-cs-cols">
            <div>
              <div className="sl-cs-head">{t('Slides in presentation')}</div>
              <div className="sl-cs-list sl-cs-all" role="listbox" aria-multiselectable="true">
                {outline.map((o, i) => (
                  <button key={i} type="button" role="option" aria-selected={picked.includes(i)} className={`sl-cs-item${picked.includes(i) ? ' on' : ''}`} data-slide={i}
                    onClick={() => setPicked((p) => (p.includes(i) ? p.filter((x) => x !== i) : [...p, i].sort((a, b) => a - b)))}>
                    {label(outline, i)}
                  </button>
                ))}
              </div>
            </div>
            <div className="sl-cs-mid">
              <Button label={t('Add')} className="sl-cs-add" disabled={!picked.length} onClick={() => { setEditing({ ...editing, slides: [...editing.slides, ...picked] }); setPicked([]); }} />
              <Button label={t('Remove')} className="sl-cs-remove" disabled={inSel < 0} onClick={() => { setEditing({ ...editing, slides: editing.slides.filter((_, k) => k !== inSel) }); setInSel(-1); }} />
            </div>
            <div>
              <div className="sl-cs-head">{t('Slides in custom show')}</div>
              <div className="sl-cs-list sl-cs-in" role="listbox">
                {editing.slides.map((i, k) => (
                  <button key={`${i}-${k}`} type="button" role="option" aria-selected={k === inSel} className={`sl-cs-item${k === inSel ? ' on' : ''}`} data-slide={i} onClick={() => setInSel(k)}>
                    {`${k + 1}. ${outline[i]?.title || t('Slide {number}', { number: i + 1 })}`}
                  </button>
                ))}
              </div>
              <div className="sl-cs-order">
                <Button icon="chevronUp" title={t('Earlier in the show')} disabled={inSel <= 0} onClick={() => move(-1)} />
                <Button icon="chevronDown" title={t('Later in the show')} disabled={inSel < 0 || inSel >= editing.slides.length - 1} onClick={() => move(1)} />
              </div>
            </div>
          </div>
          {error ? <div className="sl-cs-error">{error}</div> : null}
        </div>
      </Dialog>
    );
  }

  const current = list[sel] || null;
  return (
    <Dialog title={t('Custom Shows')} width={460} onClose={onClose} actions={<Button label={t('Close')} onClick={onClose} />}>
      <div className="sl-cs sl-cs-main">
        <div className="sl-cs-list sl-cs-shows" role="listbox" aria-label={t('Custom shows')}>
          {list.length ? list.map((s, i) => (
            <button key={`${s.id}-${s.name}`} type="button" role="option" aria-selected={i === sel} className={`sl-cs-item${i === sel ? ' on' : ''}`} data-name={s.name} onClick={() => setSel(i)} onDoubleClick={() => onPlay(s)}>
              {s.name} <span className="sl-cs-count">{tn(s.slides.length, '{count} slide', '{count} slides')}</span>
            </button>
          )) : <div className="sl-cs-empty">{t('No custom shows yet. New makes one.')}</div>}
        </div>
        <div className="sl-cs-buttons">
          <Button label={t('New…')} className="sl-cs-new" onClick={() => { setEditing({ index: null, id: undefined, name: uniqueName(t('Custom Show')), slides: [] }); setPicked([]); setInSel(-1); }} />
          <Button label={t('Edit…')} disabled={!current} onClick={() => { setEditing({ index: sel, id: current.id, name: current.name, slides: [...current.slides] }); setPicked([]); setInSel(-1); }} />
          <Button label={t('Remove')} className="sl-cs-delete" disabled={!current} onClick={async () => { const next = list.filter((_, i) => i !== sel); if (await save(next)) setSel(next.length ? Math.min(sel, next.length - 1) : -1); }} />
          <Button label={t('Copy')} disabled={!current} onClick={async () => { const next = [...list, { name: uniqueName(t('Copy of {name}', { name: current.name })), slides: [...current.slides] }]; if (await save(next)) setSel(next.length - 1); }} />
          <Button primary label={t('Show')} className="sl-cs-play" disabled={!current || !current.slides.length} onClick={() => onPlay(current)} />
        </div>
        {error ? <div className="sl-cs-error">{error}</div> : null}
      </div>
    </Dialog>
  );
}

export const CUSTOM_SHOWS_CSS = `
.sl-cs { display: flex; flex-direction: column; gap: 8px; font-size: 12.5px; }
.sl-cs-main { flex-direction: row; align-items: stretch; }
.sl-cs-main .sl-cs-list { flex: 1; }
.sl-cs-buttons { display: flex; flex-direction: column; gap: 6px; min-width: 96px; }
.sl-cs-buttons .rw-btn { justify-content: center; border: 1px solid var(--line); }
.sl-cs-list { height: 220px; overflow: auto; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); }
.sl-cs-item { display: flex; justify-content: space-between; gap: 8px; width: 100%; border: 0; border-bottom: 1px solid var(--line-soft); background: none; font: inherit; color: var(--ink); text-align: left; padding: 5px 8px; cursor: pointer; }
.sl-cs-item:hover { background: var(--hover); }
.sl-cs-item.on { background: var(--selected); color: var(--accent); }
.sl-cs-count { color: var(--ink-3); font-size: 11.5px; }
.sl-cs-empty { padding: 10px; color: var(--ink-3); }
.sl-cs-cols { display: grid; grid-template-columns: 1fr auto 1fr; gap: 10px; align-items: start; }
.sl-cs-head { font-size: 11.5px; color: var(--ink-2); margin-bottom: 4px; }
.sl-cs-mid { display: flex; flex-direction: column; gap: 6px; margin-top: 90px; }
.sl-cs-mid .rw-btn, .sl-cs-order .rw-btn { border: 1px solid var(--line); }
.sl-cs-order { display: flex; gap: 4px; margin-top: 6px; }
.sl-cs-error { color: var(--bad); }
`;
