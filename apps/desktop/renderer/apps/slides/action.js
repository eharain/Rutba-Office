// Insert → Action: what a click on the selected shape does in the show —
// PowerPoint's Action Settings, its Mouse Click side. Kept on the shape as
// PowerPoint keeps it, so Office runs the same show.

import React, { useState } from 'react';
import { Button, Dialog, t } from '@rutba/office-ui';

const TARGETS = [
  ['next', t('slides.action.next-slide')],
  ['previous', t('slides.action.previous-slide')],
  ['first', t('slides.action.first-slide')],
  ['last', t('slides.action.last-slide')],
  ['end', t('slides.action.end-show')],
  ['slide', t('slides.action.slide-2')],
  ['url', t('slides.action.url')],
];

/**
 * @param {{ action, outline, onApply, onClose }} props
 *   action   the shape's action now, as the deck reads it, or null
 *   outline  the deck's slides, `{ index, title }`, for "Slide…"
 */
export function ActionDialog({ action, outline = [], onApply, onClose }) {
  const known = action && TARGETS.some(([k]) => k === action.kind);
  const [on, setOn] = useState(Boolean(known));
  const [kind, setKind] = useState(known ? action.kind : 'next');
  const [slide, setSlide] = useState(action?.kind === 'slide' ? action.slide : 0);
  const [url, setUrl] = useState(action?.kind === 'url' ? action.url : 'https://');
  const urlOk = /^(https?:\/\/|mailto:)\S+$/i.test(url.trim()) && !/^https?:\/\/$/i.test(url.trim());
  const ok = !on || kind !== 'url' || urlOk;
  const apply = () => {
    if (!ok) return;
    if (!on) return onApply(null);
    if (kind === 'slide') return onApply({ kind, slide: Number(slide) });
    if (kind === 'url') return onApply({ kind, url: url.trim() });
    return onApply({ kind });
  };

  return (
    <Dialog
      title={t('slides.action.action-settings')}
      width={440}
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.ok')} className="sl-action-ok" disabled={!ok} onClick={apply} /></>}
    >
      <div className="sl-action">
        <div className="sl-action-lead">{t('slides.action.click-show')}</div>
        {action?.kind === 'other' ? (
          <p className="sl-action-note">{t('slides.action.this-shape-runs')} {action.action ? <code>{action.action}</code> : t('slides.action.something')} {t('slides.action.power-point-macro')}</p>
        ) : null}
        <label className="sl-action-row">
          <input type="radio" name="sl-action-on" className="sl-action-none" checked={!on} onChange={() => setOn(false)} />
          {t('common.none')}
        </label>
        <label className="sl-action-row">
          <input type="radio" name="sl-action-on" className="sl-action-link" checked={on} onChange={() => setOn(true)} />
          {t('slides.action.hyperlink-to')}
          <select className="rw-input sl-action-kind" value={kind} disabled={!on} onChange={(e) => setKind(e.target.value)}>
            {TARGETS.map(([k, label]) => <option key={k} value={k}>{label}</option>)}
          </select>
        </label>
        {on && kind === 'slide' ? (
          <label className="sl-action-row sl-action-sub">
            {t('slides.action.slide')}
            <select className="rw-input sl-action-slide" value={slide} onChange={(e) => setSlide(e.target.value)}>
              {outline.map((o) => <option key={o.index} value={o.index}>{`${o.index + 1}. ${o.title || t('slides.slide-number-2', { number: o.index + 1 })}`}</option>)}
            </select>
          </label>
        ) : null}
        {on && kind === 'url' ? (
          <label className="sl-action-row sl-action-sub">
            {t('common.address-2')}
            <input className="rw-input sl-action-url" value={url} spellCheck={false} onChange={(e) => setUrl(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') apply(); }} />
          </label>
        ) : null}
        {on && kind === 'url' && !urlOk ? <div className="sl-action-hint">{t('slides.action.web-address-starts')}</div> : null}
      </div>
    </Dialog>
  );
}

export const ACTION_CSS = `
.sl-action { display: flex; flex-direction: column; gap: 6px; font-size: 12.5px; }
.sl-action-lead { color: var(--ink-2); margin-bottom: 2px; }
.sl-action-row { display: flex; align-items: center; gap: 8px; }
.sl-action-row select, .sl-action-row input.rw-input { flex: 1; min-width: 0; }
.sl-action-sub { padding-left: 24px; }
.sl-action-note { margin: 0 0 4px; color: var(--ink-2); line-height: 1.45; }
.sl-action-hint { padding-left: 24px; color: var(--ink-3); font-size: 11.5px; }
`;
