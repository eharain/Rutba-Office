// Slide Show → Set Up Slide Show: who the show is for, how it runs, which
// slides it shows and how it moves on — PowerPoint's dialog, written into the
// deck as PowerPoint keeps it (ppt/presProps.xml), so Office runs the show
// the same way.

import React, { useState } from 'react';
import { Button, Dialog, Field, t } from '@rutba/office-ui';

const TYPES = [
  ['present', t('slides.setup.presented-speaker-full')],
  ['browse', t('slides.setup.browsed-individual')],
  ['kiosk', t('slides.setup.browsed-kiosk-full')],
];

export function SetUpShowDialog({ settings, count, onClose, onApply }) {
  const [s, setS] = useState(() => ({
    type: settings?.type || 'present',
    loop: Boolean(settings?.loop),
    narration: settings?.narration !== false,
    animation: settings?.animation !== false,
    useTimings: settings?.useTimings !== false,
    pen: settings?.pen || '#FF0000',
    some: Boolean(settings?.range),
    from: settings?.range?.from ?? 1,
    to: settings?.range?.to ?? count,
  }));
  const set = (patch) => setS((v) => ({ ...v, ...patch }));
  const kiosk = s.type === 'kiosk';
  const from = Math.max(1, Math.min(count, Number(s.from) || 1));
  const to = Math.max(from, Math.min(count, Number(s.to) || count));
  const apply = () => onApply({
    type: s.type,
    loop: kiosk ? true : s.loop,
    narration: s.narration,
    animation: s.animation,
    useTimings: s.useTimings,
    pen: s.pen,
    range: s.some ? { from, to } : null,
  });

  return (
    <Dialog
      title={t('slides.setup.set-up-show')}
      width={520}
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.ok')} className="sl-setup-ok" onClick={apply} /></>}
    >
      <div className="sl-setup">
        <fieldset>
          <legend>{t('slides.setup.show-type')}</legend>
          {TYPES.map(([value, label]) => (
            <label key={value} className="sl-setup-row">
              <input type="radio" name="sl-setup-type" value={value} checked={s.type === value} onChange={() => set({ type: value })} />
              {label}
            </label>
          ))}
        </fieldset>
        <fieldset>
          <legend>{t('slides.setup.show-options')}</legend>
          <label className="sl-setup-row" title={kiosk ? t('slides.setup.kiosk-always-loops') : undefined}>
            <input type="checkbox" className="sl-setup-loop" checked={kiosk || s.loop} disabled={kiosk} onChange={(e) => set({ loop: e.target.checked })} />
            {t('slides.setup.loop-continuously-until')}
          </label>
          <label className="sl-setup-row">
            <input type="checkbox" checked={!s.narration} onChange={(e) => set({ narration: !e.target.checked })} />
            {t('slides.setup.show-without-narration')}
          </label>
          <label className="sl-setup-row">
            <input type="checkbox" className="sl-setup-noanim" checked={!s.animation} onChange={(e) => set({ animation: !e.target.checked })} />
            {t('slides.setup.show-without')}
          </label>
          <Field label={t('slides.setup.pen-colour')}>
            <input type="color" className="sl-setup-pen" value={s.pen} onChange={(e) => set({ pen: e.target.value.toUpperCase() })} />
          </Field>
        </fieldset>
        <fieldset>
          <legend>{t('slides.setup.show-slides')}</legend>
          <label className="sl-setup-row">
            <input type="radio" name="sl-setup-slides" checked={!s.some} onChange={() => set({ some: false })} />
            {t('common.all')}
          </label>
          <label className="sl-setup-row">
            <input type="radio" name="sl-setup-slides" className="sl-setup-some" checked={s.some} onChange={() => set({ some: true })} />
            {t('common.from')}
            <input type="number" className="rw-input sl-setup-from" min={1} max={count} value={s.from} disabled={!s.some} onChange={(e) => set({ from: e.target.value })} />
            {t('slides.setup.to')}
            <input type="number" className="rw-input sl-setup-to" min={1} max={count} value={s.to} disabled={!s.some} onChange={(e) => set({ to: e.target.value })} />
          </label>
        </fieldset>
        <fieldset>
          <legend>{t('slides.setup.advance-slides')}</legend>
          <label className="sl-setup-row">
            <input type="radio" name="sl-setup-advance" checked={!s.useTimings} onChange={() => set({ useTimings: false })} />
            {t('slides.setup.manually')}
          </label>
          <label className="sl-setup-row">
            <input type="radio" name="sl-setup-advance" checked={s.useTimings} onChange={() => set({ useTimings: true })} />
            {t('slides.setup.using-timings-present')}
          </label>
        </fieldset>
      </div>
    </Dialog>
  );
}

export const SETUP_CSS = `
.sl-setup { display: grid; grid-template-columns: 1fr 1fr; gap: 10px 14px; }
.sl-setup fieldset { border: 1px solid var(--line); border-radius: var(--r-1); padding: 6px 10px 8px; margin: 0; }
.sl-setup legend { font-size: 11.5px; color: var(--ink-2); padding: 0 4px; }
.sl-setup-row { display: flex; align-items: center; gap: 6px; font-size: 12.5px; padding: 2px 0; }
.sl-setup-row input[type="number"] { width: 56px; }
`;
