// Design → Slide Size: Standard, Widescreen, or the Slide Size dialog with
// PowerPoint's own list of sizes, a width and height of your own and the
// orientation. When the new shape is smaller one way, PowerPoint asks
// whether to Maximize the content (it keeps its size and may run off an
// edge) or Ensure Fit (it scales down to stay on the slide); so does this.

import React, { useState } from 'react';
import { Button, Dialog, Field, Select, t, numberTag } from '@rutba/office-ui';

const EMU_PER_CM = 360000;

/** PowerPoint's sizes, landscape, in EMU. */
export const SLIDE_SIZES = [
  ['wide', t('Widescreen'), 12192000, 6858000],
  ['standard', t('Standard (4:3)'), 9144000, 6858000],
  ['screen1610', t('On-screen Show (16:10)'), 9144000, 5715000],
  ['a4', t('A4 Paper (210 x 297 mm)'), 9906000, 6858000],
  ['letter', t('Letter Paper (8.5 x 11 in)'), 9144000, 6858000],
  ['35mm', t('35mm Slides'), 10287000, 6858000],
  ['banner', t('Banner'), 7315200, 914400],
];

/** Whether going from one size to another needs Maximize or Ensure Fit: the shape changes and the content no longer fits one way. */
export function needsFitChoice(from, cx, cy) {
  if (!from?.cx || !from?.cy) return false;
  const sx = cx / from.cx;
  const sy = cy / from.cy;
  return Math.abs(sx - sy) > 0.001 && Math.min(sx, sy) < 0.999;
}

const cm = (emu) => String(Math.round((emu / EMU_PER_CM) * 100) / 100);
const shown = (emu) => (Math.round((emu / EMU_PER_CM) * 100) / 100).toLocaleString(numberTag(), { maximumFractionDigits: 2 });
const presetOf = (cx, cy) => {
  const hit = SLIDE_SIZES.find(([, , w, h]) => (w === cx && h === cy) || (w === cy && h === cx));
  return hit ? hit[0] : 'custom';
};

/**
 * `size` the slides' size now ({ cx, cy }); `start` a size to begin from
 * ({ cx, cy }), or null for the deck's own. `onApply({ cx, cy, fit })`.
 */
export function SlideSizeDialog({ size, start = null, onClose, onApply }) {
  const begin = start || size;
  const [w, setW] = useState(() => cm(begin.cx));
  const [h, setH] = useState(() => cm(begin.cy));
  const [preset, setPreset] = useState(() => presetOf(begin.cx, begin.cy));
  // A side shown as one of the sizes known here (the deck's own, a preset's)
  // is that size exactly: 33.87 cm is a widescreen slide's 12,192,000 EMU,
  // not the 12,193,200 two places of centimetres would make it.
  const known = [begin.cx, begin.cy, size.cx, size.cy, ...SLIDE_SIZES.flatMap(([, , a, b]) => [a, b])];
  const emuOf = (text) => known.find((v) => cm(v) === String(text).trim()) ?? Math.round(Number(String(text).replace(',', '.')) * EMU_PER_CM);
  const cx = emuOf(w);
  const cy = emuOf(h);
  const valid = cx >= 914400 && cy >= 914400 && cx <= 51206400 && cy <= 51206400;
  const portrait = cy > cx;
  const choose = needsFitChoice(size, cx, cy);
  const pick = (key) => {
    setPreset(key);
    const hit = SLIDE_SIZES.find(([k]) => k === key);
    if (!hit) return;
    const [, , pw, ph] = hit;
    setW(cm(portrait ? ph : pw));
    setH(cm(portrait ? pw : ph));
  };
  const turn = (toPortrait) => {
    if (toPortrait === portrait) return;
    setW(h);
    setH(w);
  };
  const typed = (setter) => (e) => { setter(e.target.value); setPreset('custom'); };
  const go = (fit) => valid && onApply({ cx, cy, fit });

  return (
    <Dialog
      title={t('Slide Size')}
      width={460}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          {choose ? (
            <>
              <Button label={t('Maximize')} className="sl-size-maximize" disabled={!valid} title={t('Maximize: the content keeps its size and may run off the edges of the new slide')} onClick={() => go('maximize')} />
              <Button primary label={t('Ensure Fit')} className="sl-size-fit" disabled={!valid} title={t('Ensure Fit: the content scales down so all of it stays on the new slide')} onClick={() => go('ensure')} />
            </>
          ) : (
            <Button primary label={t('OK')} className="sl-size-ok" disabled={!valid} onClick={() => go('ensure')} />
          )}
        </>
      }
    >
      <div className="sl-size">
        <Field label={t('Slides sized for')}>
          <Select className="sl-size-preset" value={preset} onChange={(e) => pick(e.target.value)}>
            {SLIDE_SIZES.map(([key, label]) => <option key={key} value={key}>{label}</option>)}
            <option value="custom">{t('Custom')}</option>
          </Select>
        </Field>
        <div className="sl-size-dims">
          <Field label={t('Width (cm)')}>
            <input className="rw-input sl-size-w" inputMode="decimal" value={w} onChange={typed(setW)} />
          </Field>
          <Field label={t('Height (cm)')}>
            <input className="rw-input sl-size-h" inputMode="decimal" value={h} onChange={typed(setH)} />
          </Field>
        </div>
        <fieldset className="sl-size-turn">
          <legend>{t('Slides')}</legend>
          <label><input type="radio" name="sl-size-turn" checked={!portrait} onChange={() => turn(false)} />{t('Landscape')}</label>
          <label><input type="radio" name="sl-size-turn" className="sl-size-portrait" checked={portrait} onChange={() => turn(true)} />{t('Portrait')}</label>
        </fieldset>
        <div className="sl-size-note">
          {!valid
            ? t('A slide is between 2.54 cm and 142.24 cm each way.')
            : choose
              ? t('The new size is smaller one way. Maximize keeps the content as large as it is, so some may run off the edges; Ensure Fit scales it down so all of it stays on the slide.')
              : t('Now {width} by {height} cm. Everything on the slides is scaled and centred to suit the new size.', { width: shown(size.cx), height: shown(size.cy) })}
        </div>
      </div>
    </Dialog>
  );
}

export const SLIDE_SIZE_CSS = `
.sl-size { display: flex; flex-direction: column; gap: 10px; }
.sl-size-dims { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.sl-size-turn { border: 1px solid var(--line); border-radius: var(--r-1); padding: 6px 10px 8px; margin: 0; display: flex; gap: 18px; }
.sl-size-turn legend { font-size: 11.5px; color: var(--ink-2); padding: 0 4px; }
.sl-size-turn label { display: flex; align-items: center; gap: 6px; font-size: 12.5px; }
.sl-size-note { font-size: 12px; color: var(--ink-2); line-height: 1.45; }
`;
