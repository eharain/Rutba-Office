// Set up's right column (A5): the selected source's place, size and shape, its
// crop and mirror, and what is its own: a camera's device, a title's words, a
// picture's file, a sound's level. With nothing selected, Live's Back soon and
// Auto, or how to begin.

import React, { useEffect, useState } from 'react';
import { Button, Field, Icon, Input, Select, basename, t } from '@rutba/office-ui';
import { Seg } from './parts.js';
import { Toggle } from './toggle.js';
import { KIND_ICON, itemName, kindName } from './list.js';

// Each Place, a small box in the frame as it would sit.
const PLACE_BOX = {
  fill: [0, 0, 100, 100], centre: [20, 20, 60, 60], left: [0, 0, 50, 100], right: [50, 0, 50, 100],
  tl: [4, 6, 34, 34], tr: [62, 6, 34, 34], bl: [4, 60, 34, 34], br: [62, 60, 34, 34],
};
const placeName = (p) => ({ fill: t('Fill'), centre: t('Centre'), left: t('Left side'), right: t('Right side'), tl: t('Top left'), tr: t('Top right'), bl: t('Bottom left'), br: t('Bottom right') }[p] || p);

/** A number field in percent: typed as a person types, applied when it is a number. */
function Percent({ label, value, onChange, className }) {
  const [text, setText] = useState(String(Math.round(value * 1000) / 10));
  useEffect(() => { setText(String(Math.round(value * 1000) / 10)); }, [value]);
  return (
    <Field label={label}>
      <Input className={`rw-input ${className}`} inputMode="decimal" value={text} onChange={(e) => {
        setText(e.target.value);
        const n = Number(e.target.value);
        if (e.target.value.trim() !== '' && Number.isFinite(n)) onChange(n / 100);
      }} />
    </Field>
  );
}

function SoundProps({ s, sound }) {
  const { barAct } = s;
  return (
    <div className="st-bar" role="group" aria-label={kindName(sound.kind)} data-kind={sound.kind}>
      <div className="st-col-actions">
        <Button icon={sound.live ? 'stop' : 'play'} label={sound.live ? t('Stop') : t('Start')} className="st-bar-run" onClick={() => barAct.sound({ live: !sound.live })} />
        <Button icon={sound.muted ? 'micOff' : 'volume'} label={sound.muted ? t('Unmute') : t('Mute')} className="st-bar-mute" pressed={sound.muted} onClick={() => barAct.sound({ muted: !sound.muted })} />
      </div>
      <Field label={t('Level')}>
        <input type="range" className="st-level" min={0} max={2} step={0.05} value={sound.gain} aria-label={t('Level')} data-tip={`${Math.round(sound.gain * 100)}%`} onChange={(e) => barAct.sound({ gain: Number(e.target.value) })} />
      </Field>
    </div>
  );
}

function ItemProps({ s, item }) {
  const { barAct, clipping } = s;
  const capture = ['screen', 'window', 'camera'].includes(item.kind);
  const shapes = [['rect', t('Rectangle')], ['rounded', t('Rounded')], ['circle', t('Circle')]];
  const box = item.box;
  const setBox = (patch) => s.changeFrame((st) => s.withBox(st, item.id, patch));
  const size = s.media.sizeOf(item);
  return (
    <div className="st-bar" role="group" aria-label={itemName(item)} data-kind={item.kind}>
      <Field label={t('Place')}>
        <div className="st-places" role="group" aria-label={t('Place')}>
          {Object.entries(PLACE_BOX).map(([p, [x, y, w, h]]) => (
            <button key={p} type="button" className={`st-place st-place-${p}`} data-tip={placeName(p)} aria-label={placeName(p)} onClick={() => barAct.place(p)}><i style={{ left: `${x}%`, top: `${y}%`, width: `${w}%`, height: `${h}%` }} /></button>
          ))}
        </div>
      </Field>
      <div className="st-fourfields">
        <Percent label={t('Left %')} className="st-box-x" value={box.x} onChange={(x) => setBox({ x })} />
        <Percent label={t('Top %')} className="st-box-y" value={box.y} onChange={(y) => setBox({ y })} />
        <Percent label={t('Width %')} className="st-box-w" value={box.w} onChange={(w) => setBox({ w })} />
        <Percent label={t('Height %')} className="st-box-h" value={box.h} onChange={(h) => setBox({ h })} />
      </div>
      <Toggle on={item.keep} label={t('Keep its shape')} className="st-keep" onChange={(keep) => s.changeStack((st) => s.withItem(st, item.id, { keep }))} />
      <Field label={t('Shape')}>
        <Seg value={item.shape} options={shapes} label={t('Shape')} className="st-shapes" onChange={barAct.shape} />
      </Field>
      <div className="st-col-actions">
        {item.kind !== 'title' ? <Button icon="crop" label={t('Edit crop')} className="st-clip" pressed={clipping} title={t('Edit crop: drag on the frame over the part to show')} onClick={barAct.clip} /> : null}
        {item.crop ? <Button label={t('Whole picture')} className="st-whole" onClick={barAct.whole} /> : null}
        {item.kind !== 'title' ? <Button icon="flip" label={t('Mirror')} className="st-mirror-btn" pressed={item.mirror} onClick={barAct.mirror} /> : null}
      </div>
      {item.kind === 'camera' ? (
        <Field label={t('Camera')}>
          <Select className="rw-select st-camera-select" value={item.ref || s.devices.cameras[0]?.id || ''} onChange={(e) => { const c = s.devices.cameras.find((x) => x.id === e.target.value); if (c) s.chooseSource(item, c); }}>
            {item.ref && !s.devices.cameras.some((c) => c.id === item.ref) ? <option value={item.ref} disabled>{t('Not found')}</option> : null}
            {s.devices.cameras.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </Select>
        </Field>
      ) : null}
      {capture && size ? <div className="st-col-note">{t('Picture from this source: {w} × {h}', { w: size.w, h: size.h })}</div> : null}
      {item.kind === 'title' ? (
        <>
          <Field label={t('Title')}><Input className="rw-input st-title" value={item.text} placeholder={t('Your name, or the title')} onChange={(e) => barAct.text({ text: e.target.value })} /></Field>
          <Field label={t('A line under it')}><Input className="rw-input st-subtitle" value={item.subtitle} placeholder={t('A line under it')} onChange={(e) => barAct.text({ subtitle: e.target.value })} /></Field>
        </>
      ) : null}
      {item.kind === 'picture' ? <Button icon="picture" label={t('Change…')} className="st-change-picture" onClick={barAct.change} /> : null}
    </div>
  );
}

/** Back soon and Auto, in Live's Set up with nothing selected (the LiveSetup artboard). */
function LiveExtras({ s }) {
  const { stack, choosePlaceholder, chooseMusic, changeStack } = s;
  return (
    <>
      <section className="st-col-section">
        <h3>{t('Back soon')}<kbd>B</kbd></h3>
        {stack.placeholder ? (
          <div className="st-col-row"><span>{basename(stack.placeholder)}</span><Button label={t('Use words')} onClick={() => changeStack((st) => ({ ...st, placeholder: '' }))} /></div>
        ) : (
          <>
            <Field label={t('The card says')}><Input className="rw-input st-placeholder-text" value={stack.placeholderText} placeholder={t('Back soon')} onChange={(e) => changeStack((st) => ({ ...st, placeholderText: e.target.value }))} /></Field>
            <Button icon="picture" label={t('Picture…')} onClick={choosePlaceholder} />
          </>
        )}
        <div className="st-col-row"><span>{stack.music ? basename(stack.music) : t('Music: a quiet chord, made here')}</span>
          {stack.music ? <Button label={t('Use the chord')} onClick={() => changeStack((st) => ({ ...st, music: '' }))} /> : null}
          <Button label={t('Music…')} onClick={chooseMusic} />
        </div>
      </section>
      <section className="st-col-section">
        <h3>{t('Auto')}</h3>
        <Seg value={stack.auto} options={[['off', t('Off')], ['follow', t('Follow the speaker')], ['rotate', t('Rotate')]]} label={t('Auto')} className="st-auto-seg" onChange={(v) => { if (v === 'follow') s.startFollow(); else if (v === 'rotate') s.startRotate(stack.rotate.every); else s.setAuto('off'); }} />
        <p className="st-col-note">{t('Sound levels and mute change what the audience hears straight away, as on a mixing desk. Pictures and places wait for Take.')}</p>
      </section>
    </>
  );
}

/** @param {{ s: object }} p  the studio (window.js) */
export function PropertiesColumn({ s }) {
  const { selItem, selSound, isLive } = s;
  const name = selItem ? (selItem.label || itemName(selItem)) : selSound ? selSound.name : '';
  const kind = selItem?.kind || selSound?.kind;
  return (
    <aside className="st-col st-props" aria-label={t('Properties')}>
      <div className="st-col-head"><span>{t('Properties')}</span>{kind ? <span className="st-col-name" data-tip={name}><Icon name={KIND_ICON[kind]} size={13} />{name}</span> : null}</div>
      <div className="st-col-body">
        {selSound ? <SoundProps s={s} sound={selSound} /> : selItem ? <ItemProps s={s} item={selItem} /> : isLive ? <LiveExtras s={s} /> : <p className="st-col-note">{t('Select a source in the list or on the frame to change it.')}</p>}
      </div>
    </aside>
  );
}
