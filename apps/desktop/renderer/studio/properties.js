// Set up's right column (A5): the selected source's place, size and shape, its
// crop and mirror, and what is its own: a camera's device, a title's words, a
// picture's file, a sound's level. With nothing selected, Live's Back soon and
// Auto, or how to begin.

import React, { useEffect, useState } from 'react';
import { Button, Field, Icon, Input, Select, basename, t } from '@rutba/office-ui';
import { Seg } from './parts.js';
import { Toggle } from './toggle.js';
import { KIND_ICON, itemName, kindName } from './list.js';
import { SpokenTitleProps } from './spoken-run.js';

// Each Place, a small box in the frame as it would sit.
const PLACE_BOX = {
  fill: [0, 0, 100, 100], centre: [20, 20, 60, 60], left: [0, 0, 50, 100], right: [50, 0, 50, 100],
  tl: [4, 6, 34, 34], tr: [62, 6, 34, 34], bl: [4, 60, 34, 34], br: [62, 60, 34, 34],
};
const placeName = (p) => ({ fill: t('common.fill'), centre: t('common.centre'), left: t('studio.properties.left-side'), right: t('studio.properties.right-side'), tl: t('common.top-left'), tr: t('common.top-right'), bl: t('common.bottom-left'), br: t('common.bottom-right') }[p] || p);

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
        <Button icon={sound.live ? 'stop' : 'play'} label={sound.live ? t('common.stop') : t('common.start')} className="st-bar-run" onClick={() => barAct.sound({ live: !sound.live })} />
        <Button icon={sound.muted ? 'micOff' : 'volume'} label={sound.muted ? t('common.unmute') : t('common.mute')} className="st-bar-mute" pressed={sound.muted} onClick={() => barAct.sound({ muted: !sound.muted })} />
      </div>
      <Field label={t('studio.level')}>
        <input type="range" className="st-level" min={0} max={2} step={0.05} value={sound.gain} aria-label={t('studio.level')} data-tip={`${Math.round(sound.gain * 100)}%`} onChange={(e) => barAct.sound({ gain: Number(e.target.value) })} />
      </Field>
    </div>
  );
}

function ItemProps({ s, item }) {
  const { barAct, clipping } = s;
  const capture = ['screen', 'window', 'camera'].includes(item.kind);
  const shapes = [['rect', t('common.rectangle')], ['rounded', t('common.rounded')], ['circle', t('common.circle')]];
  const box = item.box;
  const setBox = (patch) => s.changeFrame((st) => s.withBox(st, item.id, patch));
  const size = s.media.sizeOf(item);
  return (
    <div className="st-bar" role="group" aria-label={itemName(item)} data-kind={item.kind}>
      <Field label={t('studio.properties.place')}>
        <div className="st-places" role="group" aria-label={t('studio.properties.place')}>
          {Object.entries(PLACE_BOX).map(([p, [x, y, w, h]]) => (
            <button key={p} type="button" className={`st-place st-place-${p}`} data-tip={placeName(p)} aria-label={placeName(p)} onClick={() => barAct.place(p)}><i style={{ left: `${x}%`, top: `${y}%`, width: `${w}%`, height: `${h}%` }} /></button>
          ))}
        </div>
      </Field>
      <div className="st-fourfields">
        <Percent label={t('studio.properties.left')} className="st-box-x" value={box.x} onChange={(x) => setBox({ x })} />
        <Percent label={t('studio.properties.top')} className="st-box-y" value={box.y} onChange={(y) => setBox({ y })} />
        <Percent label={t('studio.properties.width')} className="st-box-w" value={box.w} onChange={(w) => setBox({ w })} />
        <Percent label={t('studio.properties.height')} className="st-box-h" value={box.h} onChange={(h) => setBox({ h })} />
      </div>
      <Toggle on={item.keep} label={t('studio.properties.keep-its-shape')} className="st-keep" onChange={(keep) => s.changeStack((st) => s.withItem(st, item.id, { keep }))} />
      <Field label={t('common.shape')}>
        <Seg value={item.shape} options={shapes} label={t('common.shape')} className="st-shapes" onChange={barAct.shape} />
      </Field>
      <div className="st-col-actions">
        {item.kind !== 'title' ? <Button icon="crop" label={t('studio.properties.edit-crop')} className="st-clip" pressed={clipping} title={t('studio.properties.edit-crop-drag')} onClick={barAct.clip} /> : null}
        {item.crop ? <Button label={t('studio.properties.whole')} className="st-whole" onClick={barAct.whole} /> : null}
        {item.kind !== 'title' ? <Button icon="flip" label={t('common.mirror')} className="st-mirror-btn" pressed={item.mirror} onClick={barAct.mirror} /> : null}
      </div>
      {item.kind === 'camera' ? (
        <Field label={t('common.camera')}>
          <Select className="rw-select st-camera-select" value={item.ref || s.devices.cameras[0]?.id || ''} onChange={(e) => { const c = s.devices.cameras.find((x) => x.id === e.target.value); if (c) s.chooseSource(item, c); }}>
            {item.ref && !s.devices.cameras.some((c) => c.id === item.ref) ? <option value={item.ref} disabled>{t('studio.not-found')}</option> : null}
            {s.devices.cameras.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </Select>
        </Field>
      ) : null}
      {capture && size ? <div className="st-col-note">{t('studio.properties.picture-source-w', { w: size.w, h: size.h })}</div> : null}
      {item.kind === 'title' ? (
        <>
          <Field label={t('common.title')}><Input className="rw-input st-title" value={item.text} placeholder={t('studio.name-title')} onChange={(e) => barAct.text({ text: e.target.value })} /></Field>
          <Field label={t('studio.properties.line-under')}><Input className="rw-input st-subtitle" value={item.subtitle} placeholder={t('studio.properties.line-under')} onChange={(e) => barAct.text({ subtitle: e.target.value })} /></Field>
          {s.spoken ? <SpokenTitleProps s={s} item={item} Field={Field} /> : null}
        </>
      ) : null}
      {item.kind === 'picture' ? <Button icon="picture" label={t('common.change-2')} className="st-change-picture" onClick={barAct.change} /> : null}
    </div>
  );
}

/** Back soon and Auto, in Live's Set up with nothing selected (the LiveSetup artboard). */
function LiveExtras({ s }) {
  const { stack, choosePlaceholder, chooseMusic, changeStack } = s;
  return (
    <>
      <section className="st-col-section">
        <h3>{t('studio.back-soon')}<kbd>B</kbd></h3>
        {stack.placeholder ? (
          <div className="st-col-row"><span>{basename(stack.placeholder)}</span><Button label={t('studio.properties.use-words')} onClick={() => changeStack((st) => ({ ...st, placeholder: '' }))} /></div>
        ) : (
          <>
            <Field label={t('studio.properties.the-card-says')}><Input className="rw-input st-placeholder-text" value={stack.placeholderText} placeholder={t('studio.back-soon')} onChange={(e) => changeStack((st) => ({ ...st, placeholderText: e.target.value }))} /></Field>
            <Button icon="picture" label={t('common.picture-3')} onClick={choosePlaceholder} />
          </>
        )}
        <div className="st-col-row"><span>{stack.music ? basename(stack.music) : t('studio.properties.music-quiet-chord')}</span>
          {stack.music ? <Button label={t('studio.properties.use-the')} onClick={() => changeStack((st) => ({ ...st, music: '' }))} /> : null}
          <Button label={t('studio.properties.music')} onClick={chooseMusic} />
        </div>
      </section>
      <section className="st-col-section">
        <h3>{t('common.auto')}</h3>
        <Seg value={stack.auto} options={[['off', t('common.off')], ['follow', t('studio.follow-the-speaker')], ['rotate', t('common.rotate')]]} label={t('common.auto')} className="st-auto-seg" onChange={(v) => { if (v === 'follow') s.startFollow(); else if (v === 'rotate') s.startRotate(stack.rotate.every); else s.setAuto('off'); }} />
        <p className="st-col-note">{t('studio.properties.sound-levels-mute')}</p>
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
    <aside className="st-col st-props" aria-label={t('common.properties')}>
      <div className="st-col-head"><span>{t('common.properties')}</span>{kind ? <span className="st-col-name" data-tip={name}><Icon name={KIND_ICON[kind]} size={13} />{name}</span> : null}</div>
      <div className="st-col-body">
        {selSound ? <SoundProps s={s} sound={selSound} /> : selItem ? <ItemProps s={s} item={selItem} /> : isLive ? <LiveExtras s={s} /> : <p className="st-col-note">{t('studio.properties.select-source-list')}</p>}
      </div>
    </aside>
  );
}
