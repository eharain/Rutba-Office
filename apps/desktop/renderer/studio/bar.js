// The bar over the share: the selected source's controls (owner, 2026-11-02:
// "the control on the start stop trip change as user selects different
// viewable"). Start and Stop first, then Clip, Shape, Place and Mirror, the
// same for a screen, a window and a camera; a title has its words instead of
// a clip, a picture its file; a sound its mute and level.

import React, { useState } from 'react';
import { Icon, Input, Menu, t } from '@rutba/office-ui';
import { PLACES } from '@rutba/studio/stack';
import { KIND_ICON, itemName, kindName } from './list.js';

const PLACE_ICON = { fill: 'maximize', centre: 'mkRect', left: 'alignLeft', right: 'alignRight', tl: 'chevronUp', tr: 'chevronUp', bl: 'chevronDown', br: 'chevronDown' };
export const placeName = (p) => ({
  fill: t('Fill'), centre: t('Centre'), left: t('Left side'), right: t('Right side'), tl: t('Top left'), tr: t('Top right'), bl: t('Bottom left'), br: t('Bottom right'),
}[p] || p);

function Tool({ icon, label, on, onClick, className = '', title, disabled }) {
  return (
    <button type="button" className={`st-tool${on ? ' on' : ''}${label ? '' : ' icon'} ${className}`.trim()} aria-pressed={on === undefined ? undefined : Boolean(on)} title={title || label} aria-label={title || label} disabled={disabled} onClick={onClick}>
      {icon ? <Icon name={icon} size={15} /> : null}{label ? <span>{label}</span> : null}
    </button>
  );
}

const Sep = () => <span className="st-sep" aria-hidden="true" />;

/**
 * @param {object} p
 * @param {object|null} p.item  the selected picture item
 * @param {object|null} p.sound  the selected sound
 * @param {boolean} p.clipping
 * @param {object} p.act  live(on), clip(), whole(), shape(s), place(p), mirror(), text(patch), change(), copy(), remove(), sound(patch)
 */
export function SourceBar({ item, sound, clipping, act }) {
  const [placeMenu, setPlaceMenu] = useState(null);
  if (sound) {
    return (
      <div className="st-bar" role="toolbar" aria-label={kindName(sound.kind)} data-kind={sound.kind}>
        <span className="st-bar-name" title={sound.name || kindName(sound.kind)}><Icon name={KIND_ICON[sound.kind]} size={15} /><span>{sound.name || kindName(sound.kind)}</span></span>
        <Tool icon={sound.live ? 'stop' : 'play'} label={sound.live ? t('Stop') : t('Start')} className="st-bar-run" on={sound.live} onClick={() => act.sound({ live: !sound.live })} />
        <Sep />
        <Tool icon={sound.muted ? 'micOff' : 'volume'} label={sound.muted ? t('Unmute') : t('Mute')} className="st-bar-mute" on={sound.muted} onClick={() => act.sound({ muted: !sound.muted })} />
        <input type="range" className="st-level" min={0} max={2} step={0.05} value={sound.gain} aria-label={t('Level')} title={`${Math.round(sound.gain * 100)}%`} onChange={(e) => act.sound({ gain: Number(e.target.value) })} />
      </div>
    );
  }
  if (!item) {
    return <div className="st-bar" role="toolbar"><span className="st-bar-hint">{t('Select a source in the list or on the share to change it.')}</span></div>;
  }
  const capture = item.kind === 'screen' || item.kind === 'window' || item.kind === 'camera';
  const shapes = [['rect', 'mkRect', t('Rectangle')], ['rounded', 'mkRounded', t('Rounded')], ['circle', 'mkEllipse', t('Circle')]];
  return (
    <div className="st-bar" role="toolbar" aria-label={itemName(item)} data-kind={item.kind}>
      <span className="st-bar-name" title={itemName(item)}><Icon name={KIND_ICON[item.kind]} size={15} /><span>{itemName(item)}</span></span>
      {capture ? <Tool icon={item.live ? 'stop' : 'play'} label={item.live ? t('Stop') : t('Start')} className="st-bar-run" on={item.live} onClick={() => act.live(!item.live)} /> : null}
      {item.kind === 'title' ? (
        <>
          <Input className="st-title" value={item.text} placeholder={t('Your name, or the title')} aria-label={t('Title')} onChange={(e) => act.text({ text: e.target.value })} />
          <Input className="st-subtitle" value={item.subtitle} placeholder={t('A line under it')} aria-label={t('A line under it')} onChange={(e) => act.text({ subtitle: e.target.value })} />
        </>
      ) : null}
      {item.kind === 'picture' ? <Tool icon="picture" label={t('Change…')} onClick={act.change} /> : null}
      <Sep />
      {item.kind !== 'title' ? <Tool icon="crop" label={t('Clip')} className="st-clip" on={clipping} title={t('Clip: drag on the share over the part to show')} onClick={act.clip} /> : null}
      {item.crop ? <Tool icon="close" label={t('Whole picture')} className="st-whole" onClick={act.whole} /> : null}
      <span className="st-switch" role="radiogroup" aria-label={t('Shape')}>
        {shapes.map(([v, icon, label]) => (
          <button key={v} type="button" role="radio" aria-checked={item.shape === v} className={`st-shape-${v}${item.shape === v ? ' on' : ''}`} title={label} aria-label={label} onClick={() => act.shape(v)}><Icon name={icon} size={14} /></button>
        ))}
      </span>
      <Tool icon="layers" label={t('Place')} className="st-place" onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); setPlaceMenu({ x: r.left, y: r.bottom + 4 }); }} />
      {item.kind !== 'title' ? <Tool icon="flip" label={t('Mirror')} className="st-mirror" on={item.mirror} onClick={act.mirror} /> : null}
      <Sep />
      {capture ? <Tool icon="duplicate" title={t('Another copy, to clip or shape apart')} className="st-copy" onClick={act.copy} /> : null}
      <Tool icon="trash" title={t('Take out of the list')} className="st-remove" onClick={act.remove} />
      {placeMenu ? <Menu x={placeMenu.x} y={placeMenu.y} items={PLACES.map((p) => ({ label: placeName(p), icon: PLACE_ICON[p], run: () => act.place(p) }))} onClose={() => setPlaceMenu(null)} /> : null}
    </div>
  );
}
