// The studio's list (owner, 2026-11-02): every source as a preview, with its
// kind's icon and its buttons on it, in one scrolling list. On the share, top
// first (the order is the stacking order, and a drag changes it); not shown
// (the screens, windows and cameras there are, whether or not added yet);
// and the sounds, the microphones and the computer's sound, each with its
// meter. A click selects; the bar over the share then holds its controls.

import React, { useEffect, useRef, useState } from 'react';
import { Icon, Button, Menu, t } from '@rutba/office-ui';
import { sourceKey } from './media.js';

export const KIND_ICON = { screen: 'screen', window: 'window', camera: 'camera', title: 'textbox', picture: 'picture', mic: 'mic', system: 'speaker' };

/** A kind's name, for what reads it out and for the source bar. */
export const kindName = (kind) => ({
  screen: t('Screen'), window: t('Window'), camera: t('Camera'), title: t('Title'), picture: t('Picture'), mic: t('Microphone'), system: t('Computer sound'),
}[kind] || kind);

/** An item's name as shown: its own, or its kind's. */
export const itemName = (it) => (it.kind === 'title' ? it.text || t('Title') : it.kind === 'picture' ? (it.path ? it.path.split(/[\\/]/).pop() : t('Picture')) : it.name || kindName(it.kind));

/** A live picture of a source, or its still, its thumbnail, or its kind's icon. */
function Preview({ item, media, thumb, hidden }) {
  const ref = useRef(null);
  const stream = item ? media.streamOf(item) : null;
  useEffect(() => {
    const v = ref.current;
    if (v && v.srcObject !== stream) { v.srcObject = stream; if (stream) v.play().catch(() => {}); }
  }, [stream]);
  let body;
  if (item?.kind === 'title') body = <div className="st-pv-title">{item.text || t('Title')}{item.subtitle ? <small style={{ display: 'block', fontWeight: 500, opacity: 0.8 }}>{item.subtitle}</small> : null}</div>;
  else if (item?.kind === 'picture') body = media.imageUrl(item.path) ? <img src={media.imageUrl(item.path)} alt="" draggable={false} /> : <Icon name="picture" size={28} />;
  else if (stream) body = <video ref={ref} muted playsInline autoPlay className={item.kind === 'camera' && item.mirror ? 'mirror' : ''} />;
  else {
    const still = item ? media.stillOf(item) : null;
    body = still || thumb ? <img src={still || thumb} alt="" draggable={false} /> : <Icon name={KIND_ICON[item?.kind] || 'screen'} size={28} />;
  }
  return <div className={`st-pv${hidden ? ' hidden' : ''}`}>{body}</div>;
}

/** A sound's meter, read ten times a second. */
function Level({ media, id, live }) {
  const bar = useRef(null);
  useEffect(() => {
    if (!live) { if (bar.current) bar.current.style.width = '0%'; return undefined; }
    const timer = setInterval(() => {
      const v = media.level(id);
      const el = bar.current;
      if (!el) return;
      el.style.width = `${Math.round(v * 100)}%`;
      el.dataset.level = v > 0.9 ? 'loud' : v > 0.02 ? 'on' : 'quiet';
    }, 100);
    return () => clearInterval(timer);
  }, [media, id, live]);
  return <span className="rec-meter st-meter" role="meter" aria-label={t('Level')} aria-valuemin={0} aria-valuemax={100}><i ref={bar} /></span>;
}

/**
 * @param {object} p
 * @param {object} p.stack
 * @param {object[]} p.spare  sources there are and not yet in the stack: { kind, ref, name, thumb, aspect }
 * @param {object[]} p.spareSounds  microphones not yet in the stack: { kind: 'mic', ref, name }
 * @param {object} p.media
 * @param {string|null} p.selected
 * @param {(id: string) => void} p.onSelect
 * @param {object} p.act  what the buttons do: show(id, on), live(id, on), move(id, index), addSpare(spare, { on, live, at }), sound(id, patch), addSound(spare), menu(e)
 */
export function SourceList({ stack, spare, spareSounds, media, selected, onSelect, act, thumbs }) {
  const [drag, setDrag] = useState(null); // { id } or { spare }
  const [over, setOver] = useState(null); // { index, half }
  const [addMenu, setAddMenu] = useState(null);
  const shown = stack.items.filter((it) => it.on);
  const hidden = stack.items.filter((it) => !it.on);

  /** Where a drop lands, as an index of the list (shown then hidden), from the card under the pointer. */
  const overCard = (e, index) => {
    e.preventDefault();
    const r = e.currentTarget.getBoundingClientRect();
    const half = e.clientY > r.top + r.height / 2 ? 1 : 0;
    if (!over || over.index !== index || over.half !== half) setOver({ index, half });
  };
  const drop = (e, index) => {
    e.preventDefault();
    const at = index;
    if (drag?.id) act.move(drag.id, at);
    else if (drag?.spare) act.addSpare(drag.spare, { on: at <= shown.length, live: true, at });
    setDrag(null);
    setOver(null);
  };
  const keys = (e, it) => {
    if (e.altKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
      e.preventDefault();
      // act.move takes the gap to drop into, as a drop does: below the next one is two gaps on.
      const i = stack.items.indexOf(it);
      act.move(it.id, e.key === 'ArrowUp' ? Math.max(0, i - 1) : i + 2);
    } else if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      onSelect(it.id);
    }
  };

  const card = (it, index) => {
    const state = media.stateOf(it);
    const capture = it.kind === 'screen' || it.kind === 'window' || it.kind === 'camera';
    const err = capture ? media.errorOf(it) : null;
    const cls = ['st-card', selected === it.id ? 'sel' : '', drag?.id === it.id ? 'dragging' : '', over?.index === index ? (over.half ? 'over-bottom' : 'over-top') : ''].filter(Boolean).join(' ');
    return (
      <div
        key={it.id}
        className={cls}
        data-id={it.id}
        data-kind={it.kind}
        data-on={it.on ? '1' : '0'}
        data-state={state}
        tabIndex={0}
        role="option"
        aria-selected={selected === it.id}
        aria-label={`${kindName(it.kind)}: ${itemName(it)}${it.on ? '' : `, ${t('not shown')}`}`}
        draggable
        onDragStart={(e) => { setDrag({ id: it.id }); e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', it.id); }}
        onDragEnd={() => { setDrag(null); setOver(null); }}
        onDragOver={(e) => overCard(e, index)}
        onDrop={(e) => drop(e, index + (over?.index === index && over.half ? 1 : 0))}
        onClick={() => onSelect(it.id)}
        onKeyDown={(e) => keys(e, it)}
      >
        <Preview item={it} media={media} thumb={thumbs[it.ref]} hidden={!it.on} />
        <span className="st-pv-name"><Icon name={KIND_ICON[it.kind]} size={13} />{itemName(it)}</span>
        {capture && state !== 'live' ? <span className="st-pv-state">{state === 'opening' ? t('Starting…') : state === 'error' ? t('Not available') : t('Stopped')}</span> : null}
        <span className="st-pv-btns">
          <button type="button" className={`st-pv-btn st-eye${it.on ? ' on' : ''}`} title={it.on ? t('Take off the share') : t('Show on the share')} aria-label={it.on ? t('Take off the share') : t('Show on the share')} onClick={(e) => { e.stopPropagation(); act.show(it.id, !it.on); }}>
            <Icon name={it.on ? 'eye' : 'eyeOff'} size={14} />
          </button>
          {capture ? (
            <button type="button" className={`st-pv-btn st-run${it.live ? ' stop' : ''}`} title={it.live ? t('Stop') : t('Start')} aria-label={it.live ? t('Stop') : t('Start')} onClick={(e) => { e.stopPropagation(); act.live(it.id, !it.live); }}>
              <Icon name={it.live ? 'stop' : 'play'} size={14} />
            </button>
          ) : null}
        </span>
        {err ? <div className="st-err">{err}</div> : null}
      </div>
    );
  };

  const spareCard = (s, index) => (
    <div
      key={`spare:${s.kind}:${s.ref}`}
      className={`st-card st-spare${over?.index === index ? (over.half ? ' over-bottom' : ' over-top') : ''}`}
      data-spare={s.ref}
      data-kind={s.kind}
      data-tip={s.name}
      tabIndex={0}
      role="option"
      aria-selected={false}
      aria-label={`${kindName(s.kind)}: ${s.name}, ${t('not shown')}`}
      draggable
      onDragStart={(e) => { setDrag({ spare: s }); e.dataTransfer.effectAllowed = 'copy'; e.dataTransfer.setData('text/plain', s.ref); }}
      onDragEnd={() => { setDrag(null); setOver(null); }}
      onDragOver={(e) => overCard(e, index)}
      onDrop={(e) => drop(e, index)}
      onClick={() => act.addSpare(s, { on: false, live: true, select: true })}
      onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); act.addSpare(s, { on: true, live: true, select: true }); } }}
    >
      <div className="st-pv hidden">{s.thumb ? <img src={s.thumb} alt="" draggable={false} /> : <Icon name={KIND_ICON[s.kind]} size={28} />}</div>
      <span className="st-pv-name"><Icon name={KIND_ICON[s.kind]} size={13} />{s.name}</span>
      <span className="st-pv-btns">
        <button type="button" className="st-pv-btn st-eye" title={t('Show on the share')} aria-label={t('Show on the share')} onClick={(e) => { e.stopPropagation(); act.addSpare(s, { on: true, live: true, select: true }); }}>
          <Icon name="eyeOff" size={14} />
        </button>
        <button type="button" className="st-pv-btn st-run" title={t('Start')} aria-label={t('Start')} onClick={(e) => { e.stopPropagation(); act.addSpare(s, { on: false, live: true }); }}>
          <Icon name="play" size={14} />
        </button>
      </span>
    </div>
  );

  const soundCard = (s) => {
    const state = media.soundState(s.id);
    const err = media.soundError(s.id);
    const name = s.name || kindName(s.kind);
    return (
      <div key={s.id} className={`st-card st-snd${s.live ? '' : ' off'}${selected === s.id ? ' sel' : ''}`} data-id={s.id} data-kind={s.kind} data-state={state} tabIndex={0} role="option" aria-selected={selected === s.id} aria-label={`${kindName(s.kind)}: ${name}`}
        onClick={() => onSelect(s.id)} onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onSelect(s.id); } }}>
        <Icon name={KIND_ICON[s.kind]} size={18} />
        <span className="st-snd-name">{name}</span>
        <Level media={media} id={s.id} live={s.live && state === 'live' && !s.muted} />
        <span className="st-snd-btns">
          <button type="button" className={`st-pv-btn st-mute${s.muted ? '' : ' on'}`} title={s.muted ? t('Unmute') : t('Mute')} aria-label={s.muted ? t('Unmute') : t('Mute')} aria-pressed={s.muted} onClick={(e) => { e.stopPropagation(); act.sound(s.id, { muted: !s.muted }); }}>
            <Icon name={s.muted ? 'micOff' : 'volume'} size={14} />
          </button>
          <button type="button" className={`st-pv-btn st-run${s.live ? ' stop' : ''}`} title={s.live ? t('Stop') : t('Start')} aria-label={s.live ? t('Stop') : t('Start')} onClick={(e) => { e.stopPropagation(); act.sound(s.id, { live: !s.live }); }}>
            <Icon name={s.live ? 'stop' : 'play'} size={14} />
          </button>
        </span>
        {err ? <div className="st-err" style={{ gridColumn: '1 / -1', padding: 0 }}>{err}</div> : null}
      </div>
    );
  };

  const spareSound = (s) => (
    <div key={`spare:${s.ref}`} className="st-card st-snd off st-spare" data-spare={s.ref} data-kind="mic" tabIndex={0} role="option" aria-selected={false} aria-label={`${kindName('mic')}: ${s.name}`}
      onClick={() => act.addSound(s)} onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); act.addSound(s); } }}>
      <Icon name="mic" size={18} />
      <span className="st-snd-name">{s.name}</span>
      <span className="rec-meter st-meter"><i /></span>
      <span className="st-snd-btns">
        <button type="button" className="st-pv-btn st-run" title={t('Start')} aria-label={t('Start')} onClick={(e) => { e.stopPropagation(); act.addSound(s); }}><Icon name="play" size={14} /></button>
      </span>
    </div>
  );

  // The drop zone at the foot of each part, so a drag can land last in it.
  const zone = (index, key) => (
    <div key={key} className={`st-drop${over?.index === index && over.zone ? ' over' : ''}`}
      onDragOver={(e) => { e.preventDefault(); if (!over || over.index !== index || !over.zone) setOver({ index, half: 0, zone: true }); }}
      onDrop={(e) => drop(e, index)} />
  );

  return (
    <div className="st-list" role="listbox" aria-label={t('Sources')}>
      <h3>{t('On the share')}<span>{t('Top first')}</span></h3>
      {shown.map((it, i) => card(it, i))}
      {zone(shown.length, 'z-shown')}
      <h3>{t('Not shown')}</h3>
      {hidden.map((it, i) => card(it, shown.length + i))}
      {spare.map((s, i) => spareCard(s, shown.length + hidden.length + i))}
      {zone(stack.items.length, 'z-hidden')}
      <h3>{t('Sound')}</h3>
      {stack.sounds.map(soundCard)}
      {spareSounds.map(spareSound)}
      <Button icon="plus" label={t('Add')} className="st-add" onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); setAddMenu({ x: r.left, y: r.bottom + 4 }); }} />
      {addMenu ? <Menu x={addMenu.x} y={addMenu.y} items={act.addItems()} onClose={() => setAddMenu(null)} /> : null}
    </div>
  );
}
