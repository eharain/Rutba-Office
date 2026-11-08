// Set up's left column (A5): the sources in this scene, top of the list in
// front, each a row with its picture, its name, an eye and a menu; what cannot
// be found says so and offers Choose. Then the sounds, each with its level.

import React, { useEffect, useRef, useState } from 'react';
import { Button, Icon, Menu, t } from '@rutba/office-ui';
import { Meter, NameDialog } from './parts.js';
import { KIND_ICON, kindName, itemName } from './list.js';

/** A source's picture, small: live, or its still, its thumbnail, or its kind's icon. */
function Thumb({ item, media, thumb }) {
  const ref = useRef(null);
  const stream = media.streamOf(item);
  useEffect(() => {
    const v = ref.current;
    if (v && v.srcObject !== stream) { v.srcObject = stream; if (stream) v.play().catch(() => {}); }
  }, [stream]);
  if (item.kind === 'title') return <span className="st-src-pic"><Icon name="textbox" size={14} /></span>;
  if (item.kind === 'picture') return <span className="st-src-pic">{media.imageUrl(item.path) ? <img src={media.imageUrl(item.path)} alt="" draggable={false} /> : <Icon name="picture" size={14} />}</span>;
  const still = media.stillOf(item);
  return (
    <span className="st-src-pic">
      {stream ? <video ref={ref} muted playsInline autoPlay /> : still || thumb ? <img src={still || thumb} alt="" draggable={false} /> : <Icon name={KIND_ICON[item.kind]} size={14} />}
    </span>
  );
}

const size = (x) => (x?.w ? `${x.w} × ${x.h}` : '');

/** @param {{ s: object, thumbs: object }} p  the studio (window.js); the sources' thumbnails by id */
export function SourcesColumn({ s, thumbs }) {
  const { named: stack, media, selected, setSelected, act, missing } = s;
  const [menu, setMenu] = useState(null);
  const [renaming, setRenaming] = useState(null);
  const [drag, setDrag] = useState(null);
  const [over, setOver] = useState(null);
  const shown = stack.items.filter((it) => it.on);
  const hidden = stack.items.filter((it) => !it.on);
  const lost = (id) => missing.find((m) => m.id === id && m.scope === 'item') || null;

  const place = (e) => { const r = e.currentTarget.getBoundingClientRect(); return { x: r.left, y: r.bottom + 4 }; };
  const sameKind = (it) => {
    if (it.kind === 'camera') return s.devices.cameras.map((c) => ({ id: c.id, name: c.name }));
    return (s.sources || []).filter((x) => x.kind === it.kind);
  };
  const choices = (it) => {
    const list = sameKind(it);
    return list.length ? list.map((x) => ({ label: x.name, checked: x.id === it.ref, run: () => s.chooseSource(it, x) })) : [{ label: t('None to choose from'), disabled: true }];
  };
  const rowMenu = (it) => [
    ...(['screen', 'window', 'camera'].includes(it.kind) ? [
      { label: t('Choose…'), icon: KIND_ICON[it.kind], items: choices(it) },
      { label: it.live ? t('Stop') : t('Start'), icon: it.live ? 'stop' : 'play', run: () => act.live(it.id, !it.live) },
    ] : []),
    { label: t('Duplicate'), icon: 'duplicate', run: () => act.copyOf(it.id) },
    { label: t('Rename…'), icon: 'textbox', run: () => setRenaming(it) },
    { label: t('Remove'), icon: 'trash', disabled: s.onAirIds.includes(it.id), title: s.onAirIds.includes(it.id) ? t('It is on air') : undefined, run: () => act.removeItem(it.id) },
  ];

  const keys = (e, it) => {
    if (e.altKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
      e.preventDefault();
      const i = stack.items.indexOf(it);
      act.move(it.id, e.key === 'ArrowUp' ? Math.max(0, i - 1) : i + 2);
    } else if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      setSelected(it.id);
    }
  };

  const row = (it, index) => {
    const miss = lost(it.id);
    const state = media.stateOf(it);
    const capture = ['screen', 'window', 'camera'].includes(it.kind);
    const unset = it.kind === 'window' && !it.ref;
    const err = capture && !miss ? media.errorOf(it) : null;
    const name = it.label || itemName(it);
    const cls = ['st-card st-src', selected === it.id ? 'sel' : '', miss ? 'miss' : '', it.on ? '' : 'off', drag === it.id ? 'dragging' : '', over === index ? 'over-top' : ''].filter(Boolean).join(' ');
    return (
      <div key={it.id} className={cls} data-id={it.id} data-kind={it.kind} data-on={it.on ? '1' : '0'} data-state={state} data-tip={name} tabIndex={0} role="option" aria-selected={selected === it.id}
        aria-label={`${kindName(it.kind)}: ${name}${it.on ? '' : `, ${t('not shown')}`}`} draggable
        onDragStart={(e) => { setDrag(it.id); e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', it.id); }}
        onDragEnd={() => { setDrag(null); setOver(null); }}
        onDragOver={(e) => { e.preventDefault(); if (over !== index) setOver(index); }}
        onDrop={(e) => { e.preventDefault(); if (drag) act.move(drag, index); setDrag(null); setOver(null); }}
        onClick={() => setSelected(it.id)} onKeyDown={(e) => keys(e, it)}>
        <Thumb item={it} media={media} thumb={thumbs[it.ref]} />
        <span className="st-src-text">
          {miss ? <span className="st-src-miss">{t('Not found')}</span> : null}
          <span className="st-src-name">{unset ? t('A window, not chosen yet') : name}</span>
          {it.was ? <small>{t('Was {name}', { name: it.was })}</small> : it.kind === 'screen' && size(it.size) ? <small>{size(it.size)}</small> : null}
          {capture && !miss && state === 'error' ? <small className="st-src-err">{err}</small> : capture && !miss && !unset && !it.live ? <small>{t('Stopped')}</small> : null}
          {miss || unset ? <button type="button" className="st-src-choose" data-tip={t('Choose what to show')} onClick={(e) => { e.stopPropagation(); setMenu({ ...place(e), items: choices(it) }); }}>{t('Choose…')}</button> : null}
        </span>
        <button type="button" className={`st-pv-btn st-eye${it.on ? ' on' : ''}`} data-tip={it.on ? t('Take off the frame') : t('Show on the frame')} aria-label={it.on ? t('Take off the frame') : t('Show on the frame')} onClick={(e) => { e.stopPropagation(); act.show(it.id, !it.on); }}><Icon name={it.on ? 'eye' : 'eyeOff'} size={14} /></button>
        {capture && !miss && !unset ? <button type="button" className={`st-pv-btn st-run${it.live ? ' stop' : ''}`} data-tip={it.live ? t('Stop') : t('Start')} aria-label={it.live ? t('Stop') : t('Start')} onClick={(e) => { e.stopPropagation(); act.live(it.id, !it.live); }}><Icon name={it.live ? 'stop' : 'play'} size={13} /></button> : null}
        <button type="button" className="st-pv-btn st-more" aria-haspopup="menu" data-tip={t('More')} aria-label={t('More')} onClick={(e) => { e.stopPropagation(); setMenu({ ...place(e), items: rowMenu(it) }); }}><Icon name="more" size={14} /></button>
      </div>
    );
  };

  const soundRow = (x) => {
    const state = media.soundState(x.id);
    const err = media.soundError(x.id);
    return (
      <div key={x.id} className={`st-card st-snd${x.live ? '' : ' off'}${selected === x.id ? ' sel' : ''}`} data-id={x.id} data-kind={x.kind} data-state={state} tabIndex={0} role="option" aria-selected={selected === x.id} aria-label={`${kindName(x.kind)}: ${x.name}`} onClick={() => setSelected(x.id)} onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setSelected(x.id); } }}>
        <Icon name={KIND_ICON[x.kind]} size={16} />
        <span className="st-snd-name" data-tip={x.name}>{x.name}</span>
        <Meter media={media} id={x.id} live={x.live && state === 'live' && !x.muted} />
        <span className="st-snd-btns">
          <button type="button" className={`st-pv-btn st-mute${x.muted ? '' : ' on'}`} data-tip={x.muted ? t('Unmute') : t('Mute')} aria-label={x.muted ? t('Unmute') : t('Mute')} aria-pressed={x.muted} onClick={(e) => { e.stopPropagation(); act.sound(x.id, { muted: !x.muted }); }}><Icon name={x.muted ? 'micOff' : 'volume'} size={13} /></button>
          <button type="button" className={`st-pv-btn st-run${x.live ? ' stop' : ''}`} data-tip={x.live ? t('Stop') : t('Start')} aria-label={x.live ? t('Stop') : t('Start')} onClick={(e) => { e.stopPropagation(); act.sound(x.id, { live: !x.live }); }}><Icon name={x.live ? 'stop' : 'play'} size={13} /></button>
        </span>
        {err ? <div className="st-err" style={{ gridColumn: '1 / -1', padding: 0 }}>{err}</div> : null}
      </div>
    );
  };

  return (
    <aside className="st-col st-sources" aria-label={t('Sources')}>
      <div className="st-col-head">
        <span>{t('Sources')}</span>
        <Button icon="plus" label={t('Add')} dropdown className="st-add" onClick={(e) => setMenu({ ...place(e), items: act.addItems() })} />
      </div>
      <div className="st-list" role="listbox" aria-label={t('Sources')}>
        {shown.map((it) => row(it, stack.items.indexOf(it)))}
        {hidden.length ? <h3>{t('Not shown')}</h3> : null}
        {hidden.map((it) => row(it, stack.items.indexOf(it)))}
        <p className="st-col-note">{t('Top of the list is in front. Drag to change the order.')}</p>
        <h3>{t('Sound')}</h3>
        {stack.sounds.map(soundRow)}
      </div>
      {menu ? <Menu x={menu.x} y={menu.y} items={menu.items} onClose={() => setMenu(null)} /> : null}
      {renaming ? <NameDialog title={t('Rename')} label={t('Name')} value={renaming.label || itemName(renaming)} okLabel={t('Rename')} onOk={(name) => { act.rename(renaming.id, name); setRenaming(null); }} onClose={() => setRenaming(null)} /> : null}
    </aside>
  );
}
