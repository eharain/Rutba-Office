// The scenes, as large buttons with a live picture and a number (A2, A3): a
// click or the number key changes the scene while recording or live. A scene
// whose source is missing shows an amber mark, and a click on it opens Set up
// at that source.

import React, { useEffect, useRef } from 'react';
import { Icon, Menu, t } from '@rutba/office-ui';
import { applyScene, BACK_SOON, SOUND_ONLY } from '@rutba/studio/stack';
import { Program } from './program.js';

/** Each scene drawn small, a few times a second: the live thumbnails. */
function useThumbs(ref, media, stack, scenes) {
  const live = useRef({ stack, scenes });
  live.current = { stack, scenes };
  useEffect(() => {
    const painters = new Map();
    const timer = setInterval(() => {
      const root = ref.current;
      if (!root) return;
      for (const c of painters.keys()) if (!c.isConnected) painters.delete(c);
      for (const canvas of root.querySelectorAll('canvas[data-thumb]')) {
        const sc = live.current.scenes.find((x) => x.key === canvas.dataset.thumb);
        if (!sc) continue;
        let p = painters.get(canvas);
        if (!p) { p = new Program({ media, canvas }); painters.set(canvas, p); }
        p.stack = applyScene(live.current.stack, sc);
        try { p.draw(); } catch { /* the next tick draws it */ }
      }
    }, 450);
    return () => clearInterval(timer);
  }, [ref, media]);
}

/**
 * @param {object} p
 * @param {object} p.s  the studio (window.js)
 * @param {boolean} [p.editing]  in Set up: right-clicking changes a scene, and + saves the frame as one
 * @param {boolean} [p.small]  a little smaller, under the frame in Set up
 */
export function SceneStrip({ s, editing = false, small = false }) {
  const ref = useRef(null);
  const { scenes, stack, media } = s;
  useThumbs(ref, media, stack, scenes);
  const numbered = scenes.filter((sc) => sc.key !== BACK_SOON && sc.key !== SOUND_ONLY);
  return (
    <div ref={ref} className={`st-scenes${small ? ' small' : ''}`} role="toolbar" aria-label={t('common.scenes')}>
      {scenes.map((sc) => {
        const n = numbered.indexOf(sc);
        const lost = s.sceneLost(sc);
        const air = s.isLive && s.onAirKey === sc.key;
        const next = s.isLive && s.holding && s.nextKey === sc.key && !air;
        const on = s.nextKey === sc.key;
        const cls = ['st-scene', on ? 'on' : '', air ? 'air' : '', next ? 'next' : '', lost.length ? 'lost' : '', sc.back ? 'back' : '', sc.soundOnly ? 'sound' : ''].filter(Boolean).join(' ');
        return (
          <button key={sc.key} type="button" className={cls} data-scene={sc.key} data-tip={sc.name} aria-pressed={on}
            onClick={() => (lost.length ? s.openSetupAt(lost[0].id) : s.pickScene(sc))}
            onContextMenu={editing && s.sceneMenuItems(sc).length ? (e) => { e.preventDefault(); s.setSceneMenu({ x: e.clientX, y: e.clientY, sc }); } : undefined}
            draggable={editing && !sc.back && !sc.soundOnly}
            onDragStart={(e) => { e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', `scene:${sc.key}`); }}
            onDragOver={(e) => { if (editing && !sc.back && !sc.soundOnly) e.preventDefault(); }}
            onDrop={(e) => {
              if (!editing) return;
              const from = e.dataTransfer.getData('text/plain').replace(/^scene:/, '');
              e.preventDefault();
              if (from && from !== sc.key) s.moveSceneBy(from, sc.key);
            }}
          >
            <span className={`st-thumb${sc.soundOnly ? ' none' : ''}`}>
              {sc.soundOnly ? <Icon name="mic" size={26} /> : <canvas data-thumb={sc.key} className="st-thumb-canvas" width={320} height={180} />}
              {air ? <span className="st-tag air">{t('studio.scene-strip.on-air')}</span> : next ? <span className="st-tag next">{t('studio.scene-strip.next')}</span> : null}
              {lost.length ? <span className="st-lost-mark" data-tip={t('studio.scene-strip.name-not-found', { name: lost[0].name || t('studio.scene-strip.a-source') })}><Icon name="info" size={13} /></span> : null}
              {s.unsetScene(sc) ? <span className="st-thumb-note">{t('studio.scene-strip.choose-a-window')}</span> : null}
            </span>
            <span className="st-scene-foot">
              {n >= 0 && n < 9 ? <kbd>{n + 1}</kbd> : sc.back ? <kbd>B</kbd> : null}
              <span className="st-scene-name">{lost.length ? t('studio.scene-strip.name-not-found', { name: lost[0].name || sc.name }) : sc.name}</span>
              {s.isLive && (stack.follow[sc.key] || stack.rotate.keys.includes(sc.key)) ? (
                <span className="st-scene-tags">
                  {stack.follow[sc.key] ? <Icon name="mic" size={11} /> : null}
                  {stack.rotate.keys.includes(sc.key) ? <Icon name="refresh" size={11} /> : null}
                </span>
              ) : null}
            </span>
          </button>
        );
      })}
      {editing ? (
        <button type="button" className="st-scene st-scene-save" data-tip={t('studio.scene-strip.save-frame-shows')} aria-label={t('studio.scene-strip.save-frame-shows')} disabled={!s.anyShown} onClick={s.saveAsScene}>
          <Icon name="plus" size={15} />{t('studio.scene-strip.save-as-scene')}
        </button>
      ) : null}
      {s.sceneMenu ? <Menu x={s.sceneMenu.x} y={s.sceneMenu.y} items={s.sceneMenuItems(s.sceneMenu.sc)} onClose={() => s.setSceneMenu(null)} /> : null}
    </div>
  );
}

/** A row of screens or windows to choose from, under the scenes (A2): the pick is kept in the setup. */
export function PickRow({ s }) {
  const { pickRow } = s;
  if (!pickRow) return null;
  const items = (s.sources || []).filter((x) => x.kind === pickRow.kind);
  return (
    <div className="st-pick" role="listbox" aria-label={pickRow.kind === 'window' ? t('common.windows') : t('common.screens')}>
      <span className="st-pick-label">{pickRow.kind === 'window' ? t('studio.scene-strip.which-window') : t('studio.scene-strip.which-screen')}</span>
      {items.map((x) => (
        <button key={x.id} type="button" className={`st-pick-item${x.id === pickRow.current ? ' on' : ''}`} role="option" aria-selected={x.id === pickRow.current} data-source={x.id} data-tip={x.name} onClick={() => s.choosePick(x)}>
          {x.thumbnail ? <img src={x.thumbnail} alt="" draggable={false} /> : <Icon name={pickRow.kind === 'window' ? 'window' : 'screen'} size={22} />}
          <span>{x.name}</span>
        </button>
      ))}
      <button type="button" className="st-pick-close" aria-label={t('common.close')} data-tip={t('common.close')} onClick={() => s.setPickRow(null)}><Icon name="close" size={14} /></button>
    </div>
  );
}
