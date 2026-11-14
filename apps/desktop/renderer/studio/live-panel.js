// The Live panel (A3): a bar across the top with the streams and the one Go
// live button, which stays in its place and becomes End broadcast; On air
// large with the scenes under it; Preview first puts a smaller Next left of
// it, and Take (Space) sends it on air. Auto is one menu.
//
// The bar is the window's (live-bar.js, kept going while Set up is open, or the
// broadcast would end with the panel) and only shown here.

import React, { useLayoutEffect, useRef, useState } from 'react';
import { Button, Icon, Menu, t } from '@rutba/office-ui';
import { ROTATE_EVERY } from '@rutba/studio/stack';
import { Frame, Seg } from './parts.js';
import { SceneStrip, PickRow } from './scene-strip.js';
import { BottomBar } from './bottom-bar.js';
import { clock, PictureStrip, PictureButton } from './shared.js';
import { Toggle } from './toggle.js';
import { FlipStrip } from './flip-strip.js';
import { SpokenStrip } from './spoken-run.js';

const AUTO_LABEL = (auto, every) => (auto === 'follow' ? t('studio.live-panel.auto-follow-speaker') : auto === 'rotate' ? t('studio.live-panel.auto-rotate-every', { n: every }) : t('studio.live-panel.auto-off'));

/** The strip across the top: where it goes, and the one button. */
function LiveBarSlot({ host }) {
  const ref = useRef(null);
  useLayoutEffect(() => {
    if (host && host.parentNode !== ref.current) ref.current.appendChild(host);
  }, [host]);
  return <div ref={ref} className="st-livebar" />;
}

/** @param {{ s: object }} p  the studio (window.js) */
export function LivePanel({ s }) {
  const [autoMenu, setAutoMenu] = useState(null);
  const { stack, holding, live } = s;
  const autoItems = [
    { label: t('common.off'), checked: stack.auto === 'off', run: () => s.setAuto('off') },
    { label: t('studio.follow-the-speaker'), checked: stack.auto === 'follow', run: s.startFollow },
    ...ROTATE_EVERY.map((n) => ({ label: t('studio.live-panel.rotate-every-n', { n }), checked: stack.auto === 'rotate' && stack.rotate.every === n, run: () => s.startRotate(n) })),
  ];
  const onAirLabel = holding ? t('studio.live-panel.on-air') : live ? t('studio.live-panel.on-air') : t('studio.live-panel.picture-go-out');
  const transition = <Seg value={stack.transition} options={[['cut', t('common.cut')], ['fade', t('common.fade')]]} label={t('common.transition')} className="st-transition" onChange={(v) => s.changeStack((st) => ({ ...st, transition: v }))} />;
  return (
    <div className="st-panel" data-panel="live">
      <LiveBarSlot host={s.liveBarHost} />
      <div className="st-stage">
        <div className={`st-share${holding ? ' duo' : ''}`}>
          {holding ? (
            <>
              <div className="st-pane st-next-pane">
                <span className="st-pane-label">{t('common.next')}</span>
                <div className="st-pane-frame"><Frame canvas={s.canvases.next} className="st-next" label={t('common.next')} /></div>
              </div>
              <div className="st-take">
                <button type="button" className="st-act primary st-take-btn" data-tip={t('studio.live-panel.take-send-next')} onClick={() => s.take(stack.transition)}><Icon name="chevronRight" size={16} />{t('studio.live-panel.take')}</button>
                <kbd>{t('common.space')}</kbd>
                {transition}
              </div>
            </>
          ) : null}
          <div className="st-pane on-air">
            <span className="st-pane-label"><i />{onAirLabel}</span>
            <div className="st-pane-frame">
              <Frame canvas={s.canvases.air} className={`st-air${live ? ' live' : ''}`} label={onAirLabel}>
                {live ? <span className="st-live-badge" data-phase={s.livePhase}><i />{t('studio.live')} {clock(s.liveFor)}</span> : null}
                {s.overlayShared}
              </Frame>
            </div>
          </div>
        </div>
        {/* The strip that flips a web page sits here, under On air. */}
        <div className="st-flip-slot"><FlipStrip stack={s.airStack} web={s.media.web} /></div>
        <div className="st-under-frame">
          <SpokenStrip s={s} />
          <div className="st-scenes-head">
            <span className="st-scenes-label">{t('common.scenes')}</span>
            <span className="st-scenes-hint">{holding ? t('studio.live-panel.click-puts-next') : t('studio.live-panel.click-puts-air')}</span>
            <span className="st-dock-gap" />
            <label className="st-check st-preview-btn"><input type="checkbox" checked={stack.previewFirst} onChange={s.togglePreview} />{t('studio.live-panel.preview-first')}</label>
            {!holding ? transition : null}
            <Button label={AUTO_LABEL(stack.auto, stack.rotate.every)} dropdown className={`st-auto${stack.auto !== 'off' ? ' on' : ''}`} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); setAutoMenu({ x: r.left, y: r.top - 8 - autoItems.length * 32 }); }} />
          </div>
          <SceneStrip s={s} />
          <PickRow s={s} />
          <PictureStrip s={s} />
        </div>
      </div>
      <BottomBar s={s}>
        <Toggle on={s.settings.liveCopy} label={t('studio.live-panel.keep-a-copy')} disabled={live} className="st-keep-copy" onChange={(v) => s.change({ liveCopy: v })} />
        {s.settings.liveCopy ? <Seg value={s.settings.liveCopyShape} options={[['wide', t('common.wide')], ['tall', t('common.tall')]]} label={t('studio.live-panel.shape-copy')} className="st-copy-shape" disabled={live} onChange={(v) => s.change({ liveCopyShape: v })} /> : null}
        <PictureButton s={s} />
      </BottomBar>
      {autoMenu ? <Menu x={autoMenu.x} y={Math.max(40, autoMenu.y)} items={autoItems} onClose={() => setAutoMenu(null)} /> : null}
    </div>
  );
}
