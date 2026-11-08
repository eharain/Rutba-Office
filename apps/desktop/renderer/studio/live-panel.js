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

const AUTO_LABEL = (auto, every) => (auto === 'follow' ? t('Auto: Follow the speaker') : auto === 'rotate' ? t('Auto: Rotate every {n} s', { n: every }) : t('Auto: Off'));

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
    { label: t('Off'), checked: stack.auto === 'off', run: () => s.setAuto('off') },
    { label: t('Follow the speaker'), checked: stack.auto === 'follow', run: s.startFollow },
    ...ROTATE_EVERY.map((n) => ({ label: t('Rotate every {n} s', { n }), checked: stack.auto === 'rotate' && stack.rotate.every === n, run: () => s.startRotate(n) })),
  ];
  const onAirLabel = holding ? t('On air') : live ? t('On air') : t('The picture as it will go out');
  const transition = <Seg value={stack.transition} options={[['cut', t('Cut')], ['fade', t('Fade')]]} label={t('Transition')} className="st-transition" onChange={(v) => s.changeStack((st) => ({ ...st, transition: v }))} />;
  return (
    <div className="st-panel" data-panel="live">
      <LiveBarSlot host={s.liveBarHost} />
      <div className="st-stage">
        <div className={`st-share${holding ? ' duo' : ''}`}>
          {holding ? (
            <>
              <div className="st-pane st-next-pane">
                <span className="st-pane-label">{t('Next')}</span>
                <div className="st-pane-frame"><Frame canvas={s.canvases.next} className="st-next" label={t('Next')} /></div>
              </div>
              <div className="st-take">
                <button type="button" className="st-act primary st-take-btn" data-tip={t('Take: send Next on air')} onClick={() => s.take(stack.transition)}><Icon name="chevronRight" size={16} />{t('Take')}</button>
                <kbd>{t('Space')}</kbd>
                {transition}
              </div>
            </>
          ) : null}
          <div className="st-pane on-air">
            <span className="st-pane-label"><i />{onAirLabel}</span>
            <div className="st-pane-frame">
              <Frame canvas={s.canvases.air} className={`st-air${live ? ' live' : ''}`} label={onAirLabel}>
                {live ? <span className="st-live-badge" data-phase={s.livePhase}><i />{t('LIVE')} {clock(s.liveFor)}</span> : null}
                {s.overlayShared}
              </Frame>
            </div>
          </div>
        </div>
        {/* The strip that flips a web page sits here, under On air. */}
        <div className="st-flip-slot" />
        <div className="st-under-frame">
          <div className="st-scenes-head">
            <span className="st-scenes-label">{t('Scenes')}</span>
            <span className="st-scenes-hint">{holding ? t('Click puts it in Next · 1 to 9') : t('Click puts it on air · 1 to 9')}</span>
            <span className="st-dock-gap" />
            <label className="st-check st-preview-btn"><input type="checkbox" checked={stack.previewFirst} onChange={s.togglePreview} />{t('Preview first')}</label>
            {!holding ? transition : null}
            <Button label={AUTO_LABEL(stack.auto, stack.rotate.every)} dropdown className={`st-auto${stack.auto !== 'off' ? ' on' : ''}`} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); setAutoMenu({ x: r.left, y: r.top - 8 - autoItems.length * 32 }); }} />
          </div>
          <SceneStrip s={s} />
          <PickRow s={s} />
          <PictureStrip s={s} />
        </div>
      </div>
      <BottomBar s={s}>
        <Toggle on={s.settings.liveCopy} label={t('Keep a copy')} disabled={live} className="st-keep-copy" onChange={(v) => s.change({ liveCopy: v })} />
        <PictureButton s={s} />
      </BottomBar>
      {autoMenu ? <Menu x={autoMenu.x} y={Math.max(40, autoMenu.y)} items={autoItems} onClose={() => setAutoMenu(null)} /> : null}
    </div>
  );
}
