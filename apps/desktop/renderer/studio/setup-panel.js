// The Set up panel (A5): the sources on the left, the frame they are arranged
// on in the middle with the scenes under it, the selected source's properties
// on the right. While live, what is changed here goes to Next and the picture
// on air stays as it is.

import React from 'react';
import { t } from '@rutba/office-ui';
import { Frame, Mirror } from './parts.js';
import { SceneStrip } from './scene-strip.js';
import { SourcesColumn } from './sources-column.js';
import { PropertiesColumn } from './properties.js';

/** @param {{ s: object }} p  the studio (window.js) */
export function SetupPanel({ s }) {
  const scene = s.scenes.find((sc) => sc.key === s.nextKey);
  return (
    <div className="st-panel" data-panel="setup">
      {s.live ? <div className="st-band live">{t('You are live. Changes here go to Next, and go on air when you Take.')}</div> : null}
      {s.recording && !s.live ? <div className="st-band rec">{t('Recording. Changes here show in the recording at once.')}</div> : null}
      <div className="st-setup-grid">
        <SourcesColumn s={s} thumbs={s.thumbs} />
        <div className="st-stage">
          <div className="st-frame-head">
            <b>{scene ? scene.name : t('The frame')}</b>
            <span className="st-frame-size">{s.settings.size === 720 ? '1280 × 720' : '1920 × 1080'}</span>
            <span className="st-dock-gap" />
            <span className="st-frame-hint">{t('Snaps to edges, centre and other sources. Alt drags freely.')}</span>
          </div>
          <div className="st-share">
            <Frame canvas={s.holding ? s.canvases.next : s.canvases.air} frameRef={s.editFrameRef} className="st-air st-editing" label={t('The frame')} tabIndex={0} onPointerDown={s.onFramePointer} onKeyDown={s.onFrameKey}>
              {s.overlay}
              {s.overlayShared}
              {s.live ? <div className="st-mirror"><span>{t('On air now')}</span><Mirror program={s.program} label={t('On air now')} /></div> : null}
            </Frame>
          </div>
          <div className="st-under-frame">
            <div className="st-scenes-head"><span className="st-scenes-label">{t('Scenes')}</span><span className="st-scenes-hint">{s.isLive ? t('1 to 9 switch while live. As many as you like.') : t('1 to 9 switch while recording. As many as you like.')}</span></div>
            <SceneStrip s={s} editing small />
          </div>
        </div>
        <PropertiesColumn s={s} />
      </div>
    </div>
  );
}
