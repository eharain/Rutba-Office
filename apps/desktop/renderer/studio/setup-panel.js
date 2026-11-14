// The Set up panel (A5): the sources on the left, the frame they are arranged
// on in the middle with the scenes under it, the selected source's properties
// on the right. While live, what is changed here goes to Next and the picture
// on air stays as it is.

import React from 'react';
import { t } from '@rutba/office-ui';
import { Frame, Mirror, Seg } from './parts.js';
import { TallFrame } from './tall-frame.js';
import { SceneStrip } from './scene-strip.js';
import { SourcesColumn } from './sources-column.js';
import { PropertiesColumn } from './properties.js';
import { Prompter } from './prompter.js';

/** @param {{ s: object }} p  the studio (window.js) */
export function SetupPanel({ s }) {
  const scene = s.scenes.find((sc) => sc.key === s.nextKey);
  return (
    <div className="st-panel" data-panel="setup">
      {s.live ? <div className="st-band live">{t('studio.setup-panel.live-changes-go')}</div> : null}
      {s.recording && !s.live ? <div className="st-band rec">{t('studio.setup-panel.recording-changes-show')}</div> : null}
      <div className="st-setup-grid">
        <SourcesColumn s={s} thumbs={s.thumbs} />
        <div className="st-stage">
          <Prompter s={s} />
          <div className="st-frame-head">
            <b>{scene ? scene.name : t('studio.setup-panel.the-frame')}</b>
            <span className="st-frame-size">{s.isLive && s.frameShape === 'tall' ? (s.settings.size === 720 ? '720 × 1280' : '1080 × 1920') : (s.settings.size === 720 ? '1280 × 720' : '1920 × 1080')}</span>
            {s.isLive ? <Seg value={s.frameShape} options={[['wide', t('common.wide')], ['tall', t('common.tall')]]} label={t('studio.setup-panel.shape-frame')} className="st-frame-shape" onChange={s.setFrameShape} /> : null}
            <span className="st-dock-gap" />
            <span className="st-frame-hint">{t('studio.setup-panel.snaps-edges-centre')}</span>
          </div>
          <div className="st-share">
            {s.isLive && s.frameShape === 'tall' ? <TallFrame s={s} /> : (
            <Frame canvas={s.holding ? s.canvases.next : s.canvases.air} frameRef={s.editFrameRef} className="st-air st-editing" label={t('studio.setup-panel.the-frame')} tabIndex={0} onPointerDown={s.onFramePointer} onKeyDown={s.onFrameKey}>
              {s.overlay}
              {s.overlayShared}
              {s.live ? <div className="st-mirror"><span>{t('studio.setup-panel.on-air-now')}</span><Mirror program={s.program} label={t('studio.setup-panel.on-air-now')} /></div> : null}
            </Frame>
            )}
          </div>
          <div className="st-under-frame">
            <div className="st-scenes-head"><span className="st-scenes-label">{t('common.scenes')}</span><span className="st-scenes-hint">{s.isLive ? t('studio.setup-panel.1-9-switch') : t('studio.setup-panel.1-9-switch-while')}</span></div>
            <SceneStrip s={s} editing small />
          </div>
        </div>
        <PropertiesColumn s={s} />
      </div>
    </div>
  );
}
