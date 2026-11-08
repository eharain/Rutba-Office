// The Set up panel (A5): the sources, the frame they are arranged on, and the
// scenes kept from it. While live, what is changed here goes to Next and the
// picture on air stays as it is.

import React from 'react';
import { t } from '@rutba/office-ui';
import { Frame, Mirror } from './parts.js';
import { SceneStrip } from './scene-strip.js';
import { SourceBar } from './bar.js';

/** @param {{ s: object }} p  the studio (window.js) */
export function SetupPanel({ s }) {
  return (
    <div className="st-panel" data-panel="setup">
      {s.live ? <div className="st-band live">{t('You are live. Changes here go to Next, and go on air when you Take.')}</div> : null}
      {s.recording && !s.live ? <div className="st-band rec">{t('Recording. Changes here show in the recording at once.')}</div> : null}
      <div className="st-stage">
        <SourceBar item={s.selItem} sound={s.selSound} clipping={s.clipping} act={s.barAct} />
        <div className="st-share">
          <Frame canvas={s.holding ? s.canvases.next : s.canvases.air} frameRef={s.editFrameRef} className="st-air st-editing" label={t('The frame')} tabIndex={0} onPointerDown={s.onFramePointer} onKeyDown={s.onFrameKey}>
            {s.overlay}
            {s.overlayShared}
            {s.live ? <div className="st-mirror"><span>{t('On air now')}</span><Mirror program={s.program} label={t('On air now')} /></div> : null}
          </Frame>
        </div>
        <div className="st-under-frame"><SceneStrip s={s} editing small /></div>
      </div>
    </div>
  );
}
