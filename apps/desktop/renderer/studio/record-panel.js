// The Record panel (A2): the picture as it will be recorded, large, the scenes
// under it, the bar at the foot with Record. While recording, a red edge round
// the picture and Stop with the time. After Stop, a card with what was saved.

import React from 'react';
import { Button, Icon, basename, dirname, formatBytes, t } from '@rutba/office-ui';
import { Frame } from './parts.js';
import { SceneStrip, PickRow } from './scene-strip.js';
import { BottomBar } from './bottom-bar.js';
import { clock, PictureStrip, PictureButton } from './shared.js';
import { Prompter } from './prompter.js';
import { FlipStrip } from './flip-strip.js';

/** Length as a person says it: 12 min 34 s. */
const spoken = (seconds) => {
  const s = Math.round(seconds || 0);
  const m = Math.floor(s / 60);
  return m ? t('{m} min {s} s', { m, s: s % 60 }) : t('{s} s', { s });
};

/** What was just saved, over the scenes. */
function SavedCard({ s }) {
  const { saved } = s;
  return (
    <div className="st-saved" role="status">
      <div className="st-saved-top">
        <div className="st-saved-pic">{saved.thumb ? <img src={saved.thumb} alt="" /> : <Icon name="video" size={24} />}</div>
        <div className="st-saved-text">
          <div className="st-saved-kicker">{t('Saved')}</div>
          <div className="st-saved-name" data-tip={saved.path}>{basename(saved.path)}</div>
          <div className="st-saved-meta">{spoken(saved.seconds)} · {formatBytes(saved.written)} · {dirname(saved.path).split('/').slice(-2).join('\\')}{saved.takes?.length ? ` · ${t('{n} sources kept apart', { n: saved.takes.length })}` : ''}</div>
        </div>
      </div>
      <div className="st-saved-btns">
        <Button icon="play" label={t('Play')} className="st-play" onClick={() => s.shell.win.create({ app: 'video', file: saved.path })} />
        <Button primary icon="video" label={t('Edit in Video')} className="st-edit" onClick={() => s.shell.win.create(saved.takes?.length ? { app: 'video', query: { takes: JSON.stringify(saved.takes) } } : { app: 'video', file: saved.path, query: { edit: 1 } })} />
        <Button icon="folderOpen" label={t('Show in folder')} onClick={() => s.shell.shell.showInFolder({ path: saved.path })} />
        <span className="st-dock-gap" />
        <Button label={t('Record again')} className="st-again" onClick={() => s.setSaved(null)} />
      </div>
    </div>
  );
}

/** @param {{ s: object }} p  the studio (window.js) */
export function RecordPanel({ s }) {
  const { phase, recording, paused, elapsed, count } = s;
  const counting = phase === 'countdown';
  const canStop = recording || counting;
  return (
    <div className="st-panel" data-panel="record">
      <div className="st-stage">
        <Prompter s={s} />
        <div className="st-share">
          <Frame canvas={s.canvases.air} className={`st-air${recording ? ' live' : ''}${paused ? ' held' : ''}`} label={t('The picture as it will be recorded')}>
            {recording ? <span className="rec-badge st-badge"><i />{paused ? t('Paused') : t('REC')}</span> : null}
            {counting ? <div className="rec-count" aria-live="assertive"><span key={count}>{count}</span></div> : null}
            {s.overlayShared}
          </Frame>
        </div>
        <div className="st-flip-slot"><FlipStrip stack={s.airStack} web={s.media.web} /></div>
        <div className="st-under-frame">
          {s.saved && phase === 'setup' ? <SavedCard s={s} /> : (
            <>
              <SceneStrip s={s} />
              <PickRow s={s} />
            </>
          )}
          <PictureStrip s={s} />
        </div>
      </div>
      <BottomBar s={s}>
        <PictureButton s={s} />
        {canStop ? (
          <>
            <button type="button" className={`st-round st-pause${paused ? ' on' : ''}`} aria-pressed={paused} aria-label={paused ? t('Resume') : t('Pause')} data-tip={paused ? t('Resume') : t('Pause')} disabled={phase !== 'recording'} onClick={s.togglePause}><Icon name={paused ? 'play' : 'pause'} size={16} /></button>
            <button type="button" className={`st-rec st-record on${paused ? ' held' : ''}`} aria-pressed={recording} data-tip={s.stopTip} disabled={phase === 'stopping' || phase === 'starting'} onClick={s.stop}>
              <i className="st-square" />{counting ? t('Starting in {n}', { n: count }) : `${t('Stop')} ${clock(elapsed, true)}`}
            </button>
          </>
        ) : (
          <button type="button" className="st-rec st-record rec" data-tip={s.recordLabel} disabled={!s.canRecord} onClick={s.record}>
            <i className="st-dot" />{s.recordLabel}
          </button>
        )}
      </BottomBar>
    </div>
  );
}
