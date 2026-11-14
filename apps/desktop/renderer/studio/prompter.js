// The teleprompter (A7, an introduction's scene): the scene's script, written in
// Video's scene card, scrolled above the picture at the set speed while it records.

import React, { useEffect, useRef, useState } from 'react';
import { Icon, t } from '@rutba/office-ui';
import { prompterProgress, PROMPTER_SPEEDS } from '@rutba/studio/intro';
import { Seg } from './parts.js';

/** @param {{ s: object }} p  the studio (window.js): s.brief.prompter is { title, script, speed } */
export function Prompter({ s }) {
  const prompter = s.brief?.prompter;
  const view = useRef(null);
  const text = useRef(null);
  const [shift, setShift] = useState(0);
  const script = prompter?.script || '';
  const speed = prompter?.speed || PROMPTER_SPEEDS.normal;
  const progress = s.recording ? prompterProgress(script, s.elapsed / 1000, speed) : 0;
  useEffect(() => {
    if (!view.current || !text.current) { setShift(0); return; }
    // The line being read sits a third of the way down, as on a prompter.
    setShift(Math.max(0, progress * Math.max(0, text.current.scrollHeight - view.current.clientHeight * 0.66)));
  }, [progress, script]);
  if (!prompter) return null;
  const nearest = Object.keys(PROMPTER_SPEEDS).sort((x, y) => Math.abs(PROMPTER_SPEEDS[x] - speed) - Math.abs(PROMPTER_SPEEDS[y] - speed))[0];
  return (
    <div className="rec-prompter" data-progress={progress.toFixed(3)} aria-label={t('studio.prompter.teleprompter')}>
      <div className="rec-prompter-head">
        <Icon name="prompter" size={15} />
        <b>{prompter.title}</b>
        <Seg value={nearest} className="rec-prompter-speed" label={t('studio.prompter.teleprompter')} options={[['slow', t('studio.prompter.slow')], ['normal', t('common.normal')], ['fast', t('studio.prompter.fast')]]} onChange={(k) => s.setPrompterSpeed(PROMPTER_SPEEDS[k])} />
      </div>
      <div className="rec-prompter-view" ref={view}>
        {script ? (
          <div className="rec-prompter-text" dir="auto" ref={text} style={{ transform: `translateY(${-shift}px)` }}>{script}</div>
        ) : (
          <div className="rec-prompter-empty">{t('studio.prompter.scene-no-script')}</div>
        )}
      </div>
    </div>
  );
}
