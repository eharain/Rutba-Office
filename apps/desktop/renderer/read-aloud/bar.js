// The Read Aloud bar (C7): play and pause, a sentence back and forward,
// speed, and the voice for the language being read. The apps put it where
// their windows have room; the reader (reader.js) does the reading.

import React, { useEffect, useState } from 'react';
import { Button, t } from '@rutba/office-ui';
import { SPEEDS } from './reader.js';

export function ReadAloudBar({ reader, state, onClose, className = '' }) {
  const language = state?.item?.language || null;
  const [voices, setVoices] = useState([]);
  useEffect(() => {
    let live = true;
    reader.voicesFor(language).then((v) => live && setVoices(v)).catch(() => live && setVoices([]));
    return () => { live = false; };
  }, [reader, language]);

  const current = state?.voice;
  const value = current ? (current.kind === 'added' ? `added:${current.id}` : `system:${current.name}`) : '';
  const choose = (v) => {
    if (!v) return reader.setVoice(language, null);
    const [kind, ...rest] = v.split(':');
    const name = rest.join(':');
    return reader.setVoice(language, kind === 'added' ? { kind, id: name } : { kind, name });
  };

  return (
    <div className={`read-aloud-bar ${className}`} role="toolbar" aria-label={t('word.ribbon.read-aloud')}>
      <style>{BAR_CSS}</style>
      <Button ghost icon="previous" title={t('read-aloud.bar.back-a-sentence')} onClick={() => reader.back()} />
      {state?.playing && !state?.paused
        ? <Button ghost icon="pause" title={t('common.pause')} onClick={() => reader.pause()} />
        : <Button ghost icon="play" title={t('common.play')} onClick={() => (state?.paused ? reader.resume() : reader.forward())} />}
      <Button ghost icon="next" title={t('read-aloud.bar.forward-a-sentence')} onClick={() => reader.forward()} />
      <span className="read-aloud-where">{state?.count ? t('read-aloud.bar.sentence-n-count', { n: Math.min(state.index + 1, state.count), count: state.count }) : ''}</span>
      <label className="read-aloud-field">
        <span>{t('video.speed')}</span>
        <select value={String(state?.speed || 1)} onChange={(e) => reader.setSpeed(Number(e.target.value))}>
          {SPEEDS.map((s) => <option key={s} value={String(s)}>{`${s}×`}</option>)}
        </select>
      </label>
      <label className="read-aloud-field">
        <span>{t('common.voice')}</span>
        <select value={value} onChange={(e) => choose(e.target.value)}>
          <option value="">{t('read-aloud.bar.system-voice')}</option>
          {voices.map((v) => {
            const key = v.kind === 'added' ? `added:${v.id}` : `system:${v.name}`;
            return <option key={key} value={key}>{v.kind === 'added' ? t('read-aloud.bar.name-added', { name: v.label }) : v.label}</option>;
          })}
        </select>
      </label>
      {language ? <span className="read-aloud-lang" title={t('read-aloud.bar.language-being-read')}>{language}</span> : null}
      {state?.said ? <span className="read-aloud-said" role="status">{t(state.said)}</span> : null}
      <span style={{ flex: 1 }} />
      <Button ghost icon="close" title={t('common.stop-reading')} onClick={onClose} />
    </div>
  );
}

const BAR_CSS = `
.read-aloud-bar { display: flex; align-items: center; gap: 6px; padding: 4px 10px; border-bottom: 1px solid var(--line-soft); background: var(--surface); font-size: 12px; }
.read-aloud-where { color: var(--ink-3); min-width: 110px; }
.read-aloud-field { display: inline-flex; align-items: center; gap: 4px; color: var(--ink-2); }
.read-aloud-field select { font: inherit; }
.read-aloud-lang { color: var(--ink-3); }
.read-aloud-said { color: var(--ink-2); }
`;
