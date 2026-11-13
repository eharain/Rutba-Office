// Spoken text layers (D-2026-11-13-31), the parts Video, the Recorder and
// Live Broadcast share: the added voices there are, a line made into sound
// with one of them, a line heard in the window, and the one icon that offers
// to add a voice when there is none.
//
// Only voices added from the packages speak here: the system's voices speak
// straight to the speakers and give no sound to record or export, so they
// stay Read Aloud's (the owner's Q14: the Windows voice stays its default).
// The text read never leaves the computer: the speech engine add-on makes the
// sound (`speech.say`, one sentence at a time) and the window plays it.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { Button, t } from '@rutba/office-ui';
import { sentencesOfLine, joinSamples, secondsOf, SPOKEN_SPEEDS } from '@rutba/studio/spoken';

export { SPOKEN_SPEEDS };

/** Voices and languages, where a voice is added. */
export function openVoicesPage(shell) {
  try {
    const r = shell?.win?.create?.({ app: 'home', query: { voices: 1 } });
    r?.catch?.(() => {});
  } catch {
    // no window to open
  }
}

/**
 * The added voices that can speak: [{ id, name, language, also }], looked
 * for again when the window comes back to the front (after a voice was added
 * in Voices and languages) and when the add-ons change.
 */
export function useAddedVoices(shell) {
  const [voices, setVoices] = useState(null);
  const look = useCallback(async () => {
    if (!shell?.speech?.voices) { setVoices([]); return; }
    try {
      const v = await shell.speech.voices();
      setVoices(Array.isArray(v) ? v.filter((x) => x && x.id) : []);
    } catch {
      setVoices([]);
    }
  }, [shell]);
  useEffect(() => {
    look();
    const again = () => look();
    window.addEventListener('focus', again);
    let off = null;
    try { off = shell?.on?.('addons:changed', again); } catch { off = null; }
    return () => {
      window.removeEventListener('focus', again);
      if (typeof off === 'function') off();
    };
  }, [shell, look]);
  return { voices: voices || [], looking: voices === null, ready: Boolean(voices && voices.length), look };
}

/** The voice to use: the one asked for when it is there, else the first. */
export function voiceOrFirst(voices, id) {
  return voices.find((v) => v.id === id) || voices[0] || null;
}

/**
 * One line made into sound with an added voice: its sentences one at a time,
 * joined with a little quiet. Answers { samples, rate, seconds }.
 */
export async function makeLine(shell, { text, voice, speed = 1, language = null }) {
  if (!shell?.speech?.say) throw new Error(t('No voice has been added.'));
  const parts = [];
  for (const sentence of sentencesOfLine(text, language)) {
    const out = await shell.speech.say({ text: sentence, voice, speed });
    const samples = out?.samples instanceof Float32Array ? out.samples : new Float32Array(out?.samples || []);
    for (let i = 0; i < samples.length; i++) if (!Number.isFinite(samples[i])) throw new Error(t('The voice gave no sound for this line.'));
    parts.push({ samples, rate: out?.rate || 24000 });
  }
  const joined = joinSamples(parts);
  return { ...joined, seconds: secondsOf(joined.samples, joined.rate) };
}

/**
 * Lines heard in the window: one at a time, the one before stopped. `play`
 * answers when it has finished, or false when it was stopped.
 */
export function useLinePlayer() {
  const ctx = useRef(null);
  const source = useRef(null);
  const [playing, setPlaying] = useState(null);
  const stop = useCallback(() => {
    const s = source.current;
    source.current = null;
    try { s?.stop(); } catch { /* already stopped */ }
    setPlaying(null);
  }, []);
  const play = useCallback((key, { samples, rate }) => new Promise((resolve) => {
    stop();
    const Ctx = window.AudioContext || window.webkitAudioContext;
    if (!Ctx) { resolve(false); return; }
    ctx.current ||= new Ctx();
    const c = ctx.current;
    if (c.state === 'suspended') c.resume().catch(() => {});
    const buffer = c.createBuffer(1, Math.max(1, samples.length), rate || 24000);
    buffer.copyToChannel(samples instanceof Float32Array ? samples : Float32Array.from(samples), 0);
    const s = c.createBufferSource();
    s.buffer = buffer;
    s.connect(c.destination);
    source.current = s;
    setPlaying(key);
    s.onended = () => {
      const mine = source.current === s;
      if (mine) { source.current = null; setPlaying(null); }
      resolve(mine);
    };
    s.start();
  }), [stop]);
  useEffect(() => () => {
    try { source.current?.stop(); } catch { /* stopped */ }
    try { ctx.current?.close(); } catch { /* closed */ }
  }, []);
  return { play, stop, playing };
}

/** The one icon that offers a voice when none is added: it opens Voices and languages. */
export function AddVoiceButton({ shell, className = '' }) {
  return (
    <Button
      ghost
      icon="voiceAdd"
      className={`spoken-add-voice ${className}`}
      title={t('Add a voice to have these lines read aloud')}
      onClick={() => openVoicesPage(shell)}
    />
  );
}

/** Voice and speed, for the lines of a layer. */
export function VoiceFields({ voices, voice, speed, onVoice, onSpeed, Row = DefaultRow, disabled = false }) {
  return (
    <>
      <Row label={t('Voice')}>
        <select className="rw-select spoken-voice" value={voice || ''} disabled={disabled} onChange={(e) => onVoice(e.target.value)}>
          {voices.map((v) => <option key={v.id} value={v.id}>{v.language ? `${v.name} (${v.language})` : v.name}</option>)}
        </select>
      </Row>
      <Row label={t('Speed')}>
        <select className="rw-select spoken-speed" value={String(speed || 1)} disabled={disabled} onChange={(e) => onSpeed(Number(e.target.value))}>
          {SPOKEN_SPEEDS.map((s) => <option key={s} value={String(s)}>{`${s}×`}</option>)}
        </select>
      </Row>
    </>
  );
}

function DefaultRow({ label, children }) {
  return <label className="spoken-row"><span>{label}</span>{children}</label>;
}

/** Show the text and Play the voice, the person's two choices for a layer. */
export function ShowPlay({ show, play, onChange, disabled = false }) {
  return (
    <div className="spoken-choice" role="group" aria-label={t('What the lines do')}>
      <label><input type="checkbox" className="spoken-show" checked={show} disabled={disabled} onChange={(e) => onChange({ show: e.target.checked })} />{t('Show the text')}</label>
      <label><input type="checkbox" className="spoken-play" checked={play} disabled={disabled} onChange={(e) => onChange({ play: e.target.checked })} />{t('Play the voice')}</label>
      {!show && !play ? <small className="spoken-note">{t('Neither is on: the lines keep their time but nothing is seen or heard.')}</small> : null}
    </div>
  );
}

/**
 * The lines of a layer, each with its own play-voice button, the one being
 * read or heard lit, and each one's length once it is known: the layer's
 * line-by-line timeline.
 */
export function LineList({ lines, lengths = [], current = -1, playing = null, onPlay, disabled = false }) {
  if (!lines.length) return <p className="spoken-note">{t('No lines yet: each line of the words is read in turn.')}</p>;
  return (
    <ol className="spoken-lines">
      {lines.map((text, i) => (
        <li key={`${i}:${text}`} className={`${i === current ? 'now' : ''}${playing === i ? ' heard' : ''}`} aria-current={i === current ? 'true' : undefined}>
          <Button
            ghost
            icon={playing === i ? 'stop' : 'play'}
            className="spoken-line-play"
            disabled={disabled}
            title={playing === i ? t('Stop') : t('Play the voice of this line')}
            onClick={() => onPlay(i)}
          />
          <span className="spoken-line-text" dir="auto">{text}</span>
          {Number.isFinite(lengths[i]) ? <small className="spoken-line-time">{t('{seconds} s', { seconds: Math.round(lengths[i] * 10) / 10 })}</small> : null}
        </li>
      ))}
    </ol>
  );
}

export const SPOKEN_CSS = `
.spoken-row { display: flex; align-items: center; gap: 8px; margin: 4px 0; }
.spoken-row > span { min-width: 64px; color: var(--ink-2); }
.spoken-choice { display: flex; flex-wrap: wrap; gap: 4px 12px; margin: 6px 0; }
.spoken-choice label { display: inline-flex; align-items: center; gap: 4px; }
.spoken-note { display: block; color: var(--ink-3); margin: 4px 0; }
.spoken-lines { list-style: none; margin: 6px 0; padding: 0; max-height: 220px; overflow: auto; border: 1px solid var(--line-soft); border-radius: 6px; }
.spoken-lines li { display: flex; align-items: center; gap: 4px; padding: 2px 6px 2px 2px; border-bottom: 1px solid var(--line-soft); }
.spoken-lines li:last-child { border-bottom: 0; }
.spoken-lines li.now { background: var(--accent-soft, rgba(59,130,246,0.14)); }
.spoken-line-text { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.spoken-line-time { color: var(--ink-3); }
.spoken-progress { display: flex; align-items: center; gap: 8px; margin: 6px 0; color: var(--ink-2); }
.spoken-actions { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 6px; }
`;
