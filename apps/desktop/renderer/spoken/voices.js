// Spoken text layers (D-2026-11-13-31), the parts Video, the Recorder and
// Live Broadcast share: the added voices there are, a line made into sound
// with one of them, a line heard in the window, and the one icon that offers
// to add a voice when there is none (Read Aloud's own, read-aloud/voice-offer.js).
// The three choices for spoken words, shown and read, read only or shown only
// (D-2026-11-16-01), are one control here, SpeakChoice, and so is the box
// that takes typed words and says them, SpeakBox: the same in all three.
//
// Only voices added from the packages speak here: the system's voices speak
// straight to the speakers and give no sound to record or export, so they
// stay Read Aloud's (the owner's Q14: the Windows voice stays its default).
// The text read never leaves the computer: the speech engine add-on makes the
// sound (`speech.say`, one sentence at a time) and the window plays it.

import React, { useCallback, useEffect, useId, useRef, useState } from 'react';
import { Button, t } from '@rutba/office-ui';
import { sentencesOfLine, joinSamples, secondsOf, SPOKEN_SPEEDS, SPEAK_CHOICES, speakChoiceOf, speakSwitches } from '@rutba/studio/spoken';
import { VoiceOffer } from '../read-aloud/voice-offer.js';

export { SPOKEN_SPEEDS };

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
  if (!shell?.speech?.say) throw new Error(t('common.no-voice-added'));
  const parts = [];
  for (const sentence of sentencesOfLine(text, language)) {
    let out = await shell.speech.say({ text: sentence, voice, speed });
    // Cancelled in the engine (its cancel stops every window's speech): asked once more, never kept as quiet.
    if (out?.cancelled) out = await shell.speech.say({ text: sentence, voice, speed });
    if (out?.cancelled) throw new Error(t('spoken.voices.voice-stopped-before'));
    const samples = out?.samples instanceof Float32Array ? out.samples : new Float32Array(out?.samples || []);
    for (let i = 0; i < samples.length; i++) if (!Number.isFinite(samples[i])) throw new Error(t('spoken.voices.voice-gave-no'));
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

/**
 * The one icon that offers a voice when none is added: Read Aloud's own,
 * which opens Voices and languages, in words for lines that only an added
 * voice can speak. Shown by its callers only when they found no voice.
 */
export function AddVoiceButton({ shell, className = '' }) {
  return <VoiceOffer shell={shell} voices={[]} className={`spoken-add-voice${className ? ` ${className}` : ''}`} title={t('spoken.voices.add-voice-lines')} />;
}

/** Voice and speed, for the lines of a layer. */
export function VoiceFields({ voices, voice, speed, onVoice, onSpeed, Row = DefaultRow, disabled = false }) {
  return (
    <>
      <Row label={t('common.voice')}>
        <select className="rw-select spoken-voice" value={voice || ''} disabled={disabled} onChange={(e) => onVoice(e.target.value)}>
          {voices.map((v) => <option key={v.id} value={v.id}>{v.language ? `${v.name} (${v.language})` : v.name}</option>)}
        </select>
      </Row>
      <Row label={t('video.speed')}>
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

/**
 * What the words do, the person's one choice of three: show the text and read
 * it, read it only, or show it only. `show` and `play` are the two switches
 * stored; onChange is given both. Without a voice (`canRead` false) only
 * Show only can be chosen; `allowed` (ids) narrows them further, for words
 * whose voice is in a recording's own sound. Words kept from before with
 * neither (silent and unseen) have no choice marked, and say so.
 */
export function SpeakChoice({ show, play, onChange, canRead = true, allowed = null, disabled = false }) {
  const now = speakChoiceOf({ show, play });
  // One name for the three, so the arrow keys move between them.
  const group = useId();
  const names = { both: t('spoken.voices.show-read-aloud'), read: t('spoken.voices.read-aloud-only'), show: t('spoken.voices.show-only') };
  return (
    <div className="spoken-choice" role="radiogroup" aria-label={t('spoken.voices.lines-do')}>
      {SPEAK_CHOICES.map((c) => (
        <label key={c.id} className="spoken-choice-opt">
          <input
            type="radio"
            name={group}
            className={`spoken-choice-${c.id}`}
            value={c.id}
            checked={now === c.id}
            disabled={disabled || (c.play && !canRead) || (allowed ? !allowed.includes(c.id) : false)}
            onChange={() => onChange(speakSwitches(c.id))}
          />
          {names[c.id]}
        </label>
      ))}
      {now === null ? <small className="spoken-note">{t('spoken.voices.neither-shown-nor')}</small> : null}
    </div>
  );
}

/**
 * Words typed and said, the same box in Video, the Recorder and Live: the
 * words (one line each, read one after another), the choice of three, the
 * voice and its speed, and one button that says them, which becomes Stop
 * while they are being said. `value`: { text, show, play, voice, speed };
 * onChange is given what changed. With no voice added the box offers one and
 * the words can still be shown.
 * Ctrl+Enter in the words says them.
 */
export function SpeakBox({ shell, voices, value, onChange, onSay, onStop = null, saying = false, busy = false, sayLabel, sayIcon = 'volume', Row = DefaultRow, disabled = false, autoFocus = false }) {
  const ready = voices.length > 0;
  const v = value || {};
  const show = v.show !== false;
  const play = v.play !== false;
  const voice = voiceOrFirst(voices, v.voice);
  const needsVoice = play && !ready;
  const canSay = Boolean(String(v.text || '').trim()) && !needsVoice && !busy && !disabled && speakChoiceOf({ show, play }) !== null;
  const say = () => { if (canSay) onSay(); };
  return (
    <div className="spoken-box">
      <textarea
        className="rw-input spoken-box-text"
        dir="auto"
        rows={3}
        value={v.text || ''}
        disabled={disabled}
        autoFocus={autoFocus}
        aria-label={t('spoken.voices.words-to-say')}
        placeholder={t('spoken.voices.type-words-say')}
        onChange={(e) => onChange({ text: e.target.value })}
        onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); say(); } }}
      />
      <SpeakChoice show={show} play={play} canRead={ready} disabled={disabled} onChange={onChange} />
      {ready ? (
        <VoiceFields Row={Row} voices={voices} voice={voice?.id} speed={v.speed || 1} disabled={disabled} onVoice={(id) => onChange({ voice: id })} onSpeed={(n) => onChange({ speed: n })} />
      ) : (
        <p className="spoken-note spoken-box-novoice">{t('spoken.voices.add-voice-read')}<AddVoiceButton shell={shell} /></p>
      )}
      <div className="spoken-actions">
        {saying && onStop ? (
          <Button icon="stop" label={t('common.stop')} className="spoken-box-stop" onClick={onStop} />
        ) : (
          <Button primary icon={sayIcon} label={sayLabel} className="spoken-box-say" busy={busy} disabled={!canSay} title={t('spoken.voices.shortcut-ctrl-enter')} onClick={say} />
        )}
      </div>
    </div>
  );
}

/**
 * The lines of a layer, each with its own play-voice button, the one being
 * read or heard lit, and each one's length once it is known: the layer's
 * line-by-line timeline.
 */
export function LineList({ lines, lengths = [], current = -1, playing = null, onPlay, disabled = false }) {
  if (!lines.length) return <p className="spoken-note">{t('spoken.voices.no-lines-yet')}</p>;
  return (
    <ol className="spoken-lines">
      {lines.map((text, i) => (
        <li key={`${i}:${text}`} className={`${i === current ? 'now' : ''}${playing === i ? ' heard' : ''}`} aria-current={i === current ? 'true' : undefined}>
          <Button
            ghost
            icon={playing === i ? 'stop' : 'play'}
            className="spoken-line-play"
            disabled={disabled}
            title={playing === i ? t('common.stop') : t('spoken.voices.play-voice-line')}
            onClick={() => onPlay(i)}
          />
          <span className="spoken-line-text" dir="auto">{text}</span>
          {Number.isFinite(lengths[i]) ? <small className="spoken-line-time">{t('common.seconds-s', { seconds: Math.round(lengths[i] * 10) / 10 })}</small> : null}
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
.spoken-box { display: flex; flex-direction: column; gap: 2px; }
.spoken-box-text { width: 100%; resize: vertical; min-height: 54px; }
.spoken-box-novoice { display: flex; align-items: center; gap: 6px; }
`;
