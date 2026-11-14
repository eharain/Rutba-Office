// Spoken lines in the studio (D-2026-11-13-31), the Recorder's and Live
// Broadcast's: a title source's lines read aloud one after another. Each line
// is made before it is spoken (Live's rule: a line is always made first, so
// the stream never waits on the voice), the next one is made while one plays,
// and each plays into the studio's one sound output, so the recording, the
// sources kept apart and the stream all carry it. With Show the text, the
// title's band shows the line being read in place of its own words; with
// Play the voice off, each line is shown for as long as its voice would take,
// silently.
//
// The voice goes to the output only, never to the speakers: a microphone
// beside them would record it a second time. The output's meter shows it.
//
// While recording, each line's time from the recording's start is kept, so a
// Recorder opened from Video hands them over with its takes and they land on
// Video's timeline line by line (placeSpokenCues).

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, t, tn } from '@rutba/office-ui';
import { linesOf, LINE_GAP, spokenCueLog } from '@rutba/studio/spoken';
import { useAddedVoices, voiceOrFirst, makeLine, useLinePlayer, AddVoiceButton, VoiceFields, ShowPlay, LineList } from '../spoken/voices.js';

const wait = (ms, alive) => new Promise((resolve) => {
  const started = performance.now();
  const tick = () => {
    if (!alive() || performance.now() - started >= ms) { resolve(); return; }
    setTimeout(tick, Math.min(100, Math.max(10, ms - (performance.now() - started))));
  };
  tick();
});

/**
 * The reading of a title's lines.
 * @param {object} p
 * @param {object} p.shell
 * @param {object} p.media     StudioMedia
 * @param {(id: string) => object|null} p.itemNow   the title as it is now (its choices may change while it reads)
 * @param {() => number|null} p.clock   seconds since the recording started, or null when not recording
 */
export function useSpokenRun({ shell, media, itemNow, clock }) {
  const { voices, ready } = useAddedVoices(shell);
  const [state, setState] = useState({ itemId: null, index: -1, count: 0, lengths: [], making: false });
  const token = useRef(0);
  // The lines read while recording, as cues from its start. Kept in a ref, made
  // once: useMemo is a cache React may let go of, and a log made again part
  // way would lose the cues read so far.
  const logRef = useRef(null);
  if (!logRef.current) logRef.current = spokenCueLog();
  const log = logRef.current;
  const current = useRef(null);
  const voicesRef = useRef(voices);
  voicesRef.current = voices;

  /** The line being read ends at `end` (seconds from the recording's start, null when not recording): its cue is kept. */
  const closeLine = useCallback((end) => log.end(end), [log]);

  // The speech engine's cancel is not called: it stops every window's speech, Read Aloud's too.
  // A line still being made is let finish and left unused.
  const stop = useCallback(() => {
    token.current += 1;
    media.stopSpeaking?.();
    closeLine(clock());
    if (current.current) media.setSpokenLine?.(current.current, null);
    current.current = null;
    setState({ itemId: null, index: -1, count: 0, lengths: [], making: false });
  }, [media, clock, closeLine]);

  const start = useCallback(async (itemId, { from = 0 } = {}) => {
    stop();
    const mine = (token.current += 1);
    const alive = () => token.current === mine;
    const first = itemNow(itemId);
    const lines = linesOf(first?.lines);
    if (!first || !lines.length) return false;
    current.current = itemId;
    const lengths = [];
    const make = (i) => {
      const it = itemNow(itemId) || first;
      const voice = voiceOrFirst(voicesRef.current, it.spoken?.voice);
      if (!voice) return Promise.reject(new Error(t('No voice has been added.')));
      return makeLine(shell, { text: lines[i], voice: voice.id, speed: it.spoken?.speed || 1, language: voice.language });
    };
    setState({ itemId, index: from, count: lines.length, lengths: [], making: true });
    let next = make(from);
    try {
      for (let i = from; i < lines.length; i++) {
        const sound = await next;
        if (!alive()) return false;
        next = i + 1 < lines.length ? make(i + 1) : null;
        next?.catch(() => {});
        lengths[i] = sound.seconds;
        const it = itemNow(itemId);
        // A title taken away while it reads stops reading.
        if (!it) { stop(); return false; }
        const show = it.spoken?.show !== false;
        const play = it.spoken?.play !== false;
        setState({ itemId, index: i, count: lines.length, lengths: [...lengths], making: false });
        media.setSpokenLine?.(itemId, show ? lines[i] : null);
        log.begin(clock(), lines[i], show);
        if (play) await media.speak(sound.samples, sound.rate);
        else await wait(sound.seconds * 1000, alive);
        if (!alive()) return false;
        closeLine(clock());
        if (i + 1 < lines.length) await wait(LINE_GAP * 1000, alive);
        if (!alive()) return false;
      }
    } catch (err) {
      if (alive()) {
        stop();
        throw err;
      }
      return false;
    }
    if (alive()) {
      media.setSpokenLine?.(itemId, null);
      current.current = null;
      setState({ itemId: null, index: -1, count: 0, lengths, making: false });
    }
    return true;
  }, [stop, itemNow, shell, media, clock, closeLine, log]);

  useEffect(() => () => { token.current += 1; media.stopSpeaking?.(); }, [media]);

  return {
    voices,
    ready,
    state,
    start,
    stop,
    /**
     * The lines read while recording, from its start, and forgotten: [{ start, end, text, show }].
     * A line still being read when the recording stopped ends at `end`, the recording's length.
     */
    takeCues: (end = null) => log.take(end),
    /** A recording begins: nothing kept from before, and a line being read now counts from its start. */
    clearCues: () => log.clear(),
  };
}

/** A title's lines, their voice, and the two choices, in its Properties (Set up). */
export function SpokenTitleProps({ s, item, Field }) {
  const run = s.spoken;
  const player = useLinePlayer();
  const [hearing, setHearing] = useState(false);
  const lines = linesOf(item.lines);
  const reading = run.state.itemId === item.id;
  const sp = item.spoken || {};
  const set = (patch) => s.changeStack((st) => s.withItem(st, item.id, patch));
  const setSpoken = (patch) => set({ spoken: { ...sp, ...patch } });
  const voice = voiceOrFirst(run.voices, sp.voice);
  const Row = useMemo(() => function SpokenRow({ label, children }) { return <Field label={label}>{children}</Field>; }, [Field]);

  const hear = async (i) => {
    if (player.playing === i) { player.stop(); return; }
    try {
      setHearing(true);
      const out = await makeLine(s.shell, { text: lines[i], voice: voice.id, speed: sp.speed || 1, language: voice.language });
      setHearing(false);
      await player.play(i, out);
    } catch (err) {
      setHearing(false);
      s.toast?.(t('The line could not be read: {reason}', { reason: err?.message || String(err) }), { tone: 'bad', ms: 7000 });
    }
  };

  return (
    <section className="st-col-section st-spoken">
      <h3 className="st-spoken-head">{t('Spoken lines')}{!run.ready ? <AddVoiceButton shell={s.shell} /> : null}</h3>
      <Field label={t('Lines to read')}>
        <textarea className="rw-input st-spoken-lines" dir="auto" rows={4} value={item.lines || ''} placeholder={t('One line each: each is read aloud in turn')} onChange={(e) => set({ lines: e.target.value })} />
      </Field>
      {run.ready ? (
        <>
          <VoiceFields Row={Row} voices={run.voices} voice={voice?.id} speed={sp.speed || 1} onVoice={(v) => setSpoken({ voice: v })} onSpeed={(v) => setSpoken({ speed: v })} />
          <ShowPlay show={sp.show !== false} play={sp.play !== false} onChange={setSpoken} />
          {/* A line heard here plays to the speakers, which a microphone would pick up: not while recording or live. */}
          <LineList lines={lines} lengths={reading ? run.state.lengths : []} current={reading ? run.state.index : -1} playing={player.playing} disabled={hearing || reading || Boolean(s.recording) || Boolean(s.live)} onPlay={hear} />
          {sp.play !== false && s.stack?.sound !== 'on' ? <p className="st-col-note st-spoken-quiet">{t('Sound is not on, so the voice is not in the recording or the stream.')}</p> : null}
          {!item.on && sp.show !== false ? <p className="st-col-note">{t('The title is not on the frame: show it and the lines appear in it as they are read.')}</p> : null}
          <ReadButton s={s} item={item} />
        </>
      ) : null}
    </section>
  );
}

function ReadButton({ s, item, compact = false }) {
  const run = s.spoken;
  const reading = run.state.itemId === item.id;
  const go = () => run.start(item.id).catch((err) => s.toast?.(t('The lines could not be read aloud: {reason}', { reason: err?.message || String(err) }), { tone: 'bad', ms: 8000 }));
  if (reading) {
    return (
      <Button
        icon="stop"
        label={compact ? '' : t('Stop reading')}
        className="st-spoken-stop"
        title={run.state.making ? t('Making line {n} of {count}…', { n: run.state.index + 1, count: run.state.count }) : t('Reading line {n} of {count}', { n: run.state.index + 1, count: run.state.count })}
        onClick={run.stop}
      />
    );
  }
  return <Button icon="play" label={compact ? '' : t('Read the lines')} className="st-spoken-read" disabled={!linesOf(item.lines).length} title={t('Each line read aloud in turn, shown in the title while its voice plays')} onClick={go} />;
}

/**
 * Under the frame on the Record and Live panels: each title with lines, to
 * read them on air, the line being read shown. With no voice added, the one
 * icon that offers one.
 */
export function SpokenStrip({ s }) {
  const titles = (s.stack?.items || []).filter((it) => it.kind === 'title' && linesOf(it.lines).length);
  if (!titles.length) return null;
  const run = s.spoken;
  if (!run.ready) {
    return (
      <div className="st-spoken-strip" role="group" aria-label={t('Spoken lines')}>
        <Icon name="volume" size={14} />
        <span>{t('Spoken lines')}</span>
        <AddVoiceButton shell={s.shell} />
      </div>
    );
  }
  return (
    <div className="st-spoken-strip" role="group" aria-label={t('Spoken lines')}>
      {titles.map((it) => {
        const reading = run.state.itemId === it.id;
        const lines = linesOf(it.lines);
        return (
          <span key={it.id} className={`st-spoken-title${reading ? ' on' : ''}`}>
            <ReadButton s={s} item={it} compact />
            <span className="st-spoken-name">{it.label || it.text || t('Title')}</span>
            <span className="st-spoken-now" dir="auto">
              {reading ? `${run.state.index + 1}/${run.state.count} ${lines[run.state.index] || ''}` : tn(lines.length, '{count} line', '{count} lines')}
            </span>
          </span>
        );
      })}
    </div>
  );
}

export const STUDIO_SPOKEN_CSS = `
.st-spoken-head { display: flex; align-items: center; gap: 6px; }
.st-spoken-head .spoken-add-voice { margin-left: auto; }
.st-spoken-lines { width: 100%; resize: vertical; }
.st-spoken-strip { display: flex; align-items: center; flex-wrap: wrap; gap: 6px 14px; padding: 4px 8px; font-size: 12px; color: var(--ink-2); }
.st-spoken-title { display: inline-flex; align-items: center; gap: 6px; max-width: 100%; }
.st-spoken-title.on .st-spoken-now { color: var(--ink-1, inherit); font-weight: 600; }
.st-spoken-name { font-weight: 600; }
.st-spoken-now { max-width: 340px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
`;
