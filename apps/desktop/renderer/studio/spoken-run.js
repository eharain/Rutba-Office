// Spoken lines in the studio (D-2026-11-13-31), the Recorder's and Live
// Broadcast's: a title source's lines read aloud one after another. Each line
// is made before it is spoken (Live's rule: a line is always made first, so
// the stream never waits on the voice), the next one is made while one plays,
// and each plays into the studio's one sound output, so the recording, the
// sources kept apart and the stream all carry it. The one choice of three
// (voices.js SpeakChoice) says what a line does: Show and read aloud, the
// title's band shows the line being read in place of its own words while its
// voice plays; Read aloud only, the voice alone; Show only, each line is shown
// silently for about as long as it takes to read.
//
// The voice goes to the output only, never to the speakers: a microphone
// beside them would record it a second time. The output's meter shows it.
//
// While recording, each line's time from the recording's start is kept, so a
// Recorder opened from Video hands them over with its takes and they land on
// Video's timeline line by line (placeSpokenCues).
//
// Words can also be typed and said there and then (D-2026-11-16-01), in the
// Record panel and in Live alike: read the same way, one line after another,
// with the same three choices (shown and read, read only, shown only). Their
// band is drawn across the foot of the frame over every source (program.js),
// so they reach the air whatever the stack. Shown only needs no voice: each
// line stays up for about as long as it takes to read.
//
// Typed words can come into their band by the typewriter (S20), as a caption
// does in Video: read aloud, each line types over its voice, the last letter
// as the voice ends ("With the voice"), or the voice a moment first ("Voice
// first"); shown only, at the standard typing speed, each line held long
// enough to type and read. The typing is worked out in @rutba/studio/spoken
// (studioTyping, from Video's typingWithVoice) and drawn by program.js.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Select, t, tn } from '@rutba/office-ui';
import { linesOf, typedLinesOf, LINE_GAP, spokenCueLog, readingSeconds, TYPED_LINE, studioTyping, typedShownSeconds, TYPING_TIMINGS } from '@rutba/studio/spoken';
import { useAddedVoices, voiceOrFirst, makeLine, useLinePlayer, AddVoiceButton, VoiceFields, SpeakChoice, SpeakBox, LineList } from '../spoken/voices.js';
import { wait, sayLine } from './say-line.js';

/**
 * The reading of a title's lines.
 * @param {object} p
 * @param {object} p.shell
 * @param {object} p.media     StudioMedia
 * @param {(id: string) => object|null} p.itemNow   the title as it is now (its choices may change while it reads)
 * @param {() => number|null} p.clock   seconds since the recording started, or null when not recording
 */
export function useSpokenRun({ shell, media, itemNow, clock }) {
  const { voices, ready, looking } = useAddedVoices(shell);
  // `lines`: the lines being read, as they were when it began (the words may be changed meanwhile).
  const [state, setState] = useState({ itemId: null, index: -1, count: 0, lengths: [], lines: [], making: false });
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
  // The words typed to say: kept here, so the Record and Live panels share them.
  // `typewriter`: the shown words type in; `typing`: with the voice, or the voice first (TYPING_TIMINGS).
  const [typed, setTyped] = useState({ text: '', show: true, play: true, voice: '', speed: 1, typewriter: false, typing: 'with' });
  const typedRef = useRef(typed);
  typedRef.current = typed;
  const changeTyped = useCallback((patch) => setTyped((was) => ({ ...was, ...patch })), []);
  // Typed words read as a title would be: their choices may change while they are said.
  const itemOf = useCallback((id) => (id === TYPED_LINE ? { id: TYPED_LINE, on: true, lines: typedRef.current.text, spoken: typedRef.current } : itemNow(id)), [itemNow]);

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
    setState({ itemId: null, index: -1, count: 0, lengths: [], lines: [], making: false });
  }, [media, clock, closeLine]);

  const start = useCallback(async (itemId, { from = 0 } = {}) => {
    stop();
    const mine = (token.current += 1);
    const alive = () => token.current === mine;
    const first = itemOf(itemId);
    // Typed words too long for the band are said a sentence at a time.
    const lines = itemId === TYPED_LINE ? typedLinesOf(first?.lines) : linesOf(first?.lines);
    if (!first || !lines.length) return false;
    current.current = itemId;
    const lengths = [];
    const make = (i) => {
      const it = itemOf(itemId) || first;
      const voice = voiceOrFirst(voicesRef.current, it.spoken?.voice);
      // Shown only needs no voice, and makes none: the line stays up for its reading time.
      // Typed in by the typewriter, it stays long enough to type at the standard speed and be read.
      if (it.spoken?.play === false) {
        const seconds = itemId === TYPED_LINE ? typedShownSeconds(lines[i], { speed: it.spoken?.speed || 1, typewriter: it.spoken?.typewriter === true }) : readingSeconds(lines[i], { speed: it.spoken?.speed || 1 });
        return Promise.resolve({ samples: new Float32Array(0), rate: 24000, seconds });
      }
      if (!voice) return Promise.reject(new Error(t('common.no-voice-added')));
      return makeLine(shell, { text: lines[i], voice: voice.id, speed: it.spoken?.speed || 1, language: voice.language });
    };
    setState({ itemId, index: from, count: lines.length, lengths: [], lines, making: true });
    let next = make(from);
    try {
      for (let i = from; i < lines.length; i++) {
        const sound = await next;
        if (!alive()) return false;
        next = i + 1 < lines.length ? make(i + 1) : null;
        next?.catch(() => {});
        lengths[i] = sound.seconds;
        const it = itemOf(itemId);
        // A title taken away while it reads stops reading.
        if (!it) { stop(); return false; }
        const show = it.spoken?.show !== false;
        const play = it.spoken?.play !== false;
        setState({ itemId, index: i, count: lines.length, lengths: [...lengths], lines, making: false });
        // Typed words by the typewriter: typed over the voice when it is played, else at the standard speed.
        const voiced = play && sound.samples.length > 0;
        const typing = itemId === TYPED_LINE && it.spoken?.typewriter === true ? studioTyping(lines[i], { seconds: voiced ? sound.seconds : 0, play: voiced, typing: it.spoken?.typing }) : null;
        media.setSpokenLine?.(itemId, show ? lines[i] : null, typing);
        // Whether its voice is in the sound: played, made (a line shown alone has none), with Sound on.
        log.begin(clock(), lines[i], show, play && sound.samples.length > 0 && (media.mode ?? 'on') === 'on');
        // A line made for Show only has no voice: read aloud after a change of choice, it is still held for its time.
        await sayLine({ media, sound, play, alive });
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
      setState({ itemId: null, index: -1, count: 0, lengths, lines: [], making: false });
    }
    return true;
  }, [stop, itemOf, shell, media, clock, closeLine, log]);

  useEffect(() => () => { token.current += 1; media.stopSpeaking?.(); }, [media]);

  // Hear the voice on this computer as it goes into the recording or the stream (2.26.1): kept with the profile.
  const [hear, setHear] = useState(false);
  useEffect(() => {
    let live = true;
    Promise.resolve(shell?.store?.get({ key: HEAR_KEY, fallback: false })).then((v) => { if (live) setHear(v === true); }).catch(() => {});
    return () => { live = false; };
  }, [shell]);
  useEffect(() => { media.setHearSpoken?.(hear); }, [media, hear]);
  const changeHear = useCallback((on) => {
    setHear(Boolean(on));
    Promise.resolve(shell?.store?.set({ key: HEAR_KEY, value: Boolean(on) })).catch(() => {});
  }, [shell]);

  return {
    /** Whether the voice also plays on this computer, and the switch for it. */
    hear,
    changeHear,
    voices,
    ready,
    looking,
    state,
    start,
    stop,
    /** The words typed to say, { text, show, play, voice, speed }, and a change to them. */
    typed,
    changeTyped,
    /** Say the typed words now; whatever was being read stops. Answers as start does. */
    say: () => start(TYPED_LINE),
    /**
     * The lines read while recording, from its start, and forgotten: [{ start, end, text, show }].
     * A line still being read when the recording stopped ends at `end`, the recording's length.
     */
    takeCues: (end = null) => log.take(end),
    /** A recording begins: nothing kept from before, and a line being read now counts from its start. */
    clearCues: () => log.clear(),
  };
}

const HEAR_KEY = 'studio.spokenHear';

/** The switch to hear the voice on this computer as it is recorded or sent, with what a microphone would do said. */
function HearSwitch({ run }) {
  return (
    <label className="st-spoken-hear" title={t('studio.spoken-run.voice-plays-computer')}>
      <input type="checkbox" checked={Boolean(run.hear)} onChange={(e) => run.changeHear(e.target.checked)} />
      <span>{t('studio.spoken-run.hear-voice-computer')}</span>
    </label>
  );
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
      s.toast?.(t('common.line-could-not', { reason: err?.message || String(err) }), { tone: 'bad', ms: 7000 });
    }
  };

  return (
    <section className="st-col-section st-spoken">
      <h3 className="st-spoken-head">{t('common.spoken-lines')}{!run.ready ? <AddVoiceButton shell={s.shell} /> : null}</h3>
      <Field label={t('studio.spoken-run.lines-to-read')}>
        <textarea className="rw-input st-spoken-lines" dir="auto" rows={4} value={item.lines || ''} placeholder={t('studio.spoken-run.one-line-each')} onChange={(e) => set({ lines: e.target.value })} />
      </Field>
      {run.ready ? (
        <>
          <VoiceFields Row={Row} voices={run.voices} voice={voice?.id} speed={sp.speed || 1} onVoice={(v) => setSpoken({ voice: v })} onSpeed={(v) => setSpoken({ speed: v })} />
          <SpeakChoice show={sp.show !== false} play={sp.play !== false} onChange={setSpoken} />
          {sp.play !== false ? <HearSwitch run={run} /> : null}
          {/* A line heard here plays to the speakers, which a microphone would pick up: not while recording or live. */}
          <LineList lines={lines} lengths={reading ? run.state.lengths : []} current={reading ? run.state.index : -1} playing={player.playing} disabled={hearing || reading || Boolean(s.recording) || Boolean(s.live)} onPlay={hear} />
          {sp.play !== false && (s.airStack || s.stack)?.sound !== 'on' ? <p className="st-col-note st-spoken-quiet">{t('studio.spoken-run.sound-not-voice')}</p> : null}
          {!item.on && sp.show !== false ? <p className="st-col-note">{t('studio.spoken-run.title-not-frame')}</p> : null}
          <ReadButton s={s} item={item} />
        </>
      ) : null}
    </section>
  );
}

function ReadButton({ s, item, compact = false }) {
  const run = s.spoken;
  const reading = run.state.itemId === item.id;
  const go = () => run.start(item.id).catch((err) => s.toast?.(t('common.lines-could-not', { reason: err?.message || String(err) }), { tone: 'bad', ms: 8000 }));
  if (reading) {
    return (
      <Button
        icon="stop"
        label={compact ? '' : t('common.stop-reading')}
        className="st-spoken-stop"
        title={run.state.making ? t('studio.spoken-run.making-line-n', { n: run.state.index + 1, count: run.state.count }) : t('common.reading-line-n', { n: run.state.index + 1, count: run.state.count })}
        onClick={run.stop}
      />
    );
  }
  return <Button icon="play" label={compact ? '' : t('studio.spoken-run.read-the-lines')} className="st-spoken-read" disabled={!linesOf(item.lines).length} title={t('studio.spoken-run.each-line-read')} onClick={go} />;
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
      <div className="st-spoken-strip" role="group" aria-label={t('common.spoken-lines')}>
        <Icon name="volume" size={14} />
        <span>{t('common.spoken-lines')}</span>
        <AddVoiceButton shell={s.shell} />
      </div>
    );
  }
  return (
    <div className="st-spoken-strip" role="group" aria-label={t('common.spoken-lines')}>
      {titles.map((it) => {
        const reading = run.state.itemId === it.id;
        const lines = linesOf(it.lines);
        return (
          <span key={it.id} className={`st-spoken-title${reading ? ' on' : ''}`}>
            <ReadButton s={s} item={it} compact />
            <span className="st-spoken-name">{it.label || it.text || t('common.title')}</span>
            <span className="st-spoken-now" dir="auto">
              {reading ? `${run.state.index + 1}/${run.state.count} ${lines[run.state.index] || ''}` : tn(lines.length, 'base.documents.count-line')}
            </span>
          </span>
        );
      })}
    </div>
  );
}

/**
 * Words typed and said live, under the frame on the Record and Live panels:
 * the shared box, its Say button Stop while they are said, and which line
 * is being said. A microphone near the speakers never hears it: the voice
 * goes to the output only, as a title's lines do.
 */
export function TypedSpeech({ s }) {
  const run = s.spoken;
  const [open, setOpen] = useState(false);
  const saying = run.state.itemId === TYPED_LINE;
  const say = () => run.say().catch((err) => s.toast?.(t('common.lines-could-not', { reason: err?.message || String(err) }), { tone: 'bad', ms: 8000 }));
  // With no voice added the words can only be shown.
  const { looking, ready, changeTyped } = run;
  const reads = run.typed.play !== false;
  useEffect(() => { if (!looking && !ready && reads) changeTyped({ show: true, play: false }); }, [looking, ready, reads, changeTyped]);
  if (!open && !saying) {
    return (
      <div className="st-typed st-typed-closed">
        <Button ghost icon="volume" label={t('common.text-to-speech')} className="st-typed-open" title={t('studio.spoken-run.type-words-show')} onClick={() => setOpen(true)} />
      </div>
    );
  }
  return (
    <div className="st-typed" role="group" aria-label={t('common.text-to-speech')}>
      <div className="st-typed-head">
        <span>{t('common.text-to-speech')}</span>
        {saying ? <span className="st-typed-now" dir="auto" data-making={run.state.making ? '1' : '0'}>{`${run.state.index + 1}/${run.state.count} ${run.state.lines[run.state.index] || ''}`}</span> : null}
        {!saying ? <Button ghost icon="close" className="st-typed-close" title={t('common.close')} onClick={() => setOpen(false)} /> : null}
      </div>
      <SpeakBox
        shell={s.shell}
        voices={run.voices}
        value={run.typed}
        onChange={run.changeTyped}
        onSay={say}
        onStop={run.stop}
        saying={saying}
        busy={saying && run.state.making}
        sayLabel={run.typed.play === false ? t('studio.spoken-run.show-the-words') : t('studio.spoken-run.read-words-aloud')}
      />
      {run.typed.show !== false ? <TypewriterChoice run={run} /> : null}
      {run.ready && run.typed.play !== false ? <HearSwitch run={run} /> : null}
      {/* The sound that goes out: the air's while a shot is held for it (Live). */}
      {run.typed.play !== false && (s.airStack || s.stack)?.sound !== 'on' ? <p className="st-col-note st-spoken-quiet">{t('studio.spoken-run.sound-not-voice')}</p> : null}
    </div>
  );
}

/**
 * Whether the shown words type in, and, read aloud, how their typing goes
 * with the voice: the same two timings as a caption read aloud in Video.
 * A change counts from the next line said.
 */
function TypewriterChoice({ run }) {
  const { typewriter, typing, play } = run.typed;
  const how = TYPING_TIMINGS.includes(typing) ? typing : 'with';
  return (
    <div className="st-typed-typing">
      <label className="st-check" title={t('studio.spoken-run.each-line-types')}>
        <input type="checkbox" className="st-typed-typewriter" checked={typewriter === true} onChange={(e) => run.changeTyped({ typewriter: e.target.checked })} />
        <span>{t('video.inspector.typewriter')}</span>
      </label>
      {typewriter === true ? (
        <Select value={how} className="st-typed-timing" disabled={play === false} title={play === false ? t('studio.spoken-run.shown-only-words') : undefined} onChange={(e) => run.changeTyped({ typing: e.target.value })}>
          <option value="with">{t('video.spoken.with-the-voice')}</option>
          <option value="first">{t('video.spoken.voice-first')}</option>
        </Select>
      ) : null}
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
.st-typed { padding: 4px 8px; font-size: 12px; color: var(--ink-2); max-width: 560px; }
.st-typed-head { display: flex; align-items: center; gap: 8px; font-weight: 600; }
.st-typed-now { font-weight: 400; max-width: 340px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.st-typed-close { margin-left: auto; }
.st-typed-typing { display: flex; align-items: center; gap: 10px; padding: 4px 0; }
.st-spoken-hear { display: flex; align-items: center; gap: 6px; padding: 4px 0; font-size: 12px; color: var(--ink-2); }
`;
