// Spoken lines in Video (D-2026-11-13-31): a title's or a caption's words
// read aloud line by line onto the timeline. Each line becomes its words as a
// text clip and its voice as a sound clip under it, one group a line, one
// after another; the person picks one of the three, Show and read aloud,
// Read aloud only or Show only. The voices are WAVs written into the
// project's "Made in Video" folder, under "Spoken lines", as the library's
// sound effects are, so the project, the player and Export use them as any
// other sound.
//
// Without an added voice the section is one icon that opens Voices and
// languages.
//
// Text to speech (Insert, D-2026-11-16-01) is the short way in: words typed in
// one box, the same box the Recorder and Live have, land at the playhead as a
// spoken layer, shown and read, read only or shown only. Shown only needs no
// voice: each line stays up for about as long as it takes to read.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Dialog, Icon, Select, basename, t } from '@rutba/office-ui';
import { wavBytes } from '@rutba/studio/effects';
import {
  isSpoken, spokenLayerOf, spokenLayer, newSpokenLayer, spokenChoice, setSpokenChoice, wordsToRead, setSpokenTyping, spokenTypingOf,
  placeSpokenLines, placeTypedLines, spokenFileName, typedLinesOf, readingSeconds,
} from '@rutba/studio/spoken';
import { useAddedVoices, voiceOrFirst, makeLine, useLinePlayer, AddVoiceButton, VoiceFields, SpeakChoice, SpeakBox, LineList } from '../../spoken/voices.js';

const LAST_VOICE = 'video.spokenVoice';
const STOPPED = Symbol('stopped');

/**
 * Each line's voice written as a WAV into the project's "Spoken lines"
 * folder and described into the media library, one after another, until
 * `isStopped()`. Answers { out, written, stopped }: `out` the lines as
 * placeSpokenLines takes them, `written` every file made (for tidying when
 * they are not used). A failure carries the files made so far as `written`.
 */
async function writeLineSounds({ shell, madeFolder, media, layer, lines, lineSound, isStopped, stopped, onLine }) {
  const mark = Date.now().toString(36);
  const written = [];
  const out = [];
  try {
    const dir = await madeFolder();
    const sep = dir.includes('/') && !dir.includes('\\') ? '/' : '\\';
    const folder = `${dir}${sep}${t('common.spoken-lines')}`;
    await shell.fs.mkdirp({ path: folder });
    for (let i = 0; i < lines.length; i++) {
      if (isStopped()) break;
      onLine?.(i);
      const sound = await Promise.race([lineSound(lines[i]), stopped]);
      if (sound === STOPPED || isStopped()) break;
      const file = `${folder}${sep}${spokenFileName(layer, i, mark)}`;
      const bytes = wavBytes(sound.samples, sound.rate);
      await shell.fs.write({ path: file, bytes });
      written.push(file);
      const described = await media.describe(file, { name: basename(file), size: bytes.length }).catch(() => null);
      out.push({ text: lines[i], url: described?.url || file, name: basename(file), seconds: described?.duration || sound.seconds, size: bytes.length });
    }
  } catch (err) {
    const e = err instanceof Error ? err : new Error(String(err));
    e.written = written;
    throw e;
  }
  return { out, written, stopped: isStopped() };
}

/** Files made for a reading that did not reach the timeline: to the trash. */
const tidyFiles = (shell, files) => { for (const file of files || []) shell.fs.remove?.({ path: file })?.catch?.(() => {}); };

/** Whether a clip has a Spoken lines section: a title's words, a caption, or any clip of a spoken layer. */
export const hasSpokenLines = (clip) => Boolean(clip && (clip.type === 'text' || clip.type === 'caption' || isSpoken(clip)));

/**
 * @param {object} p
 * @param {object} p.shell
 * @param {object} p.doc          the project
 * @param {object} p.clip         the clip picked
 * @param {Function} p.edit       (label, fn) one undo step
 * @param {Function} p.madeFolder the folder files made for the project go in
 * @param {object} p.media        the media library: a file written is described into it
 * @param {Function} p.toast
 * @param {Function} [p.Row]      the Inspector's row
 */
export function SpokenLines({ shell, doc, clip, edit, madeFolder, media, toast, Row }) {
  const { voices, ready, looking } = useAddedVoices(shell);
  const layer = spokenLayerOf(clip);
  const info = useMemo(() => (layer ? spokenLayer(doc, layer) : null), [doc, layer]);
  const lines = useMemo(() => wordsToRead(doc, clip.id), [doc, clip.id]);
  const choice = layer ? spokenChoice(doc, layer) : { show: true, play: true };
  // A caption's lines read aloud: how their typing goes with the voice (S19).
  const typing = layer ? spokenTypingOf(doc, layer) : null;
  const recorded = Boolean(info?.lines[0]?.text?.spoken?.recorded);
  const [remembered, setRemembered] = useState('');
  const [voiceId, setVoiceId] = useState('');
  const [speed, setSpeed] = useState(1);
  const [busy, setBusy] = useState(null); // { n, count }
  const stopRef = useRef(false);
  const player = useLinePlayer();
  const made = useRef(new Map()); // `${voice}|${speed}|${text}` -> { samples, rate, seconds }

  useEffect(() => {
    let live = true;
    shell.store?.get({ key: LAST_VOICE, fallback: '' }).then((v) => { if (live && typeof v === 'string') setRemembered(v); }).catch(() => {});
    return () => { live = false; };
  }, [shell]);
  // The layer's own voice and speed when it has them, else the last voice used.
  useEffect(() => {
    setVoiceId(info?.voice || remembered || '');
    setSpeed(info?.speed || 1);
  }, [info?.voice, info?.speed, remembered, clip.id]);
  useEffect(() => () => { stopRef.current = true; player.stop(); }, [clip.id]);

  const voice = voiceOrFirst(voices, voiceId);

  const lineSound = useCallback(async (text) => {
    const key = `${voice.id}|${speed}|${text}`;
    if (!made.current.has(key)) made.current.set(key, await makeLine(shell, { text, voice: voice.id, speed, language: voice.language }));
    return made.current.get(key);
  }, [shell, voice, speed]);

  const hear = useCallback(async (i) => {
    if (player.playing === i) { player.stop(); return; }
    try {
      player.stop();
      setBusy({ n: i + 1, count: lines.length, hearing: true });
      const out = await lineSound(lines[i]);
      setBusy(null);
      await player.play(i, out);
    } catch (err) {
      setBusy(null);
      toast(t('common.line-could-not', { reason: err?.message || String(err) }), { tone: 'bad', ms: 7000 });
    }
  }, [player, lines, lineSound, toast]);

  // Stop answers at once: the line being made is let finish in the engine and left unused.
  const stopNow = useRef(null);
  // `choice`: { show, play } to give the lines as they are read, else what they had.
  const readOnto = useCallback(async ({ show, play } = {}) => {
    if (!voice || !lines.length) return;
    player.stop();
    stopRef.current = false;
    const stopped = new Promise((resolve) => { stopNow.current = () => { stopRef.current = true; resolve(STOPPED); }; });
    const id = layer || newSpokenLayer(doc);
    let written = [];
    // A reading that does not reach the timeline leaves no files of its own behind (to the trash).
    const tidy = () => tidyFiles(shell, written);
    try {
      const made = await writeLineSounds({ shell, madeFolder, media, layer: id, lines, lineSound, isStopped: () => stopRef.current, stopped, onLine: (i) => setBusy({ n: i + 1, count: lines.length }) });
      written = made.written;
      const { out } = made;
      if (stopRef.current) { tidy(); setBusy(null); return; }
      // An edit refused (a clip on a locked track, which says so itself) leaves nothing on the timeline to use the files.
      const placed = edit(layer ? t('sheets.read-again') : t('video.spoken.read-lines-aloud'), (d) => placeSpokenLines(d, { from: clip.id, lines: out, voice: voice.id, speed, show, play }).doc);
      if (placed === false) { tidy(); setBusy(null); return; }
      shell.store?.set({ key: LAST_VOICE, value: voice.id }).catch?.(() => {});
      setBusy(null);
    } catch (err) {
      if (err?.written) written = err.written;
      tidy();
      setBusy(null);
      toast(t('common.lines-could-not', { reason: err?.message || String(err) }), { tone: 'bad', ms: 8000 });
    } finally {
      stopNow.current = null;
    }
  }, [voice, lines, player, layer, doc, madeFolder, shell, lineSound, media, edit, clip.id, speed, toast]);

  // Words with no voice yet (typed to be shown only): a choice that reads them aloud reads them, in one step.
  const hasSounds = Boolean(info?.lines.some((l) => l.sound));
  const choose = (change) => {
    if (change.play && !hasSounds && !recorded) {
      if (ready) readOnto(change);
      return;
    }
    edit(t('video.spoken.spoken-lines'), (d) => setSpokenChoice(d, layer, change));
  };

  if (looking) return null;
  // With no voice, a text not yet read is offered the one icon alone; a layer already read keeps
  // its two switches and its lines (they need no voice), and only what makes sound waits for one.
  if (!ready && !layer) {
    return (
      <section className="vd-insp-section vd-spoken vd-spoken-none">
        <h3 className="vd-spoken-head">{t('common.spoken-lines')}<AddVoiceButton shell={shell} /></h3>
      </section>
    );
  }
  const canSpeak = ready && !recorded;
  // Read while recording, the voice is in the take's own sound: it can be hidden or shown, never taken out,
  // and words shown alone then have no voice to give. Otherwise a read choice needs a voice made or one to make it with.
  const allowed = recorded ? (choice.play ? ['both', 'read'] : ['show']) : null;
  const canRead = recorded || hasSounds || ready;
  // Each line's length, from the same lines the list shows (one whose words were taken away is in neither).
  const lengths = info ? info.lines.filter((l) => String(l.text?.text || '').trim()).map((l) => l.text.timing.end - l.text.timing.start) : [];

  return (
    <section className={`vd-insp-section vd-spoken${ready ? '' : ' vd-spoken-none'}`}>
      <h3 className="vd-spoken-head"><Icon name="volume" size={13} />{t('common.spoken-lines')}{!ready && !recorded ? <AddVoiceButton shell={shell} /> : null}</h3>
      {recorded ? <p className="spoken-note">{t('video.spoken.read-while-recording')}</p> : null}
      {canSpeak ? <VoiceFields Row={Row} voices={voices} voice={voice?.id} speed={speed} disabled={Boolean(busy)} onVoice={setVoiceId} onSpeed={setSpeed} /> : null}
      <LineList
        lines={lines}
        lengths={lengths}
        playing={player.playing}
        disabled={Boolean(busy && !busy.hearing) || !canSpeak}
        onPlay={hear}
      />
      {layer ? <SpeakChoice show={choice.show} play={choice.play} canRead={canRead} allowed={allowed} onChange={choose} disabled={Boolean(busy)} /> : null}
      {layer && typing ? (
        <Row label={t('video.inspector.typing')}>
          <Select value={typing} className="vd-spoken-typing" disabled={Boolean(busy) || !choice.play} title={choice.play ? undefined : t('video.spoken.words-shown-only')} onChange={(e) => edit(t('video.inspector.typing'), (d) => setSpokenTyping(d, layer, e.target.value))}>
            <option value="with">{t('video.spoken.with-the-voice')}</option>
            <option value="first">{t('video.spoken.voice-first')}</option>
          </Select>
        </Row>
      ) : null}
      {busy && !busy.hearing ? (
        <div className="spoken-progress" role="status">
          <span>{t('common.reading-line-n', { n: busy.n, count: busy.count })}</span>
          <Button label={t('common.stop')} className="vd-spoken-stop" onClick={() => { stopRef.current = true; stopNow.current?.(); }} />
        </div>
      ) : null}
      {canSpeak ? (
        <div className="spoken-actions">
          <Button
            icon={layer ? 'refresh' : 'volume'}
            label={layer ? t('base.read-again') : t('video.spoken.read-lines-timeline')}
            className="vd-spoken-read"
            disabled={Boolean(busy) || !lines.length}
            title={layer ? t('video.spoken.each-line-made') : t('video.spoken.each-line-read')}
            onClick={() => readOnto()}
          />
        </div>
      ) : null}
    </section>
  );
}

/**
 * Text to speech (Insert): words typed, the choice of three, a voice, and Add
 * at the playhead, which makes each line's voice (with no voice the words
 * can only be shown) and lands them as one undo step. `onPlaced` is given
 * the clips made.
 */
export function SayDialog({ shell, doc, at, edit, madeFolder, media, toast, onClose, onPlaced }) {
  const { voices, ready, looking } = useAddedVoices(shell);
  const [value, setValue] = useState({ text: '', show: true, play: true, voice: '', speed: 1 });
  const [busy, setBusy] = useState(null); // { n, count }
  const stopRef = useRef(false);
  const stopNow = useRef(null);
  const change = useCallback((patch) => setValue((v) => ({ ...v, ...patch })), []);
  useEffect(() => {
    let live = true;
    shell.store?.get({ key: LAST_VOICE, fallback: '' }).then((v) => { if (live && typeof v === 'string' && v) setValue((was) => (was.voice ? was : { ...was, voice: v })); }).catch(() => {});
    return () => { live = false; stopRef.current = true; stopNow.current?.(); };
  }, [shell]);
  // With no voice the words can only be shown.
  useEffect(() => { if (!looking && !ready) setValue((v) => (v.play ? { ...v, show: true, play: false } : v)); }, [looking, ready]);

  const voice = voiceOrFirst(voices, value.voice);
  const say = async () => {
    const lines = typedLinesOf(value.text);
    if (!lines.length) return;
    const { show, play } = value;
    const speed = value.speed || 1;
    let written = [];
    stopRef.current = false;
    const stopped = new Promise((resolve) => { stopNow.current = () => { stopRef.current = true; resolve(STOPPED); }; });
    setBusy({ n: 1, count: lines.length });
    try {
      let list;
      if (voice) {
        const lineSound = (text) => makeLine(shell, { text, voice: voice.id, speed, language: voice.language });
        const made = await writeLineSounds({ shell, madeFolder, media, layer: newSpokenLayer(doc), lines, lineSound, isStopped: () => stopRef.current, stopped, onLine: (i) => setBusy({ n: i + 1, count: lines.length }) });
        written = made.written;
        if (made.stopped) { tidyFiles(shell, written); setBusy(null); return; }
        list = made.out;
      } else {
        list = lines.map((text) => ({ text, seconds: readingSeconds(text, { speed }) }));
      }
      let ids = [];
      const placed = edit(t('common.text-to-speech'), (d) => { const r = placeTypedLines(d, { at, lines: list, voice: voice?.id || null, speed, show, play: Boolean(voice) && play }); ids = r.ids; return r.doc; });
      if (placed === false || !ids.length) { tidyFiles(shell, written); setBusy(null); return; }
      if (voice) shell.store?.set({ key: LAST_VOICE, value: voice.id }).catch?.(() => {});
      setBusy(null);
      onPlaced?.(ids);
    } catch (err) {
      tidyFiles(shell, err?.written || written);
      setBusy(null);
      toast(t('common.lines-could-not', { reason: err?.message || String(err) }), { tone: 'bad', ms: 8000 });
    } finally {
      stopNow.current = null;
    }
  };

  return (
    <Dialog title={t('common.text-to-speech')} width={460} className="vd-say" onClose={() => { if (!busy) onClose(); }}>
      <p className="spoken-note">{t('video.spoken.words-land-playhead')}</p>
      <SpeakBox
        shell={shell}
        voices={voices}
        value={value}
        onChange={change}
        onSay={say}
        onStop={() => { stopRef.current = true; stopNow.current?.(); }}
        saying={Boolean(busy)}
        disabled={Boolean(busy)}
        sayLabel={t('video.library.add-playhead')}
        sayIcon="plus"
        autoFocus
      />
      {busy ? <p className="spoken-progress" role="status">{t('common.reading-line-n', { n: busy.n, count: busy.count })}</p> : null}
    </Dialog>
  );
}

export const VIDEO_SPOKEN_CSS = `
.vd-spoken-head { display: flex; align-items: center; gap: 6px; }
.vd-spoken-none .spoken-add-voice { margin-left: auto; }
`;
