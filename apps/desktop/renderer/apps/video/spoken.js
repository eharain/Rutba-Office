// Spoken lines in Video (D-2026-11-13-31): a title's or a caption's words
// read aloud line by line onto the timeline. Each line becomes its words as a
// text clip and its voice as a sound clip under it, one group a line, one
// after another; the person chooses to show the text, play the voice, or
// both. The voices are WAVs written into the project's "Made in Video"
// folder, under "Spoken lines", as the library's sound effects are, so the
// project, the player and Export use them as any other sound.
//
// Without an added voice the section is one icon that opens Voices and
// languages.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, basename, t } from '@rutba/office-ui';
import { wavBytes } from '@rutba/studio/effects';
import {
  isSpoken, spokenLayerOf, spokenLayer, newSpokenLayer, spokenChoice, setSpokenChoice, wordsToRead,
  placeSpokenLines, spokenFileName,
} from '@rutba/studio/spoken';
import { useAddedVoices, voiceOrFirst, makeLine, useLinePlayer, AddVoiceButton, VoiceFields, ShowPlay, LineList } from '../../spoken/voices.js';

const LAST_VOICE = 'video.spokenVoice';

/** Whether a clip has a Spoken lines section: a title's words, a caption, or any clip of a spoken layer. */
export const hasSpokenLines = (clip) => Boolean(clip && (clip.type === 'text' || isSpoken(clip)));

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
      toast(t('The line could not be read: {reason}', { reason: err?.message || String(err) }), { tone: 'bad', ms: 7000 });
    }
  }, [player, lines, lineSound, toast]);

  const readOnto = useCallback(async () => {
    if (!voice || !lines.length) return;
    player.stop();
    stopRef.current = false;
    const id = layer || newSpokenLayer(doc);
    const mark = Date.now().toString(36);
    try {
      const dir = await madeFolder();
      const sep = dir.includes('/') && !dir.includes('\\') ? '/' : '\\';
      const folder = `${dir}${sep}${t('Spoken lines')}`;
      await shell.fs.mkdirp({ path: folder });
      const out = [];
      for (let i = 0; i < lines.length; i++) {
        if (stopRef.current) { setBusy(null); return; }
        setBusy({ n: i + 1, count: lines.length });
        const sound = await lineSound(lines[i]);
        const file = `${folder}${sep}${spokenFileName(id, i, mark)}`;
        const bytes = wavBytes(sound.samples, sound.rate);
        await shell.fs.write({ path: file, bytes });
        const described = await media.describe(file, { name: basename(file), size: bytes.length }).catch(() => null);
        out.push({ text: lines[i], url: described?.url || file, name: basename(file), seconds: described?.duration || sound.seconds, size: bytes.length });
      }
      if (stopRef.current) { setBusy(null); return; }
      edit(layer ? t('Read Again') : t('Read the Lines Aloud'), (d) => placeSpokenLines(d, { from: clip.id, lines: out, voice: voice.id, speed }).doc);
      shell.store?.set({ key: LAST_VOICE, value: voice.id }).catch?.(() => {});
      setBusy(null);
    } catch (err) {
      setBusy(null);
      toast(t('The lines could not be read aloud: {reason}', { reason: err?.message || String(err) }), { tone: 'bad', ms: 8000 });
    }
  }, [voice, lines, player, layer, doc, madeFolder, shell, lineSound, media, edit, clip.id, speed, toast]);

  const choose = (change) => edit(t('Spoken Lines'), (d) => setSpokenChoice(d, layer, change));

  if (looking) return null;
  if (!ready) {
    return (
      <section className="vd-insp-section vd-spoken vd-spoken-none">
        <h3 className="vd-spoken-head">{t('Spoken lines')}<AddVoiceButton shell={shell} /></h3>
      </section>
    );
  }

  return (
    <section className="vd-insp-section vd-spoken">
      <h3 className="vd-spoken-head"><Icon name="volume" size={13} />{t('Spoken lines')}</h3>
      {recorded ? <p className="spoken-note">{t('Read while recording: the voice is in the recording\'s own sound.')}</p> : (
        <VoiceFields Row={Row} voices={voices} voice={voice?.id} speed={speed} disabled={Boolean(busy)} onVoice={setVoiceId} onSpeed={setSpeed} />
      )}
      <LineList
        lines={lines}
        lengths={info ? info.lines.map((l) => { const c = l.text || l.sound; return c ? c.timing.end - c.timing.start : NaN; }) : []}
        playing={player.playing}
        disabled={Boolean(busy && !busy.hearing) || recorded}
        onPlay={hear}
      />
      {layer ? <ShowPlay show={choice.show} play={choice.play} onChange={choose} disabled={Boolean(busy)} /> : null}
      {busy && !busy.hearing ? (
        <div className="spoken-progress" role="status">
          <span>{t('Reading line {n} of {count}…', { n: busy.n, count: busy.count })}</span>
          <Button label={t('Stop')} className="vd-spoken-stop" onClick={() => { stopRef.current = true; }} />
        </div>
      ) : null}
      {!recorded ? (
        <div className="spoken-actions">
          <Button
            icon={layer ? 'refresh' : 'volume'}
            label={layer ? t('Read again') : t('Read the lines onto the timeline')}
            className="vd-spoken-read"
            disabled={Boolean(busy) || !lines.length}
            title={layer ? t('Each line made again from its words now, with this voice and speed') : t('Each line read aloud, its words shown while its voice plays, one after another')}
            onClick={readOnto}
          />
        </div>
      ) : null}
    </section>
  );
}

export const VIDEO_SPOKEN_CSS = `
.vd-spoken-head { display: flex; align-items: center; gap: 6px; }
.vd-spoken-none .spoken-add-voice { margin-left: auto; }
`;
