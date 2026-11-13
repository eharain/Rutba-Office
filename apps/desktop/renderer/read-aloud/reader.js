// The reader the apps share (C7): it takes pieces of text, each with where it
// came from and its language, splits them into sentences, reads them one by
// one and says which sentence is being read, so the app can highlight it.
//
// It speaks through the system's voices, as Read Aloud did before (the
// Windows voice stays the default), or, when the person has picked an added
// voice, hands each sentence to the speech engine add-on
// (`speech.say`, C6) and plays the samples it answers. Everything about how
// a voice turns words into sound is the engine's: the reader never sees a
// phoneme. When the engine fails, the rest of the session reads with the
// system's voice, and the reader says so once.
//
// It remembers the voice chosen for each language, and the speed, in the
// profile.

const VOICES_KEY = 'readAloud.voices';
const SPEED_KEY = 'readAloud.speed';
export const SPEEDS = [0.75, 1, 1.25, 1.5, 1.75, 2];

/** The sentences of a text: { start, end, text } by character offsets, spaces left out. */
export function sentencesOf(text, language) {
  const s = String(text || '');
  const out = [];
  const push = (start, end) => {
    while (start < end && /\s/.test(s[start])) start += 1;
    while (end > start && /\s/.test(s[end - 1])) end -= 1;
    if (end > start) out.push({ start, end, text: s.slice(start, end) });
  };
  if (typeof Intl !== 'undefined' && Intl.Segmenter) {
    let seg;
    try {
      seg = new Intl.Segmenter(language || undefined, { granularity: 'sentence' });
    } catch {
      seg = new Intl.Segmenter(undefined, { granularity: 'sentence' });
    }
    for (const part of seg.segment(s)) push(part.index, part.index + part.segment.length);
    return out;
  }
  const re = /[^.!?۔؟。]+[.!?۔؟。]*/g;
  for (let m = re.exec(s); m; m = re.exec(s)) push(m.index, m.index + m[0].length);
  return out;
}

/** The sentences of every piece, in order: { piece, source, language, start, end, text }. */
export function itemsOf(pieces) {
  const items = [];
  pieces.forEach((p, piece) => {
    for (const sentence of sentencesOf(p.text, p.language)) items.push({ piece, source: p.source, language: p.language || null, ...sentence });
  });
  return items;
}

/** A voice chosen: { kind: 'system', name } or { kind: 'added', id }. */
const sameVoice = (a, b) => a && b && a.kind === b.kind && (a.kind === 'added' ? a.id === b.id : a.name === b.name);

/**
 * @param {object} o
 * @param {object} o.shell          the window's bridge (store, and speech when an engine is added)
 * @param {SpeechSynthesis} [o.synth]
 * @param {() => AudioContext} [o.audio]
 * @param {(state: object) => void} [o.onState]   every change: { playing, paused, index, count, item, voice, speed, said }
 */
export function createReader({ shell, synth = globalThis.speechSynthesis, audio = () => new AudioContext(), onState = () => {} }) {
  let items = [];
  let index = 0;
  let playing = false;
  let paused = false;
  let speed = 1;
  let voices = {};
  let engineFailed = false;
  let said = null;
  let token = 0;
  let ctx = null;
  let source = null;
  const ahead = new Map(); // index -> Promise<{ samples, rate }>
  let onSentence = () => {};
  let onEnd = () => {};

  const ready = Promise.all([
    shell.store?.get({ key: VOICES_KEY, fallback: {} }).then((v) => { voices = v && typeof v === 'object' ? v : {}; }).catch(() => {}),
    shell.store?.get({ key: SPEED_KEY, fallback: 1 }).then((v) => { if (SPEEDS.includes(Number(v))) speed = Number(v); }).catch(() => {}),
  ]);

  const state = () => ({ playing, paused, index, count: items.length, item: items[index] || null, voice: voiceFor(items[index]?.language), speed, said });
  const tell = () => onState(state());

  /** The voice for a language: the person's choice, else the system's own. */
  function voiceFor(language) {
    const key = String(language || '').toLowerCase() || '*';
    const v = voices[key] || voices[key.split('-')[0]] || voices['*'] || null;
    if (v?.kind === 'added' && engineFailed) return null;
    return v;
  }

  /** The voices there are for a language: the system's, then the added ones that read it. */
  async function voicesFor(language) {
    const lang = String(language || '').toLowerCase();
    const base = lang.split('-')[0];
    const system = (synth?.getVoices?.() || [])
      .filter((v) => !lang || String(v.lang).toLowerCase().startsWith(base))
      .map((v) => ({ kind: 'system', name: v.name, label: v.name, language: v.lang }));
    let added = [];
    if (shell.speech?.voices && !engineFailed) {
      try {
        added = (await shell.speech.voices())
          .filter((v) => !lang || [v.language, ...(v.also || [])].some((l) => String(l).toLowerCase().startsWith(base)))
          .map((v) => ({ kind: 'added', id: v.id, label: v.name, language: v.language }));
      } catch {
        added = [];
      }
    }
    return [...system, ...added];
  }

  function stopSound() {
    token += 1;
    try {
      synth?.cancel?.();
    } catch {
      // nothing playing
    }
    try {
      source?.stop();
    } catch {
      // already stopped
    }
    source = null;
  }

  const fetchSamples = (i) => {
    const it = items[i];
    const v = voiceFor(it?.language);
    if (!it || v?.kind !== 'added') return null;
    if (!ahead.has(i)) ahead.set(i, shell.speech.say({ text: it.text, voice: v.id, speed }));
    return ahead.get(i);
  };

  function fellBack(err) {
    engineFailed = true;
    ahead.clear();
    if (!said) {
      said = 'The added voice stopped working, so the rest is read with the system\'s voice.';
      tell();
    }
    return err;
  }

  async function play(i) {
    const mine = (token += 1);
    index = i;
    if (i >= items.length) {
      playing = false;
      paused = false;
      tell();
      onEnd();
      return;
    }
    playing = true;
    paused = false;
    const it = items[i];
    onSentence(it, i);
    tell();
    const v = voiceFor(it.language);
    if (v?.kind === 'added') {
      try {
        const out = await fetchSamples(i);
        if (mine !== token) return;
        fetchSamples(i + 1)?.catch(() => {});
        ctx ||= audio();
        const samples = out?.samples instanceof Float32Array ? out.samples : new Float32Array(out?.samples || []);
        const buffer = ctx.createBuffer(1, Math.max(1, samples.length), out?.rate || 24000);
        buffer.copyToChannel(samples, 0);
        source = ctx.createBufferSource();
        source.buffer = buffer;
        source.connect(ctx.destination);
        source.onended = () => { if (mine === token && !paused) { ahead.delete(i); play(i + 1); } };
        source.start();
        return;
      } catch (err) {
        if (mine !== token) return;
        fellBack(err);
      }
    }
    if (!synth) {
      playing = false;
      said = 'Speech is not available on this system.';
      tell();
      return;
    }
    const u = new SpeechSynthesisUtterance(it.text);
    const chosen = v?.kind === 'system' ? (synth.getVoices?.() || []).find((x) => x.name === v.name) : null;
    if (chosen) u.voice = chosen;
    if (it.language) u.lang = it.language;
    u.rate = speed;
    u.onend = () => { if (mine === token && !paused) play(i + 1); };
    u.onerror = () => { if (mine === token) { playing = false; tell(); } };
    synth.speak(u);
  }

  return {
    ready,
    voicesFor,
    voiceFor,
    state,
    /** Reads pieces from the first: [{ text, source, language }]. */
    async read(pieces, { onSentence: os = () => {}, onEnd: oe = () => {} } = {}) {
      await ready;
      stopSound();
      ahead.clear();
      items = itemsOf(pieces);
      onSentence = os;
      onEnd = oe;
      said = null;
      if (!items.length) {
        tell();
        oe();
        return false;
      }
      play(0);
      return true;
    },
    pause() {
      if (!playing || paused) return;
      paused = true;
      try { synth?.pause?.(); } catch { /* not speaking */ }
      ctx?.suspend?.().catch?.(() => {});
      tell();
    },
    resume() {
      if (!paused) return;
      paused = false;
      try { synth?.resume?.(); } catch { /* not speaking */ }
      ctx?.resume?.().catch?.(() => {});
      tell();
    },
    back() {
      stopSound();
      play(Math.max(0, index - 1));
    },
    forward() {
      stopSound();
      play(Math.min(items.length, index + 1));
    },
    stop() {
      stopSound();
      ahead.clear();
      playing = false;
      paused = false;
      tell();
    },
    /** A new speed, from the sentence being read. */
    async setSpeed(next) {
      if (!SPEEDS.includes(Number(next))) return;
      speed = Number(next);
      ahead.clear();
      shell.store?.set({ key: SPEED_KEY, value: speed }).catch?.(() => {});
      if (playing && !paused) { stopSound(); play(index); } else tell();
    },
    /** The voice for a language, from the sentence being read; null for the system's own. */
    async setVoice(language, voice) {
      const key = String(language || '').toLowerCase() || '*';
      if (sameVoice(voices[key], voice)) return;
      voices = { ...voices, [key]: voice || null };
      if (voice?.kind === 'added') engineFailed = false;
      ahead.clear();
      shell.store?.set({ key: VOICES_KEY, value: voices }).catch?.(() => {});
      if (playing && !paused) { stopSound(); play(index); } else tell();
    },
  };
}
