// Transitions → Sound: a few short sounds of the suite's own, made here
// rather than shipped — a chime, a click, a coin and the like, under the
// names PowerPoint gives its own — each as the WAV a transition plays.
// The same name always makes the same bytes.

const RATE = 22050;

/** A seeded noise source, so a sound is the same every time it is made. */
function noise(seed = 7) {
  let s = seed >>> 0;
  return () => {
    s = (s * 1664525 + 1013904223) >>> 0;
    return (s / 4294967296) * 2 - 1;
  };
}

/** Samples in -1…1 as a mono 16-bit PCM WAV. */
export function wavOf(samples, rate = RATE) {
  const pcm = new Int16Array(samples.length);
  for (let i = 0; i < samples.length; i++) pcm[i] = toPcm(samples[i]);
  return wavOfPcm(pcm, rate);
}

/** A sample from -1 to 1 as the 16-bit value a WAV holds, with a little headroom. */
export const toPcm = (s) => Math.round(Math.max(-1, Math.min(1, s)) * 30000);

/** A mono WAV of 16-bit samples already made. */
export function wavOfPcm(pcm, rate = RATE) {
  const n = pcm.length;
  const buf = new Uint8Array(44 + n * 2);
  const v = new DataView(buf.buffer);
  const text = (at, s) => { for (let i = 0; i < s.length; i++) buf[at + i] = s.charCodeAt(i); };
  text(0, 'RIFF'); v.setUint32(4, 36 + n * 2, true); text(8, 'WAVE');
  text(12, 'fmt '); v.setUint32(16, 16, true); v.setUint16(20, 1, true); v.setUint16(22, 1, true);
  v.setUint32(24, rate, true); v.setUint32(28, rate * 2, true); v.setUint16(32, 2, true); v.setUint16(34, 16, true);
  text(36, 'data'); v.setUint32(40, n * 2, true);
  for (let i = 0; i < n; i++) v.setInt16(44 + i * 2, pcm[i], true);
  return buf;
}

const length = (seconds) => new Float32Array(Math.round(seconds * RATE));
const tone = (out, hz, from, to, gain, decay) => {
  for (let i = Math.round(from * RATE); i < Math.min(out.length, Math.round(to * RATE)); i++) {
    const t = i / RATE - from;
    out[i] += Math.sin(2 * Math.PI * hz * t) * gain * Math.exp(-t * decay);
  }
};

/** Each sound, by the name the menu shows. */
const MAKERS = {
  Breeze: () => {
    const out = length(1.6); const r = noise(11); let low = 0;
    for (let i = 0; i < out.length; i++) { const t = i / RATE; low += (r() - low) * 0.03; out[i] = low * 4 * Math.sin((Math.PI * t) / 1.6); }
    return out;
  },
  Camera: () => {
    const out = length(0.3); const r = noise(5);
    for (let i = 0; i < out.length; i++) { const t = i / RATE; out[i] = r() * (t < 0.02 ? 0.9 : t > 0.12 && t < 0.15 ? 0.6 : 0) * Math.exp(-t * 6); }
    return out;
  },
  Chime: () => { const out = length(1.4); tone(out, 880, 0, 1.4, 0.45, 3); tone(out, 1320, 0, 1.4, 0.25, 4); tone(out, 1760, 0, 1.4, 0.15, 5); return out; },
  Click: () => {
    const out = length(0.06); const r = noise(3);
    for (let i = 0; i < out.length; i++) out[i] = r() * Math.exp(-(i / RATE) * 120);
    return out;
  },
  Coin: () => { const out = length(0.55); tone(out, 988, 0, 0.08, 0.5, 2); tone(out, 1319, 0.08, 0.55, 0.5, 6); return out; },
  'Drum Roll': () => {
    const out = length(1.3); const r = noise(9);
    for (let i = 0; i < out.length; i++) { const t = i / RATE; const hit = (t % 0.045) / 0.045; out[i] = r() * Math.exp(-hit * 6) * (0.3 + 0.5 * (t / 1.3)) * (t > 1.22 ? Math.exp(-(t - 1.22) * 40) : 1); }
    return out;
  },
  Laser: () => {
    const out = length(0.4); let phase = 0;
    for (let i = 0; i < out.length; i++) { const t = i / RATE; phase += (2 * Math.PI * (1800 - 4000 * t)) / RATE; out[i] = Math.sin(phase) * 0.5 * (1 - t / 0.4); }
    return out;
  },
  Whoosh: () => {
    const out = length(0.7); const r = noise(13); let low = 0;
    for (let i = 0; i < out.length; i++) { const t = i / RATE; low += (r() - low) * (0.05 + 0.3 * (t / 0.7)); out[i] = low * 2.2 * Math.sin((Math.PI * t) / 0.7); }
    return out;
  },
};

/** The names the Sound menu lists, in PowerPoint's alphabetical order. */
export const SOUNDS = Object.keys(MAKERS);

/** One sound's WAV bytes, or null for a name this does not make. */
export function soundWav(name) {
  const make = MAKERS[name];
  return make ? wavOf(make()) : null;
}

/** The file name a sound is kept under in the deck, as PowerPoint names its own. */
export const soundFile = (name) => `${String(name).toLowerCase().replace(/\s+/g, '')}.wav`;
