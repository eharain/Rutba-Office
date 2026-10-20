// Sound effects for Video's library (the 2.7 plan: "library and panels:
// ... sound effects"). Each one is made here from sines and noise, the same
// every time, so the suite carries no recordings and fetches none; a WAV is
// written once beside the project when one is first used.

/** The effects in the library, in the order shown: key and length in seconds. */
export const SOUND_EFFECTS = [
    { key: 'whoosh', seconds: 0.7 },
    { key: 'swoosh', seconds: 0.35 },
    { key: 'pop', seconds: 0.12 },
    { key: 'click', seconds: 0.05 },
    { key: 'chime', seconds: 1.6 },
    { key: 'ding', seconds: 1.1 },
    { key: 'rise', seconds: 1 },
    { key: 'drop', seconds: 0.8 },
    { key: 'beep', seconds: 0.25 },
    { key: 'heartbeat', seconds: 1 },
];

/** A noise source that is the same every run. */
function noise(seed = 1) {
    let s = seed >>> 0 || 1;
    return () => {
        s ^= s << 13; s >>>= 0;
        s ^= s >>> 17;
        s ^= s << 5; s >>>= 0;
        return (s / 4294967296) * 2 - 1;
    };
}

/** A sine whose frequency goes from `f0` to `f1` over `n` samples, its phase carried. */
function sweep(out, rate, f0, f1, env, curve = 'exp') {
    let phase = 0;
    const n = out.length;
    for (let i = 0; i < n; i++) {
        const p = i / n;
        const f = curve === 'exp' ? f0 * Math.pow(f1 / f0, p) : f0 + (f1 - f0) * p;
        phase += (2 * Math.PI * f) / rate;
        out[i] += Math.sin(phase) * env(p, i / rate);
    }
}

/** Bell partials, each fading on its own. */
function bell(out, rate, partials) {
    for (const [f, a, decay] of partials) {
        for (let i = 0; i < out.length; i++) {
            const t = i / rate;
            out[i] += Math.sin(2 * Math.PI * f * t) * a * Math.exp(-t / decay) * Math.min(1, t * 400);
        }
    }
}

/** Noise through a low-pass that opens and closes, rising and falling: air going past. */
function air(out, rate, seed, from, to) {
    const r = noise(seed);
    let y = 0;
    for (let i = 0; i < out.length; i++) {
        const p = i / out.length;
        const shape = Math.sin(Math.PI * p) ** 2;
        const cutoff = from + (to - from) * shape;
        const k = 1 - Math.exp((-2 * Math.PI * cutoff) / rate);
        y += k * (r() - y);
        out[i] += y * shape * 2.2;
    }
}

const MAKERS = {
    whoosh: (o, rate) => air(o, rate, 7, 300, 2600),
    swoosh: (o, rate) => air(o, rate, 11, 800, 5200),
    pop: (o, rate) => sweep(o, rate, 520, 140, (p) => Math.exp(-p * 5) * Math.min(1, p * 60)),
    click: (o, rate) => {
        const r = noise(3);
        for (let i = 0; i < o.length; i++) o[i] += r() * Math.exp(-(i / rate) * 160);
        sweep(o, rate, 2400, 1800, (p) => 0.4 * Math.exp(-p * 9));
    },
    chime: (o, rate) => bell(o, rate, [[880, 0.42, 0.6], [1320, 0.24, 0.45], [1760, 0.16, 0.3], [2640, 0.08, 0.2]]),
    ding: (o, rate) => bell(o, rate, [[1046.5, 0.5, 0.4], [2093, 0.2, 0.25], [3139.5, 0.08, 0.12]]),
    rise: (o, rate) => sweep(o, rate, 200, 1200, (p) => 0.5 * Math.min(1, p * 8) * Math.min(1, (1 - p) * 10)),
    drop: (o, rate) => sweep(o, rate, 900, 90, (p) => 0.55 * Math.min(1, p * 30) * Math.min(1, (1 - p) * 8)),
    beep: (o, rate) => sweep(o, rate, 1000, 1000, (p) => 0.45 * Math.min(1, p * 40) * Math.min(1, (1 - p) * 40), 'lin'),
    heartbeat: (o, rate) => {
        for (const at of [0, 0.28]) {
            const from = Math.round(at * rate);
            const part = o.subarray(from, from + Math.round(0.18 * rate));
            sweep(part, rate, 90, 45, (p) => (at ? 0.6 : 0.8) * Math.exp(-p * 6) * Math.min(1, p * 50));
        }
    },
};

/** An effect's samples at `rate`: mono, peaking at 0.8 (about -2 dB). */
export function effectSamples(key, rate = 48000) {
    const def = SOUND_EFFECTS.find((e) => e.key === key);
    if (!def) return null;
    const out = new Float32Array(Math.max(1, Math.round(def.seconds * rate)));
    MAKERS[key](out, rate);
    // A short ramp at each end, so nothing clicks where it starts or stops.
    const ramp = Math.min(out.length >> 2, Math.round(rate * 0.003));
    for (let i = 0; i < ramp; i++) {
        out[i] *= i / ramp;
        out[out.length - 1 - i] *= i / ramp;
    }
    let peak = 0;
    for (let i = 0; i < out.length; i++) peak = Math.max(peak, Math.abs(out[i]));
    if (peak > 0) for (let i = 0; i < out.length; i++) out[i] *= 0.8 / peak;
    return out;
}

/** Samples as a 16-bit mono WAV file. */
export function wavBytes(samples, rate = 48000) {
    const data = new Uint8Array(44 + samples.length * 2);
    const v = new DataView(data.buffer);
    const text = (at, s) => { for (let i = 0; i < s.length; i++) data[at + i] = s.charCodeAt(i); };
    text(0, 'RIFF'); v.setUint32(4, 36 + samples.length * 2, true); text(8, 'WAVE');
    text(12, 'fmt '); v.setUint32(16, 16, true); v.setUint16(20, 1, true); v.setUint16(22, 1, true);
    v.setUint32(24, rate, true); v.setUint32(28, rate * 2, true); v.setUint16(32, 2, true); v.setUint16(34, 16, true);
    text(36, 'data'); v.setUint32(40, samples.length * 2, true);
    for (let i = 0; i < samples.length; i++) v.setInt16(44 + i * 2, Math.round(Math.max(-1, Math.min(1, samples[i])) * 32767), true);
    return data;
}

/** An effect as WAV bytes, or null for a key that is not one. */
export function effectWav(key, rate = 48000) {
    const s = effectSamples(key, rate);
    return s ? wavBytes(s, rate) : null;
}
