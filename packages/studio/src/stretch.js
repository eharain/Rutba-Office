// Sound made longer or shorter without its pitch changing, and its pitch
// moved without its length changing (the 2.7 plan: "speed with the pitch
// kept" and "pitch shift"). WSOLA: windows of the source laid down one after
// another at the new rate, each one nudged within a few milliseconds to where
// it best continues the one before, so the waves line up and nothing warbles.
// The nudge is found on a mono mix and applied to every channel alike.

/** The pitch factor of `semitones`: 12 is an octave up, twice the frequency. */
export const pitchFactor = (semitones) => Math.pow(2, (Number(semitones) || 0) / 12);

/** The pitch shifts a clip can be given, in semitones: an octave down to an octave up. */
export const PITCH_STEPS = [-12, -7, -5, -4, -3, -2, -1, 0, 1, 2, 3, 4, 5, 7, 12];

/**
 * `channels` (Float32Arrays of one length) stretched to `ratio` times their
 * length, the pitch unchanged. Answers new Float32Arrays.
 */
export function stretchChannels(channels, rate, ratio) {
    const r = Number(ratio);
    const input = channels[0]?.length || 0;
    if (!(r > 0) || Math.abs(r - 1) < 1e-4 || input === 0) return channels.map((c) => Float32Array.from(c));
    const N = Math.max(64, Math.round(rate * 0.04) & ~1); // a window of 40 ms
    const Hs = N / 2; // laid down every 20 ms
    const Ha = Hs / r; // read every 20 ms over the ratio
    const tol = Math.max(8, Math.round(rate * 0.01)); // nudged up to 10 ms either way
    const outLength = Math.max(1, Math.round(input * r));
    const mono = new Float32Array(input);
    for (const c of channels) for (let i = 0; i < input; i++) mono[i] += c[i] / channels.length;
    const win = new Float32Array(N);
    for (let i = 0; i < N; i++) win[i] = 0.5 - 0.5 * Math.cos((2 * Math.PI * i) / N);
    const out = channels.map(() => new Float32Array(outLength + N));
    const norm = new Float32Array(outLength + N);
    const at = (i) => (i >= 0 && i < input ? mono[i] : 0);
    // Where the last window was read from: its natural continuation is Hs further on.
    let prev = 0;
    const frames = Math.ceil(outLength / Hs) + 1;
    const step = 2; // the match is judged on every other sample, which is plenty for speech and music
    for (let k = 0; k < frames; k++) {
        const nominal = Math.round(k * Ha);
        let best = nominal;
        if (k > 0) {
            const follow = prev + Hs;
            let bestScore = -Infinity;
            for (let d = -tol; d <= tol; d += step) {
                const cand = nominal + d;
                if (cand < 0) continue;
                let score = 0;
                for (let i = 0; i < Hs; i += step) score += at(cand + i) * at(follow + i);
                if (score > bestScore) { bestScore = score; best = cand; }
            }
        }
        prev = best;
        const o = k * Hs;
        for (let i = 0; i < N && o + i < norm.length; i++) {
            const w = win[i];
            const src = best + i;
            norm[o + i] += w;
            if (src < 0 || src >= input) continue;
            for (let c = 0; c < channels.length; c++) out[c][o + i] += channels[c][src] * w;
        }
    }
    return out.map((c) => {
        const trimmed = c.subarray(0, outLength);
        for (let i = 0; i < outLength; i++) if (norm[i] > 1e-3) trimmed[i] /= norm[i];
        return Float32Array.from(trimmed);
    });
}

/** An AudioBuffer where there is one (a window), else the same shape in plain arrays (the tests). */
function makeBuffer(channels, sampleRate) {
    const length = channels[0]?.length || 1;
    if (typeof globalThis.AudioBuffer === 'function') {
        const b = new globalThis.AudioBuffer({ length, sampleRate, numberOfChannels: Math.max(1, channels.length) });
        channels.forEach((c, i) => b.getChannelData(i).set(c));
        return b;
    }
    return { sampleRate, length, duration: length / sampleRate, numberOfChannels: channels.length, getChannelData: (i) => channels[i] };
}

const made = new WeakMap();

/**
 * A stretch of `buffer` from `offset` for `seconds`, made `ratio` times as
 * long with its pitch kept: an AudioBuffer to play from its start. The same
 * stretch asked again is the one made before.
 */
export function stretchedBuffer(buffer, offset, seconds, ratio) {
    const rate = buffer.sampleRate;
    const from = Math.max(0, Math.round(offset * rate));
    const length = Math.max(1, Math.min(buffer.length - from, Math.round(seconds * rate)));
    const key = `${from}:${length}:${Math.round(ratio * 1e4)}`;
    let byKey = made.get(buffer);
    if (!byKey) { byKey = new Map(); made.set(buffer, byKey); }
    if (byKey.has(key)) return byKey.get(key);
    const channels = [];
    for (let c = 0; c < buffer.numberOfChannels; c++) channels.push(buffer.getChannelData(c).subarray(from, from + length));
    const out = makeBuffer(stretchChannels(channels, rate, ratio), rate);
    byKey.set(key, out);
    return out;
}

/**
 * How a sound clip is played once its speed, its pitch and whether the pitch
 * is kept are taken into account: the rate the source plays at and, when the
 * waves need stretching first, by how much. `{ rate, ratio }`; ratio 1 is no
 * stretch (the plain speed, the pitch going with it).
 */
export function soundRate({ speed = 1, semitones = 0, keepPitch = true } = {}) {
    const s = Number(speed) > 0 ? Number(speed) : 1;
    const f = pitchFactor(semitones);
    const rate = keepPitch ? f : s * f;
    return { rate, ratio: rate / s };
}
