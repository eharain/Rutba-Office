/**
 * Finding the beat in a music bed, and using it to place photo cuts.
 *
 * Pure functions on purpose: they take samples and numbers and return numbers,
 * so pages/dev/timeline-fixture.js can drive them against a synthesised click
 * track without an audio device, a network, or a session.
 *
 * The method is spectral-flux-lite: short-time energy per window, then the
 * POSITIVE change between windows, then peaks in that. It finds percussive
 * onsets well enough to cut a slideshow on, which is the job — it is not a
 * tempo tracker and does not pretend to be one.
 */

/**
 * Onset times in seconds, from one channel of PCM.
 *
 * @param {Float32Array} channel  mono samples
 * @param {number} sampleRate
 * @param {object} [opts] win/hop in samples, minGap in seconds, sensitivity in
 *                        standard deviations above the mean flux
 * @returns {number[]} ascending times
 */
export function onsetTimes(channel, sampleRate, opts = {}) {
    const win = opts.win || 1024;
    const hop = opts.hop || 512;
    const minGap = opts.minGap ?? 0.3;
    const sensitivity = opts.sensitivity ?? 1.2;
    if (!channel || !channel.length || !sampleRate) return [];

    // Short-time energy, in dB-ish log space so a loud passage does not drown
    // the onsets in a quiet one.
    const frames = Math.max(0, Math.floor((channel.length - win) / hop) + 1);
    if (frames < 3) return [];
    const energy = new Float32Array(frames);
    for (let f = 0; f < frames; f++) {
        const s = f * hop;
        let sum = 0;
        for (let i = 0; i < win; i++) { const v = channel[s + i]; sum += v * v; }
        energy[f] = Math.log10(1e-8 + sum / win);
    }

    // Positive change only: a note starting is an onset, a note ending is not.
    const flux = new Float32Array(frames);
    for (let f = 1; f < frames; f++) flux[f] = Math.max(0, energy[f] - energy[f - 1]);

    let mean = 0;
    for (let f = 0; f < frames; f++) mean += flux[f];
    mean /= frames;
    let varSum = 0;
    for (let f = 0; f < frames; f++) { const d = flux[f] - mean; varSum += d * d; }
    const sd = Math.sqrt(varSum / frames);
    const threshold = mean + sensitivity * sd;

    const out = [];
    const minFrames = Math.max(1, Math.round((minGap * sampleRate) / hop));
    let last = -Infinity;
    for (let f = 1; f < frames - 1; f++) {
        if (flux[f] < threshold) continue;
        if (flux[f] < flux[f - 1] || flux[f] < flux[f + 1]) continue; // local peak only
        if (f - last < minFrames) continue;
        last = f;
        out.push(+((f * hop) / sampleRate).toFixed(3));
    }
    return out;
}

/**
 * Move each interior slot edge to the nearest onset, within reason.
 *
 * `edges` is ascending and includes both ends ([0, …, duration]); only the
 * interior ones move. An edge with no onset near it STAYS WHERE IT IS rather
 * than being dragged to a distant beat — a cut that lands on nothing is worse
 * than a cut that lands where the author put it. Order and a minimum slot
 * length are enforced left to right, so a run of close beats cannot collapse a
 * photo to nothing.
 *
 * @returns {{edges: number[], moved: number}} the new edges and how many moved
 */
export function snapEdges(edges, onsets, opts = {}) {
    const tolerance = opts.tolerance ?? 0.8;
    const minSlot = opts.minSlot ?? 1.2;
    if (!Array.isArray(edges) || edges.length < 3 || !onsets || !onsets.length) {
        return { edges: [...(edges || [])], moved: 0 };
    }
    const out = [...edges];
    const last = out.length - 1;
    let moved = 0;
    for (let i = 1; i < last; i++) {
        const want = edges[i];
        let best = null;
        let bestD = Infinity;
        for (const o of onsets) {
            const d = Math.abs(o - want);
            if (d < bestD) { bestD = d; best = o; }
        }
        if (best === null || bestD > tolerance) continue;
        // Never before the previous edge + a whole slot, never so late that the
        // remaining photos cannot each have one.
        const floor = out[i - 1] + minSlot;
        const ceil = out[last] - minSlot * (last - i);
        const v = Math.min(Math.max(best, floor), ceil);
        if (v <= out[i - 1] || v >= out[last]) continue;
        if (Math.abs(v - out[i]) > 0.01) moved++;
        out[i] = +v.toFixed(3);
    }
    return { edges: out, moved };
}

/** Cumulative edges [0, s1, s1+s2, …, total] from a list of slot lengths. */
export function edgesFromLengths(lengths) {
    const out = [0];
    for (const s of lengths) out.push(+(out[out.length - 1] + s).toFixed(3));
    return out;
}

/** The inverse: slot lengths from edges. */
export function lengthsFromEdges(edges) {
    const out = [];
    for (let i = 1; i < edges.length; i++) out.push(+(edges[i] - edges[i - 1]).toFixed(3));
    return out;
}
