/**
 * Audio-buffer surgery: decode, trim, fade, level, and write a WAV back out.
 *
 * All of it is plain arithmetic over AudioBuffers rather than a graph rendered
 * offline, because every operation here is a sample-for-sample transform with no
 * time dimension of its own - a fade is a multiply, a trim is a subarray. That
 * makes an edit exact, instant, and repeatable: the preview the user hears is
 * produced by the same call that writes the file.
 *
 * WAV is the output on purpose. It is the one format that needs no encoder, and
 * the media file server types a file by its EXTENSION - a `.webm` full of audio
 * comes back as `video/webm` and turns up in the video library.
 */

let _ctx = null;

/** The context every buffer here belongs to. AudioBuffers are bound to the
 *  context that made them, so decode, edit and preview must share one. */
export function editContext() {
    if (typeof window === "undefined") return null;
    const Ctor = window.AudioContext || window.webkitAudioContext;
    if (!Ctor) return null;
    if (!_ctx || _ctx.state === "closed") _ctx = new Ctor();
    return _ctx;
}

/** Decode any container the browser understands into an AudioBuffer. */
export async function decodeBlob(blob) {
    const ctx = editContext();
    if (!ctx) throw new Error("This browser has no Web Audio support.");
    const bytes = await blob.arrayBuffer();
    // Callback form as well as promise form: Safari still only implements the
    // callback signature and returns undefined from the promise one.
    return new Promise((resolve, reject) => {
        const p = ctx.decodeAudioData(bytes, resolve, reject);
        if (p && typeof p.then === "function") p.then(resolve, reject);
    });
}

/** Per-bin peak amplitude, for drawing a waveform strip. */
export function peaksOf(buffer, bins = 480) {
    const ch = buffer.getChannelData(0);
    const per = Math.max(1, Math.floor(ch.length / bins));
    const out = new Float32Array(bins);
    for (let i = 0; i < bins; i++) {
        let m = 0;
        // every 8th sample is plenty for a strip a few hundred pixels wide
        for (let j = i * per, end = Math.min((i + 1) * per, ch.length); j < end; j += 8) {
            const v = Math.abs(ch[j]);
            if (v > m) m = v;
        }
        out[i] = m;
    }
    return out;
}

const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));

/**
 * The whole edit, in one pass: cut to [start, end), optionally fold to mono,
 * apply gain (or normalise to just under full scale), then ramp both ends.
 *
 * Order matters. Normalising AFTER the cut measures the peak of what survives,
 * which is the only peak that means anything; fading LAST means a fade is never
 * scaled back up by the leveller into a click.
 */
export function editBuffer(buffer, opts = {}) {
    const ctx = editContext();
    const rate = buffer.sampleRate;
    const total = buffer.duration;
    const start = clamp(Number(opts.start) || 0, 0, total);
    const end = clamp(Number.isFinite(opts.end) ? opts.end : total, start, total);

    const i0 = Math.floor(start * rate);
    const frames = Math.max(1, Math.floor(end * rate) - i0);
    const srcCh = buffer.numberOfChannels;
    const fold = !!opts.mono && srcCh > 1;
    const outCh = fold ? 1 : srcCh;
    const out = ctx.createBuffer(outCh, frames, rate);

    if (fold) {
        const dst = out.getChannelData(0);
        for (let c = 0; c < srcCh; c++) {
            const s = buffer.getChannelData(c);
            for (let i = 0; i < frames; i++) dst[i] += s[i0 + i] / srcCh;
        }
    } else {
        for (let c = 0; c < outCh; c++) {
            out.getChannelData(c).set(buffer.getChannelData(c).subarray(i0, i0 + frames));
        }
    }

    const gain = Number.isFinite(opts.gain) && opts.gain > 0 ? opts.gain : 1;
    let k = gain;
    if (opts.normalize) {
        let peak = 0;
        for (let c = 0; c < outCh; c++) {
            const d = out.getChannelData(c);
            for (let i = 0; i < frames; i++) { const v = Math.abs(d[i]); if (v > peak) peak = v; }
        }
        // 0.98 rather than 1.0: a sample sitting exactly at full scale clips
        // the moment anything downstream sums it with something else.
        if (peak > 0) k = (0.98 / peak) * gain;
    }
    if (k !== 1) {
        for (let c = 0; c < outCh; c++) {
            const d = out.getChannelData(c);
            for (let i = 0; i < frames; i++) d[i] = clamp(d[i] * k, -1, 1);
        }
    }

    const half = Math.floor(frames / 2);
    const fi = clamp(Math.floor((Number(opts.fadeIn) || 0) * rate), 0, half);
    const fo = clamp(Math.floor((Number(opts.fadeOut) || 0) * rate), 0, half);
    if (fi || fo) {
        for (let c = 0; c < outCh; c++) {
            const d = out.getChannelData(c);
            for (let i = 0; i < fi; i++) d[i] *= i / fi;
            for (let i = 0; i < fo; i++) d[frames - 1 - i] *= i / fo;
        }
    }
    return out;
}

/** True when both channels carry the same signal - a mono source recorded, or
 *  mixed, into a stereo pair. Written once, it halves the file for nothing. */
function channelsAreIdentical(buf) {
    if (buf.numberOfChannels < 2) return true;
    const a = buf.getChannelData(0);
    const b = buf.getChannelData(1);
    const step = Math.max(1, Math.floor(a.length / 4000));
    for (let i = 0; i < a.length; i += step) { if (Math.abs(a[i] - b[i]) > 1e-4) return false; }
    return true;
}

/** An AudioBuffer as a 16-bit PCM WAV blob. */
export function encodeWav(buf) {
    const channels = channelsAreIdentical(buf) ? 1 : Math.min(2, buf.numberOfChannels);
    const rate = buf.sampleRate;
    const frames = buf.length;
    const data = new DataView(new ArrayBuffer(44 + frames * channels * 2));
    const text = (offset, s) => { for (let i = 0; i < s.length; i++) data.setUint8(offset + i, s.charCodeAt(i)); };

    text(0, "RIFF");
    data.setUint32(4, 36 + frames * channels * 2, true);
    text(8, "WAVEfmt ");
    data.setUint32(16, 16, true);                   // PCM header length
    data.setUint16(20, 1, true);                    // format: PCM
    data.setUint16(22, channels, true);
    data.setUint32(24, rate, true);
    data.setUint32(28, rate * channels * 2, true);  // byte rate
    data.setUint16(32, channels * 2, true);         // block align
    data.setUint16(34, 16, true);                   // bits per sample
    text(36, "data");
    data.setUint32(40, frames * channels * 2, true);

    const src = [];
    for (let c = 0; c < channels; c++) src.push(buf.getChannelData(c));
    let at = 44;
    for (let i = 0; i < frames; i++) {
        for (let c = 0; c < channels; c++) {
            const v = clamp(src[c][i], -1, 1);
            data.setInt16(at, v < 0 ? v * 0x8000 : v * 0x7fff, true);
            at += 2;
        }
    }
    return new Blob([data.buffer], { type: "audio/wav" });
}

/** Decode a recording and write it back out as WAV. Throws if the browser
 *  cannot decode its own output - callers keep the original bytes instead. */
export async function blobToWav(blob) {
    return encodeWav(await decodeBlob(blob));
}
