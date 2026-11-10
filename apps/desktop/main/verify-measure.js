// What an exported file really holds, measured in the window that made it:
// a file with an audio track can be silent, and a picture can be the wrong one.

import zlib from 'node:zlib';

/** Sound in a file: { duration, peak, at: [{ t, rms, hz }] }, each `t` measured over a tenth of a second either side. */
export function soundIn(js, file, times = []) {
  return js(`(async () => {
    const r = await window.rutbaOffice.fs.read({ path: ${JSON.stringify(file)} });
    const copy = r.bytes.buffer.slice(r.bytes.byteOffset, r.bytes.byteOffset + r.bytes.byteLength);
    const ctx = new OfflineAudioContext(1, 48000, 48000);
    const audio = await ctx.decodeAudioData(copy);
    const data = audio.getChannelData(0);
    let peak = 0;
    for (let i = 0; i < data.length; i += 7) peak = Math.max(peak, Math.abs(data[i]));
    const rate = audio.sampleRate;
    const at = ${JSON.stringify(times)}.map((t) => {
      const from = Math.max(0, Math.round((t - 0.1) * rate));
      const to = Math.min(data.length, Math.round((t + 0.1) * rate));
      let sum = 0; let crossings = 0;
      for (let i = from; i < to; i++) { sum += data[i] * data[i]; if (i > from && (data[i - 1] < 0) !== (data[i] < 0)) crossings++; }
      const seconds = Math.max(1, to - from) / rate;
      return { t, rms: Math.round(Math.sqrt(sum / Math.max(1, to - from)) * 1000) / 1000, hz: Math.round(crossings / 2 / seconds) };
    });
    return { duration: audio.duration, peak: Math.round(peak * 1000) / 1000, at };
  })()`).catch((err) => ({ error: err.message }));
}

/** The average colour of a video's frame at each of `times`: [{ t, rgb }]. */
export function framesIn(js, file, times) {
  return js(`(async () => {
    const r = await window.rutbaOffice.fs.read({ path: ${JSON.stringify(file)} });
    const v = document.createElement('video');
    v.muted = true;
    v.src = URL.createObjectURL(new Blob([r.bytes]));
    await new Promise((res) => { v.onloadeddata = res; v.onerror = res; setTimeout(res, 6000); });
    const c = document.createElement('canvas'); c.width = 64; c.height = 36;
    const g = c.getContext('2d', { willReadFrequently: true });
    const out = [];
    for (const t of ${JSON.stringify(times)}) {
      v.currentTime = t;
      await new Promise((res) => { v.onseeked = res; setTimeout(res, 5000); });
      if (v.requestVideoFrameCallback) await new Promise((res) => { v.requestVideoFrameCallback(() => res()); setTimeout(res, 800); });
      g.drawImage(v, 0, 0, 64, 36);
      const px = g.getImageData(0, 0, 64, 36).data;
      const rgb = [0, 0, 0];
      for (let i = 0; i < px.length; i += 4) { rgb[0] += px[i]; rgb[1] += px[i + 1]; rgb[2] += px[i + 2]; }
      out.push({ t, rgb: rgb.map((x) => Math.round(x / (px.length / 4))) });
    }
    return out;
  })()`).catch((err) => [{ error: err.message }]);
}

/** Whether two average colours are the same picture, allowing for what encoding moves them. */
export const sameColour = (a, b, within = 30) => Boolean(a && b) && a.every((v, i) => Math.abs(v - b[i]) <= within);

/**
 * Whether words are anywhere in a PDF, by a reading that shares nothing with
 * the code that made it: the file's bytes, and every stream inflated.
 */
export function anyStreamHolds(bytes, words) {
  const buf = Buffer.from(bytes);
  if (buf.includes(words, 0, 'latin1')) return true;
  const start = Buffer.from('stream');
  const stop = Buffer.from('endstream');
  for (let at = buf.indexOf(start); at >= 0; at = buf.indexOf(start, at + start.length)) {
    if (buf.subarray(Math.max(0, at - 3), at).toString('latin1') === 'end') continue;
    let from = at + start.length;
    if (buf[from] === 13) from++;
    if (buf[from] === 10) from++;
    const end = buf.indexOf(stop, from);
    if (end < 0) break;
    let inflated = null;
    try { inflated = zlib.inflateSync(buf.subarray(from, end), { finishFlush: zlib.constants.Z_SYNC_FLUSH }); } catch { /* not deflated */ }
    if (inflated?.includes(words, 0, 'latin1')) return true;
    at = end;
  }
  return false;
}
