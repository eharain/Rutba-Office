// The containers recordings are made in. The recorder itself is the studio's
// (renderer/studio); these are shared with Presentations and Video.

const recorderClass = () => (typeof MediaRecorder !== 'undefined' ? MediaRecorder : null);

/**
 * The types a recording for a slide is asked in, best first: MP4, as
 * PowerPoint plays it everywhere, then WebM where MP4 is not recorded.
 *
 * H.264 as avc3 first. Its key frames carry the picture's parameter sets, so
 * when the encoder starts again part way through, with new ones (a new size,
 * or perhaps one encoder taking over from another on a busy machine; not
 * reproduced), every frame after still decodes. avc1 keeps only the first set, in the file's header:
 * the frames after such a change will not decode, and Chromium says so in
 * the console ("the codec description is not supposed to change", gate entry
 * L.0002, 14 Nov 2026). avc3 opens and plays in Windows' playback engine
 * (Media Foundation), in Chromium, in Firefox and in Office's own readers; avc1 stays as the
 * fall-back for a Chromium without it.
 */
export const RECORDING_TYPES = ['video/mp4;codecs=avc3', 'video/mp4;codecs=avc1', 'video/mp4', 'video/webm;codecs=vp9', 'video/webm'];

/** The types of `list` that `R` (MediaRecorder) says it can make, in order. */
export function supportedTypes(list, R = recorderClass()) {
  return list.filter((k) => {
    try { return Boolean(R?.isTypeSupported?.(k)); } catch { return false; }
  });
}

/** The container a recording for a slide is made in: the first of RECORDING_TYPES this machine makes. */
export function recordingType(R = recorderClass()) {
  return supportedTypes(RECORDING_TYPES, R)[0] || 'video/webm';
}

/**
 * `type`, then the types after it in RECORDING_TYPES of the same container:
 * what a recorder may fall back to once a file's kind is chosen (avc3, then
 * avc1, then MP4 as the recorder likes it; never WebM in a file named .mp4).
 */
export function typesFrom(type, list = RECORDING_TYPES) {
  const base = String(type || '').split(';')[0];
  const at = list.indexOf(type);
  return [...new Set([type, ...list.slice(at < 0 ? list.length : at + 1).filter((k) => k.split(';')[0] === base)])];
}

/**
 * A MediaRecorder for `stream` in `type`, or, where its constructor will not
 * take `type` though the type was said to be supported, in the next of the
 * same container it makes: { rec, type }, the type it was made in. Throws
 * what the last refused with.
 */
export function openRecorder(stream, type, options = {}, R = recorderClass()) {
  let last = null;
  for (const [i, k] of typesFrom(type).entries()) {
    if (i > 0 && !supportedTypes([k], R).length) continue;
    try {
      return { rec: new R(stream, { ...options, mimeType: k }), type: k };
    } catch (err) {
      last = err;
    }
  }
  throw last || new Error('no recorder for this type');
}

/**
 * The Recorder's program recorder, as openRecorder makes it: { rec, type }.
 * A type in MP4 (a slide's recording) that no MP4 type of this machine's
 * recorder takes is refused, the refusal passed on: a recorder left to
 * choose for itself writes WebM, never right in a file named .mp4. A WebM
 * type its constructor will not take with these options gets the recorder
 * as it comes, which is WebM too.
 */
export function openProgramRecorder(stream, type, options = {}, R = recorderClass()) {
  try {
    return openRecorder(stream, type, options, R);
  } catch (err) {
    if (String(type || '').split(';')[0] === 'video/mp4') throw err;
    return { rec: new R(stream), type };
  }
}

/**
 * The container a take for Video is made in: WebM, which a recorder writes
 * as a stream that plays to its last piece, so a take a crash cut short
 * still opens.
 */
export function takeType() {
  const R = typeof MediaRecorder !== 'undefined' ? MediaRecorder : null;
  for (const k of ['video/webm;codecs=vp9,opus', 'video/webm;codecs=vp8,opus', 'video/webm']) {
    if (R?.isTypeSupported?.(k)) return k;
  }
  return 'video/webm';
}

/**
 * The container a voice alone is made in, and its extension: WebM with Opus,
 * written as a stream as the picture takes are, so a take a crash cut short
 * still plays.
 */
export function voiceType() {
  const R = typeof MediaRecorder !== 'undefined' ? MediaRecorder : null;
  for (const k of ['audio/webm;codecs=opus', 'audio/webm', 'audio/ogg;codecs=opus']) {
    if (R?.isTypeSupported?.(k)) return { type: k, ext: k.startsWith('audio/ogg') ? 'ogg' : 'weba' };
  }
  return { type: 'audio/webm', ext: 'weba' };
}
