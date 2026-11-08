// The containers recordings are made in. The recorder itself is the studio's
// (renderer/studio); these three are shared with Presentations and Video.

/** The container a recording for a slide is made in: MP4 first, as PowerPoint plays it everywhere. */
export function recordingType() {
  const R = typeof MediaRecorder !== 'undefined' ? MediaRecorder : null;
  for (const k of ['video/mp4;codecs=avc1', 'video/mp4', 'video/webm;codecs=vp9', 'video/webm']) {
    if (R?.isTypeSupported?.(k)) return k;
  }
  return 'video/webm';
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
