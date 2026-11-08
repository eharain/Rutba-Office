// Each source as a file of its own (A7, "isolated recording"): one recorder a
// source, started on the program's clock (program.js), so that Video can land
// the screen, the camera and the sound on lanes of their own, as it always has.
// What is recorded is what the frame shows when Record is pressed.

import { takesPlan } from '@rutba/studio/stack';
import { videoBitrate, takeFileName, AUDIO_BITRATE } from '@rutba/studio/recorder';
import { t } from '@rutba/office-ui';
import { takeType, voiceType } from '../recording/session.js';

const STEM = { screen: () => t('Screen'), camera: () => t('Camera'), voice: () => t('Voice') };

const partOf = (crop) => (crop && (crop.w < 0.999 || crop.h < 0.999) ? crop : null);

/**
 * The takes for `stack`, not started: [{ role, path, rec, contentType, ext, hasAudio, width, height, crop, inset, shape, mirror, ref }].
 * @param {object} p
 * @param {object} p.stack
 * @param {object} p.media  StudioMedia
 * @param {object} p.shell
 * @param {{ fps: number, quality: string, folder: string }} p.settings
 */
export async function makeTakes({ stack, media, shell, settings }) {
  const made = [];
  const file = async (role, ext) => (await shell.capture.take({ dir: settings.folder, name: takeFileName(STEM[role](), new Date(), ext) })).path;
  for (const p of takesPlan(stack)) {
    if (p.role === 'voice') {
      const { type, ext } = voiceType();
      const stream = new MediaStream([media.mix()]);
      let rec;
      try { rec = new MediaRecorder(stream, { mimeType: type, audioBitsPerSecond: AUDIO_BITRATE }); } catch { rec = new MediaRecorder(stream); }
      made.push({ role: 'voice', path: await file('voice', ext), rec, contentType: type.split(';')[0], ext, hasAudio: true, width: null, height: null });
      continue;
    }
    const item = stack.items.find((it) => it.id === p.itemId);
    const video = media.streamOf(item)?.getVideoTracks()[0];
    if (!video) continue;
    const type = takeType();
    const contentType = type.split(';')[0];
    const ext = contentType === 'video/mp4' ? 'mp4' : 'webm';
    const s = video.getSettings?.() || {};
    const bits = videoBitrate({ width: s.width || 1280, height: s.height || 720, fps: settings.fps, quality: settings.quality });
    // A file with no sound of its own still has a silent track: a recorder with a
    // picture alone and a window that seldom changes stopped its clock at the last change.
    const sound = p.sound ? media.mix() : media.silence();
    const stream = new MediaStream([video, ...(sound ? [sound] : [])]);
    let rec;
    try { rec = new MediaRecorder(stream, { mimeType: type, videoBitsPerSecond: bits, audioBitsPerSecond: AUDIO_BITRATE }); } catch { rec = new MediaRecorder(stream); }
    made.push({
      role: p.role, path: await file(p.role, ext), rec, contentType, ext,
      hasAudio: Boolean(p.sound && media.heard()), width: s.width || null, height: s.height || null, ref: item.ref,
      ...(p.role === 'screen' && partOf(item.crop) ? { crop: item.crop } : {}),
      ...(p.role === 'camera' ? { inset: { fx: item.box.x, fy: item.box.y, fw: item.box.w, fh: item.box.h }, shape: item.shape, mirror: item.mirror } : {}),
    });
  }
  return made;
}
