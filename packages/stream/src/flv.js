// The FLV packets RTMP carries: a video message is an FLV video tag's body,
// an audio message an audio tag's body, the metadata a script tag's.
//
// H.264 goes as AVC: its configuration (avcC, what WebCodecs gives as the
// decoder's description) once as a sequence header, then each frame as
// length-prefixed NAL units (WebCodecs' 'avc' format), key frames marked.
// AAC goes as its AudioSpecificConfig once, then raw frames.

import { encode, EcmaArray } from './amf0.js';

/** RTMP message types. */
export const AUDIO = 8;
export const VIDEO = 9;
export const DATA = 18;

const CODEC_AVC = 7;
const CODEC_AAC = 10;

/** The body of a video message: the AVC configuration, or a frame. */
export function videoBody({ data, key = false, config = false, cts = 0 }) {
  const head = Buffer.alloc(5);
  head[0] = ((key || config ? 1 : 2) << 4) | CODEC_AVC;
  head[1] = config ? 0 : 1;
  head.writeIntBE(Math.round(cts), 2, 3);
  return Buffer.concat([head, Buffer.from(data.buffer ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength) : data)]);
}

/** The body of an audio message: AAC's configuration, or a frame. */
export function audioBody({ data, config = false }) {
  // 0xAF: AAC, 44 kHz, 16-bit, stereo, which is what every AAC packet says; the configuration says the rest.
  return Buffer.concat([Buffer.from([0xaf, config ? 0 : 1]), Buffer.from(data.buffer ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength) : data)]);
}

/** The stream's metadata, as @setDataFrame onMetaData. */
export function metadataBody({ width, height, fps, videoKbps, audioKbps, sampleRate = 48000, channels = 2, encoder = 'Rutba Recorder' }) {
  return encode('@setDataFrame', 'onMetaData', new EcmaArray({
    duration: 0,
    width,
    height,
    framerate: fps,
    videodatarate: videoKbps,
    videocodecid: CODEC_AVC,
    audiodatarate: audioKbps,
    audiosamplerate: sampleRate,
    audiosamplesize: 16,
    stereo: channels > 1,
    audiocodecid: CODEC_AAC,
    encoder,
  }));
}

/** Whether a video body is a key frame (or the configuration). */
export const isKeyBody = (body) => (body[0] >> 4) === 1;
