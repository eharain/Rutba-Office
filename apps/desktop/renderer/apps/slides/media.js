// Insert → Video and Audio on a slide: a video's poster frame taken from
// the video itself, a speaker for a sound; a play bar under each on the
// editing stage, as PowerPoint draws one; and in the show a click on the
// video or the speaker plays it, with the media's own controls when Show
// Media Controls is on.

import React, { useEffect, useRef, useState } from 'react';
import { Icon } from '@rutba/office-ui';
import { isNarration } from '@rutba/presentation/narration';

/** The kinds of file each button opens, and what each is. */
export const MEDIA_FILES = {
  video: { mp4: 'video/mp4', m4v: 'video/x-m4v', mov: 'video/quicktime', webm: 'video/webm', wmv: 'video/x-ms-wmv', avi: 'video/x-msvideo' },
  audio: { mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', wma: 'audio/x-ms-wma', aac: 'audio/aac' },
};

const pngOf = async (canvas) => new Uint8Array(await (await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'))).arrayBuffer());

/** A poster for a video this window cannot decode: dark, with a play mark, in the video's usual shape. */
export async function placeholderPoster(width = 1280, height = 720) {
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const g = canvas.getContext('2d');
  g.fillStyle = '#1f2329';
  g.fillRect(0, 0, width, height);
  const r = Math.min(width, height) * 0.14;
  g.fillStyle = 'rgba(255,255,255,0.9)';
  g.beginPath();
  g.moveTo(width / 2 - r * 0.6, height / 2 - r);
  g.lineTo(width / 2 + r, height / 2);
  g.lineTo(width / 2 - r * 0.6, height / 2 + r);
  g.closePath();
  g.fill();
  return { png: await pngOf(canvas), width, height, decoded: false };
}

/** A local file as the window draws it, over rutba://file. */
export const fileSrc = (p) => `rutba://file/${btoa(unescape(encodeURIComponent(p))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')}`;

/**
 * A video's poster frame: the picture a tenth of the way in (a second at
 * most), at the video's own size, as PNG bytes — or a plain poster when
 * this window cannot decode the video (a .wmv, say), which PowerPoint can.
 * `source` is the video's bytes, or a URL it is drawn from (a file's).
 */
export async function posterFrame(source, type) {
  const owned = typeof source !== 'string';
  const url = owned ? URL.createObjectURL(new Blob([source], { type })) : source;
  const video = document.createElement('video');
  // A file's picture is read off a canvas: asked for so, it is the window's to read.
  if (!owned) video.crossOrigin = 'anonymous';
  // The waits' timers are cleared once the frame is taken, so neither holds the video for seconds after.
  const timers = [];
  try {
    video.muted = true;
    video.preload = 'auto';
    video.src = url;
    await new Promise((resolve, reject) => {
      video.onloadeddata = resolve;
      video.onerror = () => reject(new Error('undecodable'));
      timers.push(setTimeout(() => reject(new Error('timeout')), 15000));
    });
    const at = Math.min(1, (Number.isFinite(video.duration) ? video.duration : 0) / 10);
    if (at > 0) await new Promise((resolve) => { video.onseeked = resolve; video.currentTime = at; timers.push(setTimeout(resolve, 3000)); });
    const width = video.videoWidth || 1280;
    const height = video.videoHeight || 720;
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    canvas.getContext('2d').drawImage(video, 0, 0, width, height);
    return { png: await pngOf(canvas), width, height, decoded: true };
  } catch {
    return placeholderPoster();
  } finally {
    timers.forEach(clearTimeout);
    video.onloadeddata = video.onerror = video.onseeked = null;
    video.removeAttribute('src');
    video.load();
    if (owned) URL.revokeObjectURL(url);
  }
}

const clock = (s) => (Number.isFinite(s) ? `${Math.floor(s / 60)}:${String(Math.floor(s % 60)).padStart(2, '0')}` : '0:00');

/**
 * The editing stage's play bar under each video and sound, as PowerPoint
 * draws one: play or pause, where it is, how long it is. A video plays
 * over its poster; a sound plays unseen.
 */
export function StageMedia({ shapes }) {
  const media = (shapes || []).filter((s) => s.media?.url && s.geometry && !s.hidden);
  return media.map((s) => <StageClip key={s.id} shape={s} />);
}

function StageClip({ shape }) {
  const el = useRef(null);
  const [playing, setPlaying] = useState(false);
  const [time, setTime] = useState({ at: 0, length: NaN });
  const g = shape.geometry;
  const video = shape.media.kind === 'video';
  const toggle = (e) => {
    e.stopPropagation();
    const m = el.current;
    if (!m) return;
    if (m.paused) m.play().catch(() => {}); else m.pause();
  };
  const events = {
    ref: el,
    src: shape.media.url,
    preload: 'metadata',
    onPlay: () => setPlaying(true),
    onPause: () => setPlaying(false),
    onEnded: () => setPlaying(false),
    onTimeUpdate: (e) => setTime({ at: e.currentTarget.currentTime, length: e.currentTarget.duration }),
    onLoadedMetadata: (e) => setTime({ at: 0, length: e.currentTarget.duration }),
  };
  return (
    <>
      {video
        ? <video {...events} className="sl-media-video" data-shape={shape.id} style={{ left: g.x, top: g.y, width: g.w, height: g.h, opacity: playing || time.at > 0 ? 1 : 0 }} />
        : <audio {...events} data-shape={shape.id} />}
      <div className="sl-media-bar" data-shape={shape.id} style={{ left: g.x, top: g.y + g.h + 4, width: Math.max(140, g.w) }} onMouseDown={(e) => e.stopPropagation()}>
        <button type="button" className="sl-media-play" aria-label={playing ? 'Pause' : 'Play'} data-tip={playing ? 'Pause' : 'Play'} onClick={toggle}>
          <Icon name={playing ? 'pause' : 'play'} size={13} />
        </button>
        <span className="sl-media-track"><span style={{ width: `${time.length ? Math.min(100, (100 * time.at) / time.length) : 0}%` }} /></span>
        <span className="sl-media-time">{clock(time.at)}{Number.isFinite(time.length) ? ` / ${clock(time.length)}` : ''}</span>
      </div>
    </>
  );
}

/**
 * The show's media: over each video and speaker, the media itself, played
 * by a click on it — which does not move the show on — and shown only once
 * it has started, its own controls on it when `controls` is on.
 */
export function ShowMedia({ shapes, size, controls, narration = true }) {
  const media = (shapes || []).filter((s) => s.media?.url && s.geometry && !s.hidden);
  if (!media.length || !size?.width) return null;
  return (
    <div className="sl-show-media">
      {media.map((s) => (isNarration(s) ? (narration ? <Narration key={s.id} shape={s} /> : null) : <ShowClip key={s.id} shape={s} size={size} controls={controls} />))}
    </div>
  );
}

/** A slide's narration: heard as the slide comes in, while Play Narrations is on; never seen. */
function Narration({ shape }) {
  const el = useRef(null);
  useEffect(() => {
    el.current?.play().catch(() => {});
    return () => el.current?.pause?.();
  }, []);
  return <audio ref={el} src={shape.media.url} preload="auto" data-narration={shape.id} />;
}

function ShowClip({ shape, size, controls }) {
  const el = useRef(null);
  const [started, setStarted] = useState(false);
  const g = shape.geometry;
  const box = { left: `${(100 * g.x) / size.width}%`, top: `${(100 * g.y) / size.height}%`, width: `${(100 * g.w) / size.width}%`, height: `${(100 * g.h) / size.height}%` };
  useEffect(() => () => el.current?.pause?.(), []);
  const click = (e) => {
    e.stopPropagation();
    const m = el.current;
    if (!m) return;
    if (m.paused) { setStarted(true); m.play().catch(() => {}); } else m.pause();
  };
  return shape.media.kind === 'video' ? (
    <div className="sl-show-clip" data-shape={shape.id} style={box} onClick={click}>
      <video ref={el} src={shape.media.url} preload="auto" controls={controls && started} style={{ opacity: started ? 1 : 0 }} onClick={(e) => { if (controls && started) e.stopPropagation(); }} />
    </div>
  ) : (
    <div className="sl-show-clip sl-show-sound" data-shape={shape.id} style={box} onClick={click}>
      <audio ref={el} src={shape.media.url} preload="auto" />
    </div>
  );
}

export const MEDIA_CSS = `
.sl-media-video { position: absolute; object-fit: fill; background: #000; pointer-events: none; z-index: 2; }
.sl-media-bar { position: absolute; z-index: 6; display: flex; align-items: center; gap: 8px; height: 26px; padding: 0 8px; box-sizing: border-box; background: rgba(20, 22, 26, 0.86); color: #fff; border-radius: 6px; font: 11px/1 var(--font-ui, sans-serif); transform-origin: top left; }
.sl-media-play { display: inline-grid; place-items: center; width: 20px; height: 20px; border: 0; border-radius: 4px; background: transparent; color: #fff; cursor: pointer; padding: 0; }
.sl-media-play:hover { background: rgba(255,255,255,0.16); }
.sl-media-track { flex: 1; height: 3px; border-radius: 2px; background: rgba(255,255,255,0.28); overflow: hidden; }
.sl-media-track > span { display: block; height: 100%; background: #fff; }
.sl-media-time { font-variant-numeric: tabular-nums; opacity: 0.85; white-space: nowrap; }
.sl-show-media { position: absolute; inset: 0; pointer-events: none; z-index: 3; }
.sl-show-clip { position: absolute; pointer-events: auto; cursor: pointer; }
.sl-show-clip video { width: 100%; height: 100%; object-fit: fill; background: #000; display: block; }
`;
