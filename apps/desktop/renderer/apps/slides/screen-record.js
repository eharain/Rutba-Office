// Insert → Screen Recording and Record → Screen Recording: the window or
// screen picked is recorded until Stop — PowerPoint's dock, as a bar along
// the top of this window with the time and Stop and Cancel — and the
// recording comes back as a video for the slide: MP4 where this machine
// can make one, WebM where it cannot.

import React, { useEffect, useRef, useState } from 'react';
import { Button } from '@rutba/office-ui';

/** The container a recording is made in: MP4 first, as PowerPoint plays it everywhere. */
export function recordingType() {
  const R = typeof MediaRecorder !== 'undefined' ? MediaRecorder : null;
  for (const t of ['video/mp4;codecs=avc1', 'video/mp4', 'video/webm;codecs=vp9', 'video/webm']) {
    if (R?.isTypeSupported?.(t)) return t;
  }
  return 'video/webm';
}

/** The most of a recording one message carries to its scratch file. */
const SLICE = 4 * 1024 * 1024;

const clock = (ms) => `${Math.floor(ms / 60000)}:${String(Math.floor((ms / 1000) % 60)).padStart(2, '0')}`;

/**
 * Records `source` (from capture.sources) on mount, into a scratch file half
 * a second at a time, so the window holds none of it and no one message
 * carries the whole. `onDone({ path, contentType, name })` when Stop is
 * pressed, the recording in the file at `path`; `onCancel()` when Cancel
 * is, the file gone, or when the source cannot be recorded (`onError(message)`
 * first).
 */
export function ScreenRecorder({ source, shell, onDone, onCancel, onError }) {
  const [started, setStarted] = useState(null);
  const [now, setNow] = useState(Date.now());
  const rec = useRef(null);
  const cancelled = useRef(false);

  useEffect(() => {
    let stream = null;
    let live = true;
    (async () => {
      try {
        stream = await navigator.mediaDevices.getUserMedia({
          audio: false,
          video: { mandatory: { chromeMediaSource: 'desktop', chromeMediaSourceId: source.id, maxFrameRate: 30 } },
        });
        if (!live) { stream.getTracks().forEach((t) => t.stop()); return; }
        const type = recordingType();
        const contentType = type.split(';')[0];
        const ext = contentType === 'video/mp4' ? 'mp4' : 'webm';
        const { path } = await shell.fs.temp({ ext });
        if (!live) { stream.getTracks().forEach((t) => t.stop()); shell.fs.dropTemp({ path }).catch(() => {}); return; }
        const r = new MediaRecorder(stream, { mimeType: type });
        // Each piece written after the one before, in the order they came, and
        // four megabytes at a time: a WebM recorder hands one over every half
        // second, an MP4 one may keep the whole until Stop.
        let written = Promise.resolve();
        let failed = null;
        const put = async (part) => {
          for (let at = 0; at < part.size && !failed; at += SLICE) await shell.fs.append({ path, bytes: new Uint8Array(await part.slice(at, at + SLICE).arrayBuffer()) });
        };
        r.ondataavailable = (e) => {
          const part = e.data;
          if (!part?.size) return;
          written = written.then(() => (failed ? null : put(part))).catch((err) => { failed = err; });
        };
        r.onstop = async () => {
          stream.getTracks().forEach((t) => t.stop());
          await written;
          if (cancelled.current || failed) {
            shell.fs.dropTemp({ path }).catch(() => {});
            if (failed && !cancelled.current) { onError?.(`The recording could not be kept: ${failed.message || failed}`); onCancel(); }
            return;
          }
          onDone({ path, contentType, name: `Screen Recording ${new Date().toLocaleTimeString()}.${ext}` });
        };
        r.start(500);
        rec.current = r;
        setStarted(Date.now());
      } catch (err) {
        stream?.getTracks().forEach((t) => t.stop());
        onError?.(`${source.name} could not be recorded: ${err.message || err}`);
        onCancel();
      }
    })();
    return () => {
      live = false;
      if (rec.current?.state === 'recording') { cancelled.current = true; rec.current.stop(); }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [source.id]);

  useEffect(() => {
    if (!started) return undefined;
    const t = setInterval(() => setNow(Date.now()), 250);
    return () => clearInterval(t);
  }, [started]);

  return (
    <div className="sl-screenrec" role="status" aria-label="Screen Recording">
      <span className="sl-screenrec-dot" />
      <span className="sl-screenrec-what">Recording {source.name}</span>
      <b className="sl-screenrec-time">{started ? clock(now - started) : 'starting…'}</b>
      <Button primary label="Stop" className="sl-screenrec-stop" disabled={!started} onClick={() => rec.current?.stop()} />
      <Button label="Cancel" onClick={() => { cancelled.current = true; rec.current?.stop(); onCancel(); }} />
    </div>
  );
}

export const SCREENREC_CSS = `
.sl-screenrec { position: fixed; top: 10px; left: 50%; transform: translateX(-50%); z-index: 200; display: flex; align-items: center; gap: 10px; padding: 8px 10px 8px 14px; background: var(--surface); border: 1px solid var(--line); border-radius: 12px; box-shadow: var(--shadow-3); font-size: 12.5px; }
.sl-screenrec-dot { width: 10px; height: 10px; border-radius: 50%; background: #d93025; animation: sl-rec-blink 1.2s infinite; }
@keyframes sl-rec-blink { 50% { opacity: 0.35; } }
.sl-screenrec-what { max-width: 280px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--ink-2); }
.sl-screenrec-time { font-variant-numeric: tabular-nums; min-width: 44px; }
`;
