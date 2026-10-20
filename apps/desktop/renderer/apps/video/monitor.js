// Video's preview on another screen (the 2.7 plan: "preview or inspector on
// a second screen"): a window of its own showing only the picture, full
// screen on another display when there is one. It is not sent pictures: the
// editor sends it the project when it changes and where the playhead is, and
// it draws the frame itself from its own decoders, on its own clock while the
// preview plays. The editor's window keeps the sound.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { t } from '@rutba/office-ui';
import { buildPlan, paintFrame, driveVideoElements, pauseVideoElements } from '@rutba/studio/host';
import { normaliseProject, planPatches, projectMedia, emptyProject } from '@rutba/studio/project';
import { MediaLibrary } from './media.js';
import { AppFrame } from '../../shell.js';

/** The channel an editor and its monitor talk on. */
export const monitorChannel = (key) => `rutba-video-monitor:${key}`;

/**
 * The editor's side: a channel that sends the project and the playhead to a
 * monitor window and answers one that has just opened. `state()` answers
 * { doc, time, playing: { from, wall } | null } as it is now; `onGone` is
 * called when the monitor's window is closed.
 */
export function openMonitorLink(key, state, onGone = () => {}) {
  const channel = new BroadcastChannel(monitorChannel(key));
  let open = true;
  const post = (msg) => { if (open) { try { channel.postMessage(msg); } catch { /* the window has gone */ } } };
  channel.onmessage = (e) => {
    if (e.data?.type === 'bye') { onGone(); return; }
    if (e.data?.type === 'hello') {
      const s = state();
      post({ type: 'doc', doc: s.doc });
      post(s.playing ? { type: 'play', ...s.playing } : { type: 'time', time: s.time });
    }
  };
  return {
    doc: (doc) => post({ type: 'doc', doc }),
    time: (time) => post({ type: 'time', time }),
    play: (from, wall) => post({ type: 'play', from, wall }),
    stop: (time) => post({ type: 'stop', time }),
    close: () => { post({ type: 'close' }); open = false; channel.close(); },
  };
}

export default function MonitorView({ app, shell, monitorKey }) {
  const media = useMemo(() => new MediaLibrary(), []);
  const [doc, setDoc] = useState(() => emptyProject());
  const [time, setTime] = useState(0);
  const [version, setVersion] = useState(0);
  const [waiting, setWaiting] = useState(true);
  const [full, setFull] = useState(false);
  const clock = useRef(null);
  const canvasRef = useRef(null);
  const planCanvas = useMemo(() => document.createElement('canvas'), []);

  useEffect(() => media.onChange(() => setVersion((v) => v + 1)), [media]);
  useEffect(() => () => media.dispose(), [media]);

  const plan = useMemo(() => buildPlan({
    canvas: planCanvas, images: [], options: doc.options, layerPatches: planPatches(doc),
    videos: Object.fromEntries(media.videos), assets: Object.fromEntries(media.images),
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [doc, version]);
  const planRef = useRef(plan);
  planRef.current = plan;

  // What the editor says: the project, where the playhead is, and when it plays and stops.
  useEffect(() => {
    const channel = new BroadcastChannel(monitorChannel(monitorKey));
    const halt = () => {
      if (!clock.current) return;
      clearInterval(clock.current);
      clock.current = null;
      pauseVideoElements(planRef.current);
    };
    channel.onmessage = (e) => {
      const m = e.data || {};
      if (m.type === 'doc') {
        const next = normaliseProject(m.doc);
        setDoc(next);
        setWaiting(false);
        media.loadProject(projectMedia(next)).catch(() => {});
      } else if (m.type === 'time' || m.type === 'stop') {
        halt();
        setTime(Number(m.time) || 0);
      } else if (m.type === 'play') {
        halt();
        // The editor's clock, read on the wall: both windows agree on Date.now().
        const tick = () => {
          const now = Number(m.from) + (Date.now() - Number(m.wall)) / 1000;
          driveVideoElements(planRef.current, now);
          setTime(now);
        };
        tick();
        clock.current = setInterval(tick, 1000 / 60);
      } else if (m.type === 'close') {
        halt();
        shell.win.close().catch(() => window.close());
      }
    };
    channel.postMessage({ type: 'hello' });
    const bye = () => { try { channel.postMessage({ type: 'bye' }); } catch { /* closed */ } };
    window.addEventListener('beforeunload', bye);
    return () => { window.removeEventListener('beforeunload', bye); halt(); channel.close(); };
  }, [shell, media, monitorKey]);

  // On another screen, full screen there; on one screen, a window to move.
  useEffect(() => {
    (async () => {
      const displays = await shell.win.displays().catch(() => []);
      const other = displays.find((d) => !d.current);
      if (other) shell.win.fullscreen({ on: true, display: other.id }).then((r) => setFull(Boolean(r?.fullscreen))).catch(() => {});
    })();
  }, [shell]);

  useEffect(() => {
    const c = canvasRef.current;
    if (!c || !plan) return;
    if (c.width !== plan.W) c.width = plan.W;
    if (c.height !== plan.H) c.height = plan.H;
    paintFrame(c.getContext('2d'), plan, time);
  }, [plan, time]);

  const picture = (
    <div
      className={`vd-monitor-window${full ? ' full' : ''}`}
      data-time={time.toFixed(3)}
      data-clips={doc.patches.filter((p) => p.timing).length}
      onDoubleClick={() => shell.win.fullscreen({}).then((r) => setFull(Boolean(r?.fullscreen))).catch(() => {})}
      title={t('Double-click for full screen, or back')}
    >
      <canvas ref={canvasRef} className="vd-monitor-canvas" />
      {waiting ? <p className="vd-monitor-wait">{t('Waiting for the Video window')}</p> : null}
      <style>{MONITOR_CSS}</style>
    </div>
  );
  // Full screen, the picture alone; in a window, with the suite's title bar to move and close it by.
  return full ? picture : <AppFrame app={app} shell={shell} title={t('Video preview')}>{picture}</AppFrame>;
}

const MONITOR_CSS = `
.vd-monitor-window { flex: 1; min-height: 0; background: #000; display: grid; place-items: center; cursor: default; position: relative; }
.vd-monitor-window.full { position: fixed; inset: 0; z-index: 100; }
.vd-monitor-canvas { width: 100%; height: 100%; object-fit: contain; display: block; }
.vd-monitor-wait { position: absolute; color: #9aa1ad; font: 14px system-ui, sans-serif; }
`;
