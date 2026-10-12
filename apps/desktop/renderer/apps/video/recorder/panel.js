// Record: the screen recorder's window (docs/STUDIO.md section 8).
//
// One box for the whole job, as a screen recorder's setup is: on the left
// the frame as it will be recorded, live, with the camera's inset where it
// will sit (dragged to a corner before recording); on the right what to
// record (a screen, a window, or a region drawn on one), the camera, the
// microphone and the computer's own sound with their level meters, and the
// quality. Record counts down, then records until Stop: from the box, from
// the global key while another program is in front, or from the tray. Pause
// holds every file at once. While it records, the disk is watched, and the
// recording is written down every few seconds for the launcher to offer back
// after a crash (onProgress).
//
// Video lands the takes on its timeline; Presentations (mode 'slide') puts
// the screen's take on the slide, with no camera.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Dialog, Icon, Select, Spinner, t } from '@rutba/office-ui';
import {
  normaliseSettings, insetRect, cornerOf, regionCrop, bytesPerSecond, diskState, takeFileName,
  FRAME_RATES, SIZE_CAPS, LOW_DISK_MINUTES,
} from '@rutba/studio/recorder';
import { prompterProgress, PROMPTER_SPEEDS } from '@rutba/studio/intro';
import { RecorderRig } from './session.js';
import { RECORDER_CSS } from './css.js';
import { INTRO_CSS } from '../intro/css.js';

const STORE_KEY = 'recorder.settings';
const COPY_EVERY_MS = 5000;
const DISK_EVERY_MS = 5000;

const clock = (ms) => {
  const s = Math.floor(ms / 1000);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const ss = String(s % 60).padStart(2, '0');
  return h ? `${h}:${String(m).padStart(2, '0')}:${ss}` : `${m}:${ss}`;
};

/** A switch: a labelled checkbox drawn as a toggle. */
function Toggle({ on, onChange, label, icon, disabled, className = '' }) {
  return (
    <label className={`rec-toggle${disabled ? ' off' : ''}${className ? ` ${className}` : ''}`}>
      {icon ? <Icon name={icon} size={16} /> : null}
      <span className="rec-toggle-label">{label}</span>
      <input type="checkbox" checked={Boolean(on)} disabled={disabled} onChange={(e) => onChange(e.target.checked)} />
      <span className="rec-switch" aria-hidden="true" />
    </label>
  );
}

/** A row of choices, one pressed. */
function Segments({ value, options, onChange, disabled, className = '' }) {
  return (
    <div className={`rec-seg${className ? ` ${className}` : ''}`} role="radiogroup">
      {options.map((o) => (
        <button
          key={o.value}
          type="button"
          role="radio"
          aria-checked={value === o.value}
          className={value === o.value ? 'on' : ''}
          data-value={o.value}
          data-tip={o.tip || o.label}
          disabled={disabled}
          onClick={() => onChange(o.value)}
        >
          {o.icon ? <Icon name={o.icon} size={16} /> : null}
          <span>{o.label}</span>
        </button>
      ))}
    </div>
  );
}

/** A level meter, read from the rig ten times a second. */
function Meter({ rig, name, active }) {
  const bar = useRef(null);
  useEffect(() => {
    if (!active) { if (bar.current) bar.current.style.width = '0%'; return undefined; }
    const timer = setInterval(() => {
      const v = rig.level(name);
      const el = bar.current;
      if (!el) return;
      el.style.width = `${Math.round(v * 100)}%`;
      el.dataset.level = v > 0.9 ? 'loud' : v > 0.02 ? 'on' : 'quiet';
      el.parentElement.dataset.peak = String(Math.max(Number(el.parentElement.dataset.peak || 0), Math.round(v * 100)));
    }, 100);
    return () => clearInterval(timer);
  }, [rig, name, active]);
  return (
    <span className={`rec-meter rec-meter-${name}`} role="meter" aria-label={name === 'mic' ? t('Microphone level') : t('Computer sound level')} aria-valuemin={0} aria-valuemax={100}>
      <i ref={bar} />
    </span>
  );
}

/** A live stream in a video element. */
function Live({ stream, className, mirror, videoRef = null }) {
  const own = useRef(null);
  const ref = videoRef || own;
  useEffect(() => {
    const v = ref.current;
    if (!v) return;
    if (v.srcObject !== stream) v.srcObject = stream || null;
    if (stream) v.play?.().catch(() => {});
  }, [stream, ref]);
  return <video ref={ref} className={className} muted playsInline autoPlay style={mirror ? { transform: 'scaleX(-1)' } : undefined} />;
}

/** Where a picture of w by h sits inside a box of W by H, letterboxed. */
const containIn = (w, h, W, H) => {
  if (!w || !h || !W || !H) return { x: 0, y: 0, w: W, h: H };
  const k = Math.min(W / w, H / h);
  return { x: (W - w * k) / 2, y: (H - h * k) / 2, w: w * k, h: h * k };
};

/**
 * @param {{ shell, mode?: 'video' | 'slide', toast, onDone: (result) => void, onCancel: () => void, onProgress?: (takes, info) => void,
 *   prompter?: { title, script, speed, onSpeed }, force?: object }} props
 * `prompter`: a scene's script, scrolled above the frame at `speed` words a
 * minute while it records (an introduction's teleprompter, X3). `force`:
 * settings the scene needs (the camera on), over the ones kept.
 * onDone({ takes, crop, inset, settings }): takes as files, each
 * { role, path, contentType, hasAudio, width, height, duration }.
 */
export function RecorderPanel({ shell, mode = 'video', toast, onDone, onCancel, onProgress = null, prompter = null, force = null }) {
  const slide = mode === 'slide';
  const [settings, setSettings] = useState(() => normaliseSettings(null));
  const [loaded, setLoaded] = useState(false);
  const [sources, setSources] = useState(null);
  const [sourceError, setSourceError] = useState(null);
  const [sourceId, setSourceId] = useState(null);
  const [crop, setCrop] = useState(null);
  const [devices, setDevices] = useState({ cameras: [], mics: [] });
  const [, setTick] = useState(0);
  const [phase, setPhase] = useState('setup'); // setup | countdown | recording | stopping
  const [count, setCount] = useState(0);
  const [elapsed, setElapsed] = useState(0);
  const [disk, setDisk] = useState(null);
  const [keys, setKeys] = useState(null);
  const [folder, setFolder] = useState('');
  const [drag, setDrag] = useState(null); // the inset while dragged: { fx, fy }
  const [drawing, setDrawing] = useState(null); // a region being drawn: { x0, y0, x1, y1 } in fractions of the picture
  const rig = useMemo(() => new RecorderRig({ shell, onChange: () => setTick((n) => n + 1) }), [shell]);
  const frameRef = useRef(null);
  const screenRef = useRef(null);
  const phaseRef = useRef(phase);
  phaseRef.current = phase;
  const warned = useRef(false);
  const stopRef = useRef(null);
  const pauseRef = useRef(null);
  // Whether the main process is writing the pointer down for this recording (R7).
  const pointerOn = useRef(false);
  const [frameBox, setFrameBox] = useState({ w: 0, h: 0 });
  const [picture, setPicture] = useState({ w: 0, h: 0 });

  useEffect(() => () => rig.dispose(), [rig]);

  // The settings this person last recorded with, and where recordings go.
  useEffect(() => {
    let live = true;
    (async () => {
      const stored = await shell.store.get({ key: STORE_KEY, fallback: null }).catch(() => null);
      const paths = await shell.app.paths().catch(() => null);
      if (!live) return;
      const s = normaliseSettings(force ? { ...(stored || {}), ...force } : stored);
      setSettings(s);
      const sep = paths?.separator || '\\';
      setFolder(s.folder || (paths?.videos ? `${paths.videos}${sep}${t('Rutba Recordings')}` : ''));
      setLoaded(true);
    })();
    return () => { live = false; };
  }, [shell]);

  const change = useCallback((patch) => {
    setSettings((s) => {
      const next = normaliseSettings({ ...s, ...patch });
      shell.store.set({ key: STORE_KEY, value: next }).catch(() => {});
      return next;
    });
  }, [shell]);

  // What there is to record.
  const loadSources = useCallback(async () => {
    setSourceError(null);
    try {
      const list = (await shell.capture.sources()) || [];
      setSources(list);
    } catch (err) {
      setSourceError(err?.message || String(err));
      setSources([]);
    }
  }, [shell]);
  useEffect(() => { loadSources(); }, [loadSources]);

  const listed = useMemo(() => {
    const all = sources || [];
    if (settings.source === 'window') return all.filter((s) => s.kind === 'window');
    if (settings.source === 'region') return [...all.filter((s) => s.kind === 'screen'), ...all.filter((s) => s.kind === 'window')];
    return all.filter((s) => s.kind === 'screen');
  }, [sources, settings.source]);

  // The first of the kind chosen, until one is picked.
  useEffect(() => {
    if (!sources) return;
    if (!listed.some((s) => s.id === sourceId)) setSourceId(listed[0]?.id || null);
  }, [listed, sources, sourceId]);
  const source = listed.find((s) => s.id === sourceId) || null;

  // The sources open as the settings say, while setting up.
  useEffect(() => {
    if (!loaded || phase !== 'setup') return;
    rig.setScreen(source, { fps: settings.fps, cap: settings.cap, system: settings.system });
  }, [rig, loaded, phase, source, settings.fps, settings.cap, settings.system]);
  useEffect(() => {
    if (!loaded || phase !== 'setup') return;
    rig.setCamera(settings.camera && !slide, { cameraId: settings.cameraId, cameraHeight: settings.cameraHeight });
  }, [rig, loaded, phase, slide, settings.camera, settings.cameraId, settings.cameraHeight]);
  useEffect(() => {
    if (!loaded || phase !== 'setup') return;
    rig.setMic(settings.mic, { micId: settings.micId }).then(() => RecorderRig.devices().then(setDevices));
  }, [rig, loaded, phase, settings.mic, settings.micId]);
  useEffect(() => { RecorderRig.devices().then(setDevices); }, [rig.camera.stream]);

  // A region is the screen's, so another screen starts without one.
  useEffect(() => { setCrop(null); }, [sourceId, settings.source]);

  // The frame's size on screen, and the picture's own, for the overlays.
  useEffect(() => {
    const el = frameRef.current;
    if (!el) return undefined;
    const measure = () => { const r = el.getBoundingClientRect(); setFrameBox({ w: r.width, h: r.height }); };
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);
  useEffect(() => {
    const v = screenRef.current;
    if (!v) return undefined;
    const read = () => setPicture({ w: v.videoWidth, h: v.videoHeight });
    v.addEventListener('loadedmetadata', read);
    v.addEventListener('resize', read);
    read();
    return () => { v.removeEventListener('loadedmetadata', read); v.removeEventListener('resize', read); };
  }, [rig.screen.stream]);

  const camSettings = rig.camera.stream?.getVideoTracks()[0]?.getSettings?.() || {};
  const camAspect = camSettings.width && camSettings.height ? camSettings.height / camSettings.width : 9 / 16;
  const inset = useMemo(() => insetRect({ corner: settings.corner, size: settings.insetSize, shape: settings.shape, W: 1920, H: 1080, aspect: camAspect }), [settings.corner, settings.insetSize, settings.shape, camAspect]);
  const shownInset = drag ? { ...inset, fx: drag.fx, fy: drag.fy } : inset;
  const showCamera = !slide && settings.camera;

  // What the disk holds at this quality.
  const rate = useMemo(() => bytesPerSecond(settings, {
    screen: picture.w ? { width: picture.w, height: picture.h } : source?.display || { width: 1920, height: 1080 },
    camera: showCamera ? { width: camSettings.width || 1280, height: camSettings.height || 720 } : null,
  }), [settings, picture.w, picture.h, source, showCamera, camSettings.width, camSettings.height]);
  const checkDisk = useCallback(async () => {
    if (!folder) return null;
    try {
      const { free } = await shell.capture.space({ path: folder });
      const state = { ...diskState(free, rate), free };
      setDisk(state);
      return state;
    } catch {
      return null;
    }
  }, [shell, folder, rate]);
  useEffect(() => { if (phase === 'setup') checkDisk(); }, [phase, checkDisk]);

  /* ── recording ─────────────────────────────────────────────────────── */

  const fileFor = useCallback(async (role, ext) => {
    if (slide) return (await shell.fs.temp({ ext })).path;
    const stem = role === 'camera' ? t('Camera') : t('Screen');
    return (await shell.capture.take({ dir: folder, name: takeFileName(stem, new Date(), ext) })).path;
  }, [shell, slide, folder]);

  const begin = useCallback(async () => {
    try {
      await rig.record({ settings, file: fileFor, slide });
      pointerOn.current = false;
      if (!slide && settings.pointer && source) {
        const pt = await shell.capture.pointer({ on: true, id: source.id, fps: settings.fps }).catch(() => null);
        pointerOn.current = Boolean(pt?.tracking);
        if (!pt?.tracking) toast?.(t('The pointer is not written down over another program\'s window. Record its screen to keep it.'), { ms: 7000 });
      }
      setPhase('recording');
      warned.current = false;
      const c = await shell.capture.controls({ on: true, paused: false, minimise: settings.minimise }).catch(() => null);
      setKeys(c);
    } catch (err) {
      setPhase('setup');
      toast?.(err?.message || String(err), { tone: 'bad', ms: 6000 });
    }
  }, [rig, settings, fileFor, slide, shell, toast, source]);

  const record = useCallback(async () => {
    if (!rig.screen.stream) return;
    const state = await checkDisk();
    if (state?.level === 'stop') {
      toast?.(t('There is not enough space on the disk to record.'), { tone: 'bad', ms: 6000 });
      return;
    }
    if (settings.countdown > 0) {
      setCount(settings.countdown);
      setPhase('countdown');
    } else {
      setPhase('starting');
      begin();
    }
  }, [rig, checkDisk, settings.countdown, begin, toast]);

  // The countdown, a second a step.
  useEffect(() => {
    if (phase !== 'countdown') return undefined;
    if (count <= 0) { setPhase('starting'); begin(); return undefined; }
    const timer = setTimeout(() => setCount((n) => n - 1), 1000);
    return () => clearTimeout(timer);
  }, [phase, count, begin]);

  const finish = useCallback(async () => {
    if (phaseRef.current !== 'recording') return;
    setPhase('stopping');
    shell.capture.controls({ on: false }).catch(() => {});
    try {
      const pointer = pointerOn.current ? await shell.capture.pointer({ on: false }).catch(() => null) : null;
      pointerOn.current = false;
      const takes = await rig.stop();
      const camera = takes.find((k) => k.role === 'camera');
      const kept = pointer && (pointer.samples?.length || pointer.clicks?.length) ? pointer : null;
      onDone({
        takes: takes.map((k) => ({ ...k, ...(k.role === 'screen' && crop ? { crop } : {}), ...(k.role === 'screen' && kept ? { pointer: kept } : {}), ...(k.role === 'camera' ? { inset, shape: settings.shape, mirror: settings.mirror } : {}) })),
        crop,
        inset: camera ? inset : null,
        settings,
      });
    } catch (err) {
      toast?.(err?.message || String(err), { tone: 'bad', ms: 8000 });
      onCancel();
    }
  }, [rig, shell, crop, inset, settings, onDone, onCancel, toast]);
  stopRef.current = finish;

  const togglePause = useCallback(() => {
    if (phaseRef.current !== 'recording') return;
    if (rig.paused) rig.resume();
    else rig.pause();
    shell.capture.controls({ on: true, paused: rig.paused }).catch(() => {});
    if (pointerOn.current) shell.capture.pointer({ on: true, paused: rig.paused }).catch(() => {});
  }, [rig, shell]);
  pauseRef.current = togglePause;

  // The global keys and the tray come back here.
  useEffect(() => shell.on('capture:command', ({ command } = {}) => {
    if (command === 'stop') stopRef.current?.();
    else if (command === 'pause') pauseRef.current?.();
  }), [shell]);

  // The clock on screen, the disk watched, and the work written down for after a crash.
  useEffect(() => {
    if (phase !== 'recording') return undefined;
    const tick = setInterval(() => setElapsed(rig.elapsed), 250);
    const watch = setInterval(async () => {
      const state = await checkDisk();
      if (!state || phaseRef.current !== 'recording') return;
      if (state.level === 'stop') {
        toast?.(t('The disk is nearly full: the recording was stopped and kept.'), { tone: 'bad', ms: 8000 });
        stopRef.current?.();
      } else if (state.level === 'low' && !warned.current) {
        warned.current = true;
        toast?.(t('About {minutes} minutes of space are left on the disk.', { minutes: Math.max(1, Math.floor(state.minutes)) }), { ms: 8000 });
      }
    }, DISK_EVERY_MS);
    const copy = onProgress ? setInterval(() => {
      if (phaseRef.current === 'recording') onProgress(rig.snapshot().map((k) => ({ ...k, ...(k.role === 'screen' && crop ? { crop } : {}), ...(k.role === 'camera' ? { inset, shape: settings.shape, mirror: settings.mirror } : {}) })));
    }, COPY_EVERY_MS) : null;
    return () => { clearInterval(tick); clearInterval(watch); if (copy) clearInterval(copy); };
  }, [phase, rig, checkDisk, onProgress, toast, crop, inset, settings.shape, settings.mirror]);

  const cancel = useCallback(async () => {
    if (phaseRef.current === 'recording' || phaseRef.current === 'stopping') {
      shell.capture.controls({ on: false }).catch(() => {});
      if (pointerOn.current) shell.capture.pointer({ on: false }).catch(() => {});
      pointerOn.current = false;
      await rig.cancel();
    }
    onCancel();
  }, [rig, shell, onCancel]);

  // Gone in the middle of a recording (the window closed): the keys and the tray are let go.
  useEffect(() => () => {
    if (phaseRef.current === 'recording') shell.capture.controls({ on: false }).catch(() => {});
    if (pointerOn.current) shell.capture.pointer({ on: false }).catch(() => {});
  }, [shell]);

  /* ── the teleprompter ──────────────────────────────────────────────── */

  const promptView = useRef(null);
  const promptText = useRef(null);
  const script = prompter?.script || '';
  const speed = prompter?.speed || PROMPTER_SPEEDS.normal;
  const progress = phase === 'recording' || phase === 'stopping' ? prompterProgress(script, elapsed / 1000, speed) : 0;
  const [shift, setShift] = useState(0);
  useEffect(() => {
    const view = promptView.current;
    const text = promptText.current;
    if (!view || !text) { setShift(0); return; }
    // The line being read sits a third of the way down, as on a prompter.
    setShift(Math.max(0, progress * Math.max(0, text.scrollHeight - view.clientHeight * 0.66)));
  }, [progress, script]);
  // The speed offered nearest the one kept.
  const speedKey = Object.keys(PROMPTER_SPEEDS).sort((x, y) => Math.abs(PROMPTER_SPEEDS[x] - speed) - Math.abs(PROMPTER_SPEEDS[y] - speed))[0];

  /* ── the frame: dragging the camera, drawing a region ─────────────── */

  const toFrame = (e) => {
    const r = frameRef.current.getBoundingClientRect();
    return { x: (e.clientX - r.left) / r.width, y: (e.clientY - r.top) / r.height };
  };
  const pic = containIn(picture.w, picture.h, frameBox.w, frameBox.h);
  const picFrac = frameBox.w ? { x: pic.x / frameBox.w, y: pic.y / frameBox.h, w: pic.w / frameBox.w, h: pic.h / frameBox.h } : { x: 0, y: 0, w: 1, h: 1 };
  const toPicture = (p) => ({ x: (p.x - picFrac.x) / picFrac.w, y: (p.y - picFrac.y) / picFrac.h });

  const grabInset = (e) => {
    if (phase !== 'setup' || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const start = toFrame(e);
    const from = { fx: inset.fx, fy: inset.fy };
    const move = (ev) => {
      const p = toFrame(ev);
      setDrag({
        fx: Math.max(0, Math.min(1 - inset.fw, from.fx + p.x - start.x)),
        fy: Math.max(0, Math.min(1 - inset.fh, from.fy + p.y - start.y)),
      });
    };
    const up = (ev) => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      const p = toFrame(ev);
      const at = { fx: from.fx + p.x - start.x, fy: from.fy + p.y - start.y, fw: inset.fw, fh: inset.fh };
      setDrag(null);
      change({ corner: cornerOf(at) });
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const drawRegion = (e) => {
    if (phase !== 'setup' || settings.source !== 'region' || e.button !== 0 || !picture.w) return;
    e.preventDefault();
    const a = toPicture(toFrame(e));
    setDrawing({ x0: a.x, y0: a.y, x1: a.x, y1: a.y });
    const move = (ev) => {
      const b = toPicture(toFrame(ev));
      setDrawing({ x0: a.x, y0: a.y, x1: b.x, y1: b.y });
    };
    const up = (ev) => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      const b = toPicture(toFrame(ev));
      setDrawing(null);
      const next = regionCrop({ x: a.x, y: a.y, w: b.x - a.x, h: b.y - a.y }, { width: 1, height: 1 });
      if (next) setCrop(next);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const region = drawing
    ? regionCrop({ x: drawing.x0, y: drawing.y0, w: drawing.x1 - drawing.x0, h: drawing.y1 - drawing.y0 }, { width: 1, height: 1 })
    : crop;
  const regionBox = region ? {
    left: `${(picFrac.x + region.x * picFrac.w) * 100}%`,
    top: `${(picFrac.y + region.y * picFrac.h) * 100}%`,
    width: `${region.w * picFrac.w * 100}%`,
    height: `${region.h * picFrac.h * 100}%`,
  } : null;

  /* ── the box ───────────────────────────────────────────────────────── */

  const busy = phase !== 'setup';
  const recording = phase === 'recording' || phase === 'stopping';
  const paused = rig.paused;
  const canRecord = phase === 'setup' && Boolean(rig.screen.stream) && disk?.level !== 'stop';
  const kindOptions = [
    { value: 'screen', label: t('Screen'), icon: 'screen', tip: t('Screen: the whole of one screen') },
    { value: 'window', label: t('Window'), icon: 'window', tip: t('Window: one window, wherever it is') },
    { value: 'region', label: t('Region'), icon: 'region', tip: t('Region: a part you draw, kept as a crop you can change later') },
  ];
  const fpsLabel = (n) => t('{n} frames a second', { n });
  const capLabel = (c) => (c ? t('Up to {n}p', { n: c }) : t('As the screen'));
  const minutes = disk ? Math.floor(disk.minutes) : null;
  const diskNote = disk == null ? null
    : disk.level === 'stop' ? t('There is not enough space on the disk to record.')
      : disk.minutes < 120 ? t('Room for about {minutes} minutes on the disk', { minutes: Math.max(0, minutes) })
        : t('Room for about {hours} hours on the disk', { hours: Math.floor(disk.minutes / 60) });

  const title = slide ? t('Screen Recording') : t('Record');

  const actions = (
    <>
      <div className="rec-foot-info">
        {recording ? (
          <>
            <span className={`rec-dot${paused ? ' held' : ''}`} aria-hidden="true" />
            <b className="rec-time" aria-live="off">{clock(elapsed)}</b>
            {paused ? <span className="rec-chip">{t('Paused')}</span> : null}
            {settings.mic ? <Meter rig={rig} name="mic" active={recording} /> : null}
            {keys?.stop ? <span className="rec-keys">{t('{key} stops it from anywhere', { key: keys.stop })}</span> : null}
          </>
        ) : (
          <span className={`rec-disk${disk && disk.minutes < LOW_DISK_MINUTES ? ' low' : ''}`}>{diskNote}</span>
        )}
      </div>
      {recording ? (
        <>
          <Button icon={paused ? 'play' : 'pause'} label={paused ? t('Resume') : t('Pause')} className="rec-pause" disabled={phase !== 'recording'} title={keys?.pause ? `${paused ? t('Resume') : t('Pause')} (${keys.pause})` : undefined} onClick={togglePause} />
          <Button primary icon="stop" label={t('Stop')} className="rec-stop" disabled={phase !== 'recording'} title={keys?.stop ? `${t('Stop')} (${keys.stop})` : undefined} onClick={finish} />
        </>
      ) : (
        <>
          <Button label={t('Cancel')} onClick={cancel} />
          <Button primary icon="record" label={phase === 'countdown' ? t('Starting…') : t('Record')} className="rec-go" disabled={!canRecord} onClick={record} />
        </>
      )}
    </>
  );

  return (
    <Dialog title={title} className="rec-dialog" onClose={busy ? undefined : cancel} actions={actions}>
      <style>{RECORDER_CSS}</style>
      {prompter ? <style>{INTRO_CSS}</style> : null}
      <div className={`rec-body${slide ? ' slide' : ''}`} data-phase={phase}>
        <div className={`rec-stage${prompter ? ' with-prompter' : ''}`}>
          {prompter ? (
            <div className="rec-prompter" data-progress={progress.toFixed(3)} aria-label={t('Teleprompter')}>
              <div className="rec-prompter-head">
                <Icon name="prompter" size={15} />
                <b>{prompter.title}</b>
                <Segments
                  value={speedKey}
                  className="rec-prompter-speed"
                  onChange={(k) => prompter.onSpeed?.(PROMPTER_SPEEDS[k])}
                  options={[{ value: 'slow', label: t('Slow') }, { value: 'normal', label: t('Normal') }, { value: 'fast', label: t('Fast') }]}
                />
              </div>
              <div className="rec-prompter-view" ref={promptView}>
                {script ? (
                  <div className="rec-prompter-text" dir="auto" ref={promptText} style={{ transform: `translateY(${-shift}px)` }}>{script}</div>
                ) : (
                  <div className="rec-prompter-empty">{t('This scene has no script. Write one in its card and it scrolls here while you record.')}</div>
                )}
              </div>
            </div>
          ) : null}
          <div
            ref={frameRef}
            className={`rec-frame${recording ? ' live' : ''}${paused ? ' held' : ''}${settings.source === 'region' && phase === 'setup' ? ' drawing' : ''}`}
            dir="ltr"
            tabIndex={0}
            aria-label={t('Preview')}
            onPointerDown={drawRegion}
          >
            <Live stream={rig.screen.stream} className="rec-screen" videoRef={screenRef} />
            {!rig.screen.stream ? (
              <div className="rec-empty">
                {rig.screen.error || sourceError ? (
                  <><Icon name="info" size={22} /><span>{rig.screen.error || sourceError}</span></>
                ) : sources == null || source ? (
                  <><Spinner /><span>{t('Opening the screen…')}</span></>
                ) : (
                  <><Icon name="screen" size={26} /><span>{t('Choose a screen or a window to record.')}</span></>
                )}
              </div>
            ) : null}
            {regionBox ? <div className="rec-region" style={regionBox}><span>{t('Region')}</span></div> : null}
            {settings.source === 'region' && !region && rig.screen.stream && phase === 'setup' ? (
              <div className="rec-hint">{t('Drag on the picture to choose the part to record.')}</div>
            ) : null}
            {showCamera ? (
              <div
                className={`rec-inset shape-${settings.shape}${drag ? ' dragging' : ''}`}
                style={{ left: `${shownInset.fx * 100}%`, top: `${shownInset.fy * 100}%`, width: `${shownInset.fw * 100}%`, height: `${shownInset.fh * 100}%` }}
                data-corner={settings.corner}
                onPointerDown={grabInset}
                data-tip={phase === 'setup' ? t('Drag the camera to a corner') : undefined}
              >
                {rig.camera.stream ? <Live stream={rig.camera.stream} className="rec-cam" mirror={settings.mirror} /> : (
                  <span className="rec-cam-off"><Icon name="camera" size={20} />{rig.camera.error ? <small>{rig.camera.error}</small> : null}</span>
                )}
              </div>
            ) : null}
            {recording ? <span className="rec-badge"><i />{paused ? t('Paused') : t('Recording')}</span> : null}
            {phase === 'countdown' ? (
              <div className="rec-count" aria-live="assertive"><span key={count}>{count}</span></div>
            ) : null}
          </div>
          <div className="rec-under">
            <span>
              {source ? source.name : ''}
              {picture.w ? ` · ${picture.w} × ${picture.h}` : ''}
              {region ? ` · ${t('Region {w} × {h}', { w: Math.round(region.w * picture.w), h: Math.round(region.h * picture.h) })}` : ''}
            </span>
            {crop && phase === 'setup' ? <Button ghost icon="close" label={t('Whole picture')} className="rec-clear-region" onClick={() => setCrop(null)} /> : null}
          </div>
        </div>

        <aside className="rec-side" aria-disabled={busy}>
          <section className="rec-sec">
            <h3>{t('Record')}</h3>
            <Segments value={settings.source} options={kindOptions} disabled={busy} onChange={(v) => change({ source: v })} className="rec-kinds" />
            <div className="rec-sources">
              {sources == null ? <div className="rec-wait"><Spinner /> {t('Looking at what is open…')}</div> : null}
              {sources && !listed.length ? <div className="rec-wait">{settings.source === 'window' ? t('No other window is open to record.') : t('No screen can be recorded.')}</div> : null}
              {listed.map((s) => (
                <button key={s.id} type="button" className={`rec-source${s.id === sourceId ? ' on' : ''}`} data-source={s.id} data-tip={s.name} disabled={busy} onClick={() => setSourceId(s.id)}>
                  <img src={s.thumbnail} alt="" draggable={false} />
                  <span>{s.name}</span>
                </button>
              ))}
            </div>
            <Button ghost icon="refresh" label={t('Refresh the list')} className="rec-refresh" disabled={busy} onClick={loadSources} />
            {!slide ? (
              <Toggle on={settings.pointer} icon="pointer" label={t('Pointer and clicks')} disabled={busy} onChange={(v) => change({ pointer: v })} className="rec-pointer-toggle" />
            ) : null}
          </section>

          {!slide ? (
            <section className="rec-sec">
              <Toggle on={settings.camera} icon="camera" label={t('Camera')} disabled={busy} onChange={(v) => change({ camera: v })} className="rec-camera-toggle" />
              {settings.camera ? (
                <div className="rec-sub">
                  {devices.cameras.length > 1 ? (
                    <Select value={settings.cameraId} disabled={busy} aria-label={t('Camera')} onChange={(e) => change({ cameraId: e.target.value })}>
                      <option value="">{t('The usual camera')}</option>
                      {devices.cameras.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}
                    </Select>
                  ) : null}
                  <Segments
                    value={settings.shape}
                    disabled={busy}
                    className="rec-shapes"
                    onChange={(v) => change({ shape: v })}
                    options={[{ value: 'circle', label: t('Circle') }, { value: 'rounded', label: t('Rounded') }, { value: 'rect', label: t('Rectangle') }]}
                  />
                  <Segments
                    value={settings.insetSize}
                    disabled={busy}
                    className="rec-sizes"
                    onChange={(v) => change({ insetSize: v })}
                    options={[{ value: 'small', label: t('Small') }, { value: 'medium', label: t('Medium') }, { value: 'large', label: t('Large') }]}
                  />
                  <Toggle on={settings.mirror} label={t('Mirror the camera')} disabled={busy} onChange={(v) => change({ mirror: v })} />
                </div>
              ) : null}
            </section>
          ) : null}

          {(
            <section className="rec-sec">
              <Toggle on={settings.mic} icon="mic" label={t('Microphone')} disabled={busy} onChange={(v) => change({ mic: v })} className="rec-mic-toggle" />
              {settings.mic ? (
                <div className="rec-sub">
                  {devices.mics.length > 1 ? (
                    <Select value={settings.micId} disabled={busy} aria-label={t('Microphone')} onChange={(e) => change({ micId: e.target.value })}>
                      <option value="">{t('The usual microphone')}</option>
                      {devices.mics.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}
                    </Select>
                  ) : null}
                  {rig.mic.error ? <small className="rec-err">{rig.mic.error}</small> : <Meter rig={rig} name="mic" active={Boolean(rig.mic.stream)} />}
                </div>
              ) : null}
              <Toggle on={settings.system} icon="speaker" label={t('Computer sound')} disabled={busy} onChange={(v) => change({ system: v })} className="rec-system-toggle" />
              {settings.system ? (
                <div className="rec-sub">
                  {rig.system.error ? <small className="rec-err">{rig.system.error}</small> : <Meter rig={rig} name="system" active={Boolean(rig.nodes.system)} />}
                </div>
              ) : null}
            </section>
          )}

          <section className="rec-sec">
            <h3>{t('Quality')}</h3>
            <label className="rec-field">
              <span>{t('Frame rate')}</span>
              <Select value={String(settings.fps)} disabled={busy} onChange={(e) => change({ fps: Number(e.target.value) })}>
                {FRAME_RATES.map((n) => <option key={n} value={String(n)}>{fpsLabel(n)}</option>)}
              </Select>
            </label>
            <label className="rec-field">
              <span>{t('Size')}</span>
              <Select value={String(settings.cap)} disabled={busy} onChange={(e) => change({ cap: Number(e.target.value) })}>
                {SIZE_CAPS.map((c) => <option key={c} value={String(c)}>{capLabel(c)}</option>)}
              </Select>
            </label>
            <label className="rec-field">
              <span>{t('Quality')}</span>
              <Select value={settings.quality} disabled={busy} onChange={(e) => change({ quality: e.target.value })}>
                <option value="small">{t('Smaller files')}</option>
                <option value="standard">{t('Standard')}</option>
                <option value="high">{t('Sharper')}</option>
              </Select>
            </label>
            <label className="rec-field">
              <span>{t('Countdown')}</span>
              <Select value={String(settings.countdown)} disabled={busy} onChange={(e) => change({ countdown: Number(e.target.value) })} className="rec-countdown">
                <option value="0">{t('None')}</option>
                {[3, 5, 10].map((n) => <option key={n} value={String(n)}>{t('{seconds} s', { seconds: n })}</option>)}
              </Select>
            </label>
            <Toggle on={settings.minimise} label={t('Minimise this window while recording')} disabled={busy} onChange={(v) => change({ minimise: v })} />
            <small className="rec-rate">{t('About {mb} MB a minute', { mb: Math.max(1, Math.round((rate * 60) / 1e6)) })}</small>
          </section>
        </aside>
      </div>
    </Dialog>
  );
}
