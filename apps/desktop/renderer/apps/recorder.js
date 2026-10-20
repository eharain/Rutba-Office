// Rutba Recorder: record the screen, the camera and the voice as one finished
// video, made live from scenes (owner, 2026-10-19: "split the recording tool
// out to be able to record video and stream videos live"), and take still
// pictures of the same (owner, 2026-10-20: "there should be a way to capture
// pictures on camera and from screens etc.").
//
// The window is a small studio. On the left the scenes: the screen with the
// camera in a corner, the screen alone, the camera alone, the two side by
// side, and a Be right back card. In the middle the program, the picture as
// it is recorded, drawn live. On the right the sources: the screen or window,
// the camera, the microphone and the computer's sound with their levels, a
// title across the foot of the frame and a logo in a corner. A scene can be
// changed while it records. Stop keeps one file, a WebM in the recordings
// folder, to play, to edit in Video, or to find in its folder.
//
// The sources are the screen recorder's own rig (Video → Record Screen), so
// the screen, window, camera and sound behave the same in both.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Ribbon, Group, Button, Rows, Select, Icon, Input, Panel, Content, Spinner, useToast, formatBytes, basename, t } from '@rutba/office-ui';
import { normaliseStudio, SCENES, STUDIO_DEFAULTS } from '@rutba/studio/compositor';
import { takeFileName, bytesPerSecond, diskState, regionCrop } from '@rutba/studio/recorder';
import { AppFrame, useAppMenu } from '../shell.js';
import { RecorderRig } from './video/recorder/session.js';
import { Toggle, Segments, Meter } from './video/recorder/panel.js';
import { RECORDER_CSS } from './video/recorder/css.js';
import { Program } from './recorder/program.js';
import { LiveEncoder, AUDIO_KBPS } from './recorder/live.js';
import { LiveDialog, LIVE_CSS } from './recorder/live-dialog.js';
import { liveBitrate } from '@rutba/stream/endpoints';
import { STUDIO_CSS } from './recorder/css.js';

const STORE_KEY = 'recorder.studio';
const DISK_EVERY_MS = 5000;

const SCENE_ICON = { 'screen-camera': 'layers', screen: 'screen', camera: 'camera', split: 'grid', brb: 'clock' };
const SCENE_NAMES = () => ({
  'screen-camera': t('Screen and camera'),
  screen: t('Screen'),
  camera: t('Camera'),
  split: t('Side by side'),
  brb: t('Be right back'),
});

const clock = (ms) => {
  const s = Math.floor(ms / 1000);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const ss = String(s % 60).padStart(2, '0');
  return h ? `${h}:${String(m).padStart(2, '0')}:${ss}` : `${m}:${ss}`;
};

export default function Recorder({ app, shell }) {
  const toast = useToast();
  const [settings, setSettings] = useState(() => normaliseStudio(null));
  const [loaded, setLoaded] = useState(false);
  const [sources, setSources] = useState(null);
  const [kind, setKind] = useState('screen');
  // Region (owner, 2026-10-20: "in recorder the region is not present as in
  // video is"): the part of the screen or window recorded, and the one being drawn.
  const [crop, setCrop] = useState(null);
  const [drawing, setDrawing] = useState(null);
  const [sourceId, setSourceId] = useState(null);
  const [devices, setDevices] = useState({ cameras: [], mics: [] });
  const [phase, setPhase] = useState('setup'); // setup | countdown | recording | stopping
  const [count, setCount] = useState(0);
  const [elapsed, setElapsed] = useState(0);
  const [paused, setPaused] = useState(false);
  const [disk, setDisk] = useState(null);
  const [keys, setKeys] = useState(null);
  const [saved, setSaved] = useState(null); // { path, written, seconds }
  // Take Picture: the last picture kept, the seconds left on its timer, and a flash each time.
  const [picture, setPicture] = useState(null); // { path, width, height, url }
  const [shotCount, setShotCount] = useState(null);
  const [flash, setFlash] = useState(0);
  const [logoUrl, setLogoUrl] = useState(null);
  const [, setTick] = useState(0);
  const [tab, setTab] = useState('home');
  // Go Live: off, starting, live or ending; the destinations as the main process reports them.
  const [livePhase, setLivePhase] = useState('off');
  const [liveStatus, setLiveStatus] = useState([]);
  const [liveSince, setLiveSince] = useState(0);
  const [liveOpen, setLiveOpen] = useState(false);
  const encoder = useRef(null);
  const liveRecord = useRef(false);
  const canvasRef = useRef(null);
  const phaseRef = useRef(phase);
  phaseRef.current = phase;
  const stopRef = useRef(null);
  const pauseRef = useRef(null);
  const warned = useRef(false);
  const rig = useMemo(() => new RecorderRig({ shell, onChange: () => setTick((n) => n + 1) }), [shell]);
  const program = useRef(null);
  const stalls = useRef(0);

  const appMenu = useAppMenu({ shell, appKey: 'recorder', onNew: () => shell.win.create({ app: 'recorder' }), onOpen: async () => {
    const files = await shell.dialog.open({ title: t('Open'), filters: [{ name: t('Video and sound'), extensions: ['webm', 'mp4', 'mov', 'mkv', 'weba', 'm4a', 'mp3', 'wav'] }] });
    if (files?.[0]) shell.win.create({ app: 'video', file: files[0] });
  } });

  // The settings this person last recorded with, and where recordings go.
  useEffect(() => {
    let live = true;
    (async () => {
      const stored = await shell.store.get({ key: STORE_KEY, fallback: null }).catch(() => null);
      const paths = await shell.app.paths().catch(() => null);
      if (!live) return;
      const s = normaliseStudio(stored);
      // The card's words in the window's language until the person writes their own.
      if (!stored?.brb || stored.brb === STUDIO_DEFAULTS.brb) s.brb = t('Be right back');
      const sep = paths?.separator || '\\';
      if (!s.folder && paths?.videos) s.folder = `${paths.videos}${sep}${t('Rutba Recordings')}`;
      setSettings(s);
      setLoaded(true);
    })();
    return () => { live = false; };
  }, [shell]);

  const change = useCallback((patch) => {
    setSettings((s) => {
      const next = normaliseStudio({ ...s, ...patch });
      shell.store.set({ key: STORE_KEY, value: next }).catch(() => {});
      return next;
    });
  }, [shell]);

  // The program, drawn from the start, and let go with the window.
  useEffect(() => {
    const p = new Program({ rig, canvas: canvasRef.current });
    program.current = p;
    // A screen that gives no picture is opened again, twice at most for each choice of source.
    p.onStall = (name) => {
      if (name !== 'screen' || stalls.current >= 2) return;
      stalls.current += 1;
      rig.refreshScreen();
    };
    p.set(settings);
    p.start();
    return () => { p.dispose(); rig.dispose(); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rig]);
  useEffect(() => { program.current?.set(settings); }, [settings]);

  // The logo, read from its file.
  useEffect(() => {
    let url = null;
    let live = true;
    if (!settings.logo) { setLogoUrl(null); program.current?.setLogo(null); return undefined; }
    shell.fs.read({ path: settings.logo }).then(({ bytes }) => {
      if (!live) return;
      url = URL.createObjectURL(new Blob([bytes]));
      setLogoUrl(url);
      program.current?.setLogo(url);
    }).catch(() => { if (live) toast(t('The logo could not be read.'), { tone: 'bad' }); });
    return () => { live = false; if (url) URL.revokeObjectURL(url); };
  }, [shell, settings.logo, toast]);

  // What there is to record.
  const loadSources = useCallback(async () => {
    try {
      setSources((await shell.capture.sources()) || []);
    } catch {
      setSources([]);
    }
  }, [shell]);
  useEffect(() => { loadSources(); }, [loadSources]);
  const listed = useMemo(() => {
    const all = sources || [];
    // A region is drawn on a screen or a window: the screens first, then the windows.
    if (kind === 'region') return [...all.filter((s) => s.kind === 'screen'), ...all.filter((s) => s.kind === 'window')];
    return all.filter((s) => s.kind === kind);
  }, [sources, kind]);
  useEffect(() => {
    if (!sources) return;
    if (!listed.some((s) => s.id === sourceId)) setSourceId(listed[0]?.id || null);
  }, [listed, sources, sourceId]);
  const source = listed.find((s) => s.id === sourceId) || null;
  useEffect(() => { stalls.current = 0; }, [sourceId]);
  // A region belongs to its source: another source, or another kind, starts whole.
  useEffect(() => { setCrop(null); }, [sourceId, kind]);
  useEffect(() => { program.current?.setCrop(kind === 'region' ? crop : null); }, [crop, kind]);

  /** A drag on the picture, read as the part of the source under it. */
  const frameRef = useRef(null);
  const drawRegion = (e) => {
    if (kind !== 'region' || phase !== 'setup' || e.button !== 0 || !program.current?.screenAt) return;
    e.preventDefault();
    const box = frameRef.current.getBoundingClientRect();
    const at = (ev) => ({ fx: (ev.clientX - box.left) / box.width, fy: (ev.clientY - box.top) / box.height });
    const a = at(e);
    const from = program.current.toSource(a.fx, a.fy);
    if (!from) return;
    setDrawing({ a, b: a });
    const move = (ev) => setDrawing({ a, b: at(ev) });
    const up = (ev) => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      setDrawing(null);
      const b = at(ev);
      const to = program.current.toSource(b.fx, b.fy);
      if (!to) return;
      const next = regionCrop({ x: from.x, y: from.y, w: to.x - from.x, h: to.y - from.y }, { width: 1, height: 1 });
      if (next) setCrop(next);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  // The sources open as the settings say. The screen stays as it is while it records.
  const needsScreen = settings.scene !== 'camera' && settings.scene !== 'brb';
  useEffect(() => {
    if (!loaded || phase !== 'setup') return;
    rig.setScreen(source, { fps: settings.fps, cap: settings.size, system: settings.system });
  }, [rig, loaded, phase, source, settings.fps, settings.size, settings.system]);
  useEffect(() => {
    if (!loaded) return;
    rig.setCamera(settings.camera, { cameraId: settings.cameraId, cameraHeight: 720 });
  }, [rig, loaded, settings.camera, settings.cameraId]);
  useEffect(() => {
    if (!loaded || phase !== 'setup') return;
    rig.setMic(settings.mic, { micId: settings.micId }).then(() => RecorderRig.devices().then(setDevices));
  }, [rig, loaded, phase, settings.mic, settings.micId]);
  useEffect(() => { RecorderRig.devices().then(setDevices); }, [rig.camera.stream]);

  // The disk, before and while it records.
  const rate = useMemo(() => bytesPerSecond({ fps: settings.fps, cap: settings.size, quality: settings.quality, camera: false, mic: settings.mic, system: settings.system }, { screen: settings.size === 720 ? { width: 1280, height: 720 } : { width: 1920, height: 1080 } }), [settings.fps, settings.size, settings.quality, settings.mic, settings.system]);
  const checkDisk = useCallback(async () => {
    if (!settings.folder) return null;
    try {
      const { free } = await shell.capture.space({ path: settings.folder });
      const state = { ...diskState(free, rate), free };
      setDisk(state);
      return state;
    } catch {
      return null;
    }
  }, [shell, settings.folder, rate]);
  useEffect(() => { if (phase === 'setup') checkDisk(); }, [phase, checkDisk]);

  /* ── recording ─────────────────────────────────────────────────────── */

  const begin = useCallback(async () => {
    try {
      setSaved(null);
      const { path } = await shell.capture.take({ dir: settings.folder, name: takeFileName(t('Recording'), new Date(), 'webm') });
      await program.current.record({ shell, path });
      setPhase('recording');
      setPaused(false);
      warned.current = false;
      const c = await shell.capture.controls({ on: true, paused: false, minimise: settings.minimise }).catch(() => null);
      setKeys(c);
    } catch (err) {
      setPhase('setup');
      toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
    }
  }, [shell, settings.folder, settings.minimise, toast]);

  const record = useCallback(async () => {
    if (phaseRef.current !== 'setup') return;
    const state = await checkDisk();
    if (state?.level === 'stop') { toast(t('There is not enough space on the disk to record.'), { tone: 'bad', ms: 6000 }); return; }
    if (settings.countdown > 0) { setCount(settings.countdown); setPhase('countdown'); }
    else { setPhase('starting'); begin(); }
  }, [checkDisk, settings.countdown, begin, toast]);

  useEffect(() => {
    if (phase !== 'countdown') return undefined;
    if (count <= 0) { setPhase('starting'); begin(); return undefined; }
    const timer = setTimeout(() => setCount((n) => n - 1), 1000);
    return () => clearTimeout(timer);
  }, [phase, count, begin]);

  const stop = useCallback(async () => {
    if (phaseRef.current !== 'recording') return;
    setPhase('stopping');
    shell.capture.controls({ on: false }).catch(() => {});
    try {
      const done = await program.current.stop();
      setSaved(done);
      toast(t('Saved {name}', { name: basename(done.path) }), { tone: 'good', ms: 5000 });
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 8000 });
    } finally {
      setPhase('setup');
      setPaused(false);
      setElapsed(0);
    }
  }, [shell, toast]);
  stopRef.current = stop;

  const togglePause = useCallback(() => {
    if (phaseRef.current !== 'recording') return;
    const p = program.current;
    if (p.paused) p.resume();
    else p.pause();
    setPaused(p.paused);
    shell.capture.controls({ on: true, paused: p.paused }).catch(() => {});
  }, [shell]);
  pauseRef.current = togglePause;

  const pictureRef = useRef(null);
  useEffect(() => shell.on('capture:command', ({ command } = {}) => {
    if (command === 'stop') stopRef.current?.();
    else if (command === 'pause') pauseRef.current?.();
    else if (command === 'picture') pictureRef.current?.();
  }), [shell]);
  // Take Picture's key, from anywhere while this window is open.
  const [pictureKey, setPictureKey] = useState('');
  useEffect(() => {
    shell.capture.pictureKey({ on: true }).then((r) => setPictureKey(r?.key || '')).catch(() => {});
    return () => { shell.capture.pictureKey({ on: false }).catch(() => {}); };
  }, [shell]);

  useEffect(() => {
    if (phase !== 'recording') return undefined;
    const tick = setInterval(() => setElapsed(program.current?.elapsed || 0), 250);
    const watch = setInterval(async () => {
      const state = await checkDisk();
      if (!state || phaseRef.current !== 'recording') return;
      if (state.level === 'stop') {
        toast(t('The disk is nearly full: the recording was stopped and kept.'), { tone: 'bad', ms: 8000 });
        stopRef.current?.();
      } else if (state.level === 'low' && !warned.current) {
        warned.current = true;
        toast(t('About {minutes} minutes of space are left on the disk.', { minutes: Math.max(1, Math.floor(state.minutes)) }), { ms: 8000 });
      }
    }, DISK_EVERY_MS);
    return () => { clearInterval(tick); clearInterval(watch); };
  }, [phase, checkDisk, toast]);

  /* ── pictures ──────────────────────────────────────────────────────── */

  const snap = useCallback(async () => {
    setShotCount(null);
    try {
      const shot = await program.current?.snapshot(settings.shot);
      if (!shot) {
        toast(settings.shot === 'camera' ? t('The camera shows nothing to take yet.') : settings.shot === 'screen' ? t('The screen shows nothing to take yet.') : t('There is nothing to take yet.'), { tone: 'bad', ms: 5000 });
        return;
      }
      const { path } = await shell.capture.take({ dir: settings.folder, name: takeFileName(t('Picture'), new Date(), 'png') });
      await shell.fs.write({ path, bytes: shot.bytes });
      setFlash((n) => n + 1);
      setPicture((old) => {
        if (old?.url) URL.revokeObjectURL(old.url);
        return { path, width: shot.width, height: shot.height, url: URL.createObjectURL(new Blob([shot.bytes], { type: 'image/png' })) };
      });
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
    }
  }, [shell, settings.shot, settings.folder, toast]);

  const takePicture = useCallback(() => {
    if (shotCount != null) return;
    if (settings.shotTimer > 0) setShotCount(settings.shotTimer);
    else snap();
  }, [shotCount, settings.shotTimer, snap]);
  pictureRef.current = takePicture;

  useEffect(() => {
    if (shotCount == null) return undefined;
    if (shotCount <= 0) { snap(); return undefined; }
    const timer = setTimeout(() => setShotCount((n) => (n == null ? n : n - 1)), 1000);
    return () => clearTimeout(timer);
  }, [shotCount, snap]);

  useEffect(() => () => { if (picture?.url) URL.revokeObjectURL(picture.url); }, [picture]);

  const copyPicture = useCallback(async () => {
    try {
      const { bytes } = await shell.fs.read({ path: picture.path });
      await shell.clipboard.writeImage({ bytes });
      toast(t('The picture is on the clipboard.'), { tone: 'good', ms: 3000 });
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
    }
  }, [shell, picture, toast]);

  /* ── going live ────────────────────────────────────────────────────── */

  // The bitrate the slowest destination can carry: the encoder follows it, and the status bar says so.
  const [liveRate, setLiveRate] = useState(null);
  useEffect(() => shell.on('live:status', ({ destinations = [], kbps, target } = {}) => {
    setLiveStatus(destinations);
    if (kbps) { setLiveRate({ kbps, target }); encoder.current?.setBitrate(kbps); }
  }), [shell]);
  const kbps = liveBitrate(settings.size, settings.fps);

  const goLive = useCallback(async () => {
    if (livePhase !== 'off') return;
    const canvas = canvasRef.current;
    let checked;
    try {
      checked = await LiveEncoder.check({ width: canvas.width, height: canvas.height, fps: settings.fps, kbps });
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 8000 });
      return;
    }
    const dests = await shell.live.destinations().catch(() => []);
    if (!dests.some((d) => d.on && d.hasKey)) { setLiveOpen(true); return; }
    setLivePhase('starting');
    try {
      const r = await shell.live.start({ video: { width: canvas.width, height: canvas.height, fps: settings.fps, kbps }, audio: { kbps: AUDIO_KBPS, sampleRate: 48000, channels: 2 } });
      if (r.skipped?.length) toast(t('Not streamed to {names}: no stream key.', { names: r.skipped.map((x) => x.name).join(', ') }), { ms: 6000 });
      const enc = new LiveEncoder({ shell, program: program.current, rig });
      await enc.start({ video: checked.video });
      encoder.current = enc;
      setLiveSince(Date.now());
      setLivePhase('live');
      if (settings.liveCopy && phaseRef.current === 'setup') {
        liveRecord.current = true;
        setPhase('starting');
        begin();
      }
    } catch (err) {
      await encoder.current?.stop().catch(() => {});
      encoder.current = null;
      await shell.live.stop().catch(() => {});
      setLivePhase('off');
      toast(err?.message || String(err), { tone: 'bad', ms: 8000 });
    }
  }, [livePhase, settings.fps, settings.liveCopy, kbps, shell, rig, toast, begin]);

  const endLive = useCallback(async () => {
    if (livePhase !== 'live' && livePhase !== 'starting') return;
    setLivePhase('ending');
    await encoder.current?.stop().catch(() => {});
    encoder.current = null;
    await shell.live.stop().catch(() => {});
    setLivePhase('off');
    toast(t('The stream has ended.'), { ms: 4000 });
    if (liveRecord.current) {
      liveRecord.current = false;
      stopRef.current?.();
    }
  }, [livePhase, shell, toast]);

  // Gone while live: the stream is ended with the window.
  useEffect(() => () => { if (encoder.current) { encoder.current.stop().catch(() => {}); shell.live.stop().catch(() => {}); } }, [shell]);
  const [, setLiveTick] = useState(0);
  useEffect(() => {
    if (livePhase !== 'live') return undefined;
    const tick = setInterval(() => setLiveTick((n) => n + 1), 1000);
    return () => clearInterval(tick);
  }, [livePhase]);

  // Gone in the middle of a recording: the keys and the tray are let go.
  useEffect(() => () => { if (phaseRef.current === 'recording') shell.capture.controls({ on: false }).catch(() => {}); }, [shell]);

  const chooseFolder = useCallback(async () => {
    const picked = await shell.dialog.open({ title: t('Save recordings in'), directory: true, defaultPath: settings.folder });
    if (picked?.[0]) change({ folder: picked[0] });
  }, [shell, settings.folder, change]);

  const chooseLogo = useCallback(async () => {
    const picked = await shell.dialog.open({ title: t('Choose a logo'), filters: [{ name: t('Pictures'), extensions: ['png', 'jpg', 'jpeg', 'webp', 'svg', 'gif'] }] });
    if (picked?.[0]) change({ logo: picked[0] });
  }, [shell, change]);

  /* ── the window ────────────────────────────────────────────────────── */

  const recording = phase === 'recording' || phase === 'stopping';
  const live = livePhase !== 'off';
  const busy = phase !== 'setup' || live;
  const liveFor = live && liveSince ? Date.now() - liveSince : 0;
  const LIVE_WORDS = { waiting: t('Waiting'), connecting: t('Connecting…'), live: t('Live'), reconnecting: t('Reconnecting…'), error: t('Not streaming'), ended: t('Ended'), closed: t('Ended') };
  const names = SCENE_NAMES();
  const haveScreen = Boolean(rig.screen.stream);
  const canRecord = phase === 'setup' && (haveScreen || rig.camera.stream || settings.scene === 'brb') && disk?.level !== 'stop';
  const canShoot = shotCount == null && Boolean(settings.folder) && (settings.shot === 'screen' ? haveScreen : settings.shot === 'camera' ? Boolean(rig.camera.stream) : Boolean(haveScreen || rig.camera.stream || settings.scene === 'brb'));
  const SHOT_NAMES = { program: t('What is shown'), screen: t('The screen'), camera: t('The camera') };
  const micName = devices.mics.find((d) => d.id === settings.micId)?.name || t('The usual microphone');
  const minutes = disk ? Math.floor(disk.minutes) : null;
  const diskNote = disk == null ? '' : disk.level === 'stop' ? t('There is not enough space on the disk to record.')
    : disk.minutes < 120 ? t('Room for about {minutes} minutes on the disk', { minutes: Math.max(0, minutes) })
      : t('Room for about {hours} hours on the disk', { hours: Math.floor(disk.minutes / 60) });

  const ribbon = (
    <Ribbon
      tabs={[{ id: 'home', label: t('Home') }]}
      active={tab}
      onTab={setTab}
      quick={recording ? <Button icon="stop" title={t('Stop')} onClick={stop} /> : <Button icon="record" title={t('Record')} disabled={!canRecord} onClick={record} />}
    >
      <Group label={t('Record')}>
        {recording ? (
          <>
            <Button tall icon="stop" label={t('Stop')} className="st-stop" title={keys?.stop ? `${t('Stop')} (${keys.stop})` : t('Stop')} disabled={phase !== 'recording'} onClick={stop} />
            <Button tall icon={paused ? 'play' : 'pause'} label={paused ? t('Resume') : t('Pause')} className="st-pause" disabled={phase !== 'recording'} onClick={togglePause} />
          </>
        ) : (
          <Button tall primary icon="record" label={phase === 'countdown' ? t('Starting…') : t('Record')} className="st-record" disabled={!canRecord} onClick={record} />
        )}
      </Group>
      <Group label={t('Picture')}>
        <Button tall icon="picture" label={shotCount != null ? t('In {seconds} s', { seconds: shotCount }) : t('Take Picture')} className="st-shoot" title={`${t('Take Picture: {of}, saved as a PNG', { of: SHOT_NAMES[settings.shot] })}${pictureKey ? ` (${pictureKey})` : ''}`} disabled={!canShoot} onClick={takePicture} />
        <Rows>
          <Select value={settings.shot} className="st-shot-of" aria-label={t('Picture of')} onChange={(e) => change({ shot: e.target.value })}>
            {Object.entries(SHOT_NAMES).map(([k, name]) => <option key={k} value={k}>{name}</option>)}
          </Select>
          <Select value={String(settings.shotTimer)} className="st-shot-timer" aria-label={t('Picture timer')} onChange={(e) => change({ shotTimer: Number(e.target.value) })}>
            <option value="0">{t('At once')}</option>
            {[3, 5, 10].map((n) => <option key={n} value={String(n)}>{t('After {seconds} s', { seconds: n })}</option>)}
          </Select>
        </Rows>
      </Group>
      <Group label={t('Live')}>
        {live ? (
          <Button tall icon="stop" label={livePhase === 'starting' ? t('Starting…') : t('End Live')} className="st-endlive" disabled={livePhase === 'ending'} onClick={endLive} />
        ) : (
          <Button tall icon="globe" label={t('Go Live')} className="st-golive" disabled={phase === 'countdown' || phase === 'starting'} onClick={goLive} />
        )}
        <Button icon="settings" label={t('Destinations…')} className="st-dests-btn" onClick={() => setLiveOpen(true)} />
      </Group>
      <Group label={t('Scenes')}>
        {SCENES.map((s) => (
          <Button key={s} tall icon={SCENE_ICON[s]} label={names[s]} pressed={settings.scene === s} className={`st-scene-btn st-scene-${s}`} onClick={() => change({ scene: s })} />
        ))}
      </Group>
      <Group label={t('Overlays')}>
        <Rows>
          <Button icon="textbox" label={t('Title')} pressed={settings.titleOn} className="st-title-btn" onClick={() => change({ titleOn: !settings.titleOn })} />
          <Button icon="picture" label={settings.logo ? t('Change Logo…') : t('Logo…')} onClick={chooseLogo} />
          {settings.logo ? <Button icon="close" label={t('Remove Logo')} onClick={() => change({ logo: '' })} /> : null}
        </Rows>
      </Group>
      <Group label={t('Output')}>
        <Rows>
          <Select value={String(settings.size)} disabled={busy} aria-label={t('Size')} onChange={(e) => change({ size: Number(e.target.value) })}>
            <option value="1080">1080p (1920 × 1080)</option>
            <option value="720">720p (1280 × 720)</option>
          </Select>
          <Select value={String(settings.fps)} disabled={busy} aria-label={t('Frame rate')} onChange={(e) => change({ fps: Number(e.target.value) })}>
            {[24, 30, 60].map((n) => <option key={n} value={String(n)}>{t('{n} frames a second', { n })}</option>)}
          </Select>
          <Select value={settings.quality} disabled={busy} aria-label={t('Quality')} onChange={(e) => change({ quality: e.target.value })}>
            <option value="small">{t('Smaller files')}</option>
            <option value="standard">{t('Standard')}</option>
            <option value="high">{t('Sharper')}</option>
          </Select>
        </Rows>
        <Button tall icon="folderOpen" label={t('Folder…')} title={settings.folder} disabled={busy} onClick={chooseFolder} />
      </Group>
    </Ribbon>
  );

  const status = (
    <span className="st-status" data-phase={phase}>
      {live ? `${t('Live')} ${clock(liveFor)} · ` : ''}
      {live && liveRate?.kbps ? `${t('{n} kb/s', { n: liveRate.kbps })}${liveRate.kbps < liveRate.target ? ` ${t('(lowered to keep up with the upload)')}` : ''} · ` : ''}
      {recording ? `${paused ? t('Paused') : t('Recording')} · ${clock(elapsed)}` : saved ? t('Saved {name}', { name: basename(saved.path) }) : t('Ready')}
      {` · ${settings.size === 720 ? '1280 × 720' : '1920 × 1080'} · ${t('{n} frames a second', { n: settings.fps })}`}
      {diskNote && !recording ? ` · ${diskNote}` : ''}
    </span>
  );

  return (
    <AppFrame app={app} shell={shell} title={live ? `${t('Live')} ${clock(liveFor)}` : recording ? `${paused ? t('Paused') : t('Recording')} ${clock(elapsed)}` : t('Recorder')} menu={appMenu} ribbon={ribbon} status={status}>
      <style>{RECORDER_CSS}</style>
      <style>{STUDIO_CSS}</style>
      <style>{LIVE_CSS}</style>
      {liveOpen ? <LiveDialog shell={shell} toast={toast} kbps={kbps} busy={live} recordCopy={settings.liveCopy} onRecordCopy={(v) => change({ liveCopy: v })} onClose={() => setLiveOpen(false)} /> : null}
      <Panel name="scenes" width={210} title={t('Scenes')}>
        <div className="st-scenes" role="radiogroup" aria-label={t('Scenes')}>
          {SCENES.map((s) => (
            <button key={s} type="button" role="radio" aria-checked={settings.scene === s} className={`st-scene${settings.scene === s ? ' on' : ''}`} data-scene={s} onClick={() => change({ scene: s })}>
              <span className="st-scene-mark"><Icon name={SCENE_ICON[s]} size={18} /></span>
              <span>{names[s]}</span>
              {recording && settings.scene === s ? <i className="st-live-dot" aria-hidden="true" /> : null}
            </button>
          ))}
        </div>
        <p className="st-hint">{t('Change the scene at any time, while it records too.')}</p>
      </Panel>

      <Content>
        <div className="st-stage">
          <div ref={frameRef} className={`st-frame${recording ? ' live' : ''}${paused ? ' held' : ''}${kind === 'region' && phase === 'setup' ? ' drawing' : ''}`} data-scene={settings.scene} onPointerDown={drawRegion}>
            <canvas ref={canvasRef} className="st-canvas" width={1920} height={1080} aria-label={t('Program')} />
            {recording ? <span className="rec-badge st-badge"><i />{paused ? t('Paused') : t('Recording')} {clock(elapsed)}</span> : null}
            {live ? <span className="st-live-badge" data-phase={livePhase}><i />{t('LIVE')} {clock(liveFor)}</span> : null}
            {phase === 'countdown' ? <div className="rec-count" aria-live="assertive"><span key={count}>{count}</span></div> : null}
            {shotCount != null && shotCount > 0 && phase !== 'countdown' ? <div className="rec-count st-shot-count" aria-live="assertive"><span key={shotCount}>{shotCount}</span></div> : null}
            {flash ? <div key={flash} className="st-flash" aria-hidden="true" /> : null}
            {drawing ? (
              <div className="rec-region st-drawn" style={{ left: `${Math.min(drawing.a.fx, drawing.b.fx) * 100}%`, top: `${Math.min(drawing.a.fy, drawing.b.fy) * 100}%`, width: `${Math.abs(drawing.b.fx - drawing.a.fx) * 100}%`, height: `${Math.abs(drawing.b.fy - drawing.a.fy) * 100}%` }}><span>{t('Region')}</span></div>
            ) : null}
            {kind === 'region' && haveScreen && phase === 'setup' && !drawing ? (
              <div className="rec-hint">{crop ? t('Drag on the picture again to choose a smaller part.') : t('Drag on the picture to choose the part to record.')}</div>
            ) : null}
            {needsScreen && !haveScreen && phase === 'setup' ? (
              <div className="st-wait">{rig.screen.error ? rig.screen.error : sources == null ? <><Spinner /> {t('Looking at what is open…')}</> : t('Choose a screen or a window to record.')}</div>
            ) : null}
          </div>
          {kind === 'region' && crop ? (
            <div className="st-region-bar">
              <span>{t('Region {w} × {h}', { w: Math.round(crop.w * (rig.screen.stream?.getVideoTracks()[0]?.getSettings?.().width || 0)), h: Math.round(crop.h * (rig.screen.stream?.getVideoTracks()[0]?.getSettings?.().height || 0)) })}</span>
              {phase === 'setup' ? <Button ghost icon="close" label={t('Whole picture')} className="st-whole" onClick={() => setCrop(null)} /> : null}
            </div>
          ) : null}
          {live && liveStatus.length ? (
            <div className="st-dests" aria-live="polite">
              {liveStatus.map((d) => (
                <span key={d.id} className="st-dest" data-state={d.state} title={d.detail || ''}>
                  <i />{d.name}: {d.state === 'error' && d.detail ? d.detail : LIVE_WORDS[d.state] || d.state}{d.reconnects && d.state === 'reconnecting' ? ` (${d.reconnects})` : ''}
                </span>
              ))}
            </div>
          ) : null}
          {picture ? (
            <div className="st-shot" role="status">
              <img src={picture.url} alt="" />
              <span className="st-shot-name">{basename(picture.path)}</span>
              <span className="st-shot-meta">{picture.width} × {picture.height}</span>
              <Button icon="picture" label={t('Open')} className="st-shot-open" onClick={() => shell.win.create({ app: 'pictures', file: picture.path })} />
              <Button icon="image" label={t('Edit in Image')} className="st-shot-edit" onClick={() => shell.win.create({ app: 'image', file: picture.path })} />
              <Button icon="copy" label={t('Copy')} className="st-shot-copy" onClick={copyPicture} />
              <Button icon="folderOpen" label={t('Show in folder')} onClick={() => shell.shell.showInFolder({ path: picture.path })} />
            </div>
          ) : null}
          {saved && phase === 'setup' ? (
            <div className="st-saved" role="status">
              <Icon name="check" size={18} />
              <span className="st-saved-name">{basename(saved.path)}</span>
              <span className="st-saved-meta">{clock(saved.seconds * 1000)} · {formatBytes(saved.written)}</span>
              <Button icon="play" label={t('Play')} className="st-play" onClick={() => shell.win.create({ app: 'video', file: saved.path })} />
              <Button icon="video" label={t('Edit in Video')} className="st-edit" onClick={() => shell.win.create({ app: 'video', file: saved.path, query: { edit: 1 } })} />
              <Button icon="folderOpen" label={t('Show in folder')} onClick={() => shell.shell.showInFolder({ path: saved.path })} />
            </div>
          ) : null}
        </div>
      </Content>

      <Panel name="sources" width={300} title={t('Sources')}>
        <div className="st-side" aria-disabled={busy}>
          <section className="rec-sec">
            <h3>{t('Screen')}</h3>
            <Segments value={kind} disabled={busy} className="st-kinds" onChange={setKind} options={[{ value: 'screen', label: t('Screen'), icon: 'screen' }, { value: 'window', label: t('Window'), icon: 'window' }, { value: 'region', label: t('Region'), icon: 'region', tip: t('Region: a part you draw, kept as a crop you can change later') }]} />
            <div className="rec-sources st-sources">
              {sources && !listed.length ? <div className="rec-wait">{kind === 'window' ? t('No other window is open to record.') : t('No screen can be recorded.')}</div> : null}
              {listed.map((s) => (
                <button key={s.id} type="button" className={`rec-source${s.id === sourceId ? ' on' : ''}`} data-source={s.id} data-tip={s.name} disabled={busy} onClick={() => setSourceId(s.id)}>
                  <img src={s.thumbnail} alt="" draggable={false} />
                  <span>{s.name}</span>
                </button>
              ))}
            </div>
            <Button ghost icon="refresh" label={t('Refresh the list')} disabled={busy} onClick={loadSources} />
          </section>

          <section className="rec-sec">
            <Toggle on={settings.camera} icon="camera" label={t('Camera')} onChange={(v) => change({ camera: v })} className="st-camera-toggle" />
            {settings.camera ? (
              <div className="rec-sub">
                {rig.camera.error ? <small className="rec-err">{rig.camera.error}</small> : null}
                {devices.cameras.length > 1 ? (
                  <Select value={settings.cameraId} aria-label={t('Camera')} onChange={(e) => change({ cameraId: e.target.value })}>
                    <option value="">{t('The usual camera')}</option>
                    {devices.cameras.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}
                  </Select>
                ) : null}
                <Segments value={settings.shape} onChange={(v) => change({ shape: v })} options={[{ value: 'circle', label: t('Circle') }, { value: 'rounded', label: t('Rounded') }, { value: 'rect', label: t('Rectangle') }]} />
                <Segments value={settings.insetSize} onChange={(v) => change({ insetSize: v })} options={[{ value: 'small', label: t('Small') }, { value: 'medium', label: t('Medium') }, { value: 'large', label: t('Large') }]} />
                <Segments value={settings.corner} className="st-corners" onChange={(v) => change({ corner: v })} options={[{ value: 'tl', label: '↖', tip: t('Top left') }, { value: 'tr', label: '↗', tip: t('Top right') }, { value: 'bl', label: '↙', tip: t('Bottom left') }, { value: 'br', label: '↘', tip: t('Bottom right') }]} />
                <Toggle on={settings.mirror} label={t('Mirror the camera')} onChange={(v) => change({ mirror: v })} />
              </div>
            ) : null}
          </section>

          <section className="rec-sec">
            <Toggle on={settings.mic} icon="mic" label={t('Microphone')} disabled={busy} onChange={(v) => change({ mic: v })} className="st-mic-toggle" />
            {settings.mic ? (
              <div className="rec-sub">
                {devices.mics.length > 1 ? (
                  <Select value={settings.micId} disabled={busy} aria-label={t('Microphone')} onChange={(e) => change({ micId: e.target.value })}>
                    <option value="">{t('The usual microphone')}</option>
                    {devices.mics.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}
                  </Select>
                ) : <small className="st-device">{micName}</small>}
                {rig.mic.error ? <small className="rec-err">{rig.mic.error}</small> : <Meter rig={rig} name="mic" active={Boolean(rig.mic.stream)} />}
              </div>
            ) : null}
            <Toggle on={settings.system} icon="speaker" label={t('Computer sound')} disabled={busy} onChange={(v) => change({ system: v })} className="st-system-toggle" />
            {settings.system ? (
              <div className="rec-sub">
                {rig.system.error ? <small className="rec-err">{rig.system.error}</small> : <Meter rig={rig} name="system" active={Boolean(rig.nodes.system)} />}
              </div>
            ) : null}
          </section>

          <section className="rec-sec">
            <Toggle on={settings.titleOn} icon="textbox" label={t('Title')} onChange={(v) => change({ titleOn: v })} className="st-title-toggle" />
            {settings.titleOn ? (
              <div className="rec-sub">
                <Input className="st-title" value={settings.title} placeholder={t('Your name, or the title')} aria-label={t('Title')} onChange={(e) => change({ title: e.target.value })} />
                <Input className="st-subtitle" value={settings.subtitle} placeholder={t('A line under it')} aria-label={t('A line under it')} onChange={(e) => change({ subtitle: e.target.value })} />
              </div>
            ) : null}
            <label className="rec-field st-brb-field">
              <span>{t('Be right back says')}</span>
              <Input className="st-brb" value={settings.brb} aria-label={t('Be right back says')} onChange={(e) => change({ brb: e.target.value })} />
            </label>
            {settings.logo ? (
              <div className="st-logo">
                {logoUrl ? <img src={logoUrl} alt="" /> : null}
                <span>{basename(settings.logo)}</span>
                <Segments value={settings.logoCorner} onChange={(v) => change({ logoCorner: v })} options={[{ value: 'tl', label: '↖', tip: t('Top left') }, { value: 'tr', label: '↗', tip: t('Top right') }, { value: 'bl', label: '↙', tip: t('Bottom left') }, { value: 'br', label: '↘', tip: t('Bottom right') }]} />
              </div>
            ) : null}
          </section>

          <section className="rec-sec">
            <label className="rec-field">
              <span>{t('Countdown')}</span>
              <Select value={String(settings.countdown)} disabled={busy} onChange={(e) => change({ countdown: Number(e.target.value) })}>
                <option value="0">{t('None')}</option>
                {[3, 5, 10].map((n) => <option key={n} value={String(n)}>{t('{seconds} s', { seconds: n })}</option>)}
              </Select>
            </label>
            <Toggle on={settings.minimise} label={t('Minimise this window while recording')} disabled={busy} onChange={(v) => change({ minimise: v })} />
            <small className="st-folder" title={settings.folder}>{t('Saved in {folder}', { folder: settings.folder || '…' })}</small>
          </section>
        </div>
      </Panel>
    </AppFrame>
  );
}
