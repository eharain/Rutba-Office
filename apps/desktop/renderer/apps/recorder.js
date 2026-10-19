// Rutba Recorder: record the screen, the camera and the voice as one finished
// video, made live from scenes (owner, 2026-10-19: "split the recording tool
// out to be able to record video and stream videos live").
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
import { takeFileName, bytesPerSecond, diskState } from '@rutba/studio/recorder';
import { AppFrame, useAppMenu } from '../shell.js';
import { RecorderRig } from './video/recorder/session.js';
import { Toggle, Segments, Meter } from './video/recorder/panel.js';
import { RECORDER_CSS } from './video/recorder/css.js';
import { Program } from './recorder/program.js';
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
  const [sourceId, setSourceId] = useState(null);
  const [devices, setDevices] = useState({ cameras: [], mics: [] });
  const [phase, setPhase] = useState('setup'); // setup | countdown | recording | stopping
  const [count, setCount] = useState(0);
  const [elapsed, setElapsed] = useState(0);
  const [paused, setPaused] = useState(false);
  const [disk, setDisk] = useState(null);
  const [keys, setKeys] = useState(null);
  const [saved, setSaved] = useState(null); // { path, written, seconds }
  const [logoUrl, setLogoUrl] = useState(null);
  const [, setTick] = useState(0);
  const [tab, setTab] = useState('home');
  const canvasRef = useRef(null);
  const phaseRef = useRef(phase);
  phaseRef.current = phase;
  const stopRef = useRef(null);
  const pauseRef = useRef(null);
  const warned = useRef(false);
  const rig = useMemo(() => new RecorderRig({ shell, onChange: () => setTick((n) => n + 1) }), [shell]);
  const program = useRef(null);

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
  const listed = useMemo(() => (sources || []).filter((s) => s.kind === kind), [sources, kind]);
  useEffect(() => {
    if (!sources) return;
    if (!listed.some((s) => s.id === sourceId)) setSourceId(listed[0]?.id || null);
  }, [listed, sources, sourceId]);
  const source = listed.find((s) => s.id === sourceId) || null;

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

  useEffect(() => shell.on('capture:command', ({ command } = {}) => {
    if (command === 'stop') stopRef.current?.();
    else if (command === 'pause') pauseRef.current?.();
  }), [shell]);

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
  const busy = phase !== 'setup';
  const names = SCENE_NAMES();
  const haveScreen = Boolean(rig.screen.stream);
  const canRecord = phase === 'setup' && (haveScreen || rig.camera.stream || settings.scene === 'brb') && disk?.level !== 'stop';
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
      {recording ? `${paused ? t('Paused') : t('Recording')} · ${clock(elapsed)}` : saved ? t('Saved {name}', { name: basename(saved.path) }) : t('Ready')}
      {` · ${settings.size === 720 ? '1280 × 720' : '1920 × 1080'} · ${t('{n} frames a second', { n: settings.fps })}`}
      {diskNote && !recording ? ` · ${diskNote}` : ''}
    </span>
  );

  return (
    <AppFrame app={app} shell={shell} title={recording ? `${paused ? t('Paused') : t('Recording')} ${clock(elapsed)}` : t('Recorder')} menu={appMenu} ribbon={ribbon} status={status}>
      <style>{RECORDER_CSS}</style>
      <style>{STUDIO_CSS}</style>
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
          <div className={`st-frame${recording ? ' live' : ''}${paused ? ' held' : ''}`} data-scene={settings.scene}>
            <canvas ref={canvasRef} className="st-canvas" width={1920} height={1080} aria-label={t('Program')} />
            {recording ? <span className="rec-badge st-badge"><i />{paused ? t('Paused') : t('Recording')} {clock(elapsed)}</span> : null}
            {phase === 'countdown' ? <div className="rec-count" aria-live="assertive"><span key={count}>{count}</span></div> : null}
            {needsScreen && !haveScreen && phase === 'setup' ? (
              <div className="st-wait">{rig.screen.error ? rig.screen.error : sources == null ? <><Spinner /> {t('Looking at what is open…')}</> : t('Choose a screen or a window to record.')}</div>
            ) : null}
          </div>
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
            <Segments value={kind} disabled={busy} className="st-kinds" onChange={setKind} options={[{ value: 'screen', label: t('Screen'), icon: 'screen' }, { value: 'window', label: t('Window'), icon: 'window' }]} />
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
