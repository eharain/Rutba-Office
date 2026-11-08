// The studio window both the Recorder and Live Broadcast are: the screen,
// windows, cameras and sound made into one picture (owner, 2026-10-19: "split
// the recording tool out to be able to record video and stream videos live";
// 2026-10-20: "there should be a way to capture pictures on camera and from
// screens etc.").
//
// One studio in which every source works the same way (owner, 2026-11-02,
// D-2026-11-02-13): "the screen and cameras all should work exactly same way
// so no separate features and functions on top. i can clip screen i should be
// able to clip camera."
//
// The Recorder records to a file; Live Broadcast goes live, a product of its
// own (D-2026-11-02-12, D-2026-11-03-01: "my problem is recorder dose not say
// it clearly its purpose from name it feels like a recording app"), with the
// shots of several cameras switched by hand, by who is speaking, or on a
// timer, and a preview of the next shot beside the one on air
// (D-2026-11-03-02).
//
// The first line holds Take Picture, Go Live (Live Broadcast's) and Record
// side by side, the last two toggles that show they are on with their time,
// then the Video switch (video, no video, placeholder) and the Sound switch
// (sound, no sound, music). In the middle is the share, with the selected
// source's bar over it and the scenes under it. On the right, one list of
// every source: on the share top first, not shown, and the sounds.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Panel, Content, Menu, Select, useToast, formatBytes, basename, t } from '@rutba/office-ui';
import { normaliseStudio, STUDIO_DEFAULTS } from '@rutba/studio/compositor';
import {
  normaliseStack, emptyStack, stackFromSettings, addItem, removeItem, duplicateItem, moveItem, showItem, placeItem, shapeItem, clipItem,
  updateItem, updateSound, frameToSource, itemAt, saveScene, removeScene, applyScene, builtInScenes, recordKind, clampBox, newId,
  followStep, nextInRotation, FADE_MS, ROTATE_EVERY,
} from '@rutba/studio/stack';
import { takeFileName, bytesPerSecond, diskState, regionCrop } from '@rutba/studio/recorder';
import { AppFrame, useAppMenu } from '../shell.js';
import { RECORDER_CSS } from '../recording/css.js';
import { Program } from './program.js';
import { StudioMedia, sourceKey } from './media.js';
import { SourceList, itemName } from './list.js';
import { SourceBar } from './bar.js';
import { StudioSettings } from './settings-dialog.js';
import { STUDIO_CSS } from './css.js';

const DISK_EVERY_MS = 5000;
const SOUND_EXT = ['webm', 'weba', 'ogg', 'mp4', 'mov', 'mkv', 'm4a', 'mp3', 'wav'];
// Follow the speaker: how loud a microphone is when someone speaks into it, and how long a pause between words still counts as speaking.
const SPEAK_LEVEL = 0.06;
const SPEAK_HANG_MS = 450;

const clock = (ms) => {
  const s = Math.floor(ms / 1000);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const ss = String(s % 60).padStart(2, '0');
  return h ? `${h}:${String(m).padStart(2, '0')}:${ss}` : `${m}:${ss}`;
};

const SCENE_NAMES = () => ({ 'screen-camera': t('Screen and camera'), screen: t('Screen'), camera: t('Camera'), split: t('Side by side'), wide: t('Everyone'), cameraN: (i) => t('Camera {n}', { n: i + 1 }) });

/** A source's picture aspect, from what is open or what the capture list says. */
function aspectOf(media, item, sources) {
  const size = media.sizeOf(item);
  if (size) return size.w / size.h;
  const d = sources?.find((s) => s.id === item.ref)?.display;
  return d?.width && d?.height ? d.width / d.height : 16 / 9;
}

/**
 * @param {object} p
 * @param {'recorder'|'broadcast'} p.kind  which app this window is
 * @param {object} [p.live]  Live Broadcast's going live: { LiveEncoder, AUDIO_KBPS, LiveDialog, LIVE_CSS, liveBitrate }
 */
export function StudioWindow({ app, shell, kind = 'recorder', live: liveKit = null }) {
  const isLive = kind === 'broadcast' && Boolean(liveKit);
  const STORE_KEY = `${kind}.studio`;
  const STACK_KEY = `${kind}.stack`;
  const appName = isLive ? t('Live Broadcast') : t('Recorder');
  const toast = useToast();
  const [settings, setSettings] = useState(() => normaliseStudio(null));
  const [stack, setStack] = useState(() => emptyStack());
  const [loaded, setLoaded] = useState(false);
  const [sources, setSources] = useState(null);
  const [devices, setDevices] = useState({ cameras: [], mics: [] });
  const [selected, setSelected] = useState(null);
  const [clipping, setClipping] = useState(false);
  const [drawing, setDrawing] = useState(null);
  const [phase, setPhase] = useState('setup'); // setup | countdown | starting | recording | stopping
  const [count, setCount] = useState(0);
  const [elapsed, setElapsed] = useState(0);
  const [paused, setPaused] = useState(false);
  const [disk, setDisk] = useState(null);
  const [keys, setKeys] = useState(null);
  const [saved, setSaved] = useState(null);
  const [picture, setPicture] = useState(null);
  const [shotCount, setShotCount] = useState(null);
  const [flash, setFlash] = useState(0);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [sceneMenu, setSceneMenu] = useState(null);
  const [lastScene, setLastScene] = useState(null);
  const [recent, setRecent] = useState([]);
  const [, setTick] = useState(0);
  // Live Broadcast's switching: the next shot shown beside the one on air, and the shot on air;
  // Follow the speaker or the timed rotation (off, follow or rotate).
  const [preview, setPreview] = useState(false);
  const [air, setAir] = useState(null);
  const [onAir, setOnAir] = useState(null);
  const [auto, setAuto] = useState('off');
  // Go Live: off, starting, live or ending; the destinations as the main process reports them.
  const [livePhase, setLivePhase] = useState('off');
  const [liveStatus, setLiveStatus] = useState([]);
  const [liveSince, setLiveSince] = useState(0);
  const [liveOpen, setLiveOpen] = useState(false);
  const encoder = useRef(null);
  const liveRecord = useRef(false);
  const canvasRef = useRef(null);
  const previewCanvasRef = useRef(null);
  const airFrameRef = useRef(null);
  const nextFrameRef = useRef(null);
  const phaseRef = useRef(phase);
  phaseRef.current = phase;
  const stackRef = useRef(stack);
  stackRef.current = stack;
  const settingsRef = useRef(settings);
  settingsRef.current = settings;
  const onAirRef = useRef(onAir);
  onAirRef.current = onAir;
  const stopRef = useRef(null);
  const pauseRef = useRef(null);
  const warned = useRef(false);
  const fadeNext = useRef(0);
  const media = useMemo(() => new StudioMedia({ shell, onChange: () => setTick((n) => n + 1) }), [shell]);
  const program = useRef(null);
  const previewProgram = useRef(null);
  // The AI tools' requests (capture:command 'remote', the Recorder's): the latest handler, the recording they asked for, and the take's path.
  const remoteRef = useRef(null);
  const remote = useRef(null);
  const takenRef = useRef(null);
  // The frame the person works on: the next shot's while the preview is up, the one on air otherwise.
  const editFrame = () => (preview ? nextFrameRef : airFrameRef).current;

  /* ── the stack and the settings, kept ──────────────────────────────── */

  const change = useCallback((patch) => {
    setSettings((s) => {
      const next = normaliseStudio({ ...s, ...patch });
      shell.store.set({ key: STORE_KEY, value: next }).catch(() => {});
      return next;
    });
  }, [shell, STORE_KEY]);

  /** The stack changed by `fn`; kept unless `keep` is false (a drag keeps it once, at its end). */
  const changeStack = useCallback((fn, keep = true) => {
    setStack((s) => {
      const next = normaliseStack(fn(s));
      if (keep) shell.store.set({ key: STACK_KEY, value: next }).catch(() => {});
      return next;
    });
  }, [shell, STACK_KEY]);

  // The settings and the stack this person last had, and where recordings go. Live Broadcast,
  // opened the first time, starts from the Recorder's sources and settings.
  const storedRef = useRef(null);
  useEffect(() => {
    let alive = true;
    (async () => {
      const get = (key) => shell.store.get({ key, fallback: null }).catch(() => null);
      let stored = await get(STORE_KEY);
      let kept = await get(STACK_KEY);
      if (kind === 'broadcast' && !stored && !kept) {
        stored = await get('recorder.studio');
        kept = await get('recorder.stack');
      }
      const paths = await shell.app.paths().catch(() => null);
      if (!alive) return;
      const s = normaliseStudio(stored);
      const sep = paths?.separator || '\\';
      if (!s.folder && paths?.videos) s.folder = `${paths.videos}${sep}${t('Rutba Recordings')}`;
      setSettings(s);
      storedRef.current = { stored, kept };
      if (kept) setStack(normaliseStack(kept));
      setLoaded(true);
    })();
    return () => { alive = false; };
  }, [shell, kind, STORE_KEY, STACK_KEY]);

  // What there is to show: the screens and windows, the cameras and microphones.
  const loadSources = useCallback(async () => {
    try { setSources((await shell.capture.sources()) || []); } catch { setSources([]); }
    setDevices(await StudioMedia.devices());
  }, [shell]);
  useEffect(() => { loadSources(); }, [loadSources]);

  // A first stack, once the screens are known: from the Recorder's settings before it was a stack.
  useEffect(() => {
    if (!loaded || !sources || storedRef.current?.kept || storedRef.current?.made) return;
    storedRef.current = { ...storedRef.current, made: true };
    const screen = sources.find((x) => x.kind === 'screen') || sources[0] || null;
    const raw = storedRef.current?.stored;
    const first = stackFromSettings({ ...(raw || {}), brb: raw?.brb && raw.brb !== STUDIO_DEFAULTS.brb ? raw.brb : t('Back soon') }, { screen });
    changeStack(() => first);
  }, [loaded, sources, changeStack]);

  // The programs, drawn from the start, and let go with the window: the one on air, and the next shot's.
  useEffect(() => {
    const p = new Program({ media, canvas: canvasRef.current });
    program.current = p;
    p.onTick = () => media.revive(stackRef.current, { fps: settingsRef.current.fps, size: settingsRef.current.size });
    p.set(settings);
    p.start();
    const next = new Program({ media, canvas: previewCanvasRef.current });
    previewProgram.current = next;
    next.set(settings);
    return () => { p.dispose(); next.dispose(); media.dispose(); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [media]);
  useEffect(() => { program.current?.set(settings); previewProgram.current?.set(settings); }, [settings]);
  const airStack = preview && air ? air : stack;
  useEffect(() => {
    program.current?.setStack(airStack, { fade: fadeNext.current });
    fadeNext.current = 0;
  }, [airStack]);
  useEffect(() => {
    const next = previewProgram.current;
    if (!next) return;
    next.setStack(stack);
    if (preview) next.start();
    else next.stopDrawing();
  }, [stack, preview]);

  // The sources open as the stacks say (the next shot's and the one on air); the sound as its sounds and the Sound switch say.
  const syncItems = preview && air ? [...stack.items, ...air.items.filter((it) => !stack.items.some((x) => x.id === it.id))] : stack.items;
  useEffect(() => {
    if (!loaded || !sources) return;
    media.sync({ ...stack, items: syncItems }, sources, { fps: settings.fps, size: settings.size });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [media, loaded, sources, stack, air, preview, settings.fps, settings.size]);
  const soundKey = JSON.stringify([stack.sounds, stack.sound, stack.music]);
  useEffect(() => {
    if (!loaded) return;
    media.syncSounds(stackRef.current).then(() => StudioMedia.devices().then(setDevices));
  }, [media, loaded, soundKey]);

  /* ── the list's actions ────────────────────────────────────────────── */

  const thumbs = useMemo(() => Object.fromEntries((sources || []).map((s) => [s.id, s.thumbnail])), [sources]);
  const firstCamera = devices.cameras[0]?.id || '';
  const firstMic = devices.mics[0]?.id || '';
  const spare = useMemo(() => {
    const has = (kind2, ref) => stack.items.some((it) => (kind2 === 'camera' ? it.kind === 'camera' && (it.ref || firstCamera) === ref : it.ref === ref));
    const out = [];
    for (const s of sources || []) if ((s.kind === 'screen' || s.kind === 'window') && !has(s.kind, s.id)) out.push({ kind: s.kind, ref: s.id, name: s.name, thumb: s.thumbnail, aspect: s.display?.width && s.display?.height ? s.display.width / s.display.height : 16 / 9 });
    for (const c of devices.cameras) if (!has('camera', c.id)) out.push({ kind: 'camera', ref: c.id, name: c.name, thumb: null, aspect: 16 / 9 });
    // The cameras first: there are few of them, and many windows.
    return [...out.filter((x) => x.kind === 'camera'), ...out.filter((x) => x.kind === 'screen'), ...out.filter((x) => x.kind === 'window')];
  }, [sources, devices, stack.items, firstCamera]);
  const spareSounds = useMemo(() => devices.mics.filter((m) => !stack.sounds.some((s) => s.kind === 'mic' && (s.ref || firstMic) === m.id)).map((m) => ({ kind: 'mic', ref: m.id, name: m.name })), [devices, stack.sounds, firstMic]);
  // The sounds' names, from the devices.
  const named = useMemo(() => ({
    ...stack,
    sounds: stack.sounds.map((s) => ({ ...s, name: s.name || (s.kind === 'system' ? t('Computer sound') : devices.mics.find((m) => m.id === (s.ref || firstMic))?.name || t('Microphone')) })),
    items: stack.items.map((it) => (it.kind === 'camera' && !it.name ? { ...it, name: devices.cameras.find((c) => c.id === (it.ref || firstCamera))?.name || '' } : it)),
  }), [stack, devices, firstMic, firstCamera]);

  const selItem = named.items.find((it) => it.id === selected) || null;
  const selSound = named.sounds.find((s) => s.id === selected) || null;
  useEffect(() => { if (!selItem) setClipping(false); }, [selItem]);

  const choosePicture = async (id) => {
    const picked = await shell.dialog.open({ title: t('Choose a picture'), filters: [{ name: t('Pictures'), extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp'] }] });
    if (!picked?.[0]) return;
    if (id) { changeStack((s) => updateItem(s, id, { path: picked[0] })); return; }
    let added;
    changeStack((s) => { const r = addItem(s, { kind: 'picture', path: picked[0] }, { aspect: 1 }); added = r.id; return r.stack; });
    setTimeout(() => setSelected(added), 0);
  };

  const barAct = {
    live: (on) => selItem && act.live(selItem.id, on),
    clip: () => setClipping((v) => !v),
    whole: () => selItem && changeStack((s) => clipItem(s, selItem.id, null, { aspect: aspectOf(media, selItem, sources) })),
    shape: (shape) => selItem && changeStack((s) => shapeItem(s, selItem.id, shape, { aspect: aspectOf(media, selItem, sources) })),
    place: (place) => selItem && changeStack((s) => placeItem(s, selItem.id, place, { aspect: aspectOf(media, selItem, sources) })),
    mirror: () => selItem && changeStack((s) => updateItem(s, selItem.id, { mirror: !selItem.mirror })),
    text: (patch) => selItem && changeStack((s) => updateItem(s, selItem.id, patch)),
    change: () => selItem && choosePicture(selItem.id),
    copy: () => { if (!selItem) return; let id; changeStack((s) => { const r = duplicateItem(s, selItem.id); id = r.id; return r.stack; }); setTimeout(() => setSelected(id), 0); },
    remove: () => { if (!selItem) return; changeStack((s) => removeItem(s, selItem.id)); setSelected(null); },
    sound: (patch) => selSound && act.sound(selSound.id, patch),
  };

  const act = {
    show: (id, on) => changeStack((s) => {
      let next = showItem(s, id, on);
      if (on) next = updateItem(next, id, { live: true });
      return next;
    }),
    live: (id, on) => changeStack((s) => {
      const it = s.items.find((x) => x.id === id);
      if (!it) return s;
      // Every copy of a source starts and stops with it: they are one capture.
      return { ...s, items: s.items.map((x) => (sourceKey(x) === sourceKey(it) ? { ...x, live: on } : x)) };
    }),
    // `at` is the gap to drop into, as a drop gives it: below the item itself is the same place.
    move: (id, at) => changeStack((s) => {
      const from = s.items.findIndex((x) => x.id === id);
      return moveItem(s, id, from >= 0 && from < at ? at - 1 : at);
    }),
    addSpare: (sp, { on = true, live = true, at = null, select = false } = {}) => {
      let added = null;
      changeStack((s) => {
        const r = addItem(s, { kind: sp.kind, ref: sp.ref, name: sp.name, on, live, shape: 'rect' }, { aspect: sp.aspect });
        added = r.id;
        return at == null ? r.stack : moveItem(r.stack, r.id, at);
      });
      if (select) setTimeout(() => { if (added) setSelected(added); }, 0);
    },
    sound: (id, patch) => changeStack((s) => updateSound(s, id, patch)),
    addSound: (sp) => changeStack((s) => ({ ...s, sounds: [...s.sounds, { id: newId('mic', s.sounds.map((x) => x.id)), kind: 'mic', ref: sp.ref, name: '', live: true, muted: false, gain: 1 }] })),
    addItems: () => [
      { label: t('Title'), icon: 'textbox', run: () => { let id; changeStack((s) => { const r = addItem(s, { kind: 'title', text: t('Your name, or the title') }); id = r.id; return r.stack; }); setTimeout(() => setSelected(id), 0); } },
      { label: t('Picture or Logo…'), icon: 'picture', run: () => choosePicture(null) },
      ...(selItem && ['screen', 'window', 'camera'].includes(selItem.kind) ? [{ label: t('Another copy of {name}', { name: itemName(selItem) }), icon: 'duplicate', run: () => barAct.copy() }] : []),
      '-',
      { label: t('Refresh the list'), icon: 'refresh', run: loadSources },
    ],
  };

  /* ── the share: select, move, resize, clip ─────────────────────────── */

  const frameAt = (e) => {
    const r = editFrame().getBoundingClientRect();
    return { fx: (e.clientX - r.left) / r.width, fy: (e.clientY - r.top) / r.height };
  };
  const snap = (v, size) => {
    for (const edge of [0, 1 - size, 0.5 - size / 2]) if (Math.abs(v - edge) < 0.012) return edge;
    return v;
  };
  const onFramePointer = (e) => {
    if (e.button !== 0 || !editFrame()) return;
    const p = frameAt(e);
    const handle = e.target?.dataset?.c;
    const s = stackRef.current;
    // Clip: a part of the selected source dragged out over its box.
    if (clipping && selItem && !handle) {
      e.preventDefault();
      const it = s.items.find((x) => x.id === selItem.id);
      const b = it.box;
      if (p.fx < b.x || p.fy < b.y || p.fx > b.x + b.w || p.fy > b.y + b.h) { setClipping(false); return; }
      setDrawing({ a: p, b: p });
      const move = (ev) => setDrawing({ a: p, b: frameAt(ev) });
      const up = (ev) => {
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', up);
        setDrawing(null);
        const q = frameAt(ev);
        const fit = { aspect: aspectOf(media, it, sources) };
        const from = frameToSource(it, p.fx, p.fy, fit);
        const to = frameToSource(it, q.fx, q.fy, fit);
        const crop = regionCrop({ x: Math.min(from.x, to.x), y: Math.min(from.y, to.y), w: Math.abs(to.x - from.x), h: Math.abs(to.y - from.y) }, { width: 1, height: 1 });
        if (crop) changeStack((st) => clipItem(st, it.id, crop, { aspect: aspectOf(media, it, sources) }));
        setClipping(false);
      };
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', up);
      return;
    }
    const hit = handle ? s.items.find((x) => x.id === selected) : itemAt(s, p.fx, p.fy);
    if (!hit) { setSelected(null); return; }
    e.preventDefault();
    editFrame().focus();
    setSelected(hit.id);
    const start = hit.box;
    const aspect = start.w / start.h;
    const move = (ev) => {
      const q = frameAt(ev);
      const dx = q.fx - p.fx;
      const dy = q.fy - p.fy;
      let box;
      if (!handle) {
        box = { ...start, x: snap(start.x + dx, start.w), y: snap(start.y + dy, start.h) };
      } else {
        // The corner opposite the one held stays where it is; the shape is kept unless Shift is held.
        const left = handle.includes('w');
        const top = handle.includes('n');
        const ax = left ? start.x + start.w : start.x;
        const ay = top ? start.y + start.h : start.y;
        let w = Math.max(0.04, left ? ax - q.fx : q.fx - ax);
        let h = Math.max(0.04, top ? ay - q.fy : q.fy - ay);
        if (!ev.shiftKey || hit.shape === 'circle') { if (w / h > aspect) h = w / aspect; else w = h * aspect; }
        box = { x: left ? ax - w : ax, y: top ? ay - h : ay, w, h };
      }
      changeStack((st) => updateItem(st, hit.id, { box: clampBox(box) }), false);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      changeStack((st) => st);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };
  const onFrameKey = (e) => {
    if (!selItem || !['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(e.key)) return;
    e.preventDefault();
    const step = e.shiftKey ? 0.05 : 0.005;
    const dx = e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0;
    const dy = e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0;
    changeStack((s) => updateItem(s, selItem.id, (it) => ({ ...it, box: clampBox({ ...it.box, x: it.box.x + dx, y: it.box.y + dy }) })));
  };

  /* ── scenes, and switching between them ────────────────────────────── */

  const builtIns = useMemo(() => builtInScenes(stack, SCENE_NAMES()), [stack]);
  const scenes = [...builtIns.map((sc) => ({ ...sc, own: false })), ...stack.scenes.map((sc, i) => ({ ...sc, key: `own-${i}`, own: true, index: i }))];
  const scenesRef = useRef(scenes);
  scenesRef.current = scenes;
  const fadeMs = () => (stackRef.current.transition === 'fade' ? FADE_MS : 0);

  /**
   * A shot put up. By hand with the preview up, it is the next shot, and goes
   * on air with Cut or Fade; otherwise, and always when it comes from Follow
   * the speaker or the rotation, it goes on air at once, cut or faded as the
   * transition says.
   */
  const cutTo = useCallback((sc, { automatic = false } = {}) => {
    if (!sc) return;
    if (preview && automatic) {
      fadeNext.current = fadeMs();
      setAir((a) => normaliseStack(applyScene(a || stackRef.current, sc)));
      setOnAir(sc.key);
      return;
    }
    if (!preview) { fadeNext.current = fadeMs(); setOnAir(sc.key); }
    changeStack((s) => applyScene(s, sc));
    setLastScene(sc.key);
  }, [preview, changeStack]);

  /** A shot picked by hand: Follow the speaker and the rotation give way to it. */
  const pickScene = useCallback((sc) => {
    if (auto !== 'off') setAuto('off');
    cutTo(sc);
  }, [auto, cutTo]);

  /** The next shot on air: at once (Cut) or dissolved (Fade). */
  const take = useCallback((how) => {
    fadeNext.current = how === 'fade' ? FADE_MS : 0;
    setAir(stackRef.current);
    setOnAir(lastScene);
  }, [lastScene]);

  const togglePreview = () => {
    // Put away, the shot on air stays on air; the one waiting is let go.
    if (preview) { fadeNext.current = 0; if (air) changeStack(() => air); setAir(null); setPreview(false); setLastScene(onAir); return; }
    setAir(stackRef.current);
    setPreview(true);
  };

  // 1 to 9 put up a scene; Space takes the next shot on air with the transition set.
  useEffect(() => {
    const onKey = (e) => {
      if (e.ctrlKey || e.metaKey || e.altKey) return;
      const tag = e.target?.tagName;
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || e.target?.isContentEditable) return;
      if (e.key === ' ' && preview && tag !== 'BUTTON') { e.preventDefault(); take(stackRef.current.transition); return; }
      const n = Number(e.key);
      if (n >= 1 && n <= 9 && scenes[n - 1]) { e.preventDefault(); pickScene(scenes[n - 1]); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  // The microphones there are to pair a shot with: those on and heard.
  const mics = named.sounds.filter((s) => s.kind === 'mic');
  const cameraShots = scenes.filter((sc) => sc.key.startsWith('cam:') || sc.key === 'camera');

  // Follow the speaker: who is speaking, read ten times a second, and the shot of the one who has spoken long enough.
  useEffect(() => {
    if (auto !== 'follow') return undefined;
    const loud = new Map();
    let state = null;
    const tick = setInterval(() => {
      const now = performance.now();
      const st = stackRef.current;
      for (const s of st.sounds) if (s.live && !s.muted && media.level(s.id) > SPEAK_LEVEL) loud.set(s.id, now);
      const speaking = [...loud].filter((e) => now - e[1] < SPEAK_HANG_MS).map(([id]) => id);
      const there = new Set(scenesRef.current.map((sc) => sc.key));
      const pairs = Object.fromEntries(Object.entries(st.follow).filter(([k]) => there.has(k)));
      const r = followStep(state, { now, speaking, pairs, wide: there.has('wide') ? 'wide' : null, on: onAirRef.current });
      state = r.state;
      if (r.cut) cutTo(scenesRef.current.find((sc) => sc.key === r.cut), { automatic: true });
    }, 100);
    return () => clearInterval(tick);
  }, [auto, media, cutTo]);

  // The timed rotation: the chosen shots in turn, or the cameras' and the wide shot when none are chosen.
  useEffect(() => {
    if (auto !== 'rotate') return undefined;
    const tick = setInterval(() => {
      const there = scenesRef.current;
      const chosen = stackRef.current.rotate.keys.filter((k) => there.some((sc) => sc.key === k));
      const cams = there.filter((sc) => sc.key.startsWith('cam:') || sc.key === 'camera' || sc.key === 'wide').map((sc) => sc.key);
      const keys = chosen.length > 1 ? chosen : cams.length > 1 ? cams : there.map((sc) => sc.key);
      const next = nextInRotation(keys, onAirRef.current);
      if (next && next !== onAirRef.current) cutTo(there.find((sc) => sc.key === next), { automatic: true });
    }, stack.rotate.every * 1000);
    return () => clearInterval(tick);
  }, [auto, stack.rotate.every, cutTo]);

  const startFollow = () => {
    if (auto === 'follow') { setAuto('off'); return; }
    const there = new Set(scenes.map((sc) => sc.key));
    let pairs = Object.fromEntries(Object.entries(stack.follow).filter(([k]) => there.has(k)));
    // Each camera with its own microphone, in their orders, when there are as many of one as of the other.
    if (!Object.keys(pairs).length && cameraShots.length > 1 && cameraShots.length === mics.length) {
      pairs = Object.fromEntries(cameraShots.map((sc, i) => [sc.key, mics[i].id]));
      changeStack((s) => ({ ...s, follow: pairs }));
      toast(cameraShots.map((sc, i) => `${sc.name}: ${mics[i].name}`).join(' · '), { ms: 6000 });
    }
    if (!Object.keys(pairs).length) {
      toast(t('Follow the speaker needs a microphone for each shot: right-click a shot to pair it with one.'), { ms: 7000 });
      return;
    }
    setAuto('follow');
  };

  const sceneMenuItems = (sc) => {
    const paired = stack.follow[sc.key] || '';
    const inRotation = stack.rotate.keys.includes(sc.key);
    return [
      ...(isLive ? [
        { label: t('Show when this microphone is heard'), icon: 'mic', items: [
          ...mics.map((m) => ({ label: m.name, checked: paired === m.id, run: () => changeStack((s) => ({ ...s, follow: { ...s.follow, [sc.key]: m.id } })) })),
          { label: t('No microphone'), checked: !paired, run: () => changeStack((s) => { const f = { ...s.follow }; delete f[sc.key]; return { ...s, follow: f }; }) },
        ] },
        { label: t('In the rotation'), icon: 'refresh', checked: inRotation, run: () => changeStack((s) => ({ ...s, rotate: { ...s.rotate, keys: inRotation ? s.rotate.keys.filter((k) => k !== sc.key) : [...s.rotate.keys, sc.key] } })) },
      ] : []),
      ...(sc.own ? [...(isLive ? ['-'] : []), { label: t('Remove the scene'), icon: 'trash', run: () => changeStack((s) => removeScene(s, sc.index)) }] : []),
    ];
  };

  /* ── the disk ──────────────────────────────────────────────────────── */

  const camerasShown = airStack.items.filter((it) => it.on && it.kind === 'camera').length;
  const rate = useMemo(() => bytesPerSecond({ fps: settings.fps, cap: settings.size, quality: settings.quality, camera: false, mic: true, system: false }, { screen: settings.size === 720 ? { width: 1280, height: 720 } : { width: 1920, height: 1080 } }), [settings.fps, settings.size, settings.quality]);
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

  // The recordings folder's latest takes, for the menu's Recent Recordings.
  const loadRecent = useCallback(async () => {
    if (!settings.folder) return;
    try {
      const list = await shell.fs.list({ path: settings.folder, filter: SOUND_EXT });
      setRecent((list || []).filter((f) => !f.dir).sort((a, b) => (b.mtime || 0) - (a.mtime || 0)).slice(0, 10));
    } catch { setRecent([]); }
  }, [shell, settings.folder]);
  useEffect(() => { loadRecent(); }, [loadRecent, saved]);

  /* ── recording ─────────────────────────────────────────────────────── */

  const recKind = recordKind(airStack);
  const begin = useCallback(async () => {
    try {
      setSaved(null);
      const soundOnly = recordKind(program.current?.stack || stackRef.current) === 'sound';
      const name = soundOnly ? takeFileName(t('Sound'), new Date(), 'weba') : takeFileName(t('Recording'), new Date(), 'webm');
      const { path } = await shell.capture.take({ dir: settingsRef.current.folder, name });
      takenRef.current = path;
      await program.current.record({ shell, path, soundOnly });
      setPhase('recording');
      setPaused(false);
      warned.current = false;
      const c = await shell.capture.controls({ on: true, paused: false, minimise: settingsRef.current.minimise }).catch(() => null);
      setKeys(c);
    } catch (err) {
      setPhase('setup');
      toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
    }
  }, [shell, toast]);

  const record = useCallback(async () => {
    if (phaseRef.current !== 'setup') return;
    if (!recordKind(program.current?.stack || stackRef.current)) { toast(t('Video and Sound are both off: there is nothing to record.'), { ms: 5000 }); return; }
    const state = await checkDisk();
    if (state?.level === 'stop') { toast(t('There is not enough space on the disk to record.'), { tone: 'bad', ms: 6000 }); return; }
    if (settingsRef.current.countdown > 0) { setCount(settingsRef.current.countdown); setPhase('countdown'); }
    else { setPhase('starting'); begin(); }
  }, [checkDisk, begin, toast]);

  useEffect(() => {
    if (phase !== 'countdown') return undefined;
    if (count <= 0) { setPhase('starting'); begin(); return undefined; }
    const timer = setTimeout(() => setCount((n) => n - 1), 1000);
    return () => clearTimeout(timer);
  }, [phase, count, begin]);

  const stop = useCallback(async () => {
    if (phaseRef.current === 'countdown') { setPhase('setup'); return; }
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
  useEffect(() => shell.on('capture:command', ({ command, ...request } = {}) => {
    if (command === 'remote') remoteRef.current?.(request);
    else if (command === 'stop') stopRef.current?.();
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

  /* ── pictures, at any time ─────────────────────────────────────────── */

  const selRef = useRef(null);
  selRef.current = selItem;
  const snapNow = useCallback(async () => {
    setShotCount(null);
    try {
      const of = settingsRef.current.shot === 'program' ? null : selRef.current;
      const shot = await program.current?.snapshot(of);
      if (!shot) { toast(t('There is nothing to take yet.'), { tone: 'bad', ms: 5000 }); return; }
      const { path } = await shell.capture.take({ dir: settingsRef.current.folder, name: takeFileName(t('Picture'), new Date(), 'png') });
      await shell.fs.write({ path, bytes: shot.bytes });
      setFlash((n) => n + 1);
      setPicture((old) => {
        if (old?.url) URL.revokeObjectURL(old.url);
        return { path, width: shot.width, height: shot.height, url: URL.createObjectURL(new Blob([shot.bytes], { type: 'image/png' })) };
      });
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
    }
  }, [shell, toast]);

  const takePicture = useCallback(() => {
    if (shotCount != null) return;
    if (settingsRef.current.shotTimer > 0) setShotCount(settingsRef.current.shotTimer);
    else snapNow();
  }, [shotCount, snapNow]);
  pictureRef.current = takePicture;

  // Home's Take a picture (?picture=1): one taken, once what was kept has had a moment to show.
  const askedPicture = useRef(new URLSearchParams(location.search).has('picture'));
  useEffect(() => {
    if (!loaded || !askedPicture.current || !stack.items.length) return;
    askedPicture.current = false;
    setTimeout(() => pictureRef.current?.(), 1500);
  }, [loaded, stack]);

  useEffect(() => {
    if (shotCount == null) return undefined;
    if (shotCount <= 0) { snapNow(); return undefined; }
    const timer = setTimeout(() => setShotCount((n) => (n == null ? n : n - 1)), 1000);
    return () => clearTimeout(timer);
  }, [shotCount, snapNow]);

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

  /* ── going live (Live Broadcast) ───────────────────────────────────── */

  // The bitrate the slowest destination can carry: the encoder follows it, and the status bar says so.
  const [liveRate, setLiveRate] = useState(null);
  useEffect(() => (isLive ? shell.on('live:status', ({ destinations = [], kbps, target } = {}) => {
    setLiveStatus(destinations);
    if (kbps) { setLiveRate({ kbps, target }); encoder.current?.setBitrate(kbps); }
  }) : undefined), [shell, isLive]);
  const kbps = isLive ? liveKit.liveBitrate(settings.size, settings.fps) : 0;

  const goLive = useCallback(async () => {
    if (!isLive || livePhase !== 'off') return;
    const { LiveEncoder, AUDIO_KBPS } = liveKit;
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
      const enc = new LiveEncoder({ shell, program: program.current, rig: media });
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
  }, [isLive, liveKit, livePhase, settings.fps, settings.liveCopy, kbps, shell, media, toast, begin]);

  // The bar on top of the Live panel runs the broadcast; here a copy is kept as it goes.
  const barPhase = useCallback((p, since) => {
    setLivePhase(p);
    if (since) setLiveSince(since);
    if (p === 'live' && settingsRef.current.liveCopy && phaseRef.current === 'setup') { liveRecord.current = true; setPhase('starting'); begin(); }
    if (p === 'off' && liveRecord.current) { liveRecord.current = false; stopRef.current?.(); }
  }, [begin]);

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
  const choosePlaceholder = useCallback(async () => {
    const picked = await shell.dialog.open({ title: t('Choose a picture'), filters: [{ name: t('Pictures'), extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp'] }] });
    if (picked?.[0]) changeStack((s) => ({ ...s, placeholder: picked[0] }));
  }, [shell, changeStack]);
  const chooseMusic = useCallback(async () => {
    const picked = await shell.dialog.open({ title: t('Choose music'), filters: [{ name: t('Sound'), extensions: ['mp3', 'm4a', 'wav', 'ogg', 'weba', 'webm', 'flac'] }] });
    if (picked?.[0]) changeStack((s) => ({ ...s, music: picked[0] }));
  }, [shell, changeStack]);

  /* ── the AI tools' Recorder (recorder_* tools, office-shell recorder-remote.js) ── */

  // A recording the tools ask for is made as a person's is: the source shown, the settings set,
  // Record pressed, without the countdown. The person's own stack and settings come back when it ends.
  // What it answers goes to the AI app, in English as every tool's answer is, not the window's words.
  // Only the Recorder takes them.
  const answer = useCallback((reply, result, error = null) => {
    shell.capture.remoteReply({ reply, result, error: error ? String(error) : null }).catch(() => {});
  }, [shell]);
  useEffect(() => {
    if (kind !== 'recorder') return undefined;
    shell.capture.remoteListen({ on: true }).catch(() => {});
    return () => { shell.capture.remoteListen({ on: false }).catch(() => {}); };
  }, [shell, kind]);
  remoteRef.current = kind !== 'recorder' ? null : async ({ op, args = {}, reply } = {}) => {
    if (op === 'sources') {
      try {
        const list = (await shell.capture.sources()) || [];
        setSources(list);
        answer(reply, list.map((x) => ({ id: x.id, name: x.name, kind: x.kind })));
      } catch (err) {
        answer(reply, null, err?.message || err);
      }
      return;
    }
    if (op === 'state') {
      answer(reply, { phase, seconds: Math.round(elapsed / 1000), disk: disk ? { level: disk.level, minutes: Math.floor(disk.minutes) } : null, saved, forTools: Boolean(remote.current) });
      return;
    }
    if (op === 'start') {
      if (phaseRef.current !== 'setup' || remote.current) { answer(reply, null, 'The Recorder is already recording.'); return; }
      let list = [];
      try { list = (await shell.capture.sources()) || []; } catch { /* none to give */ }
      const src = list.find((x) => x.id === args.source);
      if (!src) { answer(reply, null, 'That window or screen is not there to record. recorder_windows names those there are.'); return; }
      // A stack of its own: the source, the camera when asked, the title, the sounds asked for.
      let s = { ...emptyStack(), sounds: [{ id: 'mic-1', kind: 'mic', ref: '', name: '', live: Boolean(args.microphone), muted: false, gain: 1 }, { id: 'system-1', kind: 'system', ref: '', name: '', live: args.sound !== false, muted: false, gain: 1 }] };
      const screenAdded = addItem(s, { kind: src.kind === 'window' ? 'window' : 'screen', ref: src.id, name: src.name });
      s = screenAdded.stack;
      if (args.camera) s = addItem(s, { kind: 'camera', ref: '', shape: 'circle' }).stack;
      const want = builtInScenes(s).find((sc) => sc.key === (args.scene || 'screen')) || builtInScenes(s).find((sc) => sc.key === 'screen');
      s = applyScene(s, want);
      if (args.title) s = addItem(s, { kind: 'title', text: args.title }).stack;
      remote.current = { startReply: reply, kept: { stack: stackRef.current, settings: settingsRef.current }, item: screenAdded.id, source: src.id, folder: args.folder, triggered: false };
      setSources(list);
      setStack(normaliseStack(s));
      change({ size: args.size ?? settings.size, fps: args.fps ?? settings.fps, folder: args.folder, countdown: 0 });
      // A source that gives no picture: said, and the person's own put back.
      const asked = remote.current;
      setTimeout(() => {
        if (remote.current !== asked || asked.triggered) return;
        remote.current = null;
        setStack(asked.kept.stack);
        change(asked.kept.settings);
        answer(reply, null, 'The Recorder had no picture of that window to record.');
      }, 20000);
      return;
    }
    if (op === 'stop') {
      if (phaseRef.current !== 'recording') { answer(reply, null, 'Nothing is recording.'); return; }
      remote.current = { ...(remote.current || { kept: null, triggered: true }), stopReply: reply };
      stopRef.current?.();
    }
  };
  // Record pressed once the asked-for source gives its picture and the settings are in place.
  const remoteItem = remote.current ? stack.items.find((it) => it.id === remote.current.item) : null;
  const remoteReady = Boolean(remoteItem && media.stateOf(remoteItem) === 'live');
  useEffect(() => {
    const r = remote.current;
    if (!r?.startReply || r.triggered || phase !== 'setup' || !remoteReady) return;
    if (settings.folder !== r.folder || settings.countdown !== 0) return;
    r.triggered = true;
    record();
  }, [phase, remoteReady, settings, record]);
  // Its answers: started, stopped and saved, or ended by the person's own Stop.
  useEffect(() => {
    const r = remote.current;
    if (!r) return;
    if (r.startReply && r.triggered && phase === 'recording') {
      answer(r.startReply, { recording: true, path: takenRef.current, source: remoteItem?.name || r.source });
      r.startReply = null;
      return;
    }
    if (phase !== 'setup' || !r.triggered) return;
    if (r.startReply) answer(r.startReply, null, 'The recording did not start.');
    if (r.stopReply) answer(r.stopReply, saved ? { path: saved.path, seconds: Math.round(saved.seconds ?? 0), bytes: saved.written ?? null } : null, saved ? null : 'The recording could not be saved.');
    remote.current = null;
    if (r.kept) { setStack(r.kept.stack); change(r.kept.settings); }
  }, [phase, saved, remoteItem, answer, change]);

  /* ── the menu ──────────────────────────────────────────────────────── */

  // No New or Open: a second window would share the one camera, and a recording is opened in Video.
  const settingsLabel = isLive ? t('Live Broadcast settings…') : t('Recorder settings…');
  const appMenu = useAppMenu({
    shell,
    appKey: kind,
    own: [
      ...(isLive ? [{ label: t('Live Destinations…'), icon: 'globe', run: () => setLiveOpen(true) }] : []),
      { label: t('Open in Video…'), icon: 'video', run: async () => {
        const files = await shell.dialog.open({ title: t('Open in Video'), defaultPath: settings.folder, filters: [{ name: t('Video and sound'), extensions: SOUND_EXT }] });
        if (files?.[0]) shell.win.create({ app: 'video', file: files[0] });
      } },
      { label: t('Recordings Folder'), icon: 'folderOpen', run: () => { if (settings.folder) shell.shell.openPath({ path: settings.folder }).catch(() => {}); } },
      { label: t('Recent Recordings'), icon: 'clock', items: recent.length ? recent.map((f) => ({ label: f.name, icon: 'video', title: f.path, run: () => shell.win.create({ app: 'video', file: f.path }) })) : [{ label: t('No recordings yet'), disabled: true }] },
      { label: settingsLabel, icon: 'settings', run: () => setSettingsOpen(true) },
    ],
  });

  /* ── the window ────────────────────────────────────────────────────── */

  const recording = phase === 'recording' || phase === 'stopping';
  const live = livePhase !== 'off';
  const busy = phase !== 'setup' || live;
  const liveFor = live && liveSince ? Date.now() - liveSince : 0;
  const LIVE_WORDS = { waiting: t('Waiting'), connecting: t('Connecting…'), live: t('Live'), reconnecting: t('Reconnecting…'), error: t('Not streaming'), ended: t('Ended'), closed: t('Ended') };
  const anyShown = stack.items.some((it) => it.on);
  const canRecord = phase === 'setup' && Boolean(recKind) && disk?.level !== 'stop';
  const minutes = disk ? Math.floor(disk.minutes) : null;
  const diskNote = disk == null ? '' : disk.level === 'stop' ? t('There is not enough space on the disk to record.')
    : disk.minutes < 120 ? t('Room for about {minutes} minutes on the disk', { minutes: Math.max(0, minutes) })
      : t('Room for about {hours} hours on the disk', { hours: Math.floor(disk.minutes / 60) });
  const viewers = live ? liveStatus.reduce((n, d) => (Number.isFinite(d.viewers) ? (n ?? 0) + d.viewers : n), null) : null;

  const recordLabel = phase === 'countdown' ? t('Starting in {n}', { n: count })
    : recording ? `${paused ? t('Paused') : t('Recording')} ${clock(elapsed)}`
      : recKind === 'sound' ? t('Record sound') : t('Record video');
  const switchRow = (label, value, options, onPick, cls) => (
    <span className={`st-switch ${cls}`} role="radiogroup" aria-label={label}>
      {options.map(([v, icon, text]) => (
        <button key={v} type="button" role="radio" aria-checked={value === v} data-value={v} className={value === v ? 'on' : ''} onClick={() => onPick(v)}><Icon name={icon} size={14} />{text}</button>
      ))}
    </span>
  );
  const top = (
    <div className="st-top" role="toolbar" aria-label={appName}>
      <span className="st-acts">
        <button type="button" className="st-act st-shoot" title={`${t('Take Picture')}${pictureKey ? ` (${pictureKey})` : ''}`} disabled={shotCount != null || !settings.folder} onClick={takePicture}>
          <Icon name="picture" size={16} />{shotCount != null ? t('In {seconds} s', { seconds: shotCount }) : t('Picture')}
        </button>
        {isLive ? (
          <button type="button" className={`st-act st-golive${live ? ' on' : ''}`} aria-pressed={live} title={live ? t('End Live') : t('Go Live')} disabled={livePhase === 'ending' || (!live && (phase === 'countdown' || phase === 'starting'))} onClick={live ? endLive : goLive}>
            {live ? <i className="st-dot" /> : <Icon name="globe" size={16} />}
            {live ? (livePhase === 'starting' ? t('Starting…') : `${t('Live')} ${clock(liveFor)}`) : t('Go Live')}
            {live && viewers != null ? <span className="st-viewers" title={t('Watching now')}><Icon name="eye" size={13} />{viewers}</span> : null}
          </button>
        ) : null}
        <button type="button" className={`st-act st-record${recording || phase === 'countdown' ? ' on' : ' rec'}${paused ? ' held' : ''}`} aria-pressed={recording} title={recording ? (keys?.stop ? `${t('Stop')} (${keys.stop})` : t('Stop')) : recordLabel} disabled={phase === 'stopping' || phase === 'starting' || (!recording && phase !== 'countdown' && !canRecord)} onClick={recording || phase === 'countdown' ? stop : record}>
          <i className="st-dot" />{recordLabel}
        </button>
        {recording ? (
          <button type="button" className={`st-act st-act-pause st-pause${paused ? ' on held' : ''}`} title={paused ? t('Resume') : t('Pause')} aria-label={paused ? t('Resume') : t('Pause')} aria-pressed={paused} disabled={phase !== 'recording'} onClick={togglePause}>
            <Icon name={paused ? 'play' : 'pause'} size={15} />
          </button>
        ) : null}
      </span>
      {switchRow(t('Video'), stack.video, [['on', 'video', t('Video')], ['off', 'eyeOff', t('No video')], ['placeholder', 'picture', t('Placeholder')]], (v) => changeStack((s) => ({ ...s, video: v })), 'st-video')}
      {switchRow(t('Sound'), stack.sound, [['on', 'mic', t('Sound')], ['off', 'micOff', t('No sound')], ['music', 'headphones', t('Music')]], (v) => changeStack((s) => ({ ...s, sound: v })), 'st-sound')}
      <span className="st-top-gap" />
      <button type="button" className="st-gear" title={settingsLabel.replace('…', '')} aria-label={settingsLabel.replace('…', '')} onClick={() => setSettingsOpen(true)}><Icon name="settings" size={16} /></button>
    </div>
  );

  const airName = scenes.find((sc) => sc.key === onAir)?.name;
  const status = (
    <span className="st-status" data-phase={phase}>
      {live ? `${t('Live')} ${clock(liveFor)} · ` : ''}
      {live && liveRate?.kbps ? `${t('{n} kb/s', { n: liveRate.kbps })}${liveRate.kbps < liveRate.target ? ` ${t('(lowered to keep up with the upload)')}` : ''} · ` : ''}
      {auto === 'follow' ? `${t('Following the speaker')} · ` : auto === 'rotate' ? `${t('A new shot every {n} s', { n: stack.rotate.every })} · ` : ''}
      {isLive && airName ? `${t('On air: {name}', { name: airName })} · ` : ''}
      {recording ? `${paused ? t('Paused') : t('Recording')} · ${clock(elapsed)}` : saved ? t('Saved {name}', { name: basename(saved.path) }) : t('Ready')}
      {` · ${settings.size === 720 ? '1280 × 720' : '1920 × 1080'} · ${t('{n} frames a second', { n: settings.fps })}`}
      {camerasShown > 1 ? ` · ${t('{n} cameras', { n: camerasShown })}` : ''}
      {diskNote && !recording ? ` · ${diskNote}` : ''}
    </span>
  );

  const shownBoxes = stack.video === 'on' ? stack.items.filter((it) => it.on) : [];
  const selBox = selItem && selItem.on ? selItem.box : null;
  // What the person works on, drawn over the frame that shows it: the boxes, the clip being drawn, the hints.
  const overlay = (
    <>
      <div className="st-boxes">
        {shownBoxes.map((it) => (
          <div key={it.id} className={`st-box${it.shape === 'circle' ? ' circle' : ''}${selected === it.id ? ' on' : ''}${selected === it.id && clipping ? ' clipping' : ''}`} data-id={it.id}
            style={{ left: `${it.box.x * 100}%`, top: `${it.box.y * 100}%`, width: `${it.box.w * 100}%`, height: `${it.box.h * 100}%` }}>
            {selected === it.id && !clipping ? ['nw', 'ne', 'sw', 'se'].map((c) => <span key={c} className="st-handle" data-c={c} />) : null}
          </div>
        ))}
      </div>
      {drawing ? (
        <div className="rec-region st-drawn" style={{ left: `${Math.min(drawing.a.fx, drawing.b.fx) * 100}%`, top: `${Math.min(drawing.a.fy, drawing.b.fy) * 100}%`, width: `${Math.abs(drawing.b.fx - drawing.a.fx) * 100}%`, height: `${Math.abs(drawing.b.fy - drawing.a.fy) * 100}%` }}><span>{t('Clip')}</span></div>
      ) : null}
      {clipping && selBox && !drawing ? <div className="rec-hint">{t('Drag over the part of {name} to show.', { name: itemName(selItem) })}</div> : null}
      {stack.video === 'on' && !anyShown && loaded ? <div className="st-empty">{sources == null ? t('Looking at what is open…') : t('Show a screen, a window or a camera from the list on the right.')}</div> : null}
    </>
  );
  const editing = (mine) => (mine ? { tabIndex: 0, onPointerDown: onFramePointer, onKeyDown: onFrameKey } : {});

  return (
    <AppFrame app={app} shell={shell} title={live ? `${t('Live')} ${clock(liveFor)}` : recording ? `${paused ? t('Paused') : t('Recording')} ${clock(elapsed)}` : appName} menu={appMenu} ribbon={top} status={status}>
      <style>{RECORDER_CSS}</style>
      <style>{STUDIO_CSS}</style>
      {isLive ? <style>{liveKit.LIVE_CSS}</style> : null}
      {isLive && liveOpen ? <liveKit.LiveDialog shell={shell} toast={toast} kbps={kbps} busy={live} recordCopy={settings.liveCopy} onRecordCopy={(v) => change({ liveCopy: v })} onClose={() => setLiveOpen(false)} /> : null}
      {settingsOpen ? (
        <StudioSettings settings={settings} stack={stack} busy={busy} live={isLive} title={settingsLabel.replace('…', '')} change={change} changeStack={changeStack} chooseFolder={chooseFolder} choosePlaceholder={choosePlaceholder} chooseMusic={chooseMusic}
          openDestinations={() => { setSettingsOpen(false); setLiveOpen(true); }} onClose={() => setSettingsOpen(false)} />
      ) : null}

      <Content>
        <div className="st-stage">
          {isLive ? <liveKit.LiveBar shell={shell} toast={toast} program={program} media={media} fps={settings.fps} kbps={kbps} copy={live && settings.liveCopy} blocked={phase === 'countdown' || phase === 'starting'} onPhase={barPhase} /> : null}
          <SourceBar item={selItem} sound={selSound} clipping={clipping} act={barAct} />
          <div className={`st-share${preview ? ' duo' : ''}`}>
            <div className="st-pane on-air">
              {preview ? <span className="st-pane-label"><i />{t('On air')}</span> : null}
              <div ref={airFrameRef} aria-label={preview ? t('On air') : t('The share')} className={`st-frame st-air${recording ? ' live' : ''}${paused ? ' held' : ''}`} {...editing(!preview)}>
                <canvas ref={canvasRef} className="st-canvas" width={1920} height={1080} aria-label={t('Program')} />
                {!preview ? overlay : null}
                {recording ? <span className="rec-badge st-badge"><i />{paused ? t('Paused') : t('Recording')} {clock(elapsed)}</span> : null}
                {live ? <span className="st-live-badge" data-phase={livePhase}><i />{t('LIVE')} {clock(liveFor)}</span> : null}
                {phase === 'countdown' ? <div className="rec-count" aria-live="assertive"><span key={count}>{count}</span></div> : null}
                {shotCount != null && shotCount > 0 && phase !== 'countdown' ? <div className="rec-count st-shot-count" aria-live="assertive"><span key={shotCount}>{shotCount}</span></div> : null}
                {flash ? <div key={flash} className="st-flash" aria-hidden="true" /> : null}
              </div>
            </div>
            <div className={`st-pane st-next-pane${preview ? '' : ' off'}`}>
              <span className="st-pane-label">{t('Next')}</span>
              <div ref={nextFrameRef} aria-label={t('Next')} className="st-frame st-next" {...editing(preview)}>
                <canvas ref={previewCanvasRef} className="st-canvas st-preview-canvas" width={1920} height={1080} aria-label={t('Next')} />
                {preview ? overlay : null}
              </div>
              {preview ? (
                <div className="st-take">
                  <button type="button" className="st-act st-cut" title={t('Put the next shot on air at once')} onClick={() => take('cut')}><Icon name="scissors" size={15} />{t('Cut')}</button>
                  <button type="button" className="st-act st-fade" title={t('Fade the next shot in')} onClick={() => take('fade')}><Icon name="layers" size={15} />{t('Fade')}</button>
                </div>
              ) : null}
            </div>
          </div>
          <div className="st-scenes" role="toolbar" aria-label={t('Scenes')}>
            <span className="st-scenes-label">{t('Scenes')}</span>
            {scenes.map((sc, i) => (
              <button key={sc.key} type="button" className={`st-scene${lastScene === sc.key ? ' on' : ''}${isLive && onAir === sc.key ? ' air' : ''}`} data-scene={sc.key} title={sc.name} onClick={() => pickScene(sc)}
                onContextMenu={sceneMenuItems(sc).length ? (e) => { e.preventDefault(); setSceneMenu({ x: e.clientX, y: e.clientY, sc }); } : undefined}>
                {i < 9 ? <kbd>{i + 1}</kbd> : null}{sc.name}
                {isLive && (stack.follow[sc.key] || stack.rotate.keys.includes(sc.key)) ? (
                  <span className="st-scene-tags">
                    {stack.follow[sc.key] ? <Icon name="mic" size={11} title={mics.find((m) => m.id === stack.follow[sc.key])?.name || ''} /> : null}
                    {stack.rotate.keys.includes(sc.key) ? <Icon name="refresh" size={11} /> : null}
                  </span>
                ) : null}
              </button>
            ))}
            <button type="button" className="st-scene st-scene-save" title={t('Keep what is on the share as a scene')} aria-label={t('Keep what is on the share as a scene')} disabled={!anyShown || stack.scenes.length >= 9} onClick={() => changeStack((s) => saveScene(s, t('Scene {n}', { n: s.scenes.length + 1 })))}>
              <Icon name="plus" size={14} />
            </button>
            {isLive ? (
              <span className="st-switching" role="group" aria-label={t('Switching')}>
                <button type="button" className={`st-tool st-preview-btn${preview ? ' on' : ''}`} aria-pressed={preview} title={t('Show the next shot beside the one on air, and put it on air with Cut or Fade')} onClick={togglePreview}><Icon name="layers" size={14} /><span>{t('Preview')}</span></button>
                <button type="button" className={`st-tool st-follow${auto === 'follow' ? ' on' : ''}`} aria-pressed={auto === 'follow'} title={t('Show whoever is speaking, by the microphone paired with each shot')} onClick={startFollow}><Icon name="mic" size={14} /><span>{t('Follow the speaker')}</span></button>
                <button type="button" className={`st-tool st-rotate${auto === 'rotate' ? ' on' : ''}`} aria-pressed={auto === 'rotate'} title={t('Move between the shots on a timer')} onClick={() => setAuto((a) => (a === 'rotate' ? 'off' : 'rotate'))}><Icon name="refresh" size={14} /><span>{t('Rotate')}</span></button>
                <Select value={String(stack.rotate.every)} className="st-rotate-every" aria-label={t('A new shot every')} onChange={(e) => changeStack((s) => ({ ...s, rotate: { ...s.rotate, every: Number(e.target.value) } }))}>
                  {ROTATE_EVERY.map((n) => <option key={n} value={String(n)}>{t('Every {n} s', { n })}</option>)}
                </Select>
                <Select value={stack.transition} className="st-transition" aria-label={t('Transition')} onChange={(e) => changeStack((s) => ({ ...s, transition: e.target.value }))}>
                  <option value="cut">{t('Cut')}</option>
                  <option value="fade">{t('Fade')}</option>
                </Select>
              </span>
            ) : null}
            {sceneMenu ? <Menu x={sceneMenu.x} y={sceneMenu.y} items={sceneMenuItems(sceneMenu.sc)} onClose={() => setSceneMenu(null)} /> : null}
          </div>
          <div className="st-under">
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
        </div>
      </Content>

      <Panel name={`${kind}-sources`} width={280} right resizable title={t('Sources')}>
        <SourceList stack={named} spare={spare} spareSounds={spareSounds} media={media} selected={selected} onSelect={setSelected} act={act} thumbs={thumbs} />
      </Panel>
    </AppFrame>
  );
}
