// The studio window both the Recorder and Live Broadcast are: the screen,
// windows, cameras and sound made into one picture (owner, 2026-10-19: "split
// the recording tool out to be able to record video and stream videos live";
// 2026-10-20: "there should be a way to capture pictures on camera and from
// screens etc.").
//
// One studio in which every source works the same way (D-2026-11-02-13). The
// Recorder records to a file; Live Broadcast goes live, a product of its own
// (D-2026-11-02-12, D-2026-11-03-01), with shots switched by hand, by who is
// speaking, or on a timer, and a preview of the next shot beside the one on air
// (D-2026-11-03-02).
//
// Two panels (2.23, A1): Set up, where the sources and scenes are arranged, and
// Record or Live, where the work is done with few controls. What the person has
// arranged is a named setup (A4), kept as it changes; which sources it means is
// found again each time it opens. This file holds the state and what it does.
// The panels draw it: record-panel.js, live-panel.js, setup-panel.js.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { Content, Icon, useToast, basename, t } from '@rutba/office-ui';
import { normaliseStudio } from '@rutba/studio/compositor';
import {
  normaliseStack, emptyStack, addItem, removeItem, duplicateItem, moveItem, showItem, placeItem, shapeItem, clipItem,
  updateItem, updateSound, frameToSource, itemAt, saveScene, removeScene, renameScene, duplicateScene, moveScene, updateSceneFromFrame,
  applyScene, builtInScenes, recordKind, clampBox, newId, followStep, nextInRotation, followFrame, backSoonScene, soundOnlyScene, snapMove, snapPoint,
  FADE_MS, BACK_SOON, SOUND_ONLY,
} from '@rutba/studio/stack';
import {
  SETUP_EXT, SETUP_KEYS, GLOBAL_KEYS, setupSettings, globalSettings, joinSettings, newSetup, duplicateSetup, migrate, rematch, sceneMissing,
  openingSetup, setupId, uniqueName, menuOrder, emptyHistory, historyPush, historyUndo, historyRedo,
} from '@rutba/studio/setup';
import { takeFileName, bytesPerSecond, diskState, regionCrop } from '@rutba/studio/recorder';
import { AppFrame, useAppMenu } from '../shell.js';
import { RECORDER_CSS, PROMPTER_CSS } from '../recording/css.js';
import { recordingType } from '../recording/session.js';
import { readInto, openLink } from './into.js';
import { makeTakes } from './takes.js';
import { Program } from './program.js';
import { whileStopping } from './stopping.js';
import { WebPageDialog, FLIP_CSS } from './flip-strip.js';
import { TALL_FRAME_CSS } from './tall-frame.js';
import { makeTallProgram } from './tall-program.js';
import { tallSize } from '@rutba/studio/tall';
import { StudioMedia, sourceKey } from './media.js';
import { itemName } from './list.js';
import { StudioSettings } from './settings-dialog.js';
import { STUDIO_CSS } from './css.js';
import { createSetupStore, openSetupFile, saveSetupFile, SETUP_FILTERS } from './setups-store.js';
import { PanelSwitch, NameDialog, ConfirmDialog } from './parts.js';
import { SetupsButton, OpenSetupDialog } from './setups-menu.js';
import { RecordPanel } from './record-panel.js';
import { LivePanel } from './live-panel.js';
import { SetupPanel } from './setup-panel.js';
import { clock } from './shared.js';

const DISK_EVERY_MS = 5000;
const COPY_EVERY_MS = 5000;
const SOUND_EXT = ['webm', 'weba', 'ogg', 'mp4', 'mov', 'mkv', 'm4a', 'mp3', 'wav'];
// Follow the speaker: how loud a microphone is when someone speaks into it, and how long a pause between words still counts as speaking.
const SPEAK_LEVEL = 0.06;
const SPEAK_HANG_MS = 450;
// A window's list is looked at again when this window is come back to, no more often than this.
const SOURCES_EVERY_MS = 4000;

const SCENE_NAMES = () => ({
  'screen-camera': t('Screen and camera'), screen: t('Screen'), window: t('A window'), camera: t('Camera'), split: t('Side by side'), wide: t('Everyone'),
  cameraN: (i) => t('Camera {n}', { n: i + 1 }), [BACK_SOON]: t('Back soon'), 'sound-only': t('Sound only'), backSoon: t('Back soon'), soundOnly: t('Sound only'),
});

/** A source's picture aspect, from what is open or what the capture list says. */
function aspectOf(media, item, sources) {
  const size = media.sizeOf(item);
  if (size) return size.w / size.h;
  const d = sources?.find((s) => s.id === item.ref)?.display;
  return d?.width && d?.height ? d.width / d.height : 16 / 9;
}

function makeCanvas(className, label) {
  const c = document.createElement('canvas');
  c.className = className;
  c.width = 1920;
  c.height = 1080;
  c.setAttribute('aria-label', label);
  return c;
}

/** The small picture a saved take shows on its card. */
function stillOf(canvas) {
  try {
    const c = document.createElement('canvas');
    c.width = 320;
    c.height = 180;
    c.getContext('2d').drawImage(canvas, 0, 0, 320, 180);
    return c.toDataURL('image/jpeg', 0.8);
  } catch {
    return null;
  }
}

/**
 * @param {object} p
 * @param {'recorder'|'broadcast'} p.kind  which app this window is
 * @param {object} [p.live]  Live Broadcast's going live: { LiveBar, LiveDialog, LIVE_CSS, liveBitrate }
 */
export function StudioWindow({ app, shell, boot = null, kind = 'recorder', live: liveKit = null }) {
  const isLive = kind === 'broadcast' && Boolean(liveKit);
  const appName = isLive ? t('Live Broadcast') : t('Recorder');
  const toast = useToast();
  const store = useMemo(() => createSetupStore(shell), [shell]);
  const canvases = useMemo(() => ({ air: makeCanvas('st-canvas st-air-canvas', t('Program')), next: makeCanvas('st-canvas st-preview-canvas', t('Next')) }), []);
  const [settings, setSettings] = useState(() => normaliseStudio(null));
  const [stack, setStack] = useState(() => emptyStack());
  const [setup, setSetupState] = useState(null);
  const [index, setIndex] = useState(null);
  const [panel, setPanel] = useState('main'); // setup | main (Record or Live)
  const [missing, setMissing] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [sources, setSources] = useState(null);
  const [devices, setDevices] = useState({ cameras: [], mics: [], labelled: false });
  const [selected, setSelected] = useState(null);
  const [clipping, setClipping] = useState(false);
  const [drawing, setDrawing] = useState(null);
  const [guides, setGuides] = useState([]);
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
  const [askWeb, setAskWeb] = useState(false);
  const [frameShape, setFrameShape] = useState('wide');
  // The tall picture, one for the tall streams and a tall copy; and what is recording, when that is not the wide picture.
  const tallPicture = useRef(null);
  const copyProgram = useRef(null);
  const recorderNow = () => copyProgram.current || program.current;
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [sceneMenu, setSceneMenu] = useState(null);
  const [sceneAsk, setSceneAsk] = useState(null); // { what: 'rename' | 'delete', sc }
  const [pickRow, setPickRow] = useState(null);
  const [fileOpening, setFileOpening] = useState(null);
  const [recent, setRecent] = useState([]);
  const [, setTick] = useState(0);
  const [, setHistoryTick] = useState(0);
  // What is held on air while the next shot is worked on: with Preview first, and in Set up while live.
  const [air, setAir] = useState(null);
  // Go Live: off, starting, live or ending; the destinations as the main process reports them.
  const [livePhase, setLivePhase] = useState('off');
  const [liveStatus, setLiveStatus] = useState([]);
  const [liveSince, setLiveSince] = useState(0);
  const [liveOpen, setLiveOpen] = useState(false);
  const [closeAsk, setCloseAsk] = useState(false);
  const [cameraAsk, setCameraAsk] = useState(null);
  // The bar's way in, and the element it is drawn in: kept for the window's life, so a broadcast goes on while Set up is open.
  const liveControl = useRef(null);
  const liveBarHost = useMemo(() => { const el = document.createElement('div'); el.className = 'st-livebar-host'; return el; }, []);
  const liveRecord = useRef(false);
  const editFrameRef = useRef(null);
  const phaseRef = useRef(phase);
  phaseRef.current = phase;
  const stackRef = useRef(stack);
  const settingsRef = useRef(settings);
  const setupRef = useRef(setup);
  const sourcesRef = useRef(sources);
  sourcesRef.current = sources;
  const devicesRef = useRef(devices);
  devicesRef.current = devices;
  const panelRef = useRef(panel);
  panelRef.current = panel;
  const globalRef = useRef({});
  const pathsRef = useRef(null);
  const history = useRef(emptyHistory());
  const pendingBefore = useRef(null);
  const stopRef = useRef(null);
  const pauseRef = useRef(null);
  const warned = useRef(false);
  const fadeNext = useRef(0);
  const media = useMemo(() => new StudioMedia({ shell, onChange: () => setTick((n) => n + 1) }), [shell]);
  const program = useRef(null);
  const previewProgram = useRef(null);
  // The AI tools' requests (capture:command 'remote', the Recorder's): the latest handler, the recording they asked for, and the take's path.
  const remoteRef = useRef(null);
  // Opened for another window (A7): what for, what it sent, the line back to it, and the answer waited for at Stop.
  const into = useMemo(() => (kind === 'recorder' ? readInto() : null), [kind]);
  const [brief, setBrief] = useState(null);
  const link = useRef(null);
  const handed = useRef(null);
  const finished = useRef(false);
  const pointerOn = useRef(false);
  const remote = useRef(null);
  const takenRef = useRef(null);

  const holding = Boolean(air);
  const airStack = air || stack;
  const onAirRef = useRef(null);
  onAirRef.current = airStack.scene;

  /* ── the setup, the stack and the settings, kept ───────────────────── */

  const putSetup = useCallback((doc) => { setupRef.current = doc; setSetupState(doc); }, []);
  const persist = useCallback((nextStack, nextSettings) => {
    const doc = setupRef.current;
    if (!doc || doc.tools) return;
    const fresh = { ...doc, stack: nextStack, settings: setupSettings(nextSettings) };
    setupRef.current = fresh;
    store.save(fresh);
  }, [store]);

  const change = useCallback((patch) => {
    const next = normaliseStudio({ ...settingsRef.current, ...patch });
    settingsRef.current = next;
    setSettings(next);
    if (setupRef.current?.tools) return;
    const names = Object.keys(patch);
    if (names.some((k) => SETUP_KEYS.includes(k))) persist(stackRef.current, next);
    if (names.some((k) => GLOBAL_KEYS.includes(k))) {
      globalRef.current = globalSettings(next);
      store.setGlobal(globalRef.current);
    }
  }, [persist, store]);

  const putStack = useCallback((next) => { stackRef.current = next; setStack(next); }, []);

  /**
   * The stack changed by `fn`; kept unless `keep` is false (a drag keeps it
   * once, at its end). What is changed in Set up can be undone; `record: false`
   * for what is not arranging: a scene picked, a sound changed.
   */
  const changeStack = useCallback((fn, keep = true, { record = true } = {}) => {
    const before = stackRef.current;
    const next = normaliseStack(fn(before));
    putStack(next);
    if (record && panelRef.current === 'setup') {
      if (!keep) {
        if (!pendingBefore.current) pendingBefore.current = before;
      } else {
        const from = pendingBefore.current || before;
        pendingBefore.current = null;
        if (JSON.stringify(from) !== JSON.stringify(next)) {
          history.current = historyPush(history.current, from);
          setHistoryTick((n) => n + 1);
        }
      }
    }
    if (keep) persist(next, settingsRef.current);
  }, [persist, putStack]);

  /** An arrangement changed in Set up: the frame goes with the scene it shows. */
  const changeFrame = useCallback((fn, keep = true) => changeStack((s) => followFrame(fn(s)), keep), [changeStack]);

  const undo = useCallback(() => {
    const r = historyUndo(history.current, stackRef.current);
    if (!r) return;
    history.current = r.history;
    putStack(normaliseStack(r.stack));
    persist(stackRef.current, settingsRef.current);
    setHistoryTick((n) => n + 1);
  }, [persist, putStack]);
  const redo = useCallback(() => {
    const r = historyRedo(history.current, stackRef.current);
    if (!r) return;
    history.current = r.history;
    putStack(normaliseStack(r.stack));
    persist(stackRef.current, settingsRef.current);
    setHistoryTick((n) => n + 1);
  }, [persist, putStack]);

  /** A setup made the one in use: its sources found again here, then the panel it needs. */
  const applySetup = useCallback((doc, { first = false, list = sourcesRef.current, devs = devicesRef.current } = {}) => {
    const r = rematch(doc, list || [], devs.labelled ? devs : null);
    const joined = joinSettings(r.setup.settings, globalRef.current);
    const paths = pathsRef.current;
    if (!joined.folder && paths?.videos) joined.folder = `${paths.videos}${paths.separator || '\\'}${t('Rutba Recordings')}`;
    const held = isLive && r.setup.stack.previewFirst ? r.setup.stack : null;
    history.current = emptyHistory();
    pendingBefore.current = null;
    putSetup(r.setup);
    settingsRef.current = joined;
    setSettings(joined);
    putStack(r.setup.stack);
    setAir(held);
    setMissing(r.missing);
    setPickRow(null);
    setSaved(null);
    const go = first || r.missing.some((m) => m.scope === 'item');
    setPanel(go ? 'setup' : 'main');
    setSelected(r.missing.find((m) => m.scope === 'item')?.id || null);
    persist(r.setup.stack, joined);
  }, [isLive, persist, putSetup, putStack]);

  // What there is to show: the screens and windows, the cameras and microphones.
  const lastSources = useRef(0);
  const loadSources = useCallback(async () => {
    lastSources.current = Date.now();
    let list = [];
    try { list = (await shell.capture.sources()) || []; } catch { /* none to give */ }
    setSources(list);
    const devs = await StudioMedia.devices();
    setDevices(devs);
    return { list, devs };
  }, [shell]);

  const openFileRef = useRef(null);
  // The setups as they were left, or, the first time, what the Recorder and Live Broadcast held before setups.
  useEffect(() => {
    let alive = true;
    (async () => {
      const { list, devs } = await loadSources();
      const paths = await shell.app.paths().catch(() => null);
      if (!alive) return;
      pathsRef.current = paths;
      let idx = await store.index();
      let global = await store.global();
      let first = false;
      if (!idx.setups.length) {
        const [recorderStack, recorderStudio, broadcastStack, broadcastStudio] = await Promise.all(['recorder.stack', 'recorder.studio', 'broadcast.stack', 'broadcast.studio'].map((k) => store.get(k)));
        const m = migrate({
          recorderStack, recorderStudio, broadcastStack, broadcastStudio,
          screen: list.find((x) => x.kind === 'screen') || null, hasCamera: devs.cameras.length > 0, names: SCENE_NAMES(), now: Date.now(),
          label: { recording: t('My recording setup'), live: t('My live setup'), first: t('My setup') },
        });
        for (const s of m.setups) await store.save(s);
        await store.changeIndex(() => m.index);
        if (!global) { global = m.global; await store.setGlobal(global); }
        idx = m.index;
        first = m.setups[0]?.id === 'setup-first';
      }
      if (!alive) return;
      globalRef.current = global || {};
      setIndex(idx);
      const id = openingSetup(idx, kind);
      const doc = id ? await store.load(id) : null;
      if (!alive || !doc) return;
      applySetup(doc, { first, list, devs });
      setIndex(await store.used(id, kind));
      setLoaded(true);
      if (boot?.file) openFileRef.current?.(boot.file);
    })();
    return () => { alive = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, kind, store]);

  // Sources that come and go while nothing is running: found again as they appear.
  const busyRef = useRef(false);
  useEffect(() => {
    const doc = setupRef.current;
    if (!loaded || !doc || doc.tools || busyRef.current || !sources) return;
    const devs = devices.labelled ? devices : null;
    const r = rematch({ ...doc, stack: stackRef.current }, sources, devs);
    setMissing(r.missing);
    if (JSON.stringify(r.setup.stack) !== JSON.stringify(stackRef.current)) changeStack(() => r.setup.stack, true, { record: false });
    setAir((a) => (a ? rematch({ ...doc, stack: a }, sources, devs).setup.stack : a));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sources, devices, loaded]);
  useEffect(() => {
    const onFocus = () => { if (!busyRef.current && Date.now() - lastSources.current > SOURCES_EVERY_MS) loadSources(); };
    const onDevices = () => { StudioMedia.devices().then(setDevices); };
    window.addEventListener('focus', onFocus);
    navigator.mediaDevices?.addEventListener?.('devicechange', onDevices);
    return () => {
      window.removeEventListener('focus', onFocus);
      navigator.mediaDevices?.removeEventListener?.('devicechange', onDevices);
    };
  }, [loadSources]);

  // The programs, drawn from the start, and let go with the window: the one on air, and the next shot's.
  useEffect(() => {
    const p = new Program({ media, canvas: canvases.air });
    program.current = p;
    p.onTick = () => media.revive(stackRef.current, { fps: settingsRef.current.fps, size: settingsRef.current.size });
    p.set(settingsRef.current);
    p.start();
    const next = new Program({ media, canvas: canvases.next });
    previewProgram.current = next;
    next.set(settingsRef.current);
    return () => { p.dispose(); next.dispose(); media.dispose(); };
  }, [media, canvases]);
  useEffect(() => { program.current?.set(settings); previewProgram.current?.set(settings); }, [settings]);
  useEffect(() => {
    program.current?.setStack(airStack, { fade: fadeNext.current });
    fadeNext.current = 0;
  }, [airStack]);
  useEffect(() => {
    const next = previewProgram.current;
    if (!next) return;
    next.setStack(stack);
    if (holding) next.start();
    else next.stopDrawing();
  }, [stack, holding]);

  // The sources open as the stacks say (the next shot's and the one on air). Sound goes as it does on air: a picked
  // Back soon does not bring its music in until it is taken.
  const syncItems = holding ? [...stack.items, ...air.items.filter((it) => !stack.items.some((x) => x.id === it.id))] : stack.items;
  useEffect(() => {
    if (!loaded || !sources) return;
    const items = stack.cameraOff ? syncItems.map((it) => (it.kind === 'camera' ? { ...it, live: false } : it)) : syncItems;
    media.sync({ ...stack, items }, sources, { fps: settings.fps, size: settings.size });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [media, loaded, sources, stack, air, settings.fps, settings.size]);
  const soundFrom = useMemo(() => (holding ? { ...stack, sound: air.sound, music: air.music } : stack), [holding, stack, air]);
  const soundRef = useRef(soundFrom);
  soundRef.current = soundFrom;
  const soundKey = JSON.stringify([soundFrom.sounds, soundFrom.sound, soundFrom.music]);
  useEffect(() => {
    if (!loaded) return;
    media.syncSounds(soundRef.current).then(() => StudioMedia.devices().then(setDevices));
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
    changeFrame((s) => { const r = addItem(s, { kind: 'picture', path: picked[0] }, { aspect: 1 }); added = r.id; return r.stack; });
    setSelected(added);
  };

  const barAct = {
    live: (on) => selItem && act.live(selItem.id, on),
    clip: () => setClipping((v) => !v),
    whole: () => selItem && changeFrame((s) => clipItem(s, selItem.id, null, { aspect: aspectOf(media, selItem, sources) })),
    shape: (shape) => selItem && changeFrame((s) => shapeItem(s, selItem.id, shape, { aspect: aspectOf(media, selItem, sources) })),
    place: (place) => selItem && changeFrame((s) => placeItem(s, selItem.id, place, { aspect: aspectOf(media, selItem, sources) })),
    mirror: () => selItem && changeStack((s) => updateItem(s, selItem.id, { mirror: !selItem.mirror })),
    text: (patch) => selItem && changeStack((s) => updateItem(s, selItem.id, patch)),
    change: () => selItem && choosePicture(selItem.id),
    copy: () => { if (!selItem) return; let id; changeFrame((s) => { const r = duplicateItem(s, selItem.id); id = r.id; return r.stack; }); setSelected(id); },
    remove: () => { if (!selItem) return; changeFrame((s) => removeItem(s, selItem.id)); setSelected(null); },
    sound: (patch) => selSound && act.sound(selSound.id, patch),
  };

  const act = {
    copyOf: (id) => { let added; changeFrame((s) => { const r = duplicateItem(s, id); added = r.id; return r.stack; }); setSelected(added); },
    removeItem: (id) => { changeFrame((s) => removeItem(s, id)); setSelected((sel) => (sel === id ? null : sel)); },
    rename: (id, label) => changeStack((s) => updateItem(s, id, { label })),
    show: (id, on) => changeFrame((s) => {
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
    move: (id, at) => changeFrame((s) => {
      const from = s.items.findIndex((x) => x.id === id);
      return moveItem(s, id, from >= 0 && from < at ? at - 1 : at);
    }),
    addSpare: (sp, { on = true, live = true, at = null, select = false } = {}) => {
      let added = null;
      changeFrame((s) => {
        const r = addItem(s, { kind: sp.kind, ref: sp.ref, name: sp.name, on, live, shape: 'rect' }, { aspect: sp.aspect });
        added = r.id;
        return at == null ? r.stack : moveItem(r.stack, r.id, at);
      });
      if (select && added) setSelected(added);
    },
    sound: (id, patch) => changeStack((s) => updateSound(s, id, patch), true, { record: false }),
    addSound: (sp) => changeStack((s) => ({ ...s, sounds: [...s.sounds, { id: newId('mic', s.sounds.map((x) => x.id)), kind: 'mic', ref: sp.ref, name: '', live: true, muted: false, gain: 1 }] })),
    addItems: () => [
      ...[['screen', t('Screen')], ['window', t('A window')]].map(([k, label]) => ({
        label, icon: k,
        items: spare.filter((x) => x.kind === k).length ? spare.filter((x) => x.kind === k).map((x) => ({ label: x.name, run: () => act.addSpare(x, { select: true }) })) : [{ label: t('None to add'), disabled: true }],
      })),
      { label: t('Camera'), icon: 'camera', items: spare.filter((x) => x.kind === 'camera').length ? spare.filter((x) => x.kind === 'camera').map((x) => ({ label: x.name, run: () => act.addSpare(x, { select: true }) })) : [{ label: t('None to add'), disabled: true }] },
      ...(spareSounds.length ? [{ label: t('Microphone'), icon: 'mic', items: spareSounds.map((x) => ({ label: x.name, run: () => act.addSound(x) })) }] : []),
      '-',
      { label: t('Title'), icon: 'textbox', run: () => { let id; changeFrame((s) => { const r = addItem(s, { kind: 'title', text: t('Your name, or the title') }); id = r.id; return r.stack; }); setSelected(id); } },
      { label: t('Picture or Logo…'), icon: 'picture', run: () => choosePicture(null) },
      { label: t('Web page…'), icon: 'globe', run: () => setAskWeb(true) },
      ...(selItem && ['screen', 'window', 'camera'].includes(selItem.kind) ? [{ label: t('Another copy of {name}', { name: itemName(selItem) }), icon: 'duplicate', run: () => barAct.copy() }] : []),
      '-',
      { label: t('Refresh the list'), icon: 'refresh', run: loadSources },
    ],
  };

  /** A source chosen for what could not be found, or chosen again. */
  const chooseSource = (it, x) => {
    const patch = { ref: x.id, name: x.name, was: '', size: x.display?.width ? { w: x.display.width, h: x.display.height } : null };
    const apply = () => {
      changeStack((s) => updateItem(s, it.id, patch));
      setAir((a) => (a && !liveRef.current ? normaliseStack(updateItem(a, it.id, patch)) : a));
      setMissing((list) => list.filter((m) => m.id !== it.id));
    };
    if (it.kind === 'camera' && livePhase !== 'off' && airStack.items.some((o) => o.id === it.id && o.on)) setCameraAsk({ ...x, apply });
    else apply();
  };
  const withItem = (st, id, patch) => updateItem(st, id, patch);
  /** A place typed in: a new width or height carries the other with it while the source keeps its shape. */
  const withBox = (st, id, patch) => updateItem(st, id, (it) => {
    const next = { ...it.box, ...patch };
    if (it.keep && it.box.w > 0 && it.box.h > 0) {
      const ratio = it.box.w / it.box.h;
      if (patch.w != null && patch.h == null) next.h = patch.w / ratio;
      if (patch.h != null && patch.w == null) next.w = patch.h * ratio;
    }
    return { ...it, box: clampBox(next) };
  });

  const camera = named.items.find((it) => it.kind === 'camera') || null;
  const setCameraTo = (c) => {
    if (!camera) return;
    const patch = { ref: c.id, name: c.name };
    changeStack((s) => updateItem(s, camera.id, patch), true, { record: false });
    setAir((a) => (a ? normaliseStack(updateItem(a, camera.id, patch)) : a));
  };
  const useCamera = (c) => {
    const shown = airStack.items.some((it) => it.id === camera?.id && it.on);
    if (livePhase !== 'off' && shown) setCameraAsk(c);
    else setCameraTo(c);
  };
  const useMic = (m) => changeStack((s) => {
    const mic = s.sounds.find((x) => x.kind === 'mic');
    return mic ? updateSound(s, mic.id, { ref: m.id, name: m.name, live: true }) : s;
  }, true, { record: false });
  const toggleCamera = () => {
    const off = !stackRef.current.cameraOff;
    changeStack((s) => ({ ...s, cameraOff: off }), true, { record: false });
    setAir((a) => (a ? { ...a, cameraOff: off } : a));
  };

  /* ── the frame in Set up: select, move, resize, clip ───────────────── */

  const editFrame = () => editFrameRef.current;
  const frameAt = (e) => {
    const r = editFrame().getBoundingClientRect();
    return { fx: (e.clientX - r.left) / r.width, fy: (e.clientY - r.top) / r.height };
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
        if (crop) changeFrame((st) => clipItem(st, it.id, crop, { aspect: aspectOf(media, it, sources) }));
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
    const others = s.items.filter((it) => it.on && it.id !== hit.id).map((it) => it.box);
    const move = (ev) => {
      const q = frameAt(ev);
      const dx = q.fx - p.fx;
      const dy = q.fy - p.fy;
      let box;
      // Alt lets go of the snapping; the guides show what it is pulled to.
      const free = ev.altKey;
      if (!handle) {
        const moved = { ...start, x: start.x + dx, y: start.y + dy };
        const snapped = free ? { box: moved, guides: [] } : snapMove(moved, others);
        box = snapped.box;
        setGuides(snapped.guides);
      } else {
        // The corner opposite the one held stays where it is; Keep its shape holds the proportions, and Shift goes the other way for once.
        const left = handle.includes('w');
        const top = handle.includes('n');
        const ax = left ? start.x + start.w : start.x;
        const ay = top ? start.y + start.h : start.y;
        const at = free ? { x: q.fx, y: q.fy, guides: [] } : snapPoint(q.fx, q.fy, others);
        let w = Math.max(0.04, left ? ax - at.x : at.x - ax);
        let h = Math.max(0.04, top ? ay - at.y : at.y - ay);
        if ((hit.keep !== false) !== ev.shiftKey || hit.shape === 'circle') { if (w / h > aspect) h = w / aspect; else w = h * aspect; }
        box = { x: left ? ax - w : ax, y: top ? ay - h : ay, w, h };
        setGuides(at.guides);
      }
      changeFrame((st) => updateItem(st, hit.id, { box: clampBox(box) }), false);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      setGuides([]);
      changeFrame((st) => st);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };
  const onFrameKey = (e) => {
    if (!selItem || !['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(e.key)) return;
    e.preventDefault();
    const step = e.shiftKey ? 0.02 : 0.002;
    const dx = e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0;
    const dy = e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0;
    changeFrame((s) => updateItem(s, selItem.id, (it) => ({ ...it, box: clampBox({ ...it.box, x: it.box.x + dx, y: it.box.y + dy }) })));
  };

  /* ── scenes, and switching between them ────────────────────────────── */

  const special = useMemo(() => (isLive ? [{ ...backSoonScene(SCENE_NAMES()), back: true, own: false }] : [{ ...soundOnlyScene(SCENE_NAMES()), soundOnly: true, own: false }]), [isLive]);
  const scenes = useMemo(() => [...stack.scenes.map((sc) => ({ ...sc, own: true })), ...special], [stack.scenes, special]);
  const scenesRef = useRef(scenes);
  scenesRef.current = scenes;
  const fadeMs = () => (stackRef.current.transition === 'fade' ? FADE_MS : 0);
  const sceneLost = (sc) => sceneMissing(sc, missing);
  const unsetScene = (sc) => (sc.layers || []).some((l) => { const it = stack.items.find((x) => x.id === l.id); return it?.kind === 'window' && !it.ref; });

  /**
   * A shot put up. By hand while something is held on air, it is the next
   * shot, and goes on air with Take; otherwise, and always when it comes from
   * Follow the speaker or the rotation, it goes on air at once, cut or faded
   * as the transition says.
   */
  const cutTo = useCallback((sc, { automatic = false } = {}) => {
    if (!sc) return;
    if (holding && automatic) {
      fadeNext.current = fadeMs();
      setAir((a) => normaliseStack(applyScene(a || stackRef.current, sc)));
      return;
    }
    if (!holding) fadeNext.current = fadeMs();
    changeStack((s) => applyScene(s, sc), true, { record: false });
  }, [holding, changeStack]);

  const setAuto = useCallback((auto) => changeStack((s) => ({ ...s, auto }), true, { record: false }), [changeStack]);

  /** A shot picked by hand: Follow the speaker and the rotation give way to it. A screen or a window asks which, when there is a choice. */
  const pickScene = useCallback((sc) => {
    if (stackRef.current.auto !== 'off') setAuto('off');
    cutTo(sc);
    const layers = sc.layers || [];
    const one = layers.length === 1 ? stackRef.current.items.find((x) => x.id === layers[0].id) : null;
    if (one && (one.kind === 'screen' || one.kind === 'window')) {
      const choices = (sourcesRef.current || []).filter((x) => x.kind === one.kind);
      setPickRow(!one.ref || (choices.length > 1 && !busyRef.current) ? { kind: one.kind, item: one.id, current: one.ref } : null);
    } else {
      setPickRow(null);
    }
  }, [cutTo, setAuto]);

  const choosePick = (x) => {
    const patch = { ref: x.id, name: x.name, was: '', size: x.display?.width ? { w: x.display.width, h: x.display.height } : null };
    changeStack((s) => updateItem(s, pickRow.item, patch), true, { record: false });
    setAir((a) => (a && !busyRef.current ? normaliseStack(updateItem(a, pickRow.item, patch)) : a));
    setPickRow(null);
  };

  /** The next shot on air: at once (Cut) or dissolved (Fade). */
  const take = useCallback((how) => {
    fadeNext.current = how === 'fade' ? FADE_MS : 0;
    setAir(stackRef.current);
  }, []);

  const togglePreview = () => {
    const st = stackRef.current;
    if (st.previewFirst) {
      // Put away, the shot on air stays on air; the one waiting is let go.
      fadeNext.current = 0;
      changeStack((s) => ({ ...(air || s), sounds: s.sounds, previewFirst: false }), true, { record: false });
      setAir(null);
      return;
    }
    setAir(st);
    changeStack((s) => ({ ...s, previewFirst: true }), true, { record: false });
  };

  const live = livePhase !== 'off';
  const liveRef = useRef(live);
  liveRef.current = live;

  /** Between the panels. Set up while live works on Next and leaves the picture on air alone; back from it, what was changed waits in Next for Take. */
  const goPanel = useCallback((to) => {
    if (to === panelRef.current) return;
    if (to === 'setup') {
      if (liveRef.current) setAir((a) => a || stackRef.current);
    } else {
      setClipping(false);
      const st = stackRef.current;
      if (!st.previewFirst) {
        setAir((a) => {
          if (!a) return a;
          if (!liveRef.current) return null;
          const seen = (x) => JSON.stringify([x.items, x.video, x.scene]);
          if (seen(a) === seen(st)) return null;
          changeStack((s) => ({ ...s, previewFirst: true }), true, { record: false });
          return a;
        });
      }
    }
    setPanel(to);
  }, [changeStack]);
  const openSetupAt = useCallback((id) => { goPanel('setup'); setSelected(id); }, [goPanel]);

  // The microphones there are to pair a shot with: those on and heard.
  const mics = named.sounds.filter((s) => s.kind === 'mic');
  const cameraShots = scenes.filter((sc) => sc.layers?.length === 1 && stack.items.find((it) => it.id === sc.layers[0].id)?.kind === 'camera');
  const auto = stack.auto;

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

  // The timed rotation: the chosen shots in turn, or the cameras' and the wide shot when none are chosen, or every scene.
  useEffect(() => {
    if (auto !== 'rotate') return undefined;
    const tick = setInterval(() => {
      const there = scenesRef.current.filter((sc) => sc.own);
      const items = stackRef.current.items;
      const chosen = stackRef.current.rotate.keys.filter((k) => there.some((sc) => sc.key === k));
      const cams = there.filter((sc) => sc.key === 'wide' || (sc.layers?.length === 1 && items.find((it) => it.id === sc.layers[0].id)?.kind === 'camera')).map((sc) => sc.key);
      const keys = chosen.length > 1 ? chosen : cams.length > 1 ? cams : there.map((sc) => sc.key);
      const next = nextInRotation(keys, onAirRef.current);
      if (next && next !== onAirRef.current) cutTo(there.find((sc) => sc.key === next), { automatic: true });
    }, stack.rotate.every * 1000);
    return () => clearInterval(tick);
  }, [auto, stack.rotate.every, cutTo]);

  const startFollow = () => {
    const there = new Set(scenes.map((sc) => sc.key));
    let pairs = Object.fromEntries(Object.entries(stack.follow).filter(([k]) => there.has(k)));
    // Each camera with its own microphone, in their orders, when there are as many of one as of the other.
    if (!Object.keys(pairs).length && cameraShots.length > 1 && cameraShots.length === mics.length) {
      pairs = Object.fromEntries(cameraShots.map((sc, i) => [sc.key, mics[i].id]));
      changeStack((s) => ({ ...s, follow: pairs }), true, { record: false });
      toast(cameraShots.map((sc, i) => `${sc.name}: ${mics[i].name}`).join(' · '), { ms: 6000 });
    }
    if (!Object.keys(pairs).length) {
      toast(t('Follow the speaker needs a microphone for each shot: right-click a shot in Set up to pair it with one.'), { ms: 7000 });
      return;
    }
    setAuto('follow');
  };
  const startRotate = (every) => changeStack((s) => ({ ...s, auto: 'rotate', rotate: { ...s.rotate, every } }), true, { record: false });

  const saveAsScene = () => changeStack((s) => saveScene(s, t('Scene {n}', { n: s.scenes.length + 1 })));
  const moveSceneBy = (from, onto) => changeStack((s) => moveScene(s, from, s.scenes.findIndex((x) => x.key === onto)));
  const sceneMenuItems = (sc) => {
    if (!sc.own) return [];
    const paired = stack.follow[sc.key] || '';
    const inRotation = stack.rotate.keys.includes(sc.key);
    return [
      { label: t('Rename…'), icon: 'textbox', run: () => setSceneAsk({ what: 'rename', sc }) },
      { label: t('Duplicate'), icon: 'duplicate', run: () => changeStack((s) => duplicateScene(s, sc.key, t('{name} copy', { name: sc.name }))) },
      { label: t('Update from the frame'), icon: 'refresh', run: () => changeStack((s) => updateSceneFromFrame(s, sc.key)) },
      { label: t('Delete'), icon: 'trash', run: () => setSceneAsk({ what: 'delete', sc }) },
      ...(isLive ? [
        '-',
        { label: t('Show when this microphone is heard'), icon: 'mic', items: [
          ...mics.map((m) => ({ label: m.name, checked: paired === m.id, run: () => changeStack((s) => ({ ...s, follow: { ...s.follow, [sc.key]: m.id } })) })),
          { label: t('No microphone'), checked: !paired, run: () => changeStack((s) => { const f = { ...s.follow }; delete f[sc.key]; return { ...s, follow: f }; }) },
        ] },
        { label: t('In the rotation'), icon: 'refresh', checked: inRotation, run: () => changeStack((s) => ({ ...s, rotate: { ...s.rotate, keys: inRotation ? s.rotate.keys.filter((k) => k !== sc.key) : [...s.rotate.keys, sc.key] } })) },
      ] : []),
    ];
  };

  /* ── the setups ────────────────────────────────────────────────────── */

  const screenNow = () => (sourcesRef.current || []).find((x) => x.kind === 'screen') || null;
  const takenIds = () => (index?.setups || []).map((x) => x.id);
  const takenNames = () => (index ? menuOrder(index).map((x) => x.name) : []);

  const openSetup = useCallback(async (id) => {
    const doc = await store.load(id);
    if (!doc) return;
    applySetup(doc);
    setIndex(await store.used(id, kind));
  }, [store, applySetup, kind]);

  const newSetupNamed = async (name) => {
    const doc = newSetup({ id: setupId(takenIds(), Date.now()), name: uniqueName(name, takenNames()), app: kind, now: Date.now(), screen: screenNow(), hasCamera: devicesRef.current.cameras.length > 0, names: SCENE_NAMES(), settings: setupSettings(settingsRef.current) });
    await store.add(doc);
    applySetup(doc, { first: true });
    setIndex(await store.used(doc.id, kind));
  };
  const duplicateNamed = async (name) => {
    const doc = duplicateSetup({ ...setupRef.current, stack: stackRef.current, settings: setupSettings(settingsRef.current) }, { id: setupId(takenIds(), Date.now()), name: uniqueName(name, takenNames()), now: Date.now() });
    await store.add(doc);
    applySetup(doc);
    setIndex(await store.used(doc.id, kind));
  };
  const renameSetup = async (name) => {
    const doc = { ...setupRef.current, stack: stackRef.current, name };
    putSetup(doc);
    setIndex(await store.rename(doc));
  };
  const deleteSetup = async () => {
    const id = setupRef.current.id;
    const idx = await store.remove(id);
    setIndex(idx);
    const nextId = openingSetup(idx, kind);
    if (nextId) openSetup(nextId);
  };

  const saveToFile = async () => {
    const doc = { ...setupRef.current, stack: stackRef.current, settings: setupSettings(settingsRef.current) };
    const path = await shell.dialog.save({ title: t('Save to a file'), defaultPath: `${doc.name}.${SETUP_EXT}`, filters: SETUP_FILTERS() });
    if (!path) return;
    try {
      const absent = await saveSetupFile({ shell, path: new RegExp(`\\.${SETUP_EXT}$`, 'i').test(path) ? path : `${path}.${SETUP_EXT}`, setup: doc, streams: doc.streams });
      toast(absent.length ? t('Saved. {n} files could not be included.', { n: absent.length }) : t('Saved {name}', { name: basename(path) }), { tone: absent.length ? undefined : 'good', ms: 5000 });
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 8000 });
    }
  };
  const openPath = async (path) => {
    try {
      const opened = await openSetupFile({ shell, path, store, sources: sourcesRef.current || [], devices: devicesRef.current.labelled ? devicesRef.current : null, app: kind });
      // A setup saved for the other app opens there.
      if (opened.savedApp && opened.savedApp !== kind && !boot?.handed) {
        shell.win.create({ app: opened.savedApp, file: path, query: { handed: 1 } });
        shell.win.close({ force: true });
        return;
      }
      setFileOpening(opened);
    } catch {
      toast(t('That setup file could not be read.'), { tone: 'bad', ms: 6000 });
    }
  };
  openFileRef.current = openPath;
  const openFromFile = async () => {
    const picked = await shell.dialog.open({ title: t('Open a file'), filters: SETUP_FILTERS() });
    if (picked?.[0]) openPath(picked[0]);
  };
  const addOpenedFile = async () => {
    const doc = fileOpening.setup;
    setFileOpening(null);
    await store.add(doc);
    applySetup(doc);
    setIndex(await store.used(doc.id, kind));
  };

  /* ── the disk ──────────────────────────────────────────────────────── */

  const camerasShown = airStack.cameraOff ? 0 : airStack.items.filter((it) => it.on && it.kind === 'camera').length;
  // Each camera with a file of its own writes as well as the picture of the whole.
  const ownFiles = into ? into.into !== 'slide' : Boolean(settings.isolated);
  const cameraFiles = ownFiles ? camerasShown : 0;
  const rate = useMemo(() => bytesPerSecond({ fps: settings.fps, cap: settings.size, quality: settings.quality, camera: false, mic: true, system: false }, { screen: settings.size === 720 ? { width: 1280, height: 720 } : { width: 1920, height: 1080 } })
    + cameraFiles * bytesPerSecond({ fps: 30, cap: 720, quality: settings.quality, camera: false, mic: false, system: false }, { screen: { width: 1280, height: 720 } }), [cameraFiles, settings.fps, settings.size, settings.quality]);
  const checkDisk = useCallback(async () => {
    // The drive being written: a recording into a slide goes to the temporary folder, not the recordings folder.
    const where = (phaseRef.current === 'recording' && takenRef.current) || settings.folder;
    if (!where) return null;
    try {
      const { free } = await shell.capture.space({ path: where });
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

  const getTall = useCallback(() => {
    const size = settingsRef.current.size;
    const was = tallPicture.current;
    if (was && (was.inUse() || was.canvas.width === tallSize(size).width)) return was;
    tallPicture.current = makeTallProgram({ wide: program.current, media, size });
    return tallPicture.current;
  }, [media]);

  const recKind = recordKind(airStack);
  const begin = useCallback(async () => {
    try {
      setSaved(null);
      const current = program.current?.stack || stackRef.current;
      const soundOnly = recordKind(current) === 'sound';
      const set = settingsRef.current;
      // For Video every source is a file of its own and there is no other; for a slide one file; on its own the Recorder makes one, and the sources too when asked.
      const takes = (into ? into.into !== 'slide' : set.isolated && !soundOnly) ? await makeTakes({ stack: current, media, shell, settings: set }) : [];
      let path = null;
      let type = null;
      if (into?.into === 'slide') {
        type = recordingType();
        ({ path } = await shell.fs.temp({ ext: type.startsWith('video/mp4') ? 'mp4' : 'webm' }));
      } else if (!into) {
        const name = soundOnly ? takeFileName(t('Sound'), new Date(), 'weba') : takeFileName(t('Recording'), new Date(), 'webm');
        ({ path } = await shell.capture.take({ dir: set.folder, name }));
      }
      takenRef.current = path;
      if (liveRecord.current && !soundOnly && settingsRef.current.liveCopyShape === 'tall') {
        const tall = getTall();
        tall.attach();
        copyProgram.current = tall;
      }
      // A disk that stops taking the recording ends it now: every frame after that would be lost unseen.
      // Set after the tall copy is in place, for it is that copy which is recording then.
      recorderNow().onWriteFailed = () => { if (phaseRef.current === 'recording') stopRef.current?.(); };
      await recorderNow().record({ shell, path, soundOnly, type, takes });
      pointerOn.current = false;
      const screenTake = takes.find((k) => k.role === 'screen');
      if (into && screenTake?.ref) {
        const pt = await shell.capture.pointer({ on: true, id: screenTake.ref, fps: set.fps }).catch(() => null);
        pointerOn.current = Boolean(pt?.tracking);
        if (!pt?.tracking) toast(t('The pointer is not written down over another program\'s window. Record its screen to keep it.'), { ms: 7000 });
      }
      setPhase('recording');
      setPaused(false);
      warned.current = false;
      const c = await shell.capture.controls({ on: true, paused: false, minimise: settingsRef.current.minimise }).catch(() => null);
      setKeys(c);
    } catch (err) {
      setPhase('setup');
      toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
    }
  }, [shell, toast, into, media]);

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

  /* ── for another window (A7) ───────────────────────────────────────── */

  useEffect(() => {
    if (!into) return undefined;
    const line = openLink(into.channel, (m) => {
      if (m.type === 'brief') setBrief(m);
      else if (m.type === 'took') handed.current?.(true);
      else if (m.type === 'gone') { link.current = null; handed.current?.(false); }
    });
    link.current = line;
    line.post({ type: 'ready' });
    const bye = () => { if (link.current && !finished.current) line.post({ type: 'cancel' }); };
    window.addEventListener('pagehide', bye);
    if (into.title) shell.win.setTitle({ title: t('Recording into {name}', { name: into.title }) }).catch(() => {});
    return () => { window.removeEventListener('pagehide', bye); link.current = null; line.close(); };
  }, [into, shell]);

  /** The takes given to the window that asked. Answers false if nobody took them: they are then shown here, as a recording is. */
  const handOver = useCallback((done, pointer) => new Promise((resolve) => {
    if (!link.current) { resolve(false); return; }
    const marks = pointer && (pointer.samples?.length || pointer.clicks?.length) ? pointer : null;
    const firstScreen = done.takes.findIndex((k) => k.role === 'screen');
    const takes = done.takes.length ? done.takes.map((k, i) => (i === firstScreen && marks ? { ...k, pointer: marks } : k)) : [{
      role: 'screen', path: done.path, contentType: done.type.split(';')[0], ext: done.path.split('.').pop(), hasAudio: done.heard,
      width: canvases.air.width, height: canvases.air.height, duration: done.seconds, written: done.written,
    }];
    const answered = (took) => {
      clearTimeout(timer);
      handed.current = null;
      if (took) shell.win.close({ force: true }).catch(() => {});
      resolve(took);
    };
    const timer = setTimeout(() => answered(false), 4000);
    handed.current = answered;
    finished.current = true;
    link.current.post({ type: 'done', takes });
  }), [shell, canvases]);

  const setPrompterSpeed = useCallback((speed) => {
    setBrief((b) => ({ ...b, prompter: { ...b.prompter, speed } }));
    link.current?.post({ type: 'speed', value: speed });
  }, []);

  const stop = useCallback(async () => {
    if (phaseRef.current === 'countdown') { setPhase('setup'); return; }
    if (phaseRef.current !== 'recording') return;
    setPhase('stopping');
    shell.capture.controls({ on: false }).catch(() => {});
    const thumb = stillOf(canvases.air);
    try {
      const pointer = pointerOn.current ? await shell.capture.pointer({ on: false }).catch(() => null) : null;
      pointerOn.current = false;
      const done = await recorderNow().stop();
      copyProgram.current?.detach();
      copyProgram.current = null;
      if (done.failed && into) toast(t('Writing stopped: {reason}. What was written before that is kept.', { reason: done.failed }), { tone: 'bad', ms: 10000 });
      if (into && await handOver(done, pointer)) return;
      // Kept apart with no file of the whole: the first take stands for it on the card.
      const shown = done.path ? done : { ...done, path: done.takes[0]?.path || '', written: done.takes.reduce((n, k) => n + (k.written || 0), 0) };
      setSaved({ ...shown, thumb });
      if (done.failed) toast(t('Writing stopped: {reason}. What was written before that is kept.', { reason: done.failed }), { tone: 'bad', ms: 10000 });
      else toast(t('Saved {name}', { name: basename(shown.path) }), { tone: 'good', ms: 5000 });
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 8000 });
    } finally {
      setPhase('setup');
      setPaused(false);
      setElapsed(0);
    }
  }, [shell, toast, canvases, into, handOver]);
  stopRef.current = stop;

  const togglePause = useCallback(() => {
    if (phaseRef.current !== 'recording') return;
    const p = program.current;
    if (p.paused) p.resume();
    else p.pause();
    setPaused(p.paused);
    shell.capture.controls({ on: true, paused: p.paused }).catch(() => {});
    if (pointerOn.current) shell.capture.pointer({ on: true, paused: p.paused }).catch(() => {});
  }, [shell]);
  pauseRef.current = togglePause;
  useEffect(() => () => { if (pointerOn.current) shell.capture.pointer({ on: false }).catch(() => {}); }, [shell]);

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
    const tick = setInterval(() => setElapsed(recorderNow()?.elapsed || 0), 250);
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
    // The window that asked keeps a copy of what is recorded so far, to offer back after a crash.
    const copy = link.current ? setInterval(() => { if (phaseRef.current === 'recording') link.current?.post({ type: 'progress', takes: program.current.takesSoFar() }); }, COPY_EVERY_MS) : null;
    return () => { clearInterval(tick); clearInterval(watch); if (copy) clearInterval(copy); };
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
  }, [shell, toast, into, media]);

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

  // The bitrate the slowest destination can carry: the encoder follows it, and the bar says so.
  const [liveRate, setLiveRate] = useState(null);
  useEffect(() => (isLive ? shell.on('live:status', ({ destinations = [], kbps, target } = {}) => {
    setLiveStatus(destinations);
    if (kbps) setLiveRate({ kbps, target });
  }) : undefined), [shell, isLive]);
  const kbps = isLive ? liveKit.liveBitrate(settings.size, settings.fps) : 0;

  // The bar on top of the Live panel runs the broadcast; here a copy is kept as it goes.
  const barPhase = useCallback((p, since) => {
    setLivePhase(p);
    if (since) setLiveSince(since);
    if (p === 'live' && settingsRef.current.liveCopy && phaseRef.current === 'setup') { liveRecord.current = true; setPhase('starting'); begin(); }
    if (p === 'off' && liveRecord.current) { liveRecord.current = false; stopRef.current?.(); }
  }, [begin]);

  const endLive = useCallback(async () => { await liveControl.current?.end(); }, []);

  // Which streams are on is the setup's.
  const changeLiveOn = useCallback((ids) => {
    const doc = setupRef.current;
    if (!doc) return;
    const fresh = { ...doc, liveOn: ids };
    putSetup(fresh);
    if (!doc.tools) store.save(fresh);
  }, [putSetup, store]);

  const [, setLiveTick] = useState(0);
  useEffect(() => {
    if (livePhase !== 'live') return undefined;
    const tick = setInterval(() => setLiveTick((n) => n + 1), 1000);
    return () => clearInterval(tick);
  }, [livePhase]);

  // Closing the window while live is asked first; the broadcast ends with it.
  const closing = useRef(false);
  useEffect(() => {
    const onBefore = (e) => {
      const recording = phaseRef.current === 'recording' || phaseRef.current === 'stopping';
      if ((!liveRef.current && !recording) || closing.current) return;
      e.preventDefault();
      e.returnValue = false;
      setCloseAsk(true);
    };
    window.addEventListener('beforeunload', onBefore);
    return () => window.removeEventListener('beforeunload', onBefore);
  }, []);

  // Restart and update asks first about a take still being recorded, as it does about unsaved work.
  useEffect(() => {
    shell.win.setDirty({ dirty: phase === 'recording' || phase === 'stopping', name: t('Recording') }).catch(() => {});
  }, [shell, phase]);

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

  // Opened for a slide, a scene or a voice: a setup of its own, never saved and never in the menu, as the tools' is.
  const intoReady = useRef(false);
  useEffect(() => {
    if (!into || !loaded || !sources || intoReady.current) return;
    intoReady.current = true;
    setPanel('main');
    if (into.into === 'video' && into.kind !== 'voice') return;
    const voice = into.kind === 'voice';
    let st = { ...emptyStack(), sounds: stackRef.current.sounds.map((x) => (voice && x.kind === 'mic' ? { ...x, live: true, muted: false } : x)) };
    const screen = sources.find((x) => x.kind === 'screen') || sources[0];
    if (!voice && screen) st = addItem(st, { kind: screen.kind === 'window' ? 'window' : 'screen', ref: screen.id, name: screen.name }).stack;
    if (into.into === 'intro') st = addItem(st, { kind: 'camera', ref: '', shape: 'circle' }).stack;
    if (voice) st = applyScene(st, { key: SOUND_ONLY });
    putSetup({ ...setupRef.current, tools: true, name: t('This recording') });
    putStack(normaliseStack(st));
    setAir(null);
  }, [into, loaded, sources, putSetup, putStack]);

  /* ── the AI tools' Recorder (recorder_* tools, office-shell recorder-remote.js) ── */

  // A recording the tools ask for is made as a person's is: the source shown, the settings set,
  // Record pressed, without the countdown. It runs in a setup of its own, never saved and never
  // in the menu; the person's own setup, stack and settings come back when it ends.
  // What it answers goes to the AI app, in English as every tool's answer is, not the window's words.
  // Only the Recorder takes them.
  const answer = useCallback((reply, result, error = null) => {
    shell.capture.remoteReply({ reply, result, error: error ? String(error) : null }).catch(() => {});
  }, [shell]);
  useEffect(() => {
    if (kind !== 'recorder' || into) return undefined;
    shell.capture.remoteListen({ on: true }).catch(() => {});
    return () => { shell.capture.remoteListen({ on: false }).catch(() => {}); };
  }, [shell, kind]);
  const giveBack = useCallback((kept) => {
    putSetup(kept.setup);
    putStack(kept.stack);
    change(kept.settings);
  }, [putSetup, putStack, change]);
  remoteRef.current = kind !== 'recorder' || into ? null : async ({ op, args = {}, reply } = {}) => {
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
      const kept = { setup: setupRef.current, stack: stackRef.current, settings: settingsRef.current };
      remote.current = { startReply: reply, kept, item: screenAdded.id, source: src.id, folder: args.folder, triggered: false };
      setSources(list);
      putSetup({ ...kept.setup, tools: true, name: t('For the tools') });
      putStack(normaliseStack(s));
      setAir(null);
      change({ size: args.size ?? settings.size, fps: args.fps ?? settings.fps, folder: args.folder, countdown: 0 });
      // A source that gives no picture: said, and the person's own put back.
      const asked = remote.current;
      setTimeout(() => {
        if (remote.current !== asked || asked.triggered) return;
        remote.current = null;
        giveBack(asked.kept);
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
    if (r.kept) giveBack(r.kept);
  }, [phase, saved, remoteItem, answer, giveBack]);

  /* ── the keys ──────────────────────────────────────────────────────── */

  // Ctrl+1 and Ctrl+2 change panel, Ctrl+Z and Ctrl+Y undo in Set up. 1 to 9 put up a scene, B Back soon,
  // Space takes the next shot on air with the transition set.
  useEffect(() => {
    const onKey = (e) => {
      const tag = e.target?.tagName;
      const typing = tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || e.target?.isContentEditable;
      if (e.ctrlKey || e.metaKey) {
        if (e.altKey) return;
        if (e.key === '1') { e.preventDefault(); goPanel('setup'); return; }
        if (e.key === '2') { e.preventDefault(); goPanel('main'); return; }
        if (typing || panel !== 'setup') return;
        const k = e.key.toLowerCase();
        if (k === 'z' && !e.shiftKey) { e.preventDefault(); undo(); } else if (k === 'y' || (k === 'z' && e.shiftKey)) { e.preventDefault(); redo(); }
        return;
      }
      if (e.altKey || typing || panel !== 'main') return;
      if (e.key === ' ' && holding && tag !== 'BUTTON') { e.preventDefault(); take(stackRef.current.transition); return; }
      if (isLive && e.key.toLowerCase() === 'b') { e.preventDefault(); pickScene(special[0]); return; }
      const n = Number(e.key);
      const numbered = scenes.filter((sc) => sc.own);
      if (n >= 1 && n <= 9 && numbered[n - 1]) { e.preventDefault(); pickScene(numbered[n - 1]); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

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
  const busy = phase !== 'setup' || live;
  busyRef.current = busy || Boolean(setup?.tools);
  const liveFor = live && liveSince ? Date.now() - liveSince : 0;
  const anyShown = stack.items.some((it) => it.on);
  const canRecord = phase === 'setup' && Boolean(recKind) && disk?.level !== 'stop';
  const minutes = disk ? Math.floor(disk.minutes) : null;
  const diskNote = disk == null ? '' : disk.level === 'stop' ? t('There is not enough space on the disk to record.')
    : disk.minutes < 120 ? t('Room for about {minutes} minutes on the disk', { minutes: Math.max(0, minutes) })
      : t('Room for about {hours} hours on the disk', { hours: Math.floor(disk.minutes / 60) });
  const viewers = live ? liveStatus.reduce((n, d) => (Number.isFinite(d.viewers) ? (n ?? 0) + d.viewers : n), null) : null;
  const liveRateText = live && liveRate?.kbps && liveRate.kbps < liveRate.target ? `${t('{n} kb/s', { n: liveRate.kbps })} ${t('(lowered to keep up with the upload)')}` : '';
  const recordLabel = recKind === 'sound' ? t('Record sound') : t('Record');
  const stopTip = keys?.stop ? `${t('Stop')} (${keys.stop})` : t('Stop');
  const QUALITY_WORDS = { small: t('Smaller files'), standard: t('Standard'), high: t('Sharper') };
  const mainLabel = isLive ? t('Live') : t('Record');

  const top = (
    <div className="st-top" role="toolbar" aria-label={appName}>
      <SetupsButton s={{ setup, index, busy: busy || Boolean(setup?.tools), openSetup, newSetupNamed, duplicateNamed, renameSetup, deleteSetup, saveToFile, openFromFile }} />
      {into ? <span className="st-into"><Icon name="record" size={13} />{into.title ? t('Recording into {name}', { name: into.title }) : t('Recording for a slide')}</span> : null}
      {panel === 'setup' ? (
        <span className="st-undo">
          <button type="button" className="st-hist st-undo-btn" disabled={!history.current.back.length} aria-label={t('Undo')} data-tip={t('Undo')} onClick={undo}><Icon name="undo" size={15} /></button>
          <button type="button" className="st-hist st-redo-btn" disabled={!history.current.ahead.length} aria-label={t('Redo')} data-tip={t('Redo')} onClick={redo}><Icon name="redo" size={15} /></button>
        </span>
      ) : null}
      <span className="st-top-gap" />
      {!loaded ? null : <PanelSwitch value={panel} label={t('Panels')} onChange={goPanel} options={[['setup', t('Set up'), `${t('Set up')} (Ctrl+1)`], ['main', mainLabel, `${mainLabel} (Ctrl+2)`]]} />}
      <span className="st-top-gap" />
      <button type="button" className="st-gear" data-tip={settingsLabel.replace('…', '')} aria-label={settingsLabel.replace('…', '')} onClick={() => setSettingsOpen(true)}><Icon name="settings" size={16} /></button>
    </div>
  );

  const airName = scenes.find((sc) => sc.key === airStack.scene)?.name;
  const status = (
    <span className="st-status" data-phase={phase}>
      {live ? `${t('Live')} ${clock(liveFor)} · ` : ''}
      {auto === 'follow' ? `${t('Following the speaker')} · ` : auto === 'rotate' ? `${t('A new shot every {n} s', { n: stack.rotate.every })} · ` : ''}
      {isLive && airName ? `${t('On air: {name}', { name: airName })} · ` : ''}
      {recording ? `${paused ? t('Paused') : t('Recording')} · ${clock(elapsed)}` : saved ? t('Saved {name}', { name: basename(saved.path) }) : t('Ready')}
      {` · ${settings.size === 720 ? '1280 × 720' : '1920 × 1080'} · ${t('{n} frames a second', { n: settings.fps })} · ${QUALITY_WORDS[settings.quality] || ''}`}
      {camerasShown > 1 ? ` · ${t('{n} cameras', { n: camerasShown })}` : ''}
      {diskNote && !recording ? ` · ${diskNote}` : ''}
      {settings.folder ? ` · ${t('Saves to {place}', { place: settings.folder.split(/[\\/]/).slice(-2).join('\\') })}` : ''}
    </span>
  );

  const overlayShared = (
    <>
      {shotCount != null && shotCount > 0 && phase !== 'countdown' ? <div className="rec-count st-shot-count" aria-live="assertive"><span key={shotCount}>{shotCount}</span></div> : null}
      {flash ? <div key={flash} className="st-flash" aria-hidden="true" /> : null}
    </>
  );

  const shownBoxes = stack.video === 'on' ? stack.items.filter((it) => it.on && !(stack.cameraOff && it.kind === 'camera')) : [];
  const selBox = selItem && selItem.on ? selItem.box : null;
  // What the person works on, drawn over the frame in Set up: the boxes, the clip being drawn, the hints.
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
      {guides.map((g, i) => <div key={i} className={`st-guide ${g.axis}`} style={g.axis === 'x' ? { left: `${g.at * 100}%` } : { top: `${g.at * 100}%` }} />)}
      {drawing ? (
        <div className="rec-region st-drawn" style={{ left: `${Math.min(drawing.a.fx, drawing.b.fx) * 100}%`, top: `${Math.min(drawing.a.fy, drawing.b.fy) * 100}%`, width: `${Math.abs(drawing.b.fx - drawing.a.fx) * 100}%`, height: `${Math.abs(drawing.b.fy - drawing.a.fy) * 100}%` }}><span>{t('Clip')}</span></div>
      ) : null}
      {clipping && selBox && !drawing ? <div className="rec-hint">{t('Drag over the part of {name} to show.', { name: itemName(selItem) })}</div> : null}
      {stack.video === 'on' && !anyShown && loaded ? <div className="st-empty">{sources == null ? t('Looking at what is open…') : t('Show a screen, a window or a camera from the list on the left.')}</div> : null}
    </>
  );

  // Everything the panels draw from and act on.
  const s = {
    into, brief, setPrompterSpeed,
    shell, media, program: program.current, canvases, isLive, stack, named, scenes, sources, devices, settings, setup, missing, loaded,
    holding, air, airStack, onAirKey: airStack.scene, nextKey: stack.scene, anyShown,
    phase, recording, paused, elapsed, count, busy, canRecord, recordLabel, stopTip, record, stop, togglePause,
    saved, setSaved, picture, takePicture, copyPicture, pictureKey, shotCount, overlayShared,
    live, livePhase, liveFor, liveBarHost, setLiveOpen, frameShape, setFrameShape,
    change, changeStack, changeFrame, act, barAct, take, togglePreview, setAuto, startFollow, startRotate,
    pickScene, sceneLost, unsetScene, openSetupAt, sceneMenu, setSceneMenu, sceneMenuItems, moveSceneBy, saveAsScene,
    pickRow, setPickRow, choosePick, useMic, useCamera, toggleCamera,
    selected, setSelected, selItem, selSound, thumbs, chooseSource, withItem, withBox, onAirIds: live ? airStack.items.filter((it) => it.on).map((it) => it.id) : [], clipping, overlay, editFrameRef, onFramePointer, onFrameKey, setGuides,
    fileOpening, setFileOpening, addOpenedFile, chooseFolder, choosePlaceholder, chooseMusic,
  };

  return (
    <AppFrame app={app} shell={shell} title={live ? `${t('Live')} ${clock(liveFor)}` : recording ? `${paused ? t('Paused') : t('Recording')} ${clock(elapsed)}` : appName} menu={appMenu} ribbon={top} status={status}>
      <style>{RECORDER_CSS}</style>
      {brief?.prompter ? <style>{PROMPTER_CSS}</style> : null}
      <style>{STUDIO_CSS}</style>
      {isLive ? <style>{liveKit.LIVE_CSS}</style> : null}
      {isLive && loaded ? createPortal(<liveKit.LiveBar shell={shell} toast={toast} program={program} media={media} fps={settings.fps} kbps={kbps} size={settings.size} makeTall={getTall} copy={live && settings.liveCopy} blocked={phase === 'countdown' || phase === 'starting'} on={setup ? setup.liveOn : undefined} onChangeOn={changeLiveOn} onPhase={barPhase} rateText={liveRateText} viewers={viewers} control={liveControl} />, liveBarHost) : null}
      {isLive && liveOpen ? <liveKit.LiveDialog shell={shell} toast={toast} onChanged={(list) => liveControl.current?.setList(list)} onClose={() => setLiveOpen(false)} /> : null}
      <style>{FLIP_CSS}</style>
      <style>{TALL_FRAME_CSS}</style>
      {askWeb ? <WebPageDialog allowLoopback={Boolean(window.__rutbaLocalPages)} onClose={() => setAskWeb(false)} onAdd={({ url, name }) => { setAskWeb(false); changeStack((s) => addItem(s, { kind: 'web', ref: url, name, sound: false }).stack); }} /> : null}
      {settingsOpen ? (
        <StudioSettings settings={settings} setupName={setup.name} busy={busy} live={isLive} pictureKey={pictureKey} apart={!isLive && !into} title={settingsLabel.replace('…', '')} change={change} chooseFolder={chooseFolder} onClose={() => setSettingsOpen(false)} />
      ) : null}
      <OpenSetupDialog s={s} />
      {closeAsk && liveRef.current ? <ConfirmDialog title={t('End broadcast?')} body={t('You are live. Closing this window ends the broadcast.')} okLabel={t('End broadcast and close')} onOk={async () => { closing.current = true; await endLive(); await stopRef.current?.(); shell.win.close({ force: true }); }} onClose={() => setCloseAsk(false)} /> : null}
      {closeAsk && !liveRef.current ? <ConfirmDialog title={t('Stop recording?')} body={t('A recording is running. It is stopped and kept, then this window closes.')} okLabel={t('Stop and close')} onOk={async () => { closing.current = true; await stopRef.current?.(); await whileStopping(() => phaseRef.current); shell.win.close({ force: true }); }} onClose={() => setCloseAsk(false)} /> : null}
      {cameraAsk ? <ConfirmDialog tone="plain" title={t('Change the camera on air?')} body={t('The picture on air changes to {name} at once.', { name: cameraAsk.name })} okLabel={t('Change')} onOk={() => (cameraAsk.apply ? cameraAsk.apply() : setCameraTo(cameraAsk))} onClose={() => setCameraAsk(null)} /> : null}
      {sceneAsk?.what === 'rename' ? <NameDialog title={t('Rename')} label={t('Name')} value={sceneAsk.sc.name} okLabel={t('Rename')} onOk={(name) => { changeStack((st) => renameScene(st, sceneAsk.sc.key, name)); setSceneAsk(null); }} onClose={() => setSceneAsk(null)} /> : null}
      {sceneAsk?.what === 'delete' ? <ConfirmDialog title={t('Delete {name}?', { name: sceneAsk.sc.name })} body={t('The scene goes from this setup. Undo brings it back.')} okLabel={t('Delete')} onOk={() => changeStack((st) => removeScene(st, sceneAsk.sc.key))} onClose={() => setSceneAsk(null)} /> : null}

      <Content>
        {!loaded ? null : panel === 'setup' ? <SetupPanel s={s} /> : isLive ? <LivePanel s={s} /> : <RecordPanel s={s} />}
      </Content>

    </AppFrame>
  );
}
