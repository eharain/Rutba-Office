// Rutba Video: a video editor on the studio engine.
//
// The project is a studio document (docs/STUDIO.md D1): every clip a layer
// with a window on the clock, on a lane, a video's picture and sound grouped.
// The window keeps the document and its undo history, and everything it does
// to it is the engine's arithmetic (@rutba/studio/project); the stage is the
// engine's own painter at the playhead, and an export is the engine playing
// the project into a recorder with its sound. A project is saved as .rvid
// (D2) and copied to the recovery folder while it has unsaved work.
//
// Nothing here reaches the file system but through `shell`, and media only
// through the engine's transport (media.js), so a web host could give the
// same window its own.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Spacer, Chip, Empty, Button, Panel, Content, useToast, useCommands, useMenu, menuItems, formatBytes, basename, t, tn } from '@rutba/office-ui';
import { History } from '@rutba/editing/history';
import { timecode } from '@rutba/media/timeline';
import { buildPlan, startAudioPreview, soundLayers, clipFromSoundLayer, driveVideoElements, pauseVideoElements, unsupportedReason } from '@rutba/studio/host';
import {
  emptyProject, normaliseProject, addMedia, splitAt, splitAll, trimStart, trimEnd, removeClips, closeGap, copyClips,
  pasteClips, duplicateClips, unlinkClips, linkClips, addMarker, nextMarker, previousMarker, setTransition, addTitle,
  planPatches, projectMedia, projectDuration, clipById, clipsAt, sectionOf, laneOf, partnersOf, previousOnLane, setRange,
  clearRange, TITLE_PRESETS, withPartners, placeMedia, insertLane, moveLane, splitLane, brokenLocks, laneCount,
} from '@rutba/studio/project';
import { writeRvid, readRvid } from '@rutba/studio/rvid';
import { placeTakes, settleTakes } from '@rutba/studio/recorder';
import { parseCaptions, placeCaptions, captionsOf, writeSrt, writeVtt, isCaption } from '@rutba/studio/captions';
import { pointerOf, applyAutoZoom } from '@rutba/studio/pointer';
import {
  isIntro, newIntro, recordScene, setLayout, moveScene, removeScene, addScene, updateScene, useTake, applyBrand, normaliseBrand,
  setPrompterSpeed, composeIntro, sceneSpans, sceneById, sceneAt, RECORDED, LAYOUTS,
} from '@rutba/studio/intro';
import ScenesPanel, { LAYOUT_NAMES } from './intro/scenes.js';
import { BrandDialog, BRAND_KEY } from './intro/brand.js';
import { INTRO_CSS } from './intro/css.js';
import { RecorderPanel } from './recorder/panel.js';
import { AppFrame, useAppMenu, useDirtyGuard, useFileDrop } from '../../shell.js';
import { MediaLibrary, mediaKind, MEDIA_EXTENSIONS } from './media.js';
import Timeline from './timeline.js';
import Stage from './stage.js';
import VideoRibbon, { TITLE_NAMES } from './ribbon.js';
import Inspector from './inspector.js';
import { ExportDialog, ExportProgress, exportProject, exportFormat, PRESET_TAGS } from './export.js';
import { CSS } from './css.js';

/** The timeline's share of the editor's height when nothing else is chosen: the preview has the larger part. */
export const TIMELINE_SHARE = 0.36;
/** The fewest pixels the timeline keeps (its ruler and a lane) and the preview keeps (the stage and its transport bar). */
const TIMELINE_MIN = 120;
const PREVIEW_MIN = 200;

/**
 * The timeline's height in an editor `editorH` tall at `share` of it,
 * kept between its own least and what leaves the preview its least; on an
 * editor too short for both, the preview's least gives way first.
 */
export function splitHeight(editorH, share) {
  if (!(editorH > 0)) return 260;
  const most = Math.max(TIMELINE_MIN, editorH - PREVIEW_MIN);
  return Math.round(Math.max(TIMELINE_MIN, Math.min(most, editorH * share)));
}

const FPS = 30;
const AUTOSAVE_FIRST_MS = 1500;
const AUTOSAVE_EVERY_MS = 30000;

/** A clip drawn in a box of its own, which a drag on the stage moves: a title, a shape, a picture in picture or a still not set to fill the frame. */
export const standsInBox = (p) => p.type === 'text' || p.type === 'shape' || (p.type === 'video' && Boolean(p.fw)) || (p.type === 'image' && !p.fit);

/** A clip with `change` made: a field set to undefined is taken off, so a still can stop filling the frame. */
export function withChange(p, change) {
  const next = { ...p, ...change };
  for (const [k, v] of Object.entries(change)) if (v === undefined) delete next[k];
  return next;
}

/**
 * The In to Out stretch on the preview's scrub bar: the band, and an In and
 * an Out flag that drag along the bar (owner, 2026-10-16: "video start and
 * end markers"). The bar's ends sit half a thumb in from its edges.
 */
function ScrubRange({ range, duration, onBegin, onMove, docRef }) {
  const ref = React.useRef(null);
  const at = (frac) => `calc(8px + (100% - 16px) * ${Math.max(0, Math.min(1, frac))})`;
  const drag = (edge) => (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const base = docRef.current;
    const r = ref.current.getBoundingClientRect();
    let begun = false;
    const move = (ev) => {
      if (!begun) {
        begun = true;
        onBegin(edge === 'in' ? t('Move In') : t('Move Out'));
      }
      const frac = (ev.clientX - r.left - 8) / Math.max(1, r.width - 16);
      onMove(edge, Math.round(Math.max(0, Math.min(1, frac)) * duration * 100) / 100, base);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };
  return (
    <span className="vd-scrub-range" ref={ref}>
      <span className="vd-scrub-band" style={{ left: at(range.in / duration), width: `calc((100% - 16px) * ${Math.max(0, (range.out - range.in) / duration)})` }} />
      <span className="vd-scrub-edge e-in" data-edge="in" style={{ left: at(range.in / duration) }} role="slider" aria-label={t('In point')} aria-valuetext={timecode(range.in)} data-tip={t('In point {time}: drag to move it', { time: timecode(range.in) })} onPointerDown={drag('in')} />
      <span className="vd-scrub-edge e-out" data-edge="out" style={{ left: at(range.out / duration) }} role="slider" aria-label={t('Out point')} aria-valuetext={timecode(range.out)} data-tip={t('Out point {time}: drag to move it', { time: timecode(range.out) })} onPointerDown={drag('out')} />
    </span>
  );
}

/** A key for this window's recovery copy. */
const newCopyKey = () => `v${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;

export default function VideoEditor({ app, shell, boot }) {
  const toast = useToast();
  const media = useMemo(() => new MediaLibrary(), []);
  const [doc, setDoc] = useState(() => emptyProject());
  const docRef = useRef(doc);
  docRef.current = doc;
  const history = useRef(new History({ limit: 200 }));
  const [, setHistoryTick] = useState(0);
  const [selected, setSelected] = useState([]);
  const [time, setTime] = useState(0);
  const timeRef = useRef(0);
  timeRef.current = time;
  const [playing, setPlaying] = useState(false);
  const [path, setPath] = useState(null);
  const [dirty, setDirty] = useState(false);
  const [tab, setTab] = useState('home');
  const [pps, setPps] = useState(60);
  const [follow, setFollow] = useState(true);
  const [mediaVersion, setMediaVersion] = useState(0);
  const [paintTick, setPaintTick] = useState(0);
  const [missing, setMissing] = useState(() => new Set());
  const [busy, setBusy] = useState(null);
  const [exportOpen, setExportOpen] = useState(false);
  const [recordOpen, setRecordOpen] = useState(false);
  const [exporting, setExporting] = useState(null);
  // The last export's measure, on the editor for the checks: its time, speed and encoder.
  const [lastExport, setLastExport] = useState(null);
  // The encoders are asked once, early, so the Export box knows its format when it opens.
  useEffect(() => { exportFormat(); }, []);
  // The destination last exported for, offered first next time.
  const [lastPreset, setLastPreset] = useState('project');
  useEffect(() => { shell.store.get({ key: 'video.exportPreset', fallback: 'project' }).then((v) => setLastPreset(v || 'project')).catch(() => {}); }, [shell]);
  // How a project's own layout is made again for another shape (an introduction's scenes).
  const relayoutRef = useRef(null);
  const [transitionSeconds, setTransitionSeconds] = useState(1);
  // The timeline's share of the editor's height, the preview having the rest:
  // kept on this computer, and the same share whatever the window's size.
  const [timelineShare, setTimelineShare] = useState(TIMELINE_SHARE);
  const [editorH, setEditorH] = useState(0);
  useEffect(() => { shell.store.get({ key: 'video.timelineShare', fallback: null }).then((v) => { if (Number.isFinite(v) && v > 0 && v < 1) setTimelineShare(v); }).catch(() => {}); }, [shell]);
  // The editor measured whenever it is drawn or resized, however late it first appears.
  const editorWatch = useRef(null);
  const editorRef = useCallback((el) => {
    editorWatch.current?.disconnect();
    editorWatch.current = null;
    if (!el) return;
    const measure = () => setEditorH(el.clientHeight);
    measure();
    if (typeof ResizeObserver === 'function') {
      editorWatch.current = new ResizeObserver(measure);
      editorWatch.current.observe(el);
    }
  }, []);
  const timelineH = splitHeight(editorH, timelineShare);
  const clipboard = useRef(null);
  const canvasRef = useRef(null);
  const copyKey = useRef(newCopyKey());
  const exportAbort = useRef(null);
  const appMenuOpen = useRef(null);
  const appMenuNewIntro = useRef(null);
  const appMenu = useAppMenu({
    shell,
    appKey: 'video',
    onNew: () => shell.win.create({ app: 'video' }),
    onOpen: () => appMenuOpen.current?.(),
    extra: [{ label: t('New Introduction Video'), icon: 'scenes', run: () => appMenuNewIntro.current?.() }],
  });
  // An introduction: the scene picked, the scene being recorded, the brand kit.
  const [sceneId, setSceneId] = useState(null);
  const [sceneRecord, setSceneRecord] = useState(null);
  const [brandOpen, setBrandOpen] = useState(false);
  const [brand, setBrand] = useState(() => normaliseBrand(null));
  useEffect(() => { shell.store.get({ key: BRAND_KEY, fallback: null }).then((b) => setBrand(normaliseBrand(b))).catch(() => {}); }, [shell]);

  // Every decoder ready and every thumbnail made repaints what shows it; only
  // a decoder or a picture ready rebuilds the plan, and only a sound decoded
  // while the preview plays restarts it, so that the sound is heard. A seek
  // landing ('frame') is a repaint alone: rebuilding the plan for it
  // restarted a playing preview, whose restart sought again, without end.
  useEffect(() => media.onChange((kind) => {
    if (kind === 'media') setMediaVersion((v) => v + 1);
    setPaintTick((v) => v + 1);
    if (kind === 'sound' && playback.current) restartRef.current?.(timeRef.current);
  }), [media]);
  useEffect(() => () => media.dispose(), [media]);

  /* ── the document, its history, its plan ───────────────────────────── */

  const duration = useMemo(() => projectDuration(doc), [doc]);

  const planCanvas = useMemo(() => document.createElement('canvas'), []);
  const plan = useMemo(() => {
    const videos = Object.fromEntries(media.videos);
    const assets = Object.fromEntries(media.images);
    return buildPlan({ canvas: planCanvas, images: [], options: doc.options, layerPatches: planPatches(doc), videos, assets });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [doc, mediaVersion]);
  const planRef = useRef(plan);
  planRef.current = plan;

  /** Record the document as it is now, then change it: one undo step (a `group` merges a run of the same change). */
  const edit = useCallback((label, fn, group = null, { lanes = false } = {}) => {
    const before = docRef.current;
    const next = fn(before);
    if (!next || next === before) return false;
    // A locked track's clips stay as they are, whatever the command.
    if (brokenLocks(before, next, { lanesMoved: lanes }).length) {
      toast(t('That would change a clip on a locked track. Unlock the track first.'), { ms: 5000 });
      return false;
    }
    history.current.record({ state: before, label, group, meta: selectedRef.current });
    docRef.current = next;
    setDoc(next);
    setDirty(true);
    setHistoryTick((n) => n + 1);
    return true;
  }, [toast]);
  const selectedRef = useRef(selected);
  selectedRef.current = selected;

  // A drag records once as it starts and writes as it moves.
  const beginLive = useCallback((label) => {
    history.current.record({ state: docRef.current, label, meta: selectedRef.current });
    setHistoryTick((n) => n + 1);
  }, []);
  const live = useCallback((next) => {
    docRef.current = next;
    setDoc(next);
    setDirty(true);
  }, []);
  /** Whether a clip is on a locked track: the stage does not pick it. */
  const laneLocked = useCallback((id) => {
    const d = docRef.current;
    const p = d.patches.find((x) => x.id === id);
    return Boolean(p && (d.lanes?.[sectionOf(p)] || {})[laneOf(p)]?.locked);
  }, []);

  const undo = useCallback(() => {
    const step = history.current.undo(docRef.current, selectedRef.current);
    if (!step) return;
    setDoc(step.state);
    setSelected((step.meta || []).filter((id) => clipById(step.state, id)));
    setDirty(true);
    setHistoryTick((n) => n + 1);
  }, []);
  const redo = useCallback(() => {
    const step = history.current.redo(docRef.current, selectedRef.current);
    if (!step) return;
    setDoc(step.state);
    setSelected((step.meta || []).filter((id) => clipById(step.state, id)));
    setDirty(true);
    setHistoryTick((n) => n + 1);
  }, []);

  /* ── the playhead and playback ─────────────────────────────────────── */

  const playback = useRef(null);
  const playStarts = useRef(0);

  const seek = useCallback((t0) => {
    const next = Math.max(0, Math.min(Math.max(duration, 0), Number(t0) || 0));
    timeRef.current = next;
    setTime(next);
    if (playback.current) restartRef.current?.(next);
  }, [duration]);

  const audioClips = useCallback((p) => {
    const clips = [];
    for (const layer of soundLayers(p)) {
      const buffer = media.sounds.get(layer.url);
      if (buffer) clips.push(clipFromSoundLayer(layer, buffer, p, { volume: 1, fadeIn: 0, fadeOut: 0 }));
    }
    return clips.length ? { clips } : null;
  }, [media]);

  const stop = useCallback(() => {
    const pb = playback.current;
    if (!pb) return;
    playback.current = null;
    cancelAnimationFrame(pb.raf);
    clearInterval(pb.timer);
    pb.audio?.stop();
    pauseVideoElements(planRef.current);
    setPlaying(false);
    setPaintTick((n) => n + 1);
  }, []);

  /** Play forward at `rate` (1, 2, 4), or backward (a negative rate, by stepping, without sound). */
  const play = useCallback((rate = 1) => {
    stop();
    const p = planRef.current;
    let from = timeRef.current;
    if (rate > 0 && from >= projectDuration(docRef.current) - 0.02) from = 0;
    const pb = { rate, from, t0: performance.now(), raf: 0, timer: 0, audio: null };
    playback.current = pb;
    setPlaying(true);
    if (rate === 1) pb.audio = startAudioPreview(audioClips(p), p, from);
    // Counted for the window checks: a preview left alone starts once.
    playStarts.current += 1;
    const end = projectDuration(docRef.current);
    const tick = () => {
      if (playback.current !== pb) return;
      const now = pb.from + ((performance.now() - pb.t0) / 1000) * pb.rate;
      if (now >= end || now <= 0) {
        const at = Math.max(0, Math.min(end, now));
        stop();
        timeRef.current = at;
        setTime(at);
        return;
      }
      if (pb.rate === 1) driveVideoElements(planRef.current, now);
      timeRef.current = now;
      setTime(now);
    };
    // A timer rather than animation frames: the clock is the wall's, so a
    // window that is not being drawn (behind another, or off the desktop)
    // still plays on to the end instead of stopping where it was. Faster
    // than real time and backwards, the decoder is sought frame by frame.
    pb.timer = setInterval(tick, rate === 1 ? 1000 / 60 : 1000 / FPS);
  }, [stop, audioClips]);

  const restartRef = useRef(null);
  restartRef.current = (at) => {
    const pb = playback.current;
    if (!pb) return;
    const rate = pb.rate;
    timeRef.current = at;
    play(rate);
  };

  const toggle = useCallback(() => (playback.current ? stop() : play(1)), [play, stop]);

  // J, K and L: back, stop, forward, each press of J or L faster.
  const shuttle = useCallback((dir) => {
    const pb = playback.current;
    if (dir === 0) return stop();
    const rate = pb && Math.sign(pb.rate) === dir ? Math.min(8, Math.abs(pb.rate) * 2) * dir : dir;
    play(rate);
    return undefined;
  }, [play, stop]);

  const stepFrames = useCallback((n) => {
    stop();
    seek(Math.round((timeRef.current + n / FPS) * FPS) / FPS);
  }, [stop, seek]);

  useEffect(() => () => stop(), [stop]);

  // A project changed under a playing preview plays on from where it is. The
  // project, not the plan: the plan is also rebuilt when a decoder comes
  // ready, which the playing preview picks up from planRef without starting
  // its sound again.
  useEffect(() => {
    if (playback.current) restartRef.current?.(timeRef.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [doc]);

  /* ── opening, adding, saving ───────────────────────────────────────── */

  /** The timeline zoomed so a project of `length` seconds fills it, with a little room after. */
  const fitTo = useCallback((length) => {
    const w = (document.querySelector('.vd-tl-scroll')?.clientWidth || 900) - 140;
    setPps(Math.max(4, Math.min(600, w / Math.max(1, length * 1.12))));
  }, []);

  /**
   * Put files on the timeline, one after another: `where` is 'end' (after
   * the clips on the picked clip's picture track, V1 when none is picked),
   * 'layer' (a new track from the playhead) or 'insert' (at the playhead on
   * the picked track, what follows moved on), as placeMedia does it.
   */
  const addFiles = useCallback(async (files, { initial = false, where = 'end' } = {}) => {
    const list = files.filter((f) => mediaKind(f) && mediaKind(f) !== 'project');
    if (!list.length) return;
    setBusy(tn(list.length, 'Adding {count} file', 'Adding {count} files'));
    const added = [];
    const failed = [];
    for (const file of list) {
      try {
        const stat = await shell.fs.stat({ path: file }).catch(() => null);
        const info = await media.describe(file, { name: basename(file), size: stat?.size ?? null });
        added.push(info);
      } catch {
        failed.push(basename(file));
      }
    }
    setBusy(null);
    if (added.length) {
      const wasEmpty = !docRef.current.patches.some((p) => p.timing);
      const place = (d) => {
        const picked = selectedRef.current.map((id) => clipById(d, id)).find((p) => p && sectionOf(p) === 'v');
        const { doc: out, ids } = placeMedia(d, added, { where, lane: picked ? laneOf(picked) : 0, at: timeRef.current });
        setSelected(ids.slice(0, 1));
        return out;
      };
      if (initial) {
        // A file opened in Video from outside is a project of one clip that
        // nobody has changed yet: nothing to save until it is edited.
        const next = place(docRef.current);
        docRef.current = next;
        setDoc(next);
      } else {
        edit(tn(added.length, 'Add {count} clip', 'Add {count} clips'), place);
      }
      if (wasEmpty) requestAnimationFrame(() => fitTo(projectDuration(docRef.current)));
    }
    if (failed.length) {
      toast(t('{names} could not be opened: it is not a video, sound or picture this system can play.', { names: failed.join(', ') }), { tone: 'bad', ms: 7000 });
    }
  }, [shell, media, edit, toast]);

  /** A caption file (.srt or .vtt) onto the caption track: its cues as captions, replacing any there. */
  const importCaptions = useCallback(async (given = null) => {
    let file = given;
    if (!file) {
      const picked = await shell.dialog.open({
        title: t('Import Captions'),
        filters: [{ name: t('Captions (.srt, .vtt)'), extensions: ['srt', 'vtt'] }, { name: t('All files'), extensions: ['*'] }],
        multiple: false,
      });
      file = picked?.[0] || null;
    }
    if (!file) return;
    try {
      const { bytes } = await shell.fs.read({ path: file });
      const { cues } = parseCaptions(new TextDecoder('utf-8').decode(bytes), file);
      if (!cues.length) {
        toast(t('{name} holds no captions this editor can read.', { name: basename(file) }), { tone: 'bad', ms: 6000 });
        return;
      }
      let ids = [];
      edit(t('Import Captions'), (d) => { const r = placeCaptions(d, cues, { aspect: d.options.aspect }); ids = r.ids; return r.doc; });
      setSelected(ids.slice(0, 1));
      toast(tn(cues.length, '{count} caption from {name} is on its own track.', '{count} captions from {name} are on their own track.', { name: basename(file) }), { tone: 'good', ms: 5000 });
    } catch (err) {
      toast(t('{name} could not be opened: {reason}', { name: basename(file), reason: err.message }), { tone: 'bad', ms: 8000 });
    }
  }, [shell, edit, toast]);

  /** The caption track written as a SubRip or a WebVTT file, by the name chosen. */
  const exportCaptions = useCallback(async () => {
    const cues = captionsOf(docRef.current);
    if (!cues.length) return;
    const base = (path ? basename(path) : t('Untitled video')).replace(/\.[^.]+$/, '');
    const target = await shell.dialog.save({
      title: t('Export Captions'),
      defaultPath: `${base}.srt`,
      filters: [{ name: t('SubRip captions (.srt)'), extensions: ['srt'] }, { name: t('WebVTT captions (.vtt)'), extensions: ['vtt'] }],
    });
    if (!target) return;
    const vtt = /\.vtt$/i.test(target);
    const file = /\.(srt|vtt)$/i.test(target) ? target : `${target}.srt`;
    try {
      await shell.fs.write({ path: file, bytes: new TextEncoder().encode(vtt ? writeVtt(cues) : writeSrt(cues)) });
      toast(tn(cues.length, '{count} caption written to {name}', '{count} captions written to {name}', { name: basename(file) }), { tone: 'good', ms: 5000 });
    } catch (err) {
      toast(t('The captions could not be saved: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
    }
  }, [shell, path, toast]);

  /** Read a .rvid and make it this window's project. */
  const openProject = useCallback(async (file, { bytes = null, recovered = null } = {}) => {
    stop();
    setBusy(t('Opening the project'));
    try {
      const raw = bytes || (await shell.fs.read({ path: file })).bytes;
      const read = await readRvid(raw, {
        projectPath: recovered ? recovered.path : file,
        exists: async (p) => Boolean(await shell.fs.exists({ path: p }).catch(() => false)),
      });
      const next = normaliseProject(read.doc);
      if (isIntro(next)) { setSceneId(next.intro.scenes[0]?.id || null); setTab('intro'); }
      history.current.clear();
      setDoc(next);
      docRef.current = next;
      setSelected([]);
      seek(0);
      setPath(recovered ? recovered.path : file);
      setDirty(Boolean(recovered) || read.moved.length > 0);
      const gone = new Set(read.missing);
      const failures = await media.loadProject(projectMedia(next));
      for (const f of failures) gone.add(f.url);
      setMissing(gone);
      let shown = next;
      if (next.recording?.unfinished) {
        const lengths = {};
        for (const p of next.patches) if (p.take && p.type === 'video' && media.videos.get(p.url)?.duration) lengths[p.url] = media.videos.get(p.url).duration;
        for (const p of next.patches) if (p.take && p.type === 'sound' && !lengths[p.url] && media.sounds.get(p.url)?.duration) lengths[p.url] = media.sounds.get(p.url).duration;
        shown = settleTakes(next, lengths);
        docRef.current = shown;
        setDoc(shown);
      }
      requestAnimationFrame(() => fitTo(projectDuration(shown)));
      if (!recovered) shell.app.addRecent({ path: file, app: 'video' }).catch(() => {});
      if (read.moved.length) toast(tn(read.moved.length, '{count} file had moved with the project and was found again.', '{count} files had moved with the project and were found again.'), { ms: 6000 });
      if (gone.size) toast(tn(gone.size, '{count} file of this project cannot be found. Its clips show nothing until it is back.', '{count} files of this project cannot be found. Their clips show nothing until they are back.'), { tone: 'bad', ms: 8000 });
    } catch (err) {
      toast(t('{name} could not be opened: {reason}', { name: basename(file || ''), reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [shell, media, toast, stop, seek]);

  const openFiles = useCallback(async () => {
    const files = await shell.dialog.open({
      title: t('Open'),
      filters: [
        { name: t('Video projects and media'), extensions: ['rvid', ...MEDIA_EXTENSIONS, 'srt', 'vtt'] },
        { name: t('All files'), extensions: ['*'] },
      ],
      multiple: true,
    });
    if (!files?.length) return;
    const project = files.find((f) => mediaKind(f) === 'project');
    if (project) {
      if (doc.patches.length || dirty) shell.win.create({ app: 'video', file: project });
      else openProject(project);
      return;
    }
    const captions = files.find((f) => /\.(srt|vtt)$/i.test(f));
    if (captions) importCaptions(captions);
    addFiles(files);
  }, [shell, doc, dirty, openProject, addFiles, importCaptions]);
  appMenuOpen.current = openFiles;

  const insertMedia = useCallback(async (where = 'end') => {
    const files = await shell.dialog.open({
      title: t('Add Media'),
      filters: [
        { name: t('Video, sound and pictures'), extensions: MEDIA_EXTENSIONS },
        { name: t('All files'), extensions: ['*'] },
      ],
      multiple: true,
    });
    if (files?.length) addFiles(files, { where: typeof where === 'string' ? where : 'end' });
  }, [shell, addFiles]);

  const saveTo = useCallback(async (target) => {
    try {
      const bytes = writeRvid(docRef.current, { projectPath: target });
      await shell.fs.write({ path: target, bytes });
      setPath(target);
      setDirty(false);
      shell.doc.dropCopy({ key: copyKey.current }).catch(() => {});
      shell.app.addRecent({ path: target, app: 'video' }).catch(() => {});
      toast(t('Saved {name}', { name: basename(target) }), { tone: 'good' });
      return true;
    } catch (err) {
      toast(t('The project could not be saved: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
      return false;
    }
  }, [shell, toast]);

  const saveAs = useCallback(async () => {
    const target = await shell.dialog.save({
      title: t('Save the project as'),
      defaultPath: path || t('Untitled video') + '.rvid',
      filters: [{ name: t('Rutba Video Project (.rvid)'), extensions: ['rvid'] }],
    });
    if (!target) return false;
    return saveTo(/\.rvid$/i.test(target) ? target : `${target}.rvid`);
  }, [shell, path, saveTo]);

  const save = useCallback(() => (path ? saveTo(path) : saveAs()), [path, saveTo, saveAs]);

  useDirtyGuard({ shell, dirty, name: path ? basename(path) : t('Untitled video'), onSave: save });

  // Opened on a file, or on recovered work.
  useEffect(() => {
    const recover = new URLSearchParams(location.search).get('recover');
    // The launcher's Record a screen: a new video that opens on the recorder.
    if (new URLSearchParams(location.search).get('record')) setRecordOpen(true);
    // The launcher's Introduction video, or New Introduction Video: a new introduction.
    if (new URLSearchParams(location.search).get('intro')) startIntroRef.current?.();
    if (recover) {
      (async () => {
        try {
          const copy = await shell.doc.readCopy({ file: recover });
          const m = /^window-([A-Za-z0-9_-]+)\./.exec(recover);
          if (m) copyKey.current = m[1];
          await openProject(copy.path || copy.name, { bytes: copy.bytes, recovered: copy });
          toast(t('Recovered unsaved work. Save it to keep it.'), { ms: 6000 });
        } catch (err) {
          toast(err.message, { tone: 'bad' });
        }
      })();
    } else if (boot.file) {
      if (mediaKind(boot.file) === 'project') openProject(boot.file);
      else addFiles([boot.file], { initial: true });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useFileDrop(useCallback((files) => {
    const project = files.find((f) => mediaKind(f) === 'project');
    if (project) shell.win.create({ app: 'video', file: project });
    const captions = files.find((f) => /\.(srt|vtt)$/i.test(f));
    if (captions) importCaptions(captions);
    addFiles(files);
  }, [addFiles, shell, importCaptions]));

  // Autosave: a copy in the recovery folder soon after the first unsaved
  // change, then every half minute while there is unsaved work, dropped on a
  // save and when the window closes properly. A crash leaves it for the launcher.
  const copied = useRef({ at: 0, doc: null });
  useEffect(() => {
    if (!dirty || !doc.patches.length) return undefined;
    const since = Date.now() - copied.current.at;
    const wait = copied.current.at ? Math.max(0, AUTOSAVE_EVERY_MS - since) : AUTOSAVE_FIRST_MS;
    const timer = setTimeout(() => {
      if (copied.current.doc === docRef.current) return;
      copied.current = { at: Date.now(), doc: docRef.current };
      shell.doc.keepCopy({
        key: copyKey.current,
        kind: 'video',
        name: path ? basename(path) : t('Untitled video'),
        path,
        bytes: writeRvid(docRef.current, { projectPath: path }),
      }).catch(() => {});
    }, wait);
    return () => clearTimeout(timer);
  }, [doc, dirty, path, shell]);
  useEffect(() => {
    const away = () => { shell.doc.dropCopy({ key: copyKey.current }).catch(() => {}); };
    window.addEventListener('pagehide', away);
    return () => window.removeEventListener('pagehide', away);
  }, [shell]);

  /* ── an introduction ───────────────────────────────────────────────── */

  const intro = isIntro(doc);
  /** The template's words, in this window's language. */
  const introWords = () => ({ title: t('Your name'), subtitle: t('What you do, in a line'), endTitle: t('Thanks for watching'), endSubtitle: t('Where to find you'), placeholder: t('Record this scene') });

  /** This window's project made a new introduction, in the brand kit kept for this person. */
  const startIntro = useCallback(async () => {
    stop();
    const kept = normaliseBrand(await shell.store.get({ key: BRAND_KEY, fallback: null }).catch(() => null));
    setBrand(kept);
    const next = newIntro({ brand: kept, words: introWords() });
    history.current.clear();
    docRef.current = next;
    setDoc(next);
    setDirty(false);
    setSelected([]);
    setSceneId(next.intro.scenes[0].id);
    setTab('intro');
    seek(0);
    if (kept.logo) media.image(kept.logo.url).catch(() => {});
    requestAnimationFrame(() => fitTo(projectDuration(next)));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, media, stop, seek, fitTo]);
  const startIntroRef = useRef(null);
  startIntroRef.current = startIntro;
  appMenuNewIntro.current = () => {
    if (!docRef.current.patches.some((p) => p.timing) && !dirty) startIntro();
    else shell.win.create({ app: 'video', query: { intro: 1 } });
  };

  // The brand's logo is a picture every scene draws: loaded when it changes.
  const logoUrl = doc.intro?.brand?.logo?.url || null;
  useEffect(() => { if (logoUrl) media.image(logoUrl).catch(() => {}); }, [logoUrl, media]);
  // An export to another shape lays the scenes out again for it.
  relayoutRef.current = intro ? (d, aspect) => composeIntro(d, { aspect }) : null;

  const currentScene = intro ? (sceneById(doc, sceneId) || sceneAt(doc, time) || doc.intro.scenes[0]) : null;
  const sceneIndex = currentScene ? doc.intro.scenes.findIndex((s) => s.id === currentScene.id) : -1;
  const pickScene = useCallback((id, start) => { setSceneId(id); stop(); seek(start); }, [stop, seek]);
  const sceneEdit = useCallback((label, fn, group = null) => edit(label, fn, group), [edit]);

  /** A scene's take, read like any added file and put in its scene; the scenes after it move on. */
  const landScene = useCallback(async ({ takes }) => {
    const id = sceneRecord;
    setSceneRecord(null);
    stop();
    if (!id) return;
    setBusy(t('Adding the take'));
    try {
      const files = [];
      for (const k of takes) {
        const info = await media.describe(k.path, { name: basename(k.path), size: k.written ?? null });
        files.push({
          role: k.role, url: k.path, name: basename(k.path).replace(/\.[^.]+$/, ''), duration: info.duration || k.duration,
          width: info.width || k.width, height: info.height || k.height, hasAudio: Boolean(k.hasAudio && info.hasAudio),
          ...(k.crop ? { crop: k.crop } : {}), ...(k.role === 'camera' ? { mirror: k.mirror !== false } : {}), ...(k.pointer ? { pointer: k.pointer } : {}),
        });
      }
      edit(t('Record Scene'), (d) => recordScene(d, id, files).doc);
      setSceneId(id);
      const span = sceneSpans(docRef.current).find((s) => s.id === id);
      if (span) seek(span.start);
      requestAnimationFrame(() => fitTo(projectDuration(docRef.current)));
      toast(t('The take is in its scene. Record it again for another take; every take is kept.'), { tone: 'good', ms: 5000 });
    } catch (err) {
      toast(t('The recording could not be added: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [sceneRecord, media, edit, stop, seek, fitTo, toast]);

  const saveBrand = useCallback((kit) => {
    const b = normaliseBrand(kit);
    setBrandOpen(false);
    setBrand(b);
    shell.store.set({ key: BRAND_KEY, value: b }).catch(() => {});
    if (b.logo) media.image(b.logo.url).catch(() => {});
    if (isIntro(docRef.current)) edit(t('Brand Kit'), (d) => applyBrand(d, b));
    toast(t('The brand kit is kept for you and drawn into every scene.'), { tone: 'good', ms: 4000 });
  }, [shell, media, edit, toast]);

  /* ── record ────────────────────────────────────────────────────────── */

  /** A take as the timeline places it: its file as the source. */
  const asTake = (k) => ({ ...k, url: k.path, name: basename(k.path).replace(/\.[^.]+$/, ''), size: k.written ?? null });

  // While a recording runs, the project with its takes so far is the
  // window's recovery copy: a crash leaves it for the launcher to offer.
  const recordingCopy = useCallback((takes) => {
    try {
      const { doc: withTakes } = placeTakes(docRef.current, takes.map(asTake), { unfinished: true });
      shell.doc.keepCopy({
        key: copyKey.current,
        kind: 'video',
        name: path ? basename(path) : t('Untitled video'),
        path,
        bytes: writeRvid(withTakes, { projectPath: path }),
      }).catch(() => {});
    } catch {
      /* the copy is a safety net, not the recording */
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, path]);

  /** The takes, read like any added file, placed together after what is there. */
  const landTakes = useCallback(async ({ takes }) => {
    setRecordOpen(false);
    stop();
    setBusy(t('Adding the recording'));
    try {
      const placed = [];
      for (const k of takes) {
        const info = await media.describe(k.path, { name: basename(k.path), size: k.written ?? null });
        placed.push(asTake({ ...k, duration: info.duration || k.duration, width: info.width || k.width, height: info.height || k.height, hasAudio: Boolean(k.hasAudio && info.hasAudio) }));
      }
      const wasEmpty = !docRef.current.patches.some((p) => p.timing);
      let ids = [];
      let at = 0;
      edit(t('Record'), (d) => {
        const r = placeTakes(d, placed);
        ids = r.ids;
        at = r.doc.patches.find((p) => p.id === r.ids[0])?.timing.start || 0;
        return r.doc;
      });
      setSelected(ids.slice(0, 1));
      seek(at);
      if (wasEmpty) requestAnimationFrame(() => fitTo(projectDuration(docRef.current)));
      toast(tn(placed.length, 'The recording is on the timeline: {count} take.', 'The recording is on the timeline: {count} takes.'), { tone: 'good', ms: 5000 });
    } catch (err) {
      toast(t('The recording could not be added: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
      copied.current = { at: 0, doc: null };
    }
  }, [media, edit, stop, seek, fitTo, toast]);

  /** Record closed without a recording: the copy goes back to the project as it is. */
  const recordClosed = useCallback(() => {
    setRecordOpen(false);
    copied.current = { at: 0, doc: null };
    if (dirty && docRef.current.patches.length) {
      shell.doc.keepCopy({ key: copyKey.current, kind: 'video', name: path ? basename(path) : t('Untitled video'), path, bytes: writeRvid(docRef.current, { projectPath: path }) }).catch(() => {});
    } else {
      shell.doc.dropCopy({ key: copyKey.current }).catch(() => {});
    }
  }, [shell, dirty, path]);

  /* ── export ────────────────────────────────────────────────────────── */

  const runExport = useCallback(async ({ quality, range, format, preset = 'project' }) => {
    setExportOpen(false);
    setLastPreset(preset);
    shell.store.set({ key: 'video.exportPreset', value: preset }).catch(() => {});
    const base = (path ? basename(path) : t('Untitled video')).replace(/\.[^.]+$/, '');
    const tag = PRESET_TAGS()[preset];
    const target = await shell.dialog.save({
      title: t('Export video'),
      defaultPath: `${base}${tag ? ` (${tag})` : ''}.${format.ext}`,
      filters: [{ name: format.ext === 'mp4' ? t('MP4 Video (.mp4)') : t('WebM Video (.webm)'), extensions: [format.ext] }],
    });
    if (!target) return;
    stop();
    const total = range ? range.out - range.in : projectDuration(docRef.current);
    const abort = new AbortController();
    exportAbort.current = abort;
    setExporting({ progress: 0, total, fast: Boolean(format.fast), since: null, from: 0 });
    // The box is drawn again a few times a second, not once a frame: the
    // window's own redraws would otherwise take the time the encoder needs.
    let shown = 0;
    try {
      const out = await exportProject(docRef.current, media, shell, target, {
        format,
        quality,
        range,
        preset,
        relayout: relayoutRef.current,
        signal: abort.signal,
        onProgress: (p) => {
          const now = performance.now();
          if (p < 1 && now - shown < 150) return;
          shown = now;
          // The speed is measured from the first frame encoded, not from the press:
          // reading the sources and mixing the sound come before it.
          setExporting((e) => (e ? { ...e, progress: p, since: e.since ?? now, from: e.since == null ? p : e.from } : e));
        },
      });
      setLastExport({ ...out, total });
      toast(out.fast
        ? t('Exported {name} ({size}) in {time}', { name: basename(target), size: formatBytes(out.size), time: timecode(out.seconds) })
        : t('Exported {name} ({size})', { name: basename(target), size: formatBytes(out.size) }), { tone: 'good', ms: 6000 });
    } catch (err) {
      if (err?.name !== 'AbortError') toast(t('The export stopped: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      exportAbort.current = null;
      setExporting(null);
    }
  }, [shell, media, path, toast, stop]);

  /* ── commands ──────────────────────────────────────────────────────── */

  const ripple = doc.settings?.ripple !== false;
  const sel = selected.map((id) => clipById(doc, id)).filter(Boolean);
  const first = sel[0] || null;
  const pictures = sel.filter((p) => p.type === 'video' || p.type === 'image');
  const hasClips = doc.patches.some((p) => p.timing);
  const under = clipsAt(doc, time);
  // What Trim Start and Trim End act on: a picked clip the playhead is in, else the main lane's clip there.
  const target = sel.find((p) => time > p.timing.start && time < p.timing.end) || under.find((p) => sectionOf(p) === 'v' && laneOf(p) === 0) || under[0] || null;

  const setPictures = (label, fn) => edit(label, (d) => ({ ...d, patches: d.patches.map((p) => (selected.includes(p.id) && (p.type === 'video' || p.type === 'image') ? { ...p, ...fn(p) } : p)) }));
  const soundsOfSelection = () => {
    const ids = new Set();
    for (const p of sel) {
      if (p.type === 'sound') ids.add(p.id);
      for (const q of partnersOf(doc, p.id)) if (q.type === 'sound') ids.add(q.id);
    }
    return ids;
  };
  const setSounds = (label, fn) => {
    const ids = soundsOfSelection();
    if (!ids.size) return;
    edit(label, (d) => ({ ...d, patches: d.patches.map((p) => (ids.has(p.id) ? { ...p, ...fn(p) } : p)) }));
  };
  const fitView = () => fitTo(duration);
  const titleNames = TITLE_NAMES();

  const cmds = {
    'file.open': { label: t('Open'), icon: 'open', key: 'Mod+O', run: openFiles },
    'file.save': { label: t('Save'), icon: 'save', key: 'Mod+S', run: save },
    'file.saveAs': { label: t('Save As'), icon: 'save', key: 'Mod+Shift+S', run: saveAs },
    'file.export': { label: t('Export'), icon: 'export', key: 'Mod+E', disabled: !hasClips || Boolean(exporting), tip: t('write the video file, with its sound'), run: () => setExportOpen(true) },
    'insert.media': { label: t('Add Media'), icon: 'plus', key: 'Mod+I', tip: t('videos, sound and pictures, one after another after the clips on the picked track'), run: () => insertMedia('end') },
    'insert.mediaLayer': { label: t('Add on a New Track'), icon: 'layers', tip: t('videos and pictures on a track of their own above the others, from the playhead'), run: () => insertMedia('layer') },
    'insert.mediaHere': { label: t('Insert at Playhead'), icon: 'move', tip: t('on the picked track at the playhead; what follows moves later to make room'), run: () => insertMedia('insert') },
    'captions.import': { label: t('Import Captions'), icon: 'captions', tip: t('a SubRip (.srt) or WebVTT (.vtt) file, each caption a clip on its own track'), run: () => importCaptions() },
    'captions.export': { label: t('Export Captions'), icon: 'export', disabled: !doc.patches.some(isCaption), tip: t('the caption track as a .srt or .vtt file, to send beside the video'), run: exportCaptions },
    'intro.record': { label: currentScene?.take ? t('Retake Scene') : t('Record Scene'), icon: 'record', disabled: !currentScene || !RECORDED.has(currentScene.layout) || Boolean(exporting), tip: t('the scene picked, with its script above the camera; every take is kept'), run: () => { stop(); setSceneRecord(currentScene.id); } },
    'intro.up': { label: t('Move Earlier'), icon: 'chevronUp', disabled: sceneIndex <= 0, run: () => edit(t('Move Scene'), (d) => moveScene(d, currentScene.id, -1)) },
    'intro.down': { label: t('Move Later'), icon: 'chevronDown', disabled: !intro || sceneIndex >= doc.intro.scenes.length - 1, run: () => edit(t('Move Scene'), (d) => moveScene(d, currentScene.id, 1)) },
    'intro.remove': { label: t('Remove Scene'), icon: 'trash', disabled: !intro || doc.intro.scenes.length < 2, tip: t('its takes stay on the disk'), run: () => { edit(t('Remove Scene'), (d) => removeScene(d, currentScene.id)); setSceneId(null); } },
    'intro.brand': { label: t('Brand Kit'), icon: 'brand', tip: t('a logo, colours, a font and a name and role, in every scene'), run: () => setBrandOpen(true) },
    'intro.new': { label: t('New Introduction Video'), icon: 'scenes', tip: t('scenes to record one at a time, in your brand'), run: () => appMenuNewIntro.current?.() },
    'record.screen': { label: t('Record Screen'), icon: 'record', disabled: Boolean(exporting) || recordOpen, tip: t('a screen, a window or a region, with the camera and sound, onto the timeline'), run: () => { stop(); setRecordOpen(true); } },
    'edit.undo': { label: t('Undo'), icon: 'undo', key: 'Mod+Z', disabled: !history.current.canUndo, run: undo },
    'edit.redo': { label: t('Redo'), icon: 'redo', key: 'Mod+Y', disabled: !history.current.canRedo, run: redo },
    'edit.redo2': { label: t('Redo'), key: 'Mod+Shift+Z', run: redo, hidden: true },
    'edit.cut': { label: t('Cut'), icon: 'cut', key: 'Mod+X', disabled: !sel.length, run: () => { clipboard.current = copyClips(doc, selected); edit(t('Cut'), (d) => removeClips(d, selected, { ripple })); setSelected([]); } },
    'edit.copy': { label: t('Copy'), icon: 'copy', key: 'Mod+C', disabled: !sel.length, run: () => { clipboard.current = copyClips(doc, selected); toast(tn(clipboard.current.clips.length, '{count} clip copied', '{count} clips copied')); } },
    'edit.paste': { label: t('Paste'), icon: 'paste', key: 'Mod+V', disabled: !clipboard.current, tip: t('at the playhead'), run: () => { let ids = []; edit(t('Paste'), (d) => { const r = pasteClips(d, clipboard.current, timeRef.current); ids = r.ids; return r.doc; }); setSelected(ids.slice(0, 1)); } },
    'edit.duplicate': { label: t('Duplicate'), icon: 'copy', key: 'Mod+D', disabled: !sel.length, run: () => { let ids = []; edit(t('Duplicate'), (d) => { const r = duplicateClips(d, selected); ids = r.ids; return r.doc; }); setSelected(ids.slice(0, 1)); } },
    'edit.delete': { label: t('Delete'), icon: 'trash', key: 'Delete', disabled: !sel.length, tip: ripple ? t('the clips after it close up') : t('a gap stays where it was'), run: () => { edit(t('Delete'), (d) => removeClips(d, selected, { ripple })); setSelected([]); } },
    'edit.backspace': { label: t('Delete'), key: 'Backspace', disabled: !sel.length, hidden: true, run: () => { edit(t('Delete'), (d) => removeClips(d, selected, { ripple })); setSelected([]); } },
    'edit.rippleDelete': { label: t('Ripple Delete'), icon: 'trash', key: 'Shift+Delete', disabled: !sel.length, tip: t('delete and close up what follows'), run: () => { edit(t('Ripple Delete'), (d) => removeClips(d, selected, { ripple: true })); setSelected([]); } },
    'edit.split': { label: t('Split'), icon: 'scissors', key: 's', disabled: !hasClips, tip: t('cut the clip at the playhead in two'), run: () => edit(t('Split'), (d) => splitAt(d, timeRef.current, selectedRef.current).doc) },
    'edit.splitAll': { label: t('Split All Tracks'), icon: 'scissors', key: 'Shift+S', disabled: !hasClips, tip: t('cut every clip under the playhead'), run: () => edit(t('Split All Tracks'), (d) => splitAll(d, timeRef.current).doc) },
    'edit.trimStart': { label: t('Trim Start'), icon: 'crop', key: 'i', disabled: !target, tip: t('the clip starts at the playhead'), run: () => target && edit(t('Trim Start'), (d) => trimStart(d, target.id, timeRef.current, { ripple })) },
    'edit.trimEnd': { label: t('Trim End'), icon: 'crop', key: 'o', disabled: !target, tip: t('the clip ends at the playhead'), run: () => target && edit(t('Trim End'), (d) => trimEnd(d, target.id, timeRef.current, { ripple })) },
    'edit.ripple': { label: t('Ripple'), icon: 'link', pressed: ripple, tip: t('deleting or trimming a clip closes up the clips after it'), run: () => edit(t('Ripple'), (d) => ({ ...d, settings: { ...d.settings, ripple: !ripple } })) },
    'lane.split': { label: t('Split Track'), icon: 'trSplit', tip: t('the clips on the picked track from the playhead on move up onto a new track of their own'), disabled: !first, run: () => edit(t('Split Track'), (d) => splitLane(d, sectionOf(first), laneOf(first), timeRef.current), null, { lanes: true }) },
    'lane.insert': { label: t('Insert Track'), icon: 'plus', tip: t('an empty track above the picked one'), run: () => edit(t('Insert Track'), (d) => insertLane(d, first ? sectionOf(first) : 'v', first ? laneOf(first) + 1 : laneCount(d).v), null, { lanes: true }) },
    'lane.up': { label: t('Move Track Up'), icon: 'layerUp', tip: t('the picked track moves up one, drawn over the one it passes'), disabled: !first, run: () => edit(t('Move Track Up'), (d) => moveLane(d, sectionOf(first), laneOf(first), sectionOf(first) === 'v' ? laneOf(first) + 1 : laneOf(first) - 1), null, { lanes: true }) },
    'lane.down': { label: t('Move Track Down'), icon: 'layerDown', tip: t('the picked track moves down one'), disabled: !first, run: () => edit(t('Move Track Down'), (d) => moveLane(d, sectionOf(first), laneOf(first), sectionOf(first) === 'v' ? laneOf(first) - 1 : laneOf(first) + 1), null, { lanes: true }) },
    'edit.closeGap': { label: t('Close Gap'), icon: 'chevronLeft', tip: t('the empty stretch at the playhead closes'), run: () => edit(t('Close Gap'), (d) => closeGap(d, first ? sectionOf(first) : 'v', first ? laneOf(first) : 0, timeRef.current)) },
    'edit.detach': { label: t('Detach Audio'), icon: 'volume', disabled: !sel.some((p) => p.group), tip: t('the picture and its sound move on their own'), run: () => edit(t('Detach Audio'), (d) => unlinkClips(d, selected)) },
    'edit.link': { label: t('Link'), icon: 'link', disabled: sel.length < 2, tip: t('the picked clips move together'), run: () => edit(t('Link'), (d) => linkClips(d, selected)) },
    'range.in': { label: t('Mark In'), icon: 'chevronRight', key: 'Shift+I', run: () => edit(t('Mark In'), (d) => setRange(d, { in: timeRef.current })) },
    'range.out': { label: t('Mark Out'), icon: 'chevronLeft', key: 'Shift+O', run: () => edit(t('Mark Out'), (d) => setRange(d, { out: timeRef.current })) },
    'range.clear': { label: t('Clear In and Out'), icon: 'close', key: 'Mod+Shift+X', disabled: !doc.range, run: () => edit(t('Clear In and Out'), clearRange) },
    'marker.add': { label: t('Add Marker'), icon: 'flag', key: 'm', run: () => edit(t('Add Marker'), (d) => addMarker(d, timeRef.current, t('Marker {n}', { n: (d.markers || []).length + 1 }))) },
    'marker.prev': { label: t('Previous Marker'), icon: 'chevronLeft', key: 'Shift+M', run: () => { const m = previousMarker(doc, timeRef.current); if (m) seek(m.t); } },
    'marker.next': { label: t('Next Marker'), icon: 'chevronRight', key: 'Mod+Shift+M', run: () => { const m = nextMarker(doc, timeRef.current); if (m) seek(m.t); } },
    'play.toggle': { label: t('Play'), icon: 'play', key: 'k', tip: t('or the space bar; J and L shuttle back and forward'), run: toggle },
    'play.j': { label: t('Play backward'), key: 'j', hidden: true, run: () => shuttle(-1) },
    'play.l': { label: t('Play forward'), key: 'l', hidden: true, run: () => shuttle(1) },
    'play.start': { label: t('Go to Start'), icon: 'skipBack', key: 'Home', run: () => { stop(); seek(0); } },
    'play.end': { label: t('Go to End'), icon: 'skipForward', key: 'End', hidden: true, run: () => { stop(); seek(duration); } },
    'play.prevFrame': { label: t('Previous Frame'), icon: 'chevronLeft', key: ',', run: () => stepFrames(-1) },
    'play.nextFrame': { label: t('Next Frame'), icon: 'chevronRight', key: '.', run: () => stepFrames(1) },
    'play.back': { label: t('Back a Second'), key: 'arrowleft', hidden: true, run: () => { stop(); seek(timeRef.current - 1); } },
    'play.fwd': { label: t('Forward a Second'), key: 'arrowright', hidden: true, run: () => { stop(); seek(timeRef.current + 1); } },
    'clip.fit.cover': { label: t('Fill'), icon: 'grid', disabled: !pictures.length, pressed: pictures.length > 0 && pictures.every((p) => (p.fit || 'cover') === 'cover'), tip: t('the picture fills the frame, cropped to fit'), run: () => setPictures(t('Fill'), () => ({ fit: 'cover' })) },
    'clip.fit.contain': { label: t('Fit'), icon: 'image', disabled: !pictures.length, pressed: pictures.length > 0 && pictures.every((p) => p.fit === 'contain'), tip: t('the whole picture, on black bars'), run: () => setPictures(t('Fit'), () => ({ fit: 'contain' })) },
    'clip.fit.blur': { label: t('Blur Fill'), icon: 'contrast', disabled: !pictures.length, pressed: pictures.length > 0 && pictures.every((p) => p.fit === 'blur'), tip: t('the whole picture, over a blur of itself'), run: () => setPictures(t('Blur Fill'), () => ({ fit: 'blur' })) },
    'clip.autoZoom': { label: t('Zoom on Clicks'), icon: 'zoomIn', disabled: !(first?.type === 'video' && pointerOf(doc, first.id)?.clicks?.length), tip: t('the screen zooms in where the recorded clicks are and follows them; each zoom can be changed in the properties'), run: () => edit(t('Zoom on Clicks'), (d) => applyAutoZoom(d, first.id)) },
    'clip.rotateLeft': { label: t('Rotate Left'), icon: 'rotate', disabled: !pictures.length, run: () => setPictures(t('Rotate'), (p) => ({ turn: ((((p.turn || 0) - 90) % 360) + 360) % 360 })) },
    'clip.rotateRight': { label: t('Rotate Right'), icon: 'rotate', disabled: !pictures.length, run: () => setPictures(t('Rotate'), (p) => ({ turn: ((p.turn || 0) + 90) % 360 })) },
    'clip.flipX': { label: t('Flip Horizontal'), icon: 'flip', disabled: !pictures.length, pressed: pictures.length > 0 && pictures.every((p) => p.flipX), run: () => setPictures(t('Flip'), (p) => ({ flipX: !p.flipX })) },
    'clip.flipY': { label: t('Flip Vertical'), icon: 'flip', disabled: !pictures.length, pressed: pictures.length > 0 && pictures.every((p) => p.flipY), run: () => setPictures(t('Flip'), (p) => ({ flipY: !p.flipY })) },
    'clip.mute': { label: t('Mute'), icon: 'volume', disabled: !soundsOfSelection().size, pressed: soundsOfSelection().size > 0 && [...soundsOfSelection()].every((id) => clipById(doc, id)?.muted), tip: t('the picked clips make no sound'), run: () => { const all = [...soundsOfSelection()].every((id) => clipById(doc, id)?.muted); setSounds(t('Mute'), () => ({ muted: !all })); } },
    'clip.fadeIn': { label: t('Fade In'), icon: 'trFade', disabled: !soundsOfSelection().size, tip: t('the sound rises over a second'), run: () => setSounds(t('Fade In'), (p) => ({ enter: { kind: 'fade', seconds: p.enter?.seconds ? 0 : 1 } })) },
    'clip.fadeOut': { label: t('Fade Out'), icon: 'trFade', disabled: !soundsOfSelection().size, tip: t('the sound falls away over a second'), run: () => setSounds(t('Fade Out'), (p) => ({ exit: { kind: 'fade', seconds: p.exit?.seconds ? 0 : 1 } })) },
    'view.zoomIn': { label: t('Zoom In'), icon: 'zoomIn', key: '=', run: () => setPps((v) => Math.min(600, v * 1.5)) },
    'view.zoomOut': { label: t('Zoom Out'), icon: 'zoomOut', key: '-', run: () => setPps((v) => Math.max(4, v / 1.5)) },
    'view.fit': { label: t('Fit to Window'), icon: 'maximize', key: 'Shift+Z', run: fitView },
    'view.snap': { label: t('Snapping'), icon: 'grid', pressed: doc.settings?.snap !== false, tip: t('clips snap to edges, the playhead and markers; hold Alt to drag freely'), run: () => edit(t('Snapping'), (d) => ({ ...d, settings: { ...d.settings, snap: d.settings?.snap === false } })) },
    'view.follow': { label: t('Follow Playhead'), icon: 'eye', pressed: follow, run: () => setFollow((f) => !f) },
    'view.aspect.landscape': { label: t('Landscape'), icon: 'video', pressed: doc.options.aspect === 'landscape', tip: '16:9', run: () => edit(t('Frame'), (d) => composeIntro({ ...d, options: { ...d.options, aspect: 'landscape' } })) },
    'view.aspect.vertical': { label: t('Vertical'), icon: 'video', pressed: doc.options.aspect === 'vertical', tip: '9:16', run: () => edit(t('Frame'), (d) => composeIntro({ ...d, options: { ...d.options, aspect: 'vertical' } })) },
    'view.aspect.square': { label: t('Square'), icon: 'video', pressed: doc.options.aspect === 'square', tip: '1:1', run: () => edit(t('Frame'), (d) => composeIntro({ ...d, options: { ...d.options, aspect: 'square' } })) },
  };
  for (const kind of ['none', 'fade', 'dip', 'wipe', 'slide', 'push', 'zoom']) {
    const pick = pictures[0] || null;
    const can = Boolean(pick && previousOnLane(doc, pick.id));
    const names = { none: t('None'), fade: t('Crossfade'), dip: t('Dip to Black'), wipe: t('Wipe'), slide: t('Slide'), push: t('Push'), zoom: t('Zoom') };
    const icons = { none: 'trNone', fade: 'trFade', dip: 'trDissolve', wipe: 'trWipe', slide: 'trCover', push: 'trPush', zoom: 'zoomIn' };
    cmds[`transition.${kind}`] = {
      label: names[kind],
      icon: icons[kind],
      disabled: !can,
      pressed: can && (pick.transition?.kind || 'none') === kind,
      tip: t('into the picked clip from the one before it'),
      run: () => edit(t('Transition'), (d) => setTransition(d, pick.id, { kind, seconds: transitionSeconds })),
    };
  }
  const layoutNames = LAYOUT_NAMES();
  for (const l of LAYOUTS) {
    cmds[`intro.layout.${l}`] = {
      label: layoutNames[l],
      disabled: !currentScene,
      pressed: currentScene?.layout === l,
      tip: t('the scene picked, laid out this way; a recorded one keeps its take'),
      run: () => currentScene && currentScene.layout !== l && edit(t('Layout'), (d) => setLayout(d, currentScene.id, l)),
    };
  }
  for (const preset of Object.keys(TITLE_PRESETS)) {
    cmds[`insert.title.${preset}`] = {
      label: titleNames[preset],
      tip: t('at the playhead, on a track above the clips'),
      run: () => {
        let ids = [];
        const words = { title: t('Your title here'), lowerThird: t('Name and role'), endCard: t('Thanks for watching'), callout: t('Look here') }[preset];
        edit(titleNames[preset], (d) => { const r = addTitle(d, preset, { at: timeRef.current, text: words }); ids = r.ids; return r.doc; });
        setSelected(ids.slice(-1));
        setTab('clip');
      },
    };
  }
  useCommands(cmds, [doc, selected, time, path, dirty, clipboard.current, playing, exporting, transitionSeconds, follow, duration, recordOpen, sceneId]);

  // A clip's own menu, on a right-click: the clip is picked first, and the
  // menu opens once the window has drawn it picked, so every item acts on it.
  const contextMenu = useMenu();
  const pendingMenu = useRef(null);
  const CLIP_MENU = ['edit.cut', 'edit.copy', 'edit.paste', '-', 'edit.split', 'edit.trimStart', 'edit.trimEnd', '-', 'edit.detach', 'edit.link', 'edit.duplicate', '-', 'edit.delete', 'edit.rippleDelete'];
  const openClipMenu = (e, clip) => {
    e.preventDefault();
    e.stopPropagation();
    const at = { clientX: e.clientX, clientY: e.clientY, preventDefault() {}, stopPropagation() {} };
    if (selected.includes(clip.id)) contextMenu.open(at, menuItems(cmds, CLIP_MENU));
    else {
      pendingMenu.current = at;
      setSelected([clip.id]);
    }
  };
  useEffect(() => {
    const at = pendingMenu.current;
    if (!at) return;
    pendingMenu.current = null;
    contextMenu.open(at, menuItems(cmds, CLIP_MENU));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected]);

  // The space bar plays and pauses, as in every player.
  useEffect(() => {
    const key = (e) => {
      if (e.key !== ' ' || e.ctrlKey || e.altKey || e.metaKey) return;
      const el = e.target;
      if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT' || el.tagName === 'BUTTON' || el.isContentEditable)) return;
      e.preventDefault();
      toggle();
    };
    window.addEventListener('keydown', key);
    return () => window.removeEventListener('keydown', key);
  }, [toggle]);

  /* ── the window ────────────────────────────────────────────────────── */

  const onLane = (section, lane, change) => edit(t('Track'), (d) => ({ ...d, lanes: { ...d.lanes, [section]: { ...(d.lanes?.[section] || {}), [lane]: { ...((d.lanes?.[section] || {})[lane] || {}), ...change } } } }));
  const dragBase = useRef(null);
  const kindName = (type) => ({ video: t('Video clip'), image: t('Picture'), sound: t('Sound clip'), text: t('Title'), shape: t('Title background'), caption: t('Caption'), pointer: t('Pointer') }[type] || t('Clip'));
  const reason = useMemo(() => unsupportedReason(), []);
  const name = path ? basename(path) : t('Untitled video');

  // A divider between the stage and the timeline, dragged (or moved with the
  // arrow keys) to give either more room; the share is kept for next time,
  // and a double-click puts it back as it starts.
  const keepShare = (share) => shell.store.set({ key: 'video.timelineShare', value: Math.round(share * 1000) / 1000 }).catch(() => {});
  const shareFor = (height) => (editorH > 0 ? splitHeight(editorH, height / editorH) / editorH : TIMELINE_SHARE);
  const startSplit = (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    const y0 = e.clientY;
    const h0 = timelineH;
    let share = timelineShare;
    const move = (ev) => { share = shareFor(h0 - (ev.clientY - y0)); setTimelineShare(share); };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      keepShare(share);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };
  const splitKey = (e) => {
    const by = e.key === 'ArrowUp' ? 24 : e.key === 'ArrowDown' ? -24 : e.key === 'PageUp' ? 96 : e.key === 'PageDown' ? -96 : 0;
    if (!by && e.key !== 'Home') return;
    e.preventDefault();
    const share = e.key === 'Home' ? TIMELINE_SHARE : shareFor(timelineH + by);
    setTimelineShare(share);
    keepShare(share);
  };
  const resetSplit = () => { setTimelineShare(TIMELINE_SHARE); keepShare(TIMELINE_SHARE); };

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={name}
      dirty={dirty}
      menu={appMenu}
      ribbon={
        <VideoRibbon tab={tab} setTab={setTab} cmds={cmds} playing={playing} transitionSeconds={transitionSeconds} setTransitionSeconds={setTransitionSeconds} selection={first?.type || null} intro={intro} />
      }
      status={
        <>
          <span>{path || t('Not saved yet')}</span>
          <Spacer />
          {busy ? <span role="status" aria-live="polite"><Chip>{busy}</Chip></span> : null}
          {doc.range ? <Chip>{t('In {from}, Out {to}', { from: timecode(doc.range.in), to: timecode(doc.range.out) })}</Chip> : null}
          <Chip>{tn(doc.patches.filter((p) => p.timing).length, '{count} clip', '{count} clips')}</Chip>
          <Chip>{doc.options.aspect === 'vertical' ? '9:16' : doc.options.aspect === 'square' ? '1:1' : '16:9'}</Chip>
        </>
      }
    >
      <Content>
        <div className="vd-editor" ref={editorRef} data-duration={duration} data-clips={doc.patches.filter((p) => p.timing).length} data-dirty={dirty ? '1' : '0'} data-copy={copyKey.current} data-starts={playStarts.current} data-export={lastExport ? JSON.stringify({ seconds: Number(lastExport.seconds.toFixed(2)), length: Number(lastExport.total.toFixed(2)), size: lastExport.size, fast: lastExport.fast, hardware: Boolean(lastExport.hardware), codec: lastExport.codec || null, audio: lastExport.audio || null, spent: lastExport.spent || null, sources: lastExport.sources || null }) : undefined}>
          <div className="vd-top">
            {intro ? (
              <Panel title={t('Scenes')} width={272} resizable min={200} max={420}>
                <ScenesPanel
                  doc={doc}
                  current={currentScene?.id || null}
                  speed={doc.intro.prompter?.speed || 140}
                  onPick={pickScene}
                  onRecord={(id) => { stop(); setSceneId(id); setSceneRecord(id); }}
                  onLayout={(id, l) => sceneEdit(t('Layout'), (d) => setLayout(d, id, l))}
                  onMove={(id, by) => sceneEdit(t('Move Scene'), (d) => moveScene(d, id, by))}
                  onRemove={(id) => { sceneEdit(t('Remove Scene'), (d) => removeScene(d, id)); setSceneId(null); }}
                  onAdd={(l) => { let made = null; sceneEdit(t('Add Scene'), (d) => { const r = addScene(d, l, { after: currentScene?.id || null }); made = r.id; return r.doc; }); if (made) setSceneId(made); }}
                  onChange={(id, change, group) => sceneEdit(t('Scene'), (d) => updateScene(d, id, change), group)}
                  onUseTake={(id, takeId) => sceneEdit(t('Use Take'), (d) => useTake(d, id, takeId))}
                  onBrand={() => setBrandOpen(true)}
                />
              </Panel>
            ) : null}
            <div className="vd-monitor">
              {hasClips ? (
                <Stage
                  plan={plan}
                  time={time}
                  selected={selected}
                  onSelect={setSelected}
                  canvasRef={canvasRef}
                  paintTick={paintTick}
                  isLocked={laneLocked}
                  onBegin={(label) => { dragBase.current = docRef.current; beginLive(label); }}
                  onMove={(id, dx, dy) => {
                    const base = dragBase.current || docRef.current;
                    const ids = new Set(withPartners(base, [id]));
                    live({ ...base, patches: base.patches.map((p) => (ids.has(p.id) && standsInBox(p) ? { ...p, fx: Math.round(((p.fx ?? 0.5) + dx) * 10000) / 10000, fy: Math.round(((p.fy ?? 0.5) + dy) * 10000) / 10000 } : p)) });
                  }}
                  onLive={(id, change) => {
                    const base = dragBase.current || docRef.current;
                    live({ ...base, patches: base.patches.map((p) => (p.id === id ? withChange(p, change) : p)) });
                  }}
                />
              ) : (
                <div className="vd-stage">
                  <Empty icon="video" title={t('Start a video')}>
                    {reason || t('Add videos, sound and pictures, or drop them onto this window. They go on the timeline one after another.')}
                    <span className="vd-start-actions">
                      <Button primary icon="plus" label={t('Add Media')} onClick={insertMedia} />
                      <Button icon="record" label={t('Record Screen')} className="vd-start-record" onClick={() => setRecordOpen(true)} />
                    </span>
                  </Empty>
                </div>
              )}
              <div className="vd-transport">
                <Button icon="skipBack" title={t('Go to Start')} onClick={() => { stop(); seek(0); }} />
                <Button icon="chevronLeft" title={t('Previous Frame')} onClick={() => stepFrames(-1)} />
                <Button className="vd-play" icon={playing ? 'pause' : 'play'} title={playing ? t('Pause') : t('Play')} disabled={!hasClips} onClick={toggle} />
                <Button icon="chevronRight" title={t('Next Frame')} onClick={() => stepFrames(1)} />
                <Button icon="skipForward" title={t('Go to End')} onClick={() => { stop(); seek(duration); }} />
                <span className="vd-time">{timecode(time)}</span>
                <span className="vd-scrub-wrap">
                  <input
                    className="vd-scrub"
                    type="range"
                    min={0}
                    max={duration || 0}
                    step={0.01}
                    value={Math.min(time, duration || 0)}
                    aria-label={t('Playhead')}
                    onChange={(e) => { stop(); seek(Number(e.target.value)); }}
                  />
                  {doc.range && duration ? (
                    <ScrubRange
                      range={doc.range}
                      duration={duration}
                      onBegin={beginLive}
                      onMove={(edge, at, base) => live(setRange(base, edge === 'in' ? { in: Math.min(at, base.range.out - 0.1) } : { out: Math.max(at, base.range.in + 0.1) }))}
                      docRef={docRef}
                    />
                  ) : null}
                </span>
                <span className="vd-time total">{timecode(duration)}</span>
              </div>
            </div>
            <Panel title={t('Properties')} width={286} right resizable>
              <Inspector doc={doc} selected={selected} edit={edit} onSeek={seek} missing={missing} kindName={kindName} />
            </Panel>
          </div>
          <div
            className="vd-split"
            role="separator"
            aria-orientation="horizontal"
            aria-label={t('Preview and timeline')}
            aria-valuemin={0}
            aria-valuemax={100}
            aria-valuenow={Math.round(timelineShare * 100)}
            tabIndex={0}
            data-tip={t('Preview and timeline: drag to share the height between them, or press the arrow keys; double-click puts it back')}
            onPointerDown={startSplit}
            onDoubleClick={resetSplit}
            onKeyDown={splitKey}
          />
          <div className="vd-timeline" style={{ height: timelineH }}>
            <Timeline
              doc={doc}
              time={time}
              pps={pps}
              setPps={setPps}
              selected={selected}
              onSelect={(ids) => setSelected(ids)}
              onSeek={(at) => { if (playback.current) restartRef.current?.(at); else seek(at); }}
              onBegin={beginLive}
              onLive={live}
              onLane={onLane}
              media={media}
              mediaVersion={mediaVersion}
              playing={playing}
              missing={missing}
              follow={follow}
              onContext={openClipMenu}
            />
          </div>
        </div>
      </Content>
      {contextMenu.node}
      {exportOpen ? <ExportDialog doc={doc} initialPreset={lastPreset} onClose={() => setExportOpen(false)} onExport={runExport} /> : null}
      {recordOpen ? <RecorderPanel shell={shell} mode="video" toast={toast} onDone={landTakes} onCancel={recordClosed} onProgress={recordingCopy} /> : null}
      {sceneRecord && intro && sceneById(doc, sceneRecord) ? (
        <RecorderPanel
          shell={shell}
          mode="video"
          toast={toast}
          onDone={landScene}
          onCancel={() => setSceneRecord(null)}
          force={{ camera: true }}
          prompter={{
            title: t('Scene {n}: {name}', { n: doc.intro.scenes.findIndex((s) => s.id === sceneRecord) + 1, name: LAYOUT_NAMES()[sceneById(doc, sceneRecord).layout] }),
            script: sceneById(doc, sceneRecord).script || '',
            speed: doc.intro.prompter?.speed || 140,
            onSpeed: (v) => edit(t('Teleprompter speed'), (d) => setPrompterSpeed(d, v)),
          }}
        />
      ) : null}
      {brandOpen ? <BrandDialog shell={shell} brand={doc.intro?.brand || brand} onClose={() => setBrandOpen(false)} onSave={saveBrand} /> : null}
      {exporting ? <ExportProgress progress={exporting.progress} total={exporting.total} fast={exporting.fast} since={exporting.since} from={exporting.from} onCancel={() => exportAbort.current?.abort()} /> : null}
      <style>{CSS}</style>
      {intro || brandOpen ? <style>{INTRO_CSS}</style> : null}
    </AppFrame>
  );
}
