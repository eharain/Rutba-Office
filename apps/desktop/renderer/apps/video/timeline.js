// The timeline: the project's clips on their lanes, against a ruler.
//
// Pictures above, sound below, as every editor draws them: the main picture
// lane at the bottom of the pictures (V1), titles on the lanes above it, the
// clips' sound on A1 under it and anything added on the lanes below. One
// empty lane waits at the top of the pictures and one at the foot of the
// sound, to drag a clip onto. A video clip is a strip of its own frames, a
// sound clip its waveform, a still its picture; a title its words.
//
// Everything here is a view of the project document and of the playhead: a
// drag writes through `onLive` while it moves and is one undo step, recorded
// by `onBegin` when it starts. Snapping (to clip edges, the playhead,
// markers and the In and Out points) is the engine's, at eight pixels; Alt
// while dragging lets go of it.

import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { Icon, t } from '@rutba/office-ui';
import { timecode } from '@rutba/media/timeline';
import {
  laneCount, laneOf, sectionOf, clipsOf, lengthOf, moveClip, trimStart, trimEnd, snapTime, withPartners, setRange, updateMarker, MIN_CLIP,
} from '@rutba/studio/project';

export const HEADER_W = 120;
export const RULER_H = 28;
// A track of pictures is tall enough for its frames; one of titles, or an
// empty one, is lower, so more tracks fit before the timeline scrolls.
const LANE_H = { v: 60, words: 40, a: 48 };
/** How far each can be made shorter or taller by its grip, and the names' column narrower or wider. */
const LANE_LIMITS = { v: [32, 160], words: [24, 120], a: [28, 140] };
const HEAD_LIMITS = [84, 240];
const LAYOUT_KEY = 'rutba.video.timeline';

/** The timeline's own sizes as kept on this computer: the names' column and each kind of lane. */
function keptLayout() {
  try {
    const v = JSON.parse(localStorage.getItem(LAYOUT_KEY) || 'null');
    const clamp = (x, [lo, hi], d) => (Number.isFinite(x) ? Math.max(lo, Math.min(hi, Math.round(x))) : d);
    return {
      head: clamp(v?.head, HEAD_LIMITS, HEADER_W),
      lanes: { v: clamp(v?.lanes?.v, LANE_LIMITS.v, LANE_H.v), words: clamp(v?.lanes?.words, LANE_LIMITS.words, LANE_H.words), a: clamp(v?.lanes?.a, LANE_LIMITS.a, LANE_H.a) },
    };
  } catch {
    return { head: HEADER_W, lanes: { ...LANE_H } };
  }
}
const SNAP_PX = 8;

/** The rows, top to bottom: the picture lanes from the highest down, then the sound lanes. */
export function rowsOf(doc, heights = LANE_H) {
  const n = laneCount(doc);
  const rows = [];
  const pictures = new Set(clipsOf(doc).filter((p) => p.type === 'video' || p.type === 'image').map((p) => laneOf(p)));
  for (let l = n.v; l >= 0; l--) {
    const kind = l === 0 || pictures.has(l) ? 'v' : 'words';
    rows.push({ section: 'v', lane: l, kind, h: heights[kind] });
  }
  for (let l = 0; l <= n.a; l++) rows.push({ section: 'a', lane: l, kind: 'a', h: heights.a });
  return rows;
}

/** A tick step that keeps the ruler's labels apart. */
function tickStep(pps) {
  for (const s of [0.1, 0.25, 0.5, 1, 2, 5, 10, 15, 30, 60, 120, 300, 600]) if (s * pps >= 70) return s;
  return 1200;
}

export default function Timeline({
  doc, time, pps, setPps, selected, onSelect, onSeek, onBegin, onLive, onLane, media, mediaVersion,
  playing, missing, follow = true, onContext,
}) {
  const scrollRef = useRef(null);
  const drag = useRef(null);
  const [width, setWidth] = useState(800);
  const [snapAt, setSnapAt] = useState(null);
  // The names' column and the lanes' heights, dragged by their grips and kept here.
  const [layout, setLayout] = useState(keptLayout);
  const headW = layout.head;
  const headRef = useRef(headW);
  headRef.current = headW;
  // `change` is the new layout, or a function of the one there is now.
  const keepLayout = (change) => setLayout((now) => {
    const next = typeof change === 'function' ? change(now) : change;
    try { localStorage.setItem(LAYOUT_KEY, JSON.stringify(next)); } catch { /* kept for this window alone */ }
    return next;
  });
  const rows = useMemo(() => rowsOf(doc, layout.lanes), [doc, layout.lanes]);
  const duration = useMemo(() => clipsOf(doc).reduce((m, p) => Math.max(m, p.timing.end), 0), [doc]);
  const span = Math.max(duration + 8, (width - headW) / pps);
  const trackW = Math.ceil(span * pps);

  useLayoutEffect(() => {
    const el = scrollRef.current;
    if (!el) return undefined;
    const measure = () => setWidth(el.clientWidth);
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // The playhead stays in view while playing.
  useEffect(() => {
    const el = scrollRef.current;
    if (!el || !follow || !playing) return;
    const x = time * pps;
    const view = el.clientWidth - headRef.current;
    if (x < el.scrollLeft || x > el.scrollLeft + view - 40) el.scrollLeft = Math.max(0, x - view * 0.2);
  }, [time, pps, playing, follow]);

  // Ctrl and the wheel zoom about the pointer; the wheel alone scrolls.
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return undefined;
    const wheel = (e) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      const r = el.getBoundingClientRect();
      const px = e.clientX - r.left - headRef.current + el.scrollLeft;
      const at = px / pps;
      const next = Math.max(4, Math.min(600, pps * (e.deltaY < 0 ? 1.2 : 1 / 1.2)));
      setPps(next);
      requestAnimationFrame(() => {
        el.scrollLeft = Math.max(0, at * next - (e.clientX - r.left - headRef.current));
      });
    };
    el.addEventListener('wheel', wheel, { passive: false });
    return () => el.removeEventListener('wheel', wheel);
  }, [pps, setPps]);

  /** The row under a page y, of a section. */
  const rowAt = useCallback(
    (clientY, section) => {
      const el = scrollRef.current;
      if (!el) return null;
      const r = el.getBoundingClientRect();
      let y = clientY - r.top + el.scrollTop - RULER_H;
      for (const row of rows) {
        if (y < row.h) return row.section === section ? row : null;
        y -= row.h;
      }
      return null;
    },
    [rows]
  );

  const timeAt = useCallback((clientX) => {
    const el = scrollRef.current;
    const r = el.getBoundingClientRect();
    return Math.max(0, (clientX - r.left - headRef.current + el.scrollLeft) / pps);
  }, [pps]);

  /**
   * A point on the ruler dragged along it: the In or Out point, or a marker.
   * It snaps to the clips' edges, the playhead and the other points (Alt
   * frees it), and is one undo step. A marker pressed and let go without
   * moving takes the playhead to it, as before.
   */
  const dragPoint = (e, label, apply, { onClick = null, exclude = null } = {}) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const base = doc;
    const x0 = e.clientX;
    let begun = false;
    const move = (ev) => {
      if (!begun) {
        if (Math.abs(ev.clientX - x0) < 3) return;
        begun = true;
        onBegin(label);
      }
      const raw = timeAt(ev.clientX);
      const threshold = ev.altKey || base.settings?.snap === false ? 0 : SNAP_PX / pps;
      const pointless = exclude ? exclude(base) : base;
      onLive(apply(base, snapTime(pointless, raw, { threshold, playhead: time }).t));
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      if (!begun) onClick?.();
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  /* ── dragging clips: move, trim either end, change lanes ───────────── */

  const startDrag = (e, clip, mode) => {
    if (e.button !== 0) return;
    e.stopPropagation();
    e.preventDefault();
    const additive = e.ctrlKey || e.metaKey || e.shiftKey;
    if (additive) {
      onSelect(selected.includes(clip.id) ? selected.filter((id) => id !== clip.id) : [...selected, clip.id], true);
      return;
    }
    if (!selected.includes(clip.id)) onSelect([clip.id]);
    const locked = (doc.lanes?.[sectionOf(clip)] || {})[laneOf(clip)]?.locked;
    drag.current = { mode, id: clip.id, x0: e.clientX, y0: e.clientY, start: doc, clip, moved: false, locked };
    const move = (ev) => {
      const d = drag.current;
      if (!d || d.locked) return;
      const dx = ev.clientX - d.x0;
      if (!d.moved && Math.abs(dx) < 3 && Math.abs(ev.clientY - d.y0) < 4) return;
      if (!d.moved) {
        d.moved = true;
        onBegin(d.mode === 'move' ? t('Move clip') : t('Trim clip'));
      }
      const dt = dx / pps;
      const threshold = ev.altKey || d.start.settings?.snap === false ? 0 : SNAP_PX / pps;
      const exclude = withPartners(d.start, [d.id]);
      if (!d.beats) d.beats = beatsOnClock(d.start, media, exclude);
      const snap = (v) => snapTime(d.start, v, { threshold, exclude, playhead: time, beats: d.beats });
      let next = d.start;
      if (d.mode === 'move') {
        let start = d.clip.timing.start + dt;
        const a = snap(start);
        const b = snap(start + lengthOf(d.clip));
        if (a.to !== null && (b.to === null || Math.abs(a.to - start) <= Math.abs(b.to - (start + lengthOf(d.clip))))) {
          start = a.to;
          setSnapAt(a.to);
        } else if (b.to !== null) {
          start = b.to - lengthOf(d.clip);
          setSnapAt(b.to);
        } else setSnapAt(null);
        const row = rowAt(ev.clientY, sectionOf(d.clip));
        next = moveClip(d.start, d.id, { start, lane: row ? row.lane : null });
      } else {
        const edge = d.mode === 'start' ? d.clip.timing.start : d.clip.timing.end;
        const s = snap(edge + dt);
        setSnapAt(s.to);
        const ripple = d.start.settings?.ripple !== false;
        next = d.mode === 'start' ? trimStart(d.start, d.id, s.t, { ripple }) : trimEnd(d.start, d.id, s.t, { ripple });
      }
      onLive(next);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      setSnapAt(null);
      drag.current = null;
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  /* ── scrubbing on the ruler ────────────────────────────────────────── */

  const scrub = (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    onSeek(timeAt(e.clientX));
    const move = (ev) => onSeek(timeAt(ev.clientX));
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const step = tickStep(pps);
  const ticks = [];
  for (let s = 0; s <= span; s += step) ticks.push(Math.round(s * 1000) / 1000);
  const minor = step / 5;
  const totalH = rows.reduce((n, r) => n + r.h, 0);
  const range = doc.range;

  /** A grip dragged: the names' column wider or a kind of lane taller, kept when let go; a double-click puts it back. */
  const grip = (what) => (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const x0 = e.clientX;
    const y0 = e.clientY;
    const from = layout;
    let next = from;
    const move = (ev) => {
      next = what === 'head'
        ? { ...from, head: Math.max(HEAD_LIMITS[0], Math.min(HEAD_LIMITS[1], Math.round(from.head + ev.clientX - x0))) }
        : { ...from, lanes: { ...from.lanes, [what]: Math.max(LANE_LIMITS[what][0], Math.min(LANE_LIMITS[what][1], Math.round(from.lanes[what] + ev.clientY - y0))) } };
      setLayout(next);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      keepLayout(next);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };
  const resetGrip = (what) => () => keepLayout((now) => (what === 'head' ? { ...now, head: HEADER_W } : { ...now, lanes: { ...now.lanes, [what]: LANE_H[what] } }));
  const LANE_WORDS = { v: t('Height of the picture tracks'), words: t('Height of the title tracks'), a: t('Height of the sound tracks') };

  return (
    <div className="vd-tl-scroll" ref={scrollRef}>
      <div className="vd-tl-inner" style={{ width: headW + trackW, height: RULER_H + totalH }}>
        <div className="vd-ruler" style={{ width: headW + trackW }}>
          <div className="vd-ruler-corner" style={{ width: headW }}>
            <span className="vd-ruler-time">{timecode(time)}</span>
            <div className="vd-head-grip" role="separator" aria-orientation="vertical" aria-label={t('Width of the track names')} data-tip={t('Drag to make the track names wider or narrower; double-click puts them back')} onPointerDown={grip('head')} onDoubleClick={resetGrip('head')} />
          </div>
          <div className="vd-ruler-track" style={{ width: trackW }} onPointerDown={scrub}>
            {range ? <div className="vd-range" style={{ left: range.in * pps, width: (range.out - range.in) * pps }} /> : null}
            {range ? (
              <>
                <div
                  className="vd-range-edge e-in"
                  data-edge="in"
                  style={{ left: range.in * pps }}
                  role="slider"
                  aria-label={t('In point')}
                  aria-valuenow={range.in}
                  aria-valuetext={timecode(range.in)}
                  data-tip={t('In point {time}: drag to move it (Alt: no snapping)', { time: timecode(range.in) })}
                  onPointerDown={(e) => dragPoint(e, t('Move In'), (d, at) => setRange(d, { in: Math.min(at, d.range.out - MIN_CLIP) }), { exclude: (d) => ({ ...d, range: null }) })}
                />
                <div
                  className="vd-range-edge e-out"
                  data-edge="out"
                  style={{ left: range.out * pps }}
                  role="slider"
                  aria-label={t('Out point')}
                  aria-valuenow={range.out}
                  aria-valuetext={timecode(range.out)}
                  data-tip={t('Out point {time}: drag to move it (Alt: no snapping)', { time: timecode(range.out) })}
                  onPointerDown={(e) => dragPoint(e, t('Move Out'), (d, at) => setRange(d, { out: Math.max(at, d.range.in + MIN_CLIP) }), { exclude: (d) => ({ ...d, range: null }) })}
                />
              </>
            ) : null}
            {ticks.map((s) => (
              <div key={s} className="vd-tick" style={{ left: s * pps }}>
                <span>{timecode(s).replace(/\.0$/, '')}</span>
              </div>
            ))}
            {ticks.flatMap((s) => [1, 2, 3, 4].map((k) => <div key={`${s}-${k}`} className="vd-tick minor" style={{ left: (s + k * minor) * pps }} />))}
            {(doc.markers || []).map((m) => (
              <button
                key={m.id}
                type="button"
                className="vd-marker"
                style={{ left: m.t * pps }}
                data-id={m.id}
                data-tip={m.name ? t('{name}: click to go to it, drag to move it', { name: m.name }) : t('Marker: click to go to it, drag to move it')}
                onPointerDown={(e) => dragPoint(e, t('Move marker'), (d, at) => updateMarker(d, m.id, { t: at }), { onClick: () => onSeek(m.t), exclude: (d) => ({ ...d, markers: (d.markers || []).filter((x) => x.id !== m.id) }) })}
              />
            ))}
            <div className="vd-playhead-cap" style={{ left: time * pps }} />
          </div>
        </div>

        {rows.map((row) => {
          const state = (doc.lanes?.[row.section] || {})[row.lane] || {};
          const clips = clipsOf(doc).filter((p) => sectionOf(p) === row.section && laneOf(p) === row.lane);
          const label = row.section === 'v' ? `V${row.lane + 1}` : `A${row.lane + 1}`;
          return (
            <div key={`${row.section}${row.lane}`} data-track={label} className={`vd-row ${row.section === 'v' ? 'pic' : 'snd'}${clips.length ? '' : ' empty'}${row.section === 'a' && row.lane === 0 ? ' first-sound' : ''}`} style={{ height: row.h }}>
              <div className="vd-row-head" style={{ width: headW }}>
                <div className="vd-lane-grip" role="separator" aria-orientation="horizontal" aria-label={LANE_WORDS[row.kind]} data-tip={t('Drag to make tracks like this one taller or shorter; double-click puts them back')} onPointerDown={grip(row.kind)} onDoubleClick={resetGrip(row.kind)} />
                <span className="vd-row-name">{label}</span>
                {row.section === 'v' ? (
                  <button type="button" className={`vd-row-btn${state.hidden ? ' on' : ''}`} data-tip={state.hidden ? t('Show this track') : t('Hide this track')} onClick={() => onLane(row.section, row.lane, { hidden: !state.hidden })}>
                    <Icon name={state.hidden ? 'eyeOff' : 'eye'} size={13} />
                  </button>
                ) : (
                  <button type="button" className={`vd-row-btn${state.muted ? ' on' : ''}`} data-tip={state.muted ? t('Unmute this track') : t('Mute this track')} onClick={() => onLane(row.section, row.lane, { muted: !state.muted })}>
                    <Icon name="volume" size={13} />
                    {state.muted ? <span className="vd-slash" /> : null}
                  </button>
                )}
                <button type="button" className={`vd-row-btn vd-solo${state.solo ? ' on' : ''}`} data-tip={state.solo ? t('Stop playing this track alone') : t('Play this track alone (solo): the other tracks of its kind are left out until it is switched off')} aria-pressed={Boolean(state.solo)} onClick={() => onLane(row.section, row.lane, { solo: !state.solo })}>
                  <Icon name="headphones" size={12} />
                </button>
                <button type="button" className={`vd-row-btn${state.locked ? ' on' : ''}`} data-tip={state.locked ? t('Unlock this track') : t('Lock this track')} onClick={() => onLane(row.section, row.lane, { locked: !state.locked })}>
                  <Icon name="lock" size={12} />
                </button>
              </div>
              <div className={`vd-row-track${state.hidden || state.muted ? ' dim' : ''}`} style={{ width: trackW }} onPointerDown={(e) => { if (e.target === e.currentTarget && e.button === 0) { onSelect([]); onSeek(timeAt(e.clientX)); } }}>
                {clips.map((clip) => (
                  <Clip
                    key={clip.id}
                    clip={clip}
                    pps={pps}
                    h={row.h - 6}
                    selected={selected.includes(clip.id)}
                    media={media}
                    mediaVersion={mediaVersion}
                    missing={missing?.has?.(clip.url)}
                    source={doc.sources?.[clip.url]}
                    onDown={startDrag}
                    onContext={onContext}
                  />
                ))}
              </div>
            </div>
          );
        })}

        <div className="vd-playhead" style={{ left: headW + time * pps, height: RULER_H + totalH }} />
        {snapAt !== null ? <div className="vd-snapline" style={{ left: headW + snapAt * pps, height: RULER_H + totalH }} /> : null}
      </div>
    </div>
  );
}

const KIND_ICON = { video: 'video', image: 'picture', sound: 'volume', text: 'textbox', shape: 'shape', icon: 'star', pointer: 'pointer' };
/** Words the engine draws right to left: Hebrew, Arabic and Urdu letters, as its text painter tells them. */
const RTL_TEXT = /[\u0590-\u08FF\uFB50-\uFDFF\uFE70-\uFEFC]/;
/** A recording's takes show what they are: the screen, the camera, the voice, the computer's sound. */
const TAKE_ICON = { screen: 'screen', camera: 'camera', voice: 'mic', computer: 'speaker' };

/** One clip: its picture strip, waveform or words, its name and length, its handles. */
function Clip({ clip, pps, h, selected, media, mediaVersion, missing, source, onDown, onContext }) {
  const w = Math.max(4, lengthOf(clip) * pps);
  const kind = clip.type;
  const tr = clip.transition;
  const caption = Boolean(clip.caption && kind === 'text');
  const label = kind === 'text' ? clip.text || t('Title') : kind === 'shape' ? (clip.preset ? t('Title background') : clip.name || t('Shape')) : kind === 'icon' ? clip.label || t('Icon') : kind === 'pointer' ? t('Pointer') : clip.name || '';
  return (
    <div
      className={`vd-clip k-${kind}${caption ? ' k-caption' : ''}${selected ? ' selected' : ''}${missing ? ' missing' : ''}`}
      data-id={clip.id}
      data-kind={kind}
      data-caption={caption ? '1' : undefined}
      data-dir={kind === 'text' ? (RTL_TEXT.test(clip.text || '') ? 'rtl' : 'ltr') : undefined}
      data-start={clip.timing.start}
      data-end={clip.timing.end}
      data-transition={tr ? tr.kind : ''}
      data-fit={clip.fit || ''}
      data-turn={clip.turn || 0}
      data-lane={clip.lane || 0}
      data-take={clip.take || ''}
      data-inset={clip.fw ? [clip.fx, clip.fy, clip.fw, clip.fh, clip.insetShape || ''].join(' ') : ''}
      data-crop={clip.crop ? [clip.crop.x, clip.crop.y, clip.crop.w, clip.crop.h].join(' ') : ''}
      data-zooms={clip.autoZoom ? clip.autoZoom.length : undefined}
      data-clicks={kind === 'pointer' ? (clip.clicks || []).length : undefined}
      style={{ left: clip.timing.start * pps, width: w, height: h }}
      onPointerDown={(e) => onDown(e, clip, 'move')}
      onContextMenu={(e) => onContext?.(e, clip)}
    >
      <ClipArt clip={clip} w={w} h={h} pps={pps} media={media} version={mediaVersion} source={source} />
      <div className="vd-clip-label">
        <Icon name={caption ? 'captions' : TAKE_ICON[clip.take] || KIND_ICON[kind] || 'file'} size={11} />
        <span className="vd-clip-name" dir="auto">{label}</span>
        <span className="vd-clip-len">{timecode(lengthOf(clip))}</span>
      </div>
      {clip.group ? <span className="vd-link" data-tip={t('Linked: picture and sound move together')}><Icon name="link" size={10} /></span> : null}
      {tr ? (
        <div className="vd-tr" style={{ width: Math.max(10, tr.seconds * pps) }} data-tip={t('Transition: {name}', { name: transitionName(tr.kind) })}>
          <Icon name={tr.kind === 'wipe' ? 'trWipe' : tr.kind === 'push' || tr.kind === 'slide' ? 'trPush' : 'trFade'} size={12} />
        </div>
      ) : null}
      <div className="vd-trim start" onPointerDown={(e) => onDown(e, clip, 'start')} />
      <div className="vd-trim end" onPointerDown={(e) => onDown(e, clip, 'end')} />
    </div>
  );
}

/**
 * The beats of the music under the project, on its clock: the onsets of
 * every sound file on the timeline (a song, not a clip's own sound), where
 * each of its clips plays them. Found when a drag starts, once per file.
 */
function beatsOnClock(doc, media, exclude) {
  const skip = new Set(exclude);
  const out = [];
  for (const p of clipsOf(doc)) {
    if (p.type !== 'sound' || skip.has(p.id) || doc.sources?.[p.url]?.kind !== 'audio') continue;
    for (const o of media.beatsOf(p.url)) {
      const at = p.timing.start + (o - (p.offset || 0));
      if (at >= p.timing.start && at <= p.timing.end) out.push(at);
    }
  }
  return out;
}

/** A transition's name, as the ribbon shows it. */
export function transitionName(kind) {
  return { fade: t('Crossfade'), dip: t('Dip to Black'), wipe: t('Wipe'), slide: t('Slide'), push: t('Push'), zoom: t('Zoom') }[kind] || t('None');
}

/** The canvas inside a clip: frames, a waveform, or a colour. Only the part in view is drawn wide. */
function ClipArt({ clip, w, h, pps, media, version, source }) {
  const ref = useRef(null);
  const cw = Math.min(Math.ceil(w), 8192);
  useEffect(() => {
    const c = ref.current;
    if (!c) return;
    const dpr = window.devicePixelRatio || 1;
    c.width = Math.max(1, Math.round(cw * dpr));
    c.height = Math.max(1, Math.round(h * dpr));
    const g = c.getContext('2d');
    g.setTransform(dpr, 0, 0, dpr, 0, 0);
    g.clearRect(0, 0, cw, h);
    const kind = clip.type;
    if (kind === 'video' || kind === 'image') {
      const thumbs = media.thumbs.get(clip.url);
      const first = thumbs?.frames?.find(Boolean);
      const tw = first ? Math.max(24, (h * first.width) / first.height) : h * (16 / 9);
      for (let x = 0; x < cw; x += tw) {
        let frame = first;
        if (kind === 'video' && thumbs) {
          const srcT = (clip.offset || 0) + ((x + tw / 2) / pps) * (Number(clip.speed) || 1);
          let i = Math.min(thumbs.frames.length - 1, Math.max(0, Math.floor(srcT / thumbs.step)));
          while (i > 0 && !thumbs.frames[i]) i--;
          frame = thumbs.frames[i] || first;
        }
        if (frame) g.drawImage(frame, x, 0, tw, h);
        else {
          g.fillStyle = 'rgba(255,255,255,0.05)';
          g.fillRect(x + 1, 1, tw - 2, h - 2);
        }
      }
    } else if (kind === 'pointer') {
      // Each click a dot where it was made, on the clip's own clock.
      g.fillStyle = 'rgba(255, 213, 74, 0.95)';
      for (const c of clip.clicks || []) {
        const x = (c[0] - (clip.offset || 0)) * pps;
        if (x < 0 || x > cw) continue;
        g.beginPath();
        g.arc(x, h / 2, 3.5, 0, Math.PI * 2);
        g.fill();
      }
    } else if (kind === 'sound') {
      const peaks = media.peaks.get(clip.url);
      const mid = h / 2;
      const vol = Math.min(1.4, clip.volume ?? 1);
      g.fillStyle = 'rgba(132, 236, 196, 0.85)';
      if (peaks) {
        const per = peaks.rate / pps;
        for (let x = 0; x < cw; x++) {
          const from = Math.floor(((clip.offset || 0) + (x / pps) * (Number(clip.speed) || 1)) * peaks.rate);
          let m = 0;
          for (let i = from; i < from + Math.max(1, per); i++) if (peaks.data[i] > m) m = peaks.data[i];
          const bar = Math.max(0.5, Math.min(mid - 1, m * vol * (mid - 2)));
          g.fillRect(x, mid - bar, 1, bar * 2);
        }
      } else {
        g.fillRect(0, mid - 0.5, cw, 1);
      }
      // The fades, as the ramps a mixer draws.
      const fin = (clip.enter?.seconds || 0) * pps;
      const fout = (clip.exit?.seconds || 0) * pps;
      g.fillStyle = 'rgba(0,0,0,0.35)';
      if (fin > 1) { g.beginPath(); g.moveTo(0, 0); g.lineTo(fin, 0); g.lineTo(0, h); g.closePath(); g.fill(); }
      if (fout > 1) { g.beginPath(); g.moveTo(cw, 0); g.lineTo(cw - fout, 0); g.lineTo(cw, h); g.closePath(); g.fill(); }
    }
  }, [clip, cw, h, pps, media, version, source]);
  return <canvas ref={ref} className="vd-clip-art" style={{ width: cw, height: h }} />;
}
