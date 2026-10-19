// Video's play mode (owner, 2026-10-20: "there is no video player left we
// should have play and edit modes, when opening a video it should only open
// in play mode unless user creates new or goes in edit mode").
//
// A video or a sound file opened in Video plays here: the picture as large as
// the window allows, a bar of controls that fades while it plays and comes
// back at a touch of the mouse, the keys a player answers (Space, J, K, L,
// the arrows, M, F, the digits), captions from a .srt or .vtt of the same
// name beside it, picture in picture, full screen, and the files before and
// after it in its folder. Edit Video turns it into a project on the
// timeline; a new video and a .rvid project open in the editor as before.
//
// The file is played straight from disk over rutba://file by the platform's
// own player: nothing is decoded ahead or copied, so a two-hour film opens
// as fast as a short one.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Ribbon, Group, Button, Rows, Select, Icon, Empty, basename, useToast, t } from '@rutba/office-ui';
import { timecode } from '@rutba/media/timeline';
import { parseCaptions, writeVtt } from '@rutba/studio/captions';
import { AppFrame, useAppMenu, useFileDrop } from '../../shell.js';
import { fileUrl, mediaKind, MEDIA_EXTENSIONS } from './media.js';

const STORE_KEY = 'video.player';
const SPEEDS = [0.5, 0.75, 1, 1.25, 1.5, 2];
const HIDE_AFTER_MS = 2500;

/** Whether a file plays here rather than opening in the editor. */
export const playsInPlayer = (file) => ['video', 'audio'].includes(mediaKind(file));

const dirOf = (p) => String(p).replace(/[\\/][^\\/]*$/, '');
const stemOf = (p) => String(p).replace(/\.[^.\\/]+$/, '');
const clock = (s) => timecode(Math.max(0, Number.isFinite(s) ? s : 0)).replace(/\.\d+$/, '');

export const PLAYER_CSS = `
.vp-stage { position: relative; flex: 1; min-height: 0; background: #0b0c0f; display: flex; align-items: center; justify-content: center; overflow: hidden; outline: none; }
.vp-stage.idle { cursor: none; }
.vp-video { width: 100%; height: 100%; object-fit: contain; background: #000; }
.vp-sound { position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 14px; color: rgba(255,255,255,0.86); pointer-events: none; }
.vp-sound-mark { width: 120px; height: 120px; border-radius: 28px; display: grid; place-items: center; background: linear-gradient(135deg, #c2408f, #7a3fc2); box-shadow: 0 18px 40px rgba(0,0,0,0.45); }
.vp-sound b { font-size: 18px; font-weight: 600; max-width: 80%; text-align: center; overflow-wrap: anywhere; }
.vp-big { position: absolute; width: 84px; height: 84px; border-radius: 50%; border: 0; display: grid; place-items: center; background: rgba(10,12,16,0.62); color: #fff; cursor: pointer; box-shadow: 0 8px 26px rgba(0,0,0,0.4); }
.vp-big:hover { background: rgba(194,64,143,0.9); }
.vp-big:focus { outline: none; }
.vp-big:focus-visible { box-shadow: 0 0 0 3px #c2408f, 0 8px 26px rgba(0,0,0,0.4); }
.vp-bar { position: absolute; inset-inline: 0; bottom: 0; padding: 26px 16px 12px; display: grid; gap: 6px; color: #fff;
  background: linear-gradient(to top, rgba(0,0,0,0.78), rgba(0,0,0,0)); transition: opacity 200ms ease; }
.vp-stage.idle .vp-bar { opacity: 0; pointer-events: none; }
.vp-seek { width: 100%; accent-color: #c2408f; margin: 0; cursor: pointer; }
.vp-row { display: flex; align-items: center; gap: 6px; }
.vp-row .rw-btn { color: #fff; }
.vp-row .rw-btn:hover:not(:disabled) { background: rgba(255,255,255,0.14); }
.vp-time { font-variant-numeric: tabular-nums; font-size: 12.5px; color: rgba(255,255,255,0.86); margin-inline: 6px; white-space: nowrap; }
.vp-volume { width: 90px; accent-color: #fff; }
.vp-speed { min-width: 0; height: 28px; font-size: 12px; background: rgba(255,255,255,0.1); color: #fff; border-color: rgba(255,255,255,0.2); }
.vp-speed option { color: #000; }
.vp-spacer { flex: 1; }
.vp-name { font-size: 12.5px; color: rgba(255,255,255,0.72); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 40%; }
.vp-stage:fullscreen .vp-bar { padding-bottom: 18px; }
.vp-error { color: rgba(255,255,255,0.86); }
.vp-video::cue { font-size: 1.15em; background: rgba(0,0,0,0.7); }
/* Captions rise above the bar while it shows. */
.vp-video::-webkit-media-text-track-container { transition: transform 200ms ease; }
.vp-stage:not(.idle) .vp-video::-webkit-media-text-track-container { transform: translateY(-72px); }
`;

/**
 * @param {{ app, shell, file: string, onEdit: (file: string) => void }} props
 */
export default function VideoPlayer({ app, shell, file: first, onEdit }) {
  const toast = useToast();
  const [file, setFile] = useState(first);
  const [playing, setPlaying] = useState(false);
  const [time, setTime] = useState(0);
  const [duration, setDuration] = useState(0);
  const [size, setSize] = useState(null);
  const [volume, setVolume] = useState(1);
  const [muted, setMuted] = useState(false);
  const [speed, setSpeed] = useState(1);
  const [loop, setLoop] = useState(false);
  const [captions, setCaptions] = useState(null); // { url, count } for the track beside the file
  const [showCaptions, setShowCaptions] = useState(true);
  const [error, setError] = useState(null);
  const [idle, setIdle] = useState(false);
  const [full, setFull] = useState(false);
  const [siblings, setSiblings] = useState([]);
  const [tab, setTab] = useState('play');
  const video = useRef(null);
  const stage = useRef(null);
  const hideTimer = useRef(null);
  const sound = mediaKind(file) === 'audio';

  // What this person last listened at.
  useEffect(() => {
    shell.store.get({ key: STORE_KEY, fallback: null }).then((v) => {
      if (!v) return;
      if (Number.isFinite(v.volume)) setVolume(Math.max(0, Math.min(1, v.volume)));
      setMuted(Boolean(v.muted));
      if (SPEEDS.includes(v.speed)) setSpeed(v.speed);
    }).catch(() => {});
  }, [shell]);
  const keep = useCallback((patch) => {
    shell.store.get({ key: STORE_KEY, fallback: null }).then((v) => shell.store.set({ key: STORE_KEY, value: { ...(v || {}), ...patch } })).catch(() => {});
  }, [shell]);

  useEffect(() => { const v = video.current; if (v) { v.volume = volume; v.muted = muted; } }, [volume, muted, file]);
  useEffect(() => { const v = video.current; if (v) v.playbackRate = speed; }, [speed, file]);
  useEffect(() => { const v = video.current; if (v) v.loop = loop; }, [loop, file]);

  // A new file: its place in the recent list, its captions, and the files around it.
  useEffect(() => {
    setError(null);
    setTime(0);
    setDuration(0);
    setSize(null);
    setCaptions(null);
    shell.app.addRecent({ path: file, app: 'video' }).catch(() => {});
    let live = true;
    let url = null;
    (async () => {
      for (const ext of ['vtt', 'srt']) {
        const beside = `${stemOf(file)}.${ext}`;
        if (!(await shell.fs.exists({ path: beside }).catch(() => false))) continue;
        try {
          const { bytes } = await shell.fs.read({ path: beside });
          const { cues } = parseCaptions(new TextDecoder().decode(bytes), beside);
          if (!cues.length || !live) continue;
          url = URL.createObjectURL(new Blob([writeVtt(cues)], { type: 'text/vtt' }));
          setCaptions({ url, count: cues.length, name: basename(beside) });
          break;
        } catch {
          /* captions that do not read are left out */
        }
      }
      const list = await shell.fs.list({ path: dirOf(file) }).catch(() => []);
      if (live) setSiblings(list.filter((f) => !f.dir && playsInPlayer(f.path)).map((f) => f.path).sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' })));
    })();
    return () => { live = false; if (url) URL.revokeObjectURL(url); };
  }, [shell, file]);

  useEffect(() => {
    const v = video.current;
    const tracks = v?.textTracks;
    if (!tracks) return;
    for (const tr of tracks) tr.mode = showCaptions ? 'showing' : 'hidden';
  }, [showCaptions, captions]);

  const at = siblings.indexOf(file);
  const previous = at > 0 ? siblings[at - 1] : null;
  const next = at >= 0 && at < siblings.length - 1 ? siblings[at + 1] : null;

  const toggle = useCallback(() => {
    const v = video.current;
    if (!v || error) return;
    if (v.paused || v.ended) v.play().catch((err) => setError(err.message));
    else v.pause();
  }, [error]);
  const skip = useCallback((by) => {
    const v = video.current;
    if (!v || !Number.isFinite(v.duration)) return;
    v.currentTime = Math.max(0, Math.min(v.duration, v.currentTime + by));
  }, []);
  const seekTo = useCallback((s) => {
    const v = video.current;
    if (v && Number.isFinite(v.duration)) v.currentTime = Math.max(0, Math.min(v.duration, s));
  }, []);
  const step = useCallback((frames) => {
    const v = video.current;
    if (!v) return;
    v.pause();
    v.currentTime = Math.max(0, Math.min(v.duration || 0, v.currentTime + frames / 30));
  }, []);
  const changeVolume = useCallback((x) => {
    const vol = Math.max(0, Math.min(1, x));
    setVolume(vol);
    setMuted(vol === 0);
    keep({ volume: vol, muted: vol === 0 });
  }, [keep]);
  const toggleMute = useCallback(() => setMuted((m) => { keep({ muted: !m }); return !m; }), [keep]);
  const changeSpeed = useCallback((s) => { setSpeed(s); keep({ speed: s }); }, [keep]);

  const toggleFull = useCallback(() => {
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
    else stage.current?.requestFullscreen().catch(() => {});
  }, []);
  useEffect(() => {
    const on = () => setFull(Boolean(document.fullscreenElement));
    document.addEventListener('fullscreenchange', on);
    return () => document.removeEventListener('fullscreenchange', on);
  }, []);
  const pip = useCallback(async () => {
    const v = video.current;
    if (!v || sound) return;
    try {
      if (document.pictureInPictureElement) await document.exitPictureInPicture();
      else await v.requestPictureInPicture();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [sound, toast]);

  const edit = useCallback(() => {
    video.current?.pause();
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
    onEdit(file);
  }, [onEdit, file]);

  const openOther = useCallback(async () => {
    const files = await shell.dialog.open({
      title: t('Open'),
      filters: [
        { name: t('Video projects and media'), extensions: ['rvid', ...MEDIA_EXTENSIONS] },
        { name: t('All files'), extensions: ['*'] },
      ],
    });
    const chosen = files?.[0];
    if (!chosen) return;
    if (playsInPlayer(chosen)) setFile(chosen);
    else shell.win.create({ app: 'video', file: chosen });
  }, [shell]);

  const appMenu = useAppMenu({ shell, appKey: 'video', onNew: () => shell.win.create({ app: 'video' }), onOpen: openOther });

  useFileDrop(useCallback((files) => {
    const media = files.find(playsInPlayer);
    if (media) setFile(media);
    for (const f of files) if (f !== media) shell.win.create({ app: 'video', file: f });
  }, [shell]));

  // The bar fades while the picture plays and nothing moves.
  const wake = useCallback(() => {
    setIdle(false);
    clearTimeout(hideTimer.current);
    hideTimer.current = setTimeout(() => setIdle(true), HIDE_AFTER_MS);
  }, []);
  useEffect(() => () => clearTimeout(hideTimer.current), []);
  useEffect(() => { if (!playing) { clearTimeout(hideTimer.current); setIdle(false); } else wake(); }, [playing, wake]);

  // The keys a player answers.
  useEffect(() => {
    const onKey = (e) => {
      const el = document.activeElement;
      if (el && ['INPUT', 'SELECT', 'TEXTAREA'].includes(el.tagName)) return;
      if (e.ctrlKey || e.metaKey || e.altKey) return;
      const k = e.key;
      let used = true;
      if (k === ' ' || k === 'k' || k === 'K') toggle();
      else if (k === 'j' || k === 'J') skip(-10);
      else if (k === 'l' || k === 'L') skip(10);
      else if (k === 'ArrowLeft') skip(-5);
      else if (k === 'ArrowRight') skip(5);
      else if (k === 'ArrowUp') changeVolume(volume + 0.05);
      else if (k === 'ArrowDown') changeVolume(volume - 0.05);
      else if (k === 'm' || k === 'M') toggleMute();
      else if (k === 'f' || k === 'F') toggleFull();
      else if (k === 'c' || k === 'C') setShowCaptions((s) => !s);
      else if (k === ',') step(-1);
      else if (k === '.') step(1);
      else if (k === 'Home') seekTo(0);
      else if (k === 'End') seekTo(duration);
      else if (/^[0-9]$/.test(k)) seekTo((duration * Number(k)) / 10);
      else used = false;
      if (used) { e.preventDefault(); wake(); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [toggle, skip, changeVolume, volume, toggleMute, toggleFull, step, seekTo, duration, wake]);

  const name = basename(file);
  const speedSelect = (className) => (
    <Select value={String(speed)} className={className} aria-label={t('Speed')} title={t('Speed')} onChange={(e) => changeSpeed(Number(e.target.value))}>
      {SPEEDS.map((s) => <option key={s} value={String(s)}>{s === 1 ? t('Normal speed') : `${s}×`}</option>)}
    </Select>
  );

  const ribbon = (
    <Ribbon
      tabs={[{ id: 'play', label: t('Play') }]}
      active={tab}
      onTab={setTab}
      quick={<Button icon={playing ? 'pause' : 'play'} title={playing ? t('Pause') : t('Play')} onClick={toggle} />}
    >
      <Group label={t('Playback')}>
        <Button tall icon={playing ? 'pause' : 'play'} label={playing ? t('Pause') : t('Play')} className="vp-play" title={`${playing ? t('Pause') : t('Play')} (Space)`} onClick={toggle} />
        <Rows>
          <Button icon="skipBack" label={t('Back 10 seconds')} title={`${t('Back 10 seconds')} (J)`} onClick={() => skip(-10)} />
          <Button icon="skipForward" label={t('Forward 10 seconds')} title={`${t('Forward 10 seconds')} (L)`} onClick={() => skip(10)} />
          <Button icon="refresh" label={t('Repeat')} pressed={loop} className="vp-loop" onClick={() => setLoop((x) => !x)} />
        </Rows>
      </Group>
      <Group label={t('View')}>
        <Button tall icon="maximize" label={t('Full Screen')} title={`${t('Full Screen')} (F)`} onClick={toggleFull} />
        <Rows>
          <Button icon="window" label={t('Picture in Picture')} disabled={sound || error != null} onClick={pip} />
          <Button icon="captions" label={t('Captions')} pressed={Boolean(captions) && showCaptions} disabled={!captions} className="vp-captions" title={captions ? `${t('Captions')}: ${captions.name} (C)` : t('No captions file beside this one')} onClick={() => setShowCaptions((s) => !s)} />
          {speedSelect('vp-speed-ribbon')}
        </Rows>
      </Group>
      <Group label={t('Folder')}>
        <Rows>
          <Button icon="chevronLeft" label={t('Previous File')} disabled={!previous} className="vp-prev" title={previous ? basename(previous) : undefined} onClick={() => previous && setFile(previous)} />
          <Button icon="chevronRight" label={t('Next File')} disabled={!next} className="vp-next" title={next ? basename(next) : undefined} onClick={() => next && setFile(next)} />
          <Button icon="open" label={t('Open')} onClick={openOther} />
        </Rows>
      </Group>
      <Group label={t('Edit')}>
        <Button tall primary icon="scissors" label={t('Edit Video')} className="vp-edit" title={t('Edit Video: put it on the timeline to cut, add titles, sound and more')} onClick={edit} />
      </Group>
    </Ribbon>
  );

  const status = (
    <span className="vp-status" data-mode="play">
      {t('Playing')}
      {size ? ` · ${size.width} × ${size.height}` : ''}
      {duration ? ` · ${clock(duration)}` : ''}
      {siblings.length > 1 && at >= 0 ? ` · ${t('{n} of {count} in this folder', { n: at + 1, count: siblings.length })}` : ''}
    </span>
  );

  return (
    <AppFrame app={app} shell={shell} title={name} menu={appMenu} ribbon={ribbon} status={status}>
      <style>{PLAYER_CSS}</style>
      <div
        ref={stage}
        className={`vp-stage${idle && playing ? ' idle' : ''}`}
        tabIndex={0}
        data-playing={playing ? '1' : '0'}
        data-time={time.toFixed(2)}
        data-duration={duration.toFixed(2)}
        aria-label={name}
        onMouseMove={wake}
        onDoubleClick={toggleFull}
      >
        <video
          key={file}
          ref={video}
          className="vp-video"
          src={fileUrl(file)}
          autoPlay
          playsInline
          onClick={toggle}
          onPlay={() => setPlaying(true)}
          onPause={() => setPlaying(false)}
          onEnded={() => setPlaying(false)}
          onTimeUpdate={(e) => setTime(e.currentTarget.currentTime)}
          onLoadedMetadata={(e) => {
            const v = e.currentTarget;
            setDuration(Number.isFinite(v.duration) ? v.duration : 0);
            setSize(v.videoWidth ? { width: v.videoWidth, height: v.videoHeight } : null);
            v.volume = volume;
            v.muted = muted;
            v.playbackRate = speed;
          }}
          onDurationChange={(e) => { const d = e.currentTarget.duration; if (Number.isFinite(d)) setDuration(d); }}
          onError={() => setError(t('This file cannot be played here. Edit Video can still try to read it.'))}
        >
          {captions ? <track kind="captions" src={captions.url} srcLang="en" label={captions.name} default /> : null}
        </video>
        {sound && !error ? (
          <div className="vp-sound">
            <span className="vp-sound-mark"><Icon name="volume" size={52} /></span>
            <b>{name}</b>
          </div>
        ) : null}
        {error ? (
          <Empty icon="video" title={name}>
            <span className="vp-error">{error}</span>
          </Empty>
        ) : null}
        {!playing && !error ? (
          <button type="button" className="vp-big" aria-label={t('Play')} onClick={toggle}><Icon name="play" size={34} /></button>
        ) : null}
        <div className="vp-bar" onDoubleClick={(e) => e.stopPropagation()}>
          <input
            type="range"
            className="vp-seek"
            min={0}
            max={duration || 0}
            step={0.01}
            value={Math.min(time, duration || 0)}
            aria-label={t('Position')}
            onChange={(e) => seekTo(Number(e.target.value))}
          />
          <div className="vp-row">
            <Button ghost icon={playing ? 'pause' : 'play'} title={playing ? t('Pause') : t('Play')} onClick={toggle} className="vp-bar-play" />
            <Button ghost icon="skipBack" title={t('Back 10 seconds')} onClick={() => skip(-10)} />
            <Button ghost icon="skipForward" title={t('Forward 10 seconds')} onClick={() => skip(10)} />
            <span className="vp-time" aria-live="off">{clock(time)} / {clock(duration)}</span>
            <Button ghost icon={muted || volume === 0 ? 'speaker' : 'volume'} title={muted ? t('Unmute') : t('Mute')} onClick={toggleMute} className="vp-mute" />
            <input type="range" className="vp-volume" min={0} max={1} step={0.01} value={muted ? 0 : volume} aria-label={t('Volume')} onChange={(e) => changeVolume(Number(e.target.value))} />
            <span className="vp-spacer" />
            <span className="vp-name">{name}</span>
            {speedSelect('vp-speed')}
            {captions ? <Button ghost icon="captions" pressed={showCaptions} title={t('Captions')} onClick={() => setShowCaptions((s) => !s)} /> : null}
            {!sound ? <Button ghost icon="window" title={t('Picture in Picture')} onClick={pip} /> : null}
            <Button ghost icon={full ? 'restore' : 'maximize'} title={full ? t('Exit Full Screen') : t('Full Screen')} onClick={toggleFull} />
          </div>
        </div>
      </div>
    </AppFrame>
  );
}
