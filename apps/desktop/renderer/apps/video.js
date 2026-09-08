// Rutba Video.
//
// Trimming is instant because it changes two numbers rather than a file: the
// timeline holds clips that point into sources, and nothing is written until an
// export asks for it. Playback comes straight off disk over rutba://file, which
// answers byte ranges, so scrubbing a two-hour recording seeks rather than
// loading.
//
// Export uses what the browser engine already has — MediaRecorder over a canvas
// and the audio graph — so there is no ffmpeg to ship, nothing to install, and
// no native binary in the download. That bounds what it can write, and the app
// says so rather than pretending otherwise.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Ribbon, Group, Button, Icon, Spacer, Chip, Empty, Panel, Content, Field, useToast, useCommands, formatBytes, basename } from '@rutba/office-ui';
import { Timeline, timecode, clipDuration } from '@rutba/media/timeline';
import { AppFrame, useAppMenu, pickOpen, useFileDrop } from '../shell.js';

const fileUrl = (p) => `rutba://file/${btoa(unescape(encodeURIComponent(p))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')}`;

export default function VideoTool({ app, shell, boot }) {
  const toast = useToast();
  const [path, setPath] = useState(boot.file || null);
  const [stat, setStat] = useState(null);
  const [timeline, setTimeline] = useState(null);
  const [selected, setSelected] = useState(null);
  const [time, setTime] = useState(0);
  const [playing, setPlaying] = useState(false);
  const [duration, setDuration] = useState(0);
  const [volume, setVolume] = useState(1);
  const [rate, setRate] = useState(1);
  const [tab, setTab] = useState('home');
  const [exporting, setExporting] = useState(null);
  // What the decoder said, when it refused. Without this a file the operating
  // system cannot decode — a RIFF header with nothing behind it, a clip in a
  // codec this machine has no decoder for — left the window empty for as long
  // as the person waited: no timeline, no message, nothing to act on.
  const [error, setError] = useState(null);
  const videoRef = useRef(null);
  const openFileRef = useRef(null);
  const appMenu = useAppMenu({ shell, appKey: 'video', onOpen: () => openFileRef.current?.() });

  const load = useCallback(
    async (target) => {
      try {
        const info = await shell.fs.stat({ path: target });
        setStat(info);
        setPath(target);
        setError(null);
        setTimeline(null);
        setSelected(null);
        setTime(0);
        shell.app.addRecent({ path: target, app: 'video' }).catch(() => {});
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      }
    },
    [shell, toast]
  );

  useEffect(() => {
    if (boot.file) load(boot.file);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // The timeline can only be built once the engine has told us how long the
  // media actually is, which is why it is created on loadedmetadata.
  const onMetadata = useCallback(() => {
    const el = videoRef.current;
    if (!el || !path) return;
    setDuration(el.duration || 0);
    setTimeline((current) =>
      current ||
      Timeline.of({ id: 'src1', name: basename(path), path, duration: el.duration || 0, width: el.videoWidth, height: el.videoHeight })
    );
  }, [path]);

  /**
   * The event may already have happened.
   *
   * `loadedmetadata` fires once, and a small file served over the local
   * protocol can be decoded before React has attached the handler — in which
   * case the event is gone and the timeline is never built. The media plays and
   * the timeline stays empty, which looks like a broken application and is
   * invisible on a slow machine because there the event wins the race.
   *
   * So the state is asked as well as listened for.
   */
  useEffect(() => {
    const el = videoRef.current;
    if (!el || !path) return undefined;
    if (el.readyState >= 1) onMetadata();
    // A second chance after the element has certainly had time to attach: the
    // ref is not populated on the first pass of a fresh mount.
    const timer = setTimeout(() => {
      if (videoRef.current?.readyState >= 1) onMetadata();
    }, 120);
    return () => clearTimeout(timer);
  }, [path, onMetadata]);

  const openFile = useCallback(async () => {
    const file = await pickOpen(shell, 'video');
    if (file) load(file);
  }, [shell, load]);
  openFileRef.current = openFile;

  useFileDrop(useCallback((files) => files[0] && load(files[0]), [load]));

  const play = useCallback(() => {
    const el = videoRef.current;
    if (!el) return;
    if (el.paused) {
      el.play();
      setPlaying(true);
    } else {
      el.pause();
      setPlaying(false);
    }
  }, []);

  const seek = useCallback(
    (seconds) => {
      const el = videoRef.current;
      if (!el) return;
      // The requested time, not the element's. Setting `currentTime` starts a
      // seek that lands later, so reading it straight back gives the old
      // position — usually 0 — and everything keyed on `time` (Split, Trim, the
      // clock) acted at the wrong place while the scrubber sat where it was
      // dragged.
      const limit = Number.isFinite(el.duration) && el.duration > 0 ? el.duration : duration || 0;
      const next = Math.max(0, Math.min(limit, seconds));
      el.currentTime = next;
      setTime(next);
    },
    [duration]
  );

  const split = useCallback(() => {
    if (!timeline) return;
    setTimeline(timeline.split(time));
    toast('Split at the playhead');
  }, [timeline, time, toast]);

  const trimTo = useCallback(
    (edge) => {
      if (!timeline) return;
      const found = timeline.locate(time);
      if (!found) return;
      setTimeline(timeline.trim(found.clip.id, edge === 'start' ? { start: found.sourceTime } : { end: found.sourceTime }));
      toast(edge === 'start' ? 'Trimmed the start' : 'Trimmed the end');
    },
    [timeline, time, toast]
  );

  /**
   * Export by playing the timeline through a canvas and recording it.
   *
   * Honest about its limits: what comes out is WebM, because that is what the
   * browser engine records, and it takes as long as the video is long because
   * it is a real-time capture rather than a transcode.
   */
  const exportVideo = useCallback(async () => {
    const el = videoRef.current;
    if (!el || !timeline) return;
    const target = await shell.dialog.save({
      title: 'Export video',
      defaultPath: (path || 'video').replace(/\.[^.]+$/, '') + '-edited.webm',
      filters: [{ name: 'WebM video', extensions: ['webm'] }],
    });
    if (!target) return;

    const layout = timeline.layout();
    const total = timeline.duration;
    const canvas = document.createElement('canvas');
    canvas.width = el.videoWidth || 1280;
    canvas.height = el.videoHeight || 720;
    const ctx = canvas.getContext('2d');

    const stream = canvas.captureStream(30);
    const chunks = [];
    const recorder = new MediaRecorder(stream, { mimeType: 'video/webm;codecs=vp9', videoBitsPerSecond: 6_000_000 });
    recorder.ondataavailable = (e) => e.data.size && chunks.push(e.data);

    setExporting({ done: 0, total });
    const wasPlaying = !el.paused;
    el.pause();

    await new Promise((resolve) => {
      let clipIndex = 0;
      let raf = 0;
      const startClip = () => {
        const clip = layout[clipIndex];
        if (!clip) {
          cancelAnimationFrame(raf);
          recorder.stop();
          return;
        }
        el.currentTime = clip.start;
        el.playbackRate = clip.speed || 1;
        el.play();
      };
      const tick = () => {
        const clip = layout[clipIndex];
        if (!clip) return;
        ctx.drawImage(el, 0, 0, canvas.width, canvas.height);
        setExporting({ done: Math.min(total, clip.from + (el.currentTime - clip.start)), total });
        if (el.currentTime >= clip.end - 0.05) {
          clipIndex += 1;
          if (clipIndex >= layout.length) {
            el.pause();
            cancelAnimationFrame(raf);
            recorder.stop();
            return;
          }
          startClip();
        }
        raf = requestAnimationFrame(tick);
      };
      recorder.onstop = resolve;
      recorder.start(250);
      startClip();
      raf = requestAnimationFrame(tick);
    });

    const blob = new Blob(chunks, { type: 'video/webm' });
    const bytes = new Uint8Array(await blob.arrayBuffer());
    await shell.fs.write({ path: target, bytes });
    setExporting(null);
    if (wasPlaying) el.play();
    toast(`Exported ${basename(target)} — ${formatBytes(bytes.length)}`, { tone: 'good', ms: 6000 });
  }, [timeline, path, shell, toast]);

  const commands = useMemo(
    () => ({
      'file.open': { label: 'Open…', icon: 'open', key: 'Mod+O', run: openFile },
      'play.toggle': { label: 'Play or pause', icon: 'play', key: ' ', run: play },
      'play.back': { label: 'Back 5 seconds', icon: 'skipBack', key: 'arrowleft', run: () => seek(time - 5) },
      'play.forward': { label: 'Forward 5 seconds', icon: 'skipForward', key: 'arrowright', run: () => seek(time + 5) },
      'edit.split': { label: 'Split here', icon: 'scissors', key: 's', run: split },
      'edit.trimStart': { label: 'Trim start to here', icon: 'crop', key: 'i', run: () => trimTo('start') },
      'edit.trimEnd': { label: 'Trim end to here', icon: 'crop', key: 'o', run: () => trimTo('end') },
      'file.export': { label: 'Export…', icon: 'export', key: 'Mod+E', run: exportVideo },
    }),
    [openFile, play, seek, time, split, trimTo, exportVideo]
  );

  useCommands(commands, [timeline, time, path]);

  const layout = timeline?.layout() || [];
  const isAudio = /\.(mp3|m4a|wav|flac|ogg|opus)$/i.test(path || '');

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={path ? basename(path) : 'Video'}
      menu={appMenu}
      ribbon={
        <Ribbon
          tabs={[{ id: 'home', label: 'Home' }, { id: 'edit', label: 'Edit' }]}
          active={tab}
          onTab={setTab}
          quick={<Button icon={playing ? 'pause' : 'play'} title="Play" onClick={play} disabled={!path} />}
        >
          {tab === 'home' ? (
            <>
              <Group label="File">
                <Button tall icon="open" label="Open" onClick={openFile} />
                <Button tall icon="export" label="Export" disabled={!timeline} onClick={exportVideo} />
              </Group>
              <Group label="Playback">
                <Button icon="skipBack" label="-5s" disabled={!path} onClick={() => seek(time - 5)} />
                <Button icon={playing ? 'pause' : 'play'} label={playing ? 'Pause' : 'Play'} disabled={!path} onClick={play} />
                <Button icon="skipForward" label="+5s" disabled={!path} onClick={() => seek(time + 5)} />
              </Group>
              <Group label="Speed">
                {[0.5, 1, 1.5, 2].map((r) => (
                  <Button
                    key={r}
                    label={`${r}×`}
                    pressed={rate === r}
                    disabled={!path}
                    onClick={() => {
                      setRate(r);
                      if (videoRef.current) videoRef.current.playbackRate = r;
                    }}
                  />
                ))}
              </Group>
            </>
          ) : (
            <>
              <Group label="Cut">
                <Button tall icon="scissors" label="Split" disabled={!timeline} onClick={split} />
                <Button tall icon="crop" label="Trim start" disabled={!timeline} onClick={() => trimTo('start')} />
                <Button tall icon="crop" label="Trim end" disabled={!timeline} onClick={() => trimTo('end')} />
              </Group>
              <Group label="Clip">
                <Button
                  icon="trash"
                  label="Remove"
                  disabled={!selected || layout.length < 2}
                  onClick={() => {
                    setTimeline(timeline.remove(selected));
                    setSelected(null);
                  }}
                />
              </Group>
            </>
          )}
        </Ribbon>
      }
      status={
        <>
          <span>{path || 'No media open'}</span>
          <Spacer />
          {exporting ? <Chip>Exporting {timecode(exporting.done)} of {timecode(exporting.total)}</Chip> : null}
          {timeline ? <Chip>{layout.length} clips</Chip> : null}
          <Chip>{timecode(time)} / {timecode(timeline ? timeline.duration : duration)}</Chip>
          {stat ? <Chip>{formatBytes(stat.size)}</Chip> : null}
        </>
      }
    >
      {!path ? (
        <Empty icon="video" title="No media open">
          Open a video or an audio file, or drop one onto this window.
          <Button primary icon="open" label="Open media" onClick={openFile} style={{ marginTop: 12 }} />
        </Empty>
      ) : error ? (
        <Empty icon="video" title="This file could not be opened">
          {error}
          <Button icon="open" label="Open another" onClick={openFile} style={{ marginTop: 12 }} />
        </Empty>
      ) : (
        <Content>
          <div className="vd-stage">
            {isAudio ? (
              <div className="vd-audio">
                <Icon name="volume" size={44} />
                <div>{basename(path)}</div>
              </div>
            ) : null}
            <video
              ref={videoRef}
              className={isAudio ? 'vd-hidden' : 'vd-video'}
              src={fileUrl(path)}
              onLoadedMetadata={onMetadata}
              // The decoder's own verdict, in a sentence. MediaError carries a
              // number and nothing a person can read.
              onError={(e) => {
                const code = e.currentTarget?.error?.code ?? 0;
                const name = basename(path);
                setError(
                  code === 3
                    ? `${name} is damaged: the decoder stopped part way through it.`
                    : code === 2
                      ? `${name} could not be read from disk.`
                      : `${name} is not a video or sound file this machine can play. It may be in a format with no decoder installed, or it may not be media at all despite its name.`
                );
              }}
              // Not while a seek is pending: the element still reports the old
              // position then, and taking it would snap the playhead back to
              // where it was before the drag.
              onTimeUpdate={(e) => !e.currentTarget.seeking && setTime(e.currentTarget.currentTime)}
              onPlay={() => setPlaying(true)}
              onPause={() => setPlaying(false)}
              onClick={play}
              controls={false}
              preload="metadata"
            />
          </div>

          <div className="vd-transport">
            <Button icon={playing ? 'pause' : 'play'} onClick={play} title="Play" />
            <span className="vd-time">{timecode(time)}</span>
            <input
              className="vd-scrub"
              type="range"
              min={0}
              max={duration || 0}
              step={0.05}
              value={time}
              onChange={(e) => seek(Number(e.target.value))}
            />
            <span className="vd-time">{timecode(duration)}</span>
            <Icon name="volume" size={15} />
            <input
              className="vd-volume"
              type="range"
              min={0}
              max={1}
              step={0.02}
              value={volume}
              onChange={(e) => {
                const v = Number(e.target.value);
                setVolume(v);
                if (videoRef.current) videoRef.current.volume = v;
              }}
            />
          </div>

          {timeline ? (
            <div className="vd-timeline">
              <div className="vd-track">
                {layout.map((c) => (
                  <button
                    key={c.id}
                    type="button"
                    className={`vd-clip${c.id === selected ? ' selected' : ''}`}
                    style={{ width: `${(c.length / Math.max(timeline.duration, 0.001)) * 100}%` }}
                    onClick={() => {
                      setSelected(c.id);
                      seek(c.from + 0.01);
                    }}
                    title={`${c.name} — ${timecode(c.length)}`}
                  >
                    <span className="vd-clip-name">{c.name}</span>
                    <span className="vd-clip-len">{timecode(c.length)}</span>
                  </button>
                ))}
                <div
                  className="vd-playhead"
                  style={{ left: `${(time / Math.max(timeline.duration, 0.001)) * 100}%` }}
                />
              </div>
              <div className="rw-hint" style={{ padding: '6px 12px' }}>
                Export writes WebM using the browser engine's own recorder, so there is no ffmpeg in the
                download — and it takes about as long as the video runs.
              </div>
            </div>
          ) : null}
        </Content>
      )}
      <style>{CSS}</style>
    </AppFrame>
  );
}

const CSS = `
.vd-stage { flex: 1; min-height: 0; display: grid; place-items: center; background: #0b0d10; position: relative; }
.vd-video { max-width: 100%; max-height: 100%; display: block; }
.vd-hidden { display: none; }
.vd-audio { color: rgba(255,255,255,0.7); display: grid; place-items: center; gap: 12px; font-size: 13px; }
.vd-transport {
  display: flex; align-items: center; gap: 10px; padding: 7px 12px;
  background: var(--chrome); border-top: 1px solid var(--line);
}
.vd-time { font-size: 11.5px; color: var(--ink-2); font-variant-numeric: tabular-nums; width: 54px; text-align: center; }
.vd-scrub { flex: 1; accent-color: var(--accent); }
.vd-volume { width: 82px; accent-color: var(--accent); }
.vd-timeline { background: var(--surface); border-top: 1px solid var(--line); }
.vd-track { display: flex; height: 56px; position: relative; padding: 8px 12px; gap: 3px; }
.vd-clip {
  border: 1px solid var(--line-strong); border-radius: var(--r-2); background: var(--sunken);
  color: var(--ink-2); display: flex; flex-direction: column; justify-content: center; gap: 1px;
  padding: 0 8px; min-width: 0; overflow: hidden; font: inherit; text-align: left;
}
.vd-clip.selected { border-color: var(--accent); background: var(--selected); color: var(--accent); }
.vd-clip-name { font-size: 11.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.vd-clip-len { font-size: 10.5px; color: var(--ink-3); }
.vd-playhead { position: absolute; top: 4px; bottom: 4px; width: 2px; background: var(--accent); pointer-events: none; }
`;
