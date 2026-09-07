// Rutba Pictures.
//
// A viewer whose whole job is to be instant, and to show you whatever you point
// it at. Three things make that true: the grid asks the backend for each file's
// dimensions rather than decoding it, the images themselves are fetched over
// rutba://file so Chromium decodes a 40 MP photograph straight from disk, and
// video arrives by byte range, so a two-hour recording opens as fast as a
// thumbnail does.
//
// It reads EXIF, which matters for one reason above the details panel: a phone
// records which way up the picture is, and a viewer that ignores that shows a
// third of everyone's holiday sideways.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Ribbon, Group, Button, Separator, Icon, Spacer, Chip, Empty, Spinner, Panel, Content, Search,
  Select, useToast, useMenu, useCommands, menuItems, formatBytes, formatWhen, basename, dirname, extname,
} from '@rutba/office-ui';
import { megapixels, aspectName } from '@rutba/imaging/probe';
import { timecode } from '@rutba/media/timeline';
import { AppFrame, useAppMenu, pickOpen, useFileDrop, openInApp } from '../shell.js';

/**
 * Everything this viewer will open.
 *
 * Video and audio are here deliberately. A folder of holiday files is pictures
 * and clips mixed together, and a viewer that silently skips half of them is
 * telling the person their files are missing.
 */
const STILL = ['.jpg', '.jpeg', '.jpe', '.png', '.bmp', '.tif', '.tiff', '.ico', '.svg', '.avif', '.heic', '.heif'];
const ANIMATED = ['.gif', '.webp', '.apng'];
const VIDEO = ['.mp4', '.m4v', '.mov', '.webm', '.mkv', '.ogv', '.avi'];
const AUDIO = ['.mp3', '.m4a', '.wav', '.flac', '.ogg', '.opus'];
const VIEWABLE = [...STILL, ...ANIMATED, ...VIDEO, ...AUDIO, '.pdf'];

const kindOf = (file) => {
  const ext = extname(file);
  if (VIDEO.includes(ext)) return 'video';
  if (AUDIO.includes(ext)) return 'audio';
  if (ext === '.pdf') return 'pdf';
  if (ANIMATED.includes(ext)) return 'maybe-animated';
  return 'still';
};

const fileUrl = (p) =>
  `rutba://file/${btoa(unescape(encodeURIComponent(p))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')}`;

const INTERVALS = [
  { label: '2s', value: 2000 },
  { label: '3s', value: 3000 },
  { label: '5s', value: 5000 },
  { label: '10s', value: 10000 },
  { label: '30s', value: 30000 },
];

export default function Pictures({ app, shell, boot }) {
  const toast = useToast();
  const [folder, setFolder] = useState(null);
  const [entries, setEntries] = useState([]);
  const [current, setCurrent] = useState(null);
  const [details, setDetails] = useState(null);
  const [query, setQuery] = useState('');
  const [busy, setBusy] = useState(false);
  const [tab, setTab] = useState('home');

  // View state
  const [zoom, setZoom] = useState(0); // 0 = fit to window
  const [spin, setSpin] = useState(0); // view-only rotation, in quarter turns
  const [showDetails, setShowDetails] = useState(true);
  const [filmstrip, setFilmstrip] = useState(true);

  // The auto-navigator
  const [playing, setPlaying] = useState(false);
  const [interval, setIntervalMs] = useState(3000);
  const [loop, setLoop] = useState(true);
  const [shuffle, setShuffle] = useState(false);

  // Animated stills can be held still.
  const [frozen, setFrozen] = useState(false);
  const imgRef = useRef(null);
  const freezeRef = useRef(null);
  const mediaRef = useRef(null);
  const stripRef = useRef(null);

  const menu = useMenu();
  const openFileRef = useRef(null);
  const appMenu = useAppMenu({ shell, appKey: 'pictures', onOpen: () => openFileRef.current?.() });

  /* ── the folder ──────────────────────────────────────────────────────── */

  const list = useCallback(
    async (dir) => {
      setBusy(true);
      try {
        const rows = await shell.fs.list({ path: dir, filter: VIEWABLE });
        setEntries(rows);
        setFolder(dir);
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      } finally {
        setBusy(false);
      }
    },
    [shell, toast]
  );

  useEffect(() => {
    const run = async () => {
      if (boot.file) {
        setCurrent(boot.file);
        await list(dirname(boot.file));
      } else {
        const paths = await shell.app.paths();
        await list(paths.pictures);
      }
    };
    run();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const files = useMemo(() => {
    const rows = entries.filter((e) => !e.dir);
    if (!query.trim()) return rows;
    const q = query.toLowerCase();
    return rows.filter((r) => r.name.toLowerCase().includes(q));
  }, [entries, query]);

  const at = files.findIndex((f) => f.path === current);
  const kind = current ? kindOf(current) : null;
  const isAnimated = kind === 'maybe-animated' && details?.probe?.animated;
  const isMedia = kind === 'video' || kind === 'audio';

  /* ── details, read without decoding the whole file ───────────────────── */

  useEffect(() => {
    if (!current) return undefined;
    let alive = true;
    setFrozen(false);
    (async () => {
      try {
        const { bytes, stat } = await shell.fs.read({ path: current });
        const head = bytes.slice(0, 262144);
        const { probeImage } = await import('@rutba/imaging/probe');
        const { readExif, describeExif, orientationOf } = await import('@rutba/imaging/exif');
        const { probeMedia } = await import('@rutba/media/probe');
        const probe = probeImage(head) || probeMedia(head);
        const exif = readExif(head);
        if (alive) setDetails({ stat, probe, exif, rows: describeExif(exif), orientation: orientationOf(exif) });
      } catch {
        if (alive) setDetails(null);
      }
    })();
    return () => {
      alive = false;
    };
  }, [current, shell]);

  /* ── navigation ──────────────────────────────────────────────────────── */

  const goTo = useCallback(
    (index) => {
      if (!files.length) return;
      const next = files[((index % files.length) + files.length) % files.length];
      if (next) {
        setCurrent(next.path);
        setZoom(0);
        setSpin(0);
      }
    },
    [files]
  );

  const step = useCallback(
    (delta) => {
      if (!files.length) return;
      if (shuffle && delta > 0) return goTo(Math.floor(Math.random() * files.length));
      const next = at + delta;
      // Without loop, stop at the ends rather than wrapping round.
      if (!loop && (next < 0 || next >= files.length)) {
        setPlaying(false);
        return undefined;
      }
      return goTo(next);
    },
    [files, at, loop, shuffle, goTo]
  );

  /**
   * The auto-navigator.
   *
   * A video is left to finish rather than cut off mid-sentence: the timer is
   * suspended while one plays, and the clip's own `ended` moves us on.
   */
  useEffect(() => {
    if (!playing || files.length < 2) return undefined;
    if (kind === 'video' || kind === 'audio') return undefined;
    const timer = setTimeout(() => step(1), interval);
    return () => clearTimeout(timer);
  }, [playing, interval, files.length, kind, step, current]);

  // Keep the filmstrip's current thumbnail in view as the selection moves.
  useEffect(() => {
    const strip = stripRef.current;
    const active = strip?.querySelector('.pv-strip-item.active');
    active?.scrollIntoView({ block: 'nearest', inline: 'center', behavior: 'smooth' });
  }, [current, filmstrip]);

  /* ── animated stills: hold the frame ─────────────────────────────────── */

  const freeze = useCallback(() => {
    const img = imgRef.current;
    const canvas = freezeRef.current;
    if (!img || !canvas) return;
    canvas.width = img.naturalWidth;
    canvas.height = img.naturalHeight;
    canvas.getContext('2d').drawImage(img, 0, 0);
    setFrozen(true);
  }, []);

  const toggleAnimation = useCallback(() => {
    if (frozen) setFrozen(false);
    else freeze();
  }, [frozen, freeze]);

  /* ── opening ─────────────────────────────────────────────────────────── */

  const openFolder = useCallback(async () => {
    const dirs = await shell.dialog.open({ title: 'Open a folder', directory: true });
    if (dirs[0]) {
      await list(dirs[0]);
      setCurrent(null);
    }
  }, [shell, list]);

  const openFile = useCallback(async () => {
    const file = await pickOpen(shell, 'pictures');
    if (!file) return;
    setCurrent(file);
    await list(dirname(file));
  }, [shell, list]);
  openFileRef.current = openFile;

  useFileDrop(
    useCallback(
      async (dropped) => {
        const first = dropped[0];
        if (!first) return;
        const stat = await shell.fs.stat({ path: first });
        if (stat?.dir) await list(first);
        else {
          setCurrent(first);
          await list(dirname(first));
        }
      },
      [shell, list]
    )
  );

  /* ── commands ────────────────────────────────────────────────────────── */

  const commands = useMemo(
    () => ({
      'file.open': { label: 'Open…', icon: 'open', key: 'Mod+O', run: openFile },
      'file.folder': { label: 'Open folder…', icon: 'folderOpen', run: openFolder },
      'nav.first': { label: 'First', icon: 'skipBack', key: 'Home', run: () => goTo(0) },
      'nav.prev': { label: 'Previous', icon: 'chevronLeft', key: 'arrowleft', run: () => step(-1) },
      'nav.next': { label: 'Next', icon: 'chevronRight', key: 'arrowright', run: () => step(1) },
      'nav.last': { label: 'Last', icon: 'skipForward', key: 'End', run: () => goTo(files.length - 1) },
      'nav.play': { label: playing ? 'Stop' : 'Play automatically', icon: playing ? 'pause' : 'play', key: ' ', run: () => setPlaying((p) => !p) },
      'view.fit': { label: 'Fit to window', icon: 'maximize', key: 'Mod+0', run: () => setZoom(0) },
      'view.actual': { label: 'Actual size', icon: 'check', key: 'Mod+1', run: () => setZoom(1) },
      'view.in': { label: 'Zoom in', icon: 'zoomIn', key: 'Mod+Plus', run: () => setZoom((z) => Math.min(12, (z || 1) * 1.25)) },
      'view.out': { label: 'Zoom out', icon: 'zoomOut', key: 'Mod+-', run: () => setZoom((z) => Math.max(0.05, (z || 1) / 1.25)) },
      'view.rotate': { label: 'Rotate view', icon: 'rotate', key: 'r', run: () => setSpin((s) => (s + 1) % 4) },
      'view.details': { label: 'Details panel', icon: 'info', run: () => setShowDetails((s) => !s) },
      'view.filmstrip': { label: 'Filmstrip', icon: 'list', run: () => setFilmstrip((f) => !f) },
      'view.fullscreen': { label: 'Full screen', icon: 'maximize', key: 'F11', run: () => shell.win.fullscreen({}) },
      'media.animation': { label: frozen ? 'Resume animation' : 'Hold this frame', icon: frozen ? 'play' : 'pause', run: toggleAnimation },
      'file.edit': { label: 'Edit this picture', icon: 'wand', run: () => current && shell.win.create({ app: 'image', file: current }) },
      'file.video': { label: 'Open in Video', icon: 'video', run: () => current && shell.win.create({ app: 'video', file: current }) },
      'file.reveal': { label: 'Show in folder', icon: 'folderOpen', run: () => current && shell.shell.showInFolder({ path: current }) },
      'file.copy': {
        label: 'Copy the path',
        icon: 'copy',
        run: () => current && shell.clipboard.writeText({ text: current }).then(() => toast('Path copied')),
      },
      'file.trash': {
        label: 'Move to trash',
        icon: 'trash',
        key: 'Delete',
        run: async () => {
          if (!current) return;
          await shell.fs.remove({ path: current });
          toast('Moved to the trash', { tone: 'good' });
          const dir = folder;
          setCurrent(null);
          list(dir);
        },
      },
    }),
    [openFile, openFolder, step, goTo, files.length, playing, frozen, toggleAnimation, current, folder, list, shell, toast]
  );

  useCommands(commands, [files, at, current, folder, playing, frozen]);

  const rotation = ((details?.orientation?.rotate || 0) + spin * 90) % 360;
  const mirrored = details?.orientation?.flip;

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={current ? basename(current) : folder ? basename(folder) : 'Pictures'}
      menu={appMenu}
      ribbon={
        <Ribbon
          tabs={[
            { id: 'home', label: 'Home' },
            { id: 'navigate', label: 'Navigate' },
            { id: 'view', label: 'View' },
          ]}
          active={tab}
          onTab={setTab}
          quick={
            <>
              <Button icon="chevronLeft" title="Previous" onClick={() => step(-1)} disabled={!files.length} />
              <Button icon={playing ? 'pause' : 'play'} title={playing ? 'Stop' : 'Play automatically'} onClick={() => setPlaying((p) => !p)} disabled={files.length < 2} />
              <Button icon="chevronRight" title="Next" onClick={() => step(1)} disabled={!files.length} />
            </>
          }
        >
          {tab === 'home' ? (
            <>
              <Group label="Open">
                <Button tall icon="open" label="File" onClick={openFile} />
                <Button tall icon="folderOpen" label="Folder" onClick={openFolder} />
              </Group>
              <Group label="This file">
                <Button tall icon="wand" label="Edit" disabled={!current || isMedia} onClick={() => commands['file.edit'].run()} />
                <Button tall icon="video" label="In Video" disabled={!current || !isMedia} onClick={() => commands['file.video'].run()} />
                <Button tall icon="folderOpen" label="Reveal" disabled={!current} onClick={() => commands['file.reveal'].run()} />
                <Button tall icon="trash" label="Trash" disabled={!current} onClick={() => commands['file.trash'].run()} />
              </Group>
              <Group label="Clipboard">
                <Button icon="copy" label="Copy path" disabled={!current} onClick={() => commands['file.copy'].run()} />
              </Group>
            </>
          ) : tab === 'navigate' ? (
            <>
              <Group label="Move">
                <Button tall icon="skipBack" label="First" disabled={!files.length} onClick={() => goTo(0)} />
                <Button tall icon="chevronLeft" label="Previous" disabled={!files.length} onClick={() => step(-1)} />
                <Button tall icon="chevronRight" label="Next" disabled={!files.length} onClick={() => step(1)} />
                <Button tall icon="skipForward" label="Last" disabled={!files.length} onClick={() => goTo(files.length - 1)} />
              </Group>
              <Group label="Automatic">
                <Button
                  tall
                  icon={playing ? 'pause' : 'play'}
                  label={playing ? 'Stop' : 'Play'}
                  disabled={files.length < 2}
                  onClick={() => setPlaying((p) => !p)}
                />
                <Select value={interval} onChange={(e) => setIntervalMs(Number(e.target.value))} style={{ width: 74 }} title="How long each picture is shown">
                  {INTERVALS.map((i) => (
                    <option key={i.value} value={i.value}>{i.label}</option>
                  ))}
                </Select>
                <Button icon="refresh" label="Loop" pressed={loop} onClick={() => setLoop((l) => !l)} />
                <Button icon="sort" label="Shuffle" pressed={shuffle} onClick={() => setShuffle((s) => !s)} />
              </Group>
              <Group label="Animation">
                <Button
                  tall
                  icon={frozen ? 'play' : 'pause'}
                  label={frozen ? 'Resume' : 'Hold frame'}
                  disabled={!isAnimated}
                  onClick={toggleAnimation}
                />
              </Group>
            </>
          ) : (
            <>
              <Group label="Zoom">
                <Button icon="zoomOut" label="Out" onClick={() => commands['view.out'].run()} />
                <Button icon="zoomIn" label="In" onClick={() => commands['view.in'].run()} />
                <Button icon="maximize" label="Fit" pressed={zoom === 0} onClick={() => setZoom(0)} />
                <Button icon="check" label="100%" pressed={zoom === 1} onClick={() => setZoom(1)} />
              </Group>
              <Group label="Orientation">
                <Button tall icon="rotate" label="Rotate" onClick={() => setSpin((s) => (s + 1) % 4)} />
              </Group>
              <Group label="Panels">
                <Button icon="info" label="Details" pressed={showDetails} onClick={() => setShowDetails((s) => !s)} />
                <Button icon="list" label="Filmstrip" pressed={filmstrip} onClick={() => setFilmstrip((f) => !f)} />
                <Separator />
                <Button icon="maximize" label="Full screen" onClick={() => shell.win.fullscreen({})} />
              </Group>
            </>
          )}
        </Ribbon>
      }
      status={
        <>
          <span>{folder || ''}</span>
          <Spacer />
          {details?.probe?.duration ? <Chip>{timecode(details.probe.duration)}</Chip> : null}
          {details?.probe?.width ? (
            <>
              <Chip>{details.probe.width} × {details.probe.height}</Chip>
              <Chip>{megapixels(details.probe.width, details.probe.height)} MP</Chip>
              <Chip>{aspectName(details.probe.width, details.probe.height)}</Chip>
            </>
          ) : null}
          {isAnimated ? <Chip title="This picture is animated">Animated</Chip> : null}
          {zoom ? <Chip>{Math.round(zoom * 100)}%</Chip> : <Chip>Fit</Chip>}
          {details?.stat ? <Chip>{formatBytes(details.stat.size)}</Chip> : null}
          <Chip>{files.length ? `${at + 1} of ${files.length}` : 'nothing here'}</Chip>
        </>
      }
    >
      <Panel width={228} resizable>
        <div style={{ padding: '8px 10px' }}>
          <Search value={query} onChange={setQuery} placeholder="Filter by name" />
        </div>
        {busy ? (
          <div style={{ padding: 22, display: 'grid', placeItems: 'center' }}>
            <Spinner />
          </div>
        ) : (
          <div className="pv-grid">
            {files.map((f) => (
              <Thumb key={f.path} file={f} active={f.path === current} onOpen={() => goTo(files.indexOf(f))} onMenu={(e) => menu.open(e, menuItems(commands, ['file.edit', 'file.reveal', 'file.copy', '-', 'file.trash']))} />
            ))}
          </div>
        )}
      </Panel>

      <Content>
        <div
          className="pv-stage"
          onWheel={(e) => {
            if (!e.ctrlKey) return;
            e.preventDefault();
            setZoom((z) => Math.max(0.05, Math.min(12, (z || 1) * (e.deltaY < 0 ? 1.1 : 0.9))));
          }}
          onDoubleClick={() => setZoom((z) => (z ? 0 : 1))}
          onContextMenu={(e) => menu.open(e, menuItems(commands, ['nav.prev', 'nav.next', '-', 'view.fit', 'view.actual', 'view.rotate', '-', 'file.edit', 'file.reveal', '-', 'file.trash']))}
        >
          {!current ? (
            <Empty icon="pictures" title={files.length ? 'Choose something to view' : 'Nothing here yet'}>
              {files.length
                ? 'Pick a file from the list, or press the play button to run through them.'
                : 'Open a folder, or drop pictures, video or animations onto this window.'}
            </Empty>
          ) : kind === 'pdf' ? (
            <embed src={fileUrl(current)} type="application/pdf" className="pv-pdf" />
          ) : isMedia ? (
            /*
              Video and audio play here rather than sending you to another app.
              `ended` advances the slideshow, so a folder of pictures and clips
              plays through as one sequence instead of stalling on the first
              clip.
            */
            <video
              ref={mediaRef}
              key={current}
              className={kind === 'audio' ? 'pv-audio' : 'pv-video'}
              src={fileUrl(current)}
              controls
              autoPlay={playing}
              onEnded={() => playing && step(1)}
              onError={() => toast('This file could not be played.', { tone: 'bad' })}
            />
          ) : (
            <>
              <img
                ref={imgRef}
                className={`pv-image${zoom ? ' zoomed' : ''}${frozen ? ' hidden' : ''}`}
                src={fileUrl(current)}
                alt={basename(current)}
                style={{
                  transform: `rotate(${rotation}deg)${mirrored ? ' scaleX(-1)' : ''}`,
                  width: zoom ? `${(details?.probe?.width || 0) * zoom}px` : undefined,
                  height: zoom ? 'auto' : undefined,
                }}
                draggable={false}
                onError={() => toast('This picture could not be decoded.', { tone: 'bad' })}
              />
              {/* The held frame of an animation, drawn once and shown in its place. */}
              <canvas ref={freezeRef} className={`pv-image${frozen ? '' : ' hidden'}`} style={{ transform: `rotate(${rotation}deg)` }} />
            </>
          )}
        </div>

        {/* The navigator: where you are, and every way to move. */}
        {files.length ? (
          <div className="pv-nav">
            <Button icon="skipBack" title="First" onClick={() => goTo(0)} />
            <Button icon="chevronLeft" title="Previous" onClick={() => step(-1)} />
            <Button
              icon={playing ? 'pause' : 'play'}
              primary={playing}
              title={playing ? 'Stop' : 'Play automatically'}
              disabled={files.length < 2}
              onClick={() => setPlaying((p) => !p)}
            />
            <Button icon="chevronRight" title="Next" onClick={() => step(1)} />
            <Button icon="skipForward" title="Last" onClick={() => goTo(files.length - 1)} />

            <input
              className="pv-scrub"
              type="range"
              min={0}
              max={Math.max(0, files.length - 1)}
              value={Math.max(0, at)}
              onChange={(e) => goTo(Number(e.target.value))}
              title="Scrub through the folder"
            />
            <span className="pv-counter">{files.length ? `${at + 1} / ${files.length}` : '—'}</span>

            <Separator />
            {isAnimated ? (
              <Button icon={frozen ? 'play' : 'pause'} title={frozen ? 'Resume animation' : 'Hold this frame'} onClick={toggleAnimation} />
            ) : null}
            <Button icon="rotate" title="Rotate the view" onClick={() => setSpin((s) => (s + 1) % 4)} />
            <Button icon="zoomOut" title="Zoom out" onClick={() => commands['view.out'].run()} />
            <Button icon="zoomIn" title="Zoom in" onClick={() => commands['view.in'].run()} />
            <Button icon="maximize" title="Fit to window" pressed={zoom === 0} onClick={() => setZoom(0)} />
          </div>
        ) : null}

        {filmstrip && files.length ? (
          <div className="pv-strip" ref={stripRef}>
            {files.map((f, i) => (
              <button
                key={f.path}
                type="button"
                className={`pv-strip-item${f.path === current ? ' active' : ''}`}
                onClick={() => goTo(i)}
                title={f.name}
              >
                <Preview file={f} />
              </button>
            ))}
          </div>
        ) : null}
      </Content>

      {showDetails && details ? (
        <Panel right width={230} title="Details" resizable>
          <div className="pv-details">
            <div className="pv-detail-name">{basename(current)}</div>
            <dl>
              {details.probe?.format ? (
                <>
                  <dt>Format</dt>
                  <dd>{String(details.probe.format).toUpperCase()}{details.probe.animated ? ' · animated' : ''}</dd>
                </>
              ) : null}
              {details.probe?.width ? (
                <>
                  <dt>Dimensions</dt>
                  <dd>{details.probe.width} × {details.probe.height}</dd>
                </>
              ) : null}
              {details.probe?.duration ? (
                <>
                  <dt>Duration</dt>
                  <dd>{timecode(details.probe.duration)}</dd>
                </>
              ) : null}
              <dt>Size</dt>
              <dd>{formatBytes(details.stat?.size || 0)}</dd>
              <dt>Modified</dt>
              <dd>{formatWhen(details.stat?.mtime, { long: true })}</dd>
              {details.rows.map((r) => (
                <React.Fragment key={r.label}>
                  <dt>{r.label}</dt>
                  <dd>{r.value}</dd>
                </React.Fragment>
              ))}
            </dl>
            {!isMedia ? <Button icon="wand" label="Edit this picture" onClick={() => commands['file.edit'].run()} /> : null}
          </div>
        </Panel>
      ) : null}

      {menu.node}
      <style>{CSS}</style>
    </AppFrame>
  );
}

/** A grid tile. Video gets a still poster from the file itself. */
function Thumb({ file, active, onOpen, onMenu }) {
  const kind = kindOf(file.path);
  return (
    <button type="button" className={`pv-tile${active ? ' active' : ''}`} onClick={onOpen} onContextMenu={onMenu} title={file.name}>
      <Preview file={file} />
      {kind === 'video' || kind === 'audio' ? (
        <span className="pv-tile-badge"><Icon name={kind === 'audio' ? 'volume' : 'play'} size={11} /></span>
      ) : null}
      <span className="pv-tile-name">{file.name}</span>
    </button>
  );
}

/**
 * The picture inside a tile.
 *
 * A video's poster is the video itself with `preload="metadata"`, which fetches
 * only the header and the first frame over a byte range — a folder of clips
 * costs a few kilobytes to draw, not the clips.
 */
function Preview({ file }) {
  const kind = kindOf(file.path);
  if (kind === 'pdf') return <span className="pv-tile-icon"><Icon name="pdf" size={20} /></span>;
  if (kind === 'audio') return <span className="pv-tile-icon"><Icon name="volume" size={20} /></span>;
  if (kind === 'video') return <video className="pv-thumb" src={fileUrl(file.path)} preload="metadata" muted playsInline />;
  return <img className="pv-thumb" src={fileUrl(file.path)} alt="" loading="lazy" decoding="async" />;
}

const CSS = `
.pv-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(96px, 1fr)); gap: 6px; padding: 4px 8px 12px; }
.pv-tile {
  border: 2px solid transparent; border-radius: var(--r-2); background: var(--sunken); padding: 0;
  overflow: hidden; display: flex; flex-direction: column; aspect-ratio: 1; position: relative;
}
.pv-thumb { width: 100%; height: 100%; object-fit: cover; display: block; background: var(--sunken); }
.pv-tile-icon { flex: 1; display: grid; place-items: center; color: var(--ink-3); }
.pv-tile-badge {
  position: absolute; top: 4px; right: 4px; width: 18px; height: 18px; border-radius: 50%;
  background: rgba(0,0,0,0.55); color: #fff; display: grid; place-items: center;
}
.pv-tile-name {
  position: absolute; left: 0; right: 0; bottom: 0; padding: 3px 5px; font-size: 10px;
  color: #fff; background: linear-gradient(transparent, rgba(0,0,0,0.72)); text-align: left;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.pv-tile:hover { border-color: var(--line-strong); }
.pv-tile.active { border-color: var(--accent); }

.pv-stage {
  flex: 1; min-height: 0; overflow: auto; display: grid; place-items: center;
  background: var(--sunken); padding: 16px; position: relative;
}
.pv-image { max-width: 100%; max-height: 100%; object-fit: contain; display: block; box-shadow: var(--shadow-2); }
.pv-image.zoomed { max-width: none; max-height: none; }
.pv-image.hidden { display: none; }
.pv-video { max-width: 100%; max-height: 100%; display: block; background: #000; box-shadow: var(--shadow-2); }
.pv-audio { width: min(560px, 90%); }
.pv-pdf { width: 100%; height: 100%; border: 0; }

.pv-nav {
  display: flex; align-items: center; gap: 4px; padding: 5px 10px;
  background: var(--chrome); border-top: 1px solid var(--line);
}
.pv-scrub { flex: 1; accent-color: var(--accent); margin: 0 8px; }
.pv-counter {
  font-size: 11.5px; color: var(--ink-2); font-variant-numeric: tabular-nums;
  min-width: 70px; text-align: center;
}

.pv-strip {
  display: flex; gap: 5px; padding: 6px 10px; overflow-x: auto; overflow-y: hidden;
  background: var(--surface); border-top: 1px solid var(--line-soft); scrollbar-width: thin;
}
.pv-strip-item {
  flex: none; width: 74px; height: 54px; border: 2px solid transparent; border-radius: var(--r-2);
  overflow: hidden; padding: 0; background: var(--sunken);
}
.pv-strip-item:hover { border-color: var(--line-strong); }
.pv-strip-item.active { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }

.pv-details { padding: 4px 12px 14px; display: flex; flex-direction: column; gap: 12px; }
.pv-detail-name { font-weight: 600; font-size: 12.5px; word-break: break-all; }
.pv-details dl { display: grid; grid-template-columns: auto 1fr; gap: 4px 10px; margin: 0; font-size: 11.5px; }
.pv-details dt { color: var(--ink-3); }
.pv-details dd { margin: 0; word-break: break-word; }
`;
