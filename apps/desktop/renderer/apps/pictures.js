// Rutba Pictures.
//
// A viewer whose whole job is to be instant, and to show you whatever you point
// it at. Four things make that true: a tile is a small JPEG the platform made
// once and kept (rutba://thumb), never the file itself; the grid and the
// filmstrip draw only the tiles you can see, so a folder of five thousand
// costs what a folder of fifty costs; the picture on the stage stays until
// the next one has decoded, so moving on never shows an empty frame; and
// video arrives by byte range, so a two-hour recording opens as fast as a
// thumbnail does.
//
// With nothing open the folder has the whole window — pictures, clips and
// the folders inside it — and one click opens a file with a filmstrip under
// it. Escape, or the folder button, goes back.
//
// A phone records which way up a picture is, and Chromium turns it that way
// when it decodes the file (image-orientation: from-image). The details
// panel says which way it was shot; the viewer does not turn it a second
// time.

import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import {
  Ribbon, Group, Button, Separator, Icon, Spacer, Chip, Empty, Spinner, Panel, Content, Search,
  Select, useToast, useMenu, useCommands, menuItems, formatBytes, formatWhen, basename, dirname,
} from '@rutba/office-ui';
import { megapixels, aspectName } from '@rutba/imaging/probe';
import { timecode } from '@rutba/media/timeline';
import { AppFrame, useAppMenu, pickOpen, useFileDrop } from '../shell.js';
import {
  VIEWABLE, FAMILIES, SORTS, TILE_SIZES, kindOf, arrange, countByFamily, foldersOf,
  fitColumns, gridWindow, stripWindow, stripCentre, neighbours, fileUrl, thumbUrl,
} from './pictures/library.js';
import { frameOf } from './pictures/frames.js';

const INTERVALS = [
  { label: '2s', value: 2000 },
  { label: '3s', value: 3000 },
  { label: '5s', value: 5000 },
  { label: '10s', value: 10000 },
  { label: '30s', value: 30000 },
];

const GAP = 6;
const STRIP_ITEM = 74;
const STRIP_GAP = 5;

/** What is remembered about how the window is arranged, on this machine. */
const VIEW_KEY = 'pictures.view';
const VIEW_DEFAULTS = { sort: 'name', tileSize: 'm', grid: false, filmstrip: true, details: true };
const readView = () => {
  try {
    return { ...VIEW_DEFAULTS, ...JSON.parse(localStorage.getItem(VIEW_KEY) || '{}') };
  } catch {
    return { ...VIEW_DEFAULTS };
  }
};

export default function Pictures({ app, shell, boot }) {
  const toast = useToast();
  const [folder, setFolder] = useState(null);
  const [entries, setEntries] = useState([]);
  const [current, setCurrent] = useState(null);
  const [details, setDetails] = useState(null);
  const [query, setQuery] = useState('');
  const [family, setFamily] = useState('all');
  const [busy, setBusy] = useState(false);
  const [tab, setTab] = useState('home');

  // How the window is arranged, remembered on this machine.
  const [view, setViewState] = useState(readView);
  const setView = useCallback((patch) => {
    setViewState((v) => {
      const next = { ...v, ...patch };
      try { localStorage.setItem(VIEW_KEY, JSON.stringify(next)); } catch { /* no storage, no memory */ }
      return next;
    });
  }, []);
  const { sort, tileSize, grid: showGrid, filmstrip, details: showDetails } = view;

  // View state
  const [zoom, setZoom] = useState(0); // 0 = fit to window
  const [spin, setSpin] = useState(0); // view-only rotation, in quarter turns
  // With a picture open the ribbon folds to its tabs, so the picture gets the
  // room — the top especially. Expanding it is remembered on this machine.
  const [ribbonCollapsed, setRibbonCollapsed] = useState(() => {
    try { return localStorage.getItem('pictures.ribbon') !== 'open'; } catch { return true; }
  });
  const collapseRibbon = useCallback((next) => {
    setRibbonCollapsed(next);
    try { localStorage.setItem('pictures.ribbon', next ? 'collapsed' : 'open'); } catch { /* no storage, no memory */ }
  }, []);

  // The auto-navigator
  const [playing, setPlaying] = useState(false);
  const [interval, setIntervalMs] = useState(3000);
  const [loop, setLoop] = useState(true);
  const [shuffle, setShuffle] = useState(false);

  // Animated stills can be held still.
  const [frozen, setFrozen] = useState(false);
  // The file that would not decode. A toast says it once and fades; the
  // window then holds a frame with nothing in it, which reads as the
  // application having lost the picture. The stage says so instead, and keeps
  // saying it, and the corpus check can see it.
  const [failed, setFailed] = useState(null);
  const imgRef = useRef(null);
  const freezeRef = useRef(null);
  const mediaRef = useRef(null);

  const menu = useMenu();
  const openFileRef = useRef(null);
  const appMenu = useAppMenu({ shell, appKey: 'pictures', onOpen: () => openFileRef.current?.() });

  /* ── the folder ──────────────────────────────────────────────────────── */

  // A listing replaces the last one only when it has arrived: the grid that
  // is there stays there while the next folder is read, rather than
  // emptying to a spinner and filling again.
  const list = useCallback(
    async (dir) => {
      setBusy(true);
      try {
        const rows = await shell.fs.list({ path: dir, filter: VIEWABLE });
        setEntries(rows);
        setFolder(dir);
        setFamily('all');
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
        // A folder opens as the folder; a file opens with its folder behind it.
        const stat = await shell.fs.stat({ path: boot.file }).catch(() => null);
        if (stat?.dir) {
          await list(boot.file);
        } else {
          setCurrent(boot.file);
          await list(dirname(boot.file));
        }
      } else {
        const paths = await shell.app.paths();
        await list(paths.pictures);
      }
    };
    run();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const files = useMemo(() => arrange(entries, { query, family, sort }), [entries, query, family, sort]);
  const folders = useMemo(() => (query.trim() ? [] : foldersOf(entries)), [entries, query]);
  const counts = useMemo(() => countByFamily(entries), [entries]);

  const at = files.findIndex((f) => f.path === current);
  const kind = current ? kindOf(current) : null;
  const isAnimated = kind === 'maybe-animated' && details?.probe?.animated;
  const isMedia = kind === 'video' || kind === 'audio';
  const isStill = Boolean(current) && !isMedia && kind !== 'pdf';

  /* ── details, read without decoding the whole file ───────────────────── */

  useEffect(() => {
    if (!current) return undefined;
    let alive = true;
    setFrozen(false);
    setFailed(null);
    setDetails(null);
    (async () => {
      try {
        // The header only: the picture itself reaches the screen through the
        // file URL, and its size and EXIF live in the first quarter megabyte.
        //
        // Video and sound need more of it, and sometimes from the other end.
        // An MP4 keeps its duration in the moov atom, which sits at the END
        // of every file a camera or an ordinary encoder writes unless it was
        // asked to move it to the front. A quarter megabyte of the front
        // finds nothing there, and the panel showed no length at all.
        const media = kindOf(current) === 'video' || kindOf(current) === 'audio';
        const want = media ? 4 * 1024 * 1024 : 262144;
        const { bytes: head, stat } = await shell.fs.readHead({ path: current, bytes: want });

        const { probeImage } = await import('@rutba/imaging/probe');
        const { readExif, describeExif, orientationOf } = await import('@rutba/imaging/exif');
        const { probeMedia } = await import('@rutba/media/probe');
        let probe = probeImage(head) || probeMedia(head);
        if (media && !probe?.duration && stat.size > want) {
          const { bytes: tail } = await shell.fs.readHead({ path: current, bytes: want, offset: stat.size - want });
          probe = probeMedia(tail) || probe;
        }
        const exif = readExif(head);
        if (alive) setDetails({ stat, probe, exif, rows: describeExif(exif), orientation: orientationOf(exif) });
      } catch {
        if (alive) setDetails({ stat: null, probe: null, exif: {}, rows: [], orientation: null });
      }
    })();
    return () => {
      alive = false;
    };
  }, [current, shell]);

  /* ── the stage: the picture stays until the next one has decoded ─────── */

  // `shown` is the still whose whole picture is decoded and on the stage. It
  // lags `current` by however long the next file takes to decode, and the
  // stage draws `shown` all that while — so moving through a folder never
  // paints an empty frame between two pictures, which was the flicker.
  const [shown, setShown] = useState(null);
  const [natural, setNatural] = useState(null);
  const held = useRef(new Map()); // path -> Image, the neighbours fetched ahead
  useEffect(() => {
    if (!current || !isStill) {
      setShown(current);
      setNatural(null);
      return undefined;
    }
    let alive = true;
    const img = held.current.get(current) || new Image();
    if (!img.src) img.src = fileUrl(current);
    const arrive = () => {
      if (!alive) return;
      setNatural({ w: img.naturalWidth, h: img.naturalHeight });
      setShown(current);
    };
    const refuse = () => {
      if (!alive) return;
      setFailed(current);
      setShown(current);
      toast('This picture could not be decoded.', { tone: 'bad' });
    };
    img.decode().then(arrive, () => {
      // decode() can refuse a picture Chromium will still paint (one too
      // large for its decode cache); the element's own load says which.
      if (img.naturalWidth > 0) arrive();
      else if (img.complete) refuse();
      else {
        img.onload = arrive;
        img.onerror = refuse;
      }
    });
    return () => {
      alive = false;
    };
  }, [current, isStill, toast]);

  // The next pictures either way are fetched and decoded before they are
  // asked for, so Next is instant even the first time through a folder.
  useEffect(() => {
    if (!current || shown !== current || at < 0) return;
    const keep = new Map();
    for (const i of neighbours(at, files.length, { radius: 2, loop })) {
      const f = files[i];
      if (!f || kindOf(f.path) === 'video' || kindOf(f.path) === 'audio' || kindOf(f.path) === 'pdf') continue;
      let img = held.current.get(f.path);
      if (!img) {
        img = new Image();
        img.decoding = 'async';
        img.src = fileUrl(f.path);
        img.decode().catch(() => {});
      }
      keep.set(f.path, img);
    }
    held.current = keep;
  }, [current, shown, at, files, loop]);

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

  const back = useCallback(() => {
    setPlaying(false);
    setCurrent(null);
    setShown(null);
    setDetails(null);
    setZoom(0);
    setSpin(0);
  }, []);

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
      setCurrent(null);
      await list(dirs[0]);
    }
  }, [shell, list]);

  const openFile = useCallback(async () => {
    const file = await pickOpen(shell, 'pictures');
    if (!file) return;
    setCurrent(file);
    await list(dirname(file));
  }, [shell, list]);
  openFileRef.current = openFile;

  const enterFolder = useCallback(
    async (dir) => {
      setCurrent(null);
      setQuery('');
      await list(dir);
    },
    [list]
  );

  useFileDrop(
    useCallback(
      async (dropped) => {
        const first = dropped[0];
        if (!first) return;
        const stat = await shell.fs.stat({ path: first });
        if (stat?.dir) await enterFolder(first);
        else {
          setCurrent(first);
          await list(dirname(first));
        }
      },
      [shell, list, enterFolder]
    )
  );

  // The folder above this one. A drive's root is "C:/", never "C:", which
  // Node reads as that drive's working directory.
  const parent = folder ? dirname(folder).replace(/^([A-Za-z]:)$/, '$1/') : '';
  const hasParent = Boolean(parent) && parent.replace(/\/$/, '') !== String(folder).replace(/[\\/]+$/, '').replace(/\\/g, '/');

  /* ── commands ────────────────────────────────────────────────────────── */

  const commands = useMemo(
    () => ({
      'file.open': { label: 'Open…', icon: 'open', key: 'Mod+O', run: openFile },
      'file.folder': { label: 'Open folder…', icon: 'folderOpen', run: openFolder },
      'nav.back': { label: 'Back to the folder', icon: 'grid', key: 'Escape', run: () => current && back() },
      'nav.up': { label: 'Up a folder', icon: 'chevronUp', key: 'Backspace', run: () => hasParent && !current && enterFolder(parent) },
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
      'view.details': { label: 'Details panel', icon: 'info', run: () => setView({ details: !showDetails }) },
      'view.grid': { label: 'Folder grid beside the picture', icon: 'grid', run: () => setView({ grid: !showGrid }) },
      'view.filmstrip': { label: 'Filmstrip', icon: 'list', run: () => setView({ filmstrip: !filmstrip }) },
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
          // The next file takes its place, so a folder can be weeded
          // without going back to the grid for each one.
          const next = files[at + 1] || files[at - 1] || null;
          setCurrent(next ? next.path : null);
          list(dir);
        },
      },
    }),
    [openFile, openFolder, step, goTo, back, enterFolder, parent, hasParent, files, at, playing, frozen, toggleAnimation, current, folder, list, shell, toast, setView, showDetails, showGrid, filmstrip]
  );

  useCommands(commands, [files, at, current, folder, playing, frozen]);

  const rotation = spin * 90;
  const openAt = useCallback((f) => goTo(files.indexOf(f)), [goTo, files]);
  const tileMenu = useCallback((e) => menu.open(e, menuItems(commands, ['file.edit', 'file.reveal', 'file.copy', '-', 'file.trash'])), [menu, commands]);

  const tools = (
    <Tools
      query={query}
      onQuery={setQuery}
      family={family}
      onFamily={setFamily}
      counts={counts}
      sort={sort}
      onSort={(s) => setView({ sort: s })}
      tileSize={tileSize}
      onTileSize={(s) => setView({ tileSize: s })}
      stacked={Boolean(current)}
    />
  );

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
          collapsed={Boolean(current) && ribbonCollapsed}
          onCollapse={collapseRibbon}
          quick={
            <>
              {current ? <Button icon="grid" title="Back to the folder (Esc)" onClick={back} /> : null}
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
                <Button tall icon="chevronUp" label="Up" disabled={!hasParent} onClick={() => enterFolder(parent)} />
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
                <Button tall icon="grid" label="Folder" disabled={!current} onClick={back} />
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
              <Group label="Tiles">
                {TILE_SIZES.map((s) => (
                  <Button key={s.id} label={s.label} pressed={tileSize === s.id} onClick={() => setView({ tileSize: s.id })} />
                ))}
              </Group>
              <Group label="Panels">
                <Button icon="info" label="Details" pressed={showDetails} onClick={() => setView({ details: !showDetails })} />
                <Button icon="grid" label="Grid" pressed={showGrid} title="The folder's grid beside an open picture" onClick={() => setView({ grid: !showGrid })} />
                <Button icon="list" label="Filmstrip" pressed={filmstrip} onClick={() => setView({ filmstrip: !filmstrip })} />
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
          {current && details?.probe?.duration ? <Chip>{timecode(details.probe.duration)}</Chip> : null}
          {current && details?.probe?.width ? (
            <>
              <Chip>{details.probe.width} × {details.probe.height}</Chip>
              <Chip>{megapixels(details.probe.width, details.probe.height)} MP</Chip>
              <Chip>{aspectName(details.probe.width, details.probe.height)}</Chip>
            </>
          ) : null}
          {isAnimated ? <Chip title="This picture is animated">Animated</Chip> : null}
          {current ? zoom ? <Chip>{Math.round(zoom * 100)}%</Chip> : <Chip>Fit</Chip> : null}
          {current && details?.stat ? <Chip>{formatBytes(details.stat.size)}</Chip> : null}
          <Chip>{current && files.length ? `${at + 1} of ${files.length}` : files.length ? `${files.length} ${files.length === 1 ? 'file' : 'files'}${folders.length ? `, ${folders.length} ${folders.length === 1 ? 'folder' : 'folders'}` : ''}` : 'nothing here'}</Chip>
        </>
      }
    >
      {current && showGrid ? (
        <Panel width={276} resizable>
          <div className="pv-library">
            {tools}
            <Grid files={files} folders={[]} current={current} tileSize={tileSize} busy={busy} onOpen={openAt} onMenu={tileMenu} shell={shell} />
          </div>
        </Panel>
      ) : null}

      <Content>
        {!current ? (
          /* The folder, with the whole window. */
          <div className="pv-browse">
            <div className="pv-browse-head">
              <Button icon="chevronUp" title="Up a folder (Backspace)" disabled={!hasParent} onClick={() => enterFolder(parent)} />
              <span className="pv-browse-name" title={folder || ''}>{folder ? basename(folder) || folder : 'Pictures'}</span>
              {tools}
            </div>
            {!files.length && !folders.length && !busy ? (
              <Empty icon="pictures" title={entries.length ? 'Nothing matches' : 'Nothing here yet'}>
                {entries.length
                  ? 'No file in this folder matches the filter. Clear it, or choose another kind.'
                  : 'Open a folder, or drop pictures, video or animations onto this window.'}
              </Empty>
            ) : (
              <Grid files={files} folders={folders} current={null} tileSize={tileSize} browse busy={busy} onOpen={openAt} onFolder={enterFolder} onMenu={tileMenu} shell={shell} />
            )}
          </div>
        ) : (
          <Stage
            current={current}
            shown={shown}
            kind={kind}
            failed={failed}
            playing={playing}
            zoom={zoom}
            setZoom={setZoom}
            rotation={rotation}
            natural={natural}
            frozen={frozen}
            imgRef={imgRef}
            freezeRef={freezeRef}
            mediaRef={mediaRef}
            onEnded={() => playing && step(1)}
            onMediaError={() => { setFailed(current); toast('This file could not be played.', { tone: 'bad' }); }}
            onMenu={(e) => menu.open(e, menuItems(commands, ['nav.prev', 'nav.next', 'nav.back', '-', 'view.fit', 'view.actual', 'view.rotate', '-', 'file.edit', 'file.reveal', '-', 'file.trash']))}
          />
        )}

        {/* The navigator: where you are, and every way to move. */}
        {current && files.length ? (
          <div className="pv-nav">
            <Button icon="grid" title="Back to the folder (Esc)" onClick={back} />
            <Separator />
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

        {current && filmstrip && files.length ? <Strip files={files} at={at} onGo={goTo} shell={shell} /> : null}
      </Content>

      {current && showDetails ? (
        <Panel right width={230} title="Details" resizable>
          <div className="pv-details">
            <div className="pv-detail-name">{basename(current)}</div>
            {details ? (
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
            ) : (
              <div className="pv-detail-wait">Reading the file…</div>
            )}
            {!isMedia && kind !== 'pdf' ? <Button icon="wand" label="Edit this picture" onClick={() => commands['file.edit'].run()} /> : null}
          </div>
        </Panel>
      ) : null}

      {menu.node}
      <style>{CSS}</style>
    </AppFrame>
  );
}

/* ── the tools: filter, kind, order, tile size ───────────────────────────── */

function Tools({ query, onQuery, family, onFamily, counts, sort, onSort, tileSize, onTileSize, stacked }) {
  return (
    <div className={`pv-tools${stacked ? ' stacked' : ''}`}>
      <Search value={query} onChange={onQuery} placeholder="Filter by name" style={{ minWidth: 0 }} />
      <div className="pv-kinds" role="group" aria-label="Kind">
        {FAMILIES.filter((f) => f.id === 'all' || counts[f.id] > 0).map((f) => (
          <button
            key={f.id}
            type="button"
            className="pv-kind"
            data-family={f.id}
            aria-pressed={family === f.id}
            onClick={() => onFamily(f.id)}
          >
            {f.label}
            <span className="pv-kind-n">{counts[f.id] ?? 0}</span>
          </button>
        ))}
      </div>
      <div className="pv-order">
        <Select data-role="sort" value={sort} onChange={(e) => onSort(e.target.value)} style={{ padding: '3px 8px', fontSize: 11.5 }} title="The order of the folder" aria-label="Sort by">
          {SORTS.map((s) => (
            <option key={s.id} value={s.id}>{s.label}</option>
          ))}
        </Select>
        <span className="pv-sizes" role="group" aria-label="Tile size">
          {TILE_SIZES.map((s) => (
            <button key={s.id} type="button" className="pv-size" data-size={s.id} aria-pressed={tileSize === s.id} data-tip={`${s.label} tiles`} onClick={() => onTileSize(s.id)}>
              <i />
            </button>
          ))}
        </span>
      </div>
    </div>
  );
}

/* ── the grid: only the rows in view are drawn ───────────────────────────── */

/** A scroller's size and position, kept as state so the window can be computed. */
function useScrollBox(ref) {
  const [box, setBox] = useState({ width: 0, height: 0, scrollTop: 0, scrollLeft: 0 });
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return undefined;
    const measure = () => setBox((b) => {
      const next = { width: el.clientWidth, height: el.clientHeight, scrollTop: el.scrollTop, scrollLeft: el.scrollLeft };
      return b.width === next.width && b.height === next.height && b.scrollTop === next.scrollTop && b.scrollLeft === next.scrollLeft ? b : next;
    });
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    let raf = 0;
    const onScroll = () => {
      if (raf) return;
      raf = requestAnimationFrame(() => {
        raf = 0;
        measure();
      });
    };
    el.addEventListener('scroll', onScroll, { passive: true });
    return () => {
      ro.disconnect();
      el.removeEventListener('scroll', onScroll);
      if (raf) cancelAnimationFrame(raf);
    };
  }, [ref]);
  return box;
}

function Grid({ files, folders, current, tileSize, browse = false, busy, onOpen, onFolder, onMenu, shell }) {
  const ref = useRef(null);
  const box = useScrollBox(ref);
  const padding = browse ? 12 : 8;
  const want = TILE_SIZES.find((s) => s.id === tileSize) || TILE_SIZES[1];
  const { columns, tile } = fitColumns({ width: box.width || 300, tile: browse ? want.browsePx : want.px, gap: GAP, padding });
  const count = folders.length + files.length;
  const w = gridWindow({ count, width: box.width || 300, tile, gap: GAP, scrollTop: box.scrollTop, height: box.height, overscan: 2, padding, columns });

  // The open file's tile is kept in view as the selection moves — brought to
  // the nearest edge, never re-centred, so a click on a visible tile does
  // not shift the grid under the pointer.
  const at = current ? files.findIndex((f) => f.path === current) : -1;
  useEffect(() => {
    const el = ref.current;
    if (!el || at < 0) return;
    const row = Math.floor((folders.length + at) / columns);
    const top = padding + row * w.rowHeight;
    const bottom = top + tile;
    if (top < el.scrollTop) el.scrollTo({ top: top - padding, behavior: 'smooth' });
    else if (bottom > el.scrollTop + el.clientHeight) el.scrollTo({ top: bottom - el.clientHeight + padding, behavior: 'smooth' });
  }, [at, columns, w.rowHeight, tile, padding, folders.length]);

  const items = [];
  for (let i = w.first; i < w.last; i++) {
    if (i < folders.length) {
      const d = folders[i];
      items.push(
        <button key={d.path} type="button" className="pv-tile folder" data-kind="folder" style={{ width: tile, height: tile }} title={d.name} onClick={() => onFolder?.(d.path)}>
          <span className="pv-tile-icon"><Icon name="folder" size={Math.max(22, Math.round(tile * 0.28))} /></span>
          <span className="pv-tile-name">{d.name}</span>
        </button>
      );
    } else {
      const f = files[i - folders.length];
      items.push(<Tile key={f.path} file={f} active={f.path === current} size={tile} onOpen={() => onOpen(f)} onMenu={onMenu} shell={shell} />);
    }
  }

  return (
    <div className={`pv-grid-scroll${browse ? ' browse' : ''}`} ref={ref} data-drawn={items.length} data-count={count}>
      {busy ? <div className="pv-busy" /> : null}
      <div style={{ height: w.top }} />
      <div className="pv-grid" style={{ gridTemplateColumns: `repeat(${columns}, ${tile}px)`, gap: GAP, padding: `0 ${padding}px` }}>
        {items}
      </div>
      <div style={{ height: w.bottom + padding }} />
    </div>
  );
}

/**
 * The source that arrived for each path in this window — the platform's
 * thumbnail, or the frame the window drew — so a tile drawn again starts
 * from what worked, without a fade and without asking the platform again.
 */
const arrived = new Map();
const ICON = { video: 'video', audio: 'volume', pdf: 'pdf', still: 'picture', 'maybe-animated': 'picture' };

/**
 * The picture inside a tile: the platform's thumbnail, or a frame of a clip
 * this window drew when the platform had none, or the kind's icon.
 */
function Thumb({ file, shell }) {
  const kind = kindOf(file.path);
  const url = thumbUrl(file.path);
  const [src, setSrc] = useState(() => arrived.get(file.path) || url);
  const [state, setState] = useState(() => (arrived.has(file.path) ? 'loaded' : 'loading'));
  const ref = useRef(null);
  useLayoutEffect(() => {
    // Back in view: the browser still has it, and there is nothing to fade in.
    if (ref.current?.complete && ref.current.naturalWidth > 0) setState('loaded');
  }, []);
  const onLoad = () => {
    arrived.set(file.path, src);
    setState('loaded');
  };
  const onError = () => {
    if (kind === 'video' && src === url) {
      frameOf(file.path, (bytes) => shell.thumbs.put({ path: file.path, bytes })).then((frame) => {
        if (frame) setSrc(frame);
        else setState('none');
      });
    } else setState('none');
  };
  if (state === 'none') {
    return <span className="pv-tile-icon"><Icon name={ICON[kind] || 'file'} size={22} /></span>;
  }
  return <img ref={ref} className={`pv-thumb${state === 'loaded' ? ' loaded' : ''}`} src={src} alt="" decoding="async" draggable={false} onLoad={onLoad} onError={onError} />;
}

const Tile = React.memo(function Tile({ file, active, size, onOpen, onMenu, shell }) {
  const kind = kindOf(file.path);
  return (
    <button type="button" className={`pv-tile${active ? ' active' : ''}`} data-kind={kind} data-name={file.name} style={{ width: size, height: size }} onClick={onOpen} onContextMenu={onMenu} title={file.name}>
      <Thumb file={file} shell={shell} />
      {kind === 'video' || kind === 'audio' ? (
        <span className="pv-tile-badge"><Icon name={kind === 'audio' ? 'volume' : 'play'} size={11} /></span>
      ) : kind === 'pdf' ? (
        <span className="pv-tile-badge text">PDF</span>
      ) : null}
      <span className="pv-tile-name">{file.name}</span>
    </button>
  );
});

/* ── the filmstrip: the same, sideways ───────────────────────────────────── */

function Strip({ files, at, onGo, shell }) {
  const ref = useRef(null);
  const box = useScrollBox(ref);
  const w = stripWindow({ count: files.length, itemWidth: STRIP_ITEM, gap: STRIP_GAP, scrollLeft: box.scrollLeft, width: box.width, overscan: 8 });

  // The open file's thumbnail is kept in the middle as the selection moves.
  useEffect(() => {
    const el = ref.current;
    if (!el || at < 0 || !box.width) return;
    el.scrollTo({ left: stripCentre({ index: at, itemWidth: STRIP_ITEM, gap: STRIP_GAP, width: box.width }), behavior: 'smooth' });
  }, [at, box.width]);

  const items = [];
  for (let i = w.first; i < w.last; i++) {
    const f = files[i];
    items.push(
      <button key={f.path} type="button" className={`pv-strip-item${i === at ? ' active' : ''}`} onClick={() => onGo(i)} title={f.name}>
        <Thumb file={f} shell={shell} />
      </button>
    );
  }
  return (
    <div className="pv-strip" ref={ref} data-drawn={items.length} data-count={files.length}>
      <div style={{ width: w.left, flex: 'none' }} />
      {items}
      <div style={{ width: w.right, flex: 'none' }} />
    </div>
  );
}

/* ── the stage ───────────────────────────────────────────────────────────── */

function Stage({ current, shown, kind, failed, playing, zoom, setZoom, rotation, natural, frozen, imgRef, freezeRef, mediaRef, onEnded, onMediaError, onMenu }) {
  const ref = useRef(null);
  const isMedia = kind === 'video' || kind === 'audio';
  // A zoomed picture is dragged about with the mouse, as well as scrolled.
  const drag = useRef(null);
  const onMouseDown = (e) => {
    if (!zoom || e.button !== 0) return;
    drag.current = { x: e.clientX, y: e.clientY, left: ref.current.scrollLeft, top: ref.current.scrollTop };
    e.preventDefault();
  };
  const onMouseMove = (e) => {
    const d = drag.current;
    if (!d) return;
    ref.current.scrollLeft = d.left - (e.clientX - d.x);
    ref.current.scrollTop = d.top - (e.clientY - d.y);
  };
  const stopDrag = () => { drag.current = null; };
  const waiting = current && !isMedia && kind !== 'pdf' && shown !== current && failed !== current;

  return (
    <div
      ref={ref}
      className={`pv-stage viewing${zoom ? ' zoomed' : ''}`}
      onWheel={(e) => {
        if (!e.ctrlKey) return;
        e.preventDefault();
        setZoom((z) => Math.max(0.05, Math.min(12, (z || 1) * (e.deltaY < 0 ? 1.1 : 0.9))));
      }}
      onDoubleClick={() => setZoom((z) => (z ? 0 : 1))}
      onMouseDown={onMouseDown}
      onMouseMove={onMouseMove}
      onMouseUp={stopDrag}
      onMouseLeave={stopDrag}
      onContextMenu={onMenu}
    >
      {waiting ? <div className="pv-loading" /> : null}
      {failed === current ? (
        <Empty icon="pictures" title="This file could not be opened">
          {isMedia
            ? `${basename(current)} is not a video or sound file this machine can play. It may be in a format with no decoder installed, or it may not be media at all despite its name.`
            : `${basename(current)} could not be decoded as a picture. It may be a download that stopped early, or a file with the wrong extension.`}
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
          onEnded={onEnded}
          onError={onMediaError}
        />
      ) : shown ? (
        <>
          <img
            ref={imgRef}
            className={`pv-image${zoom ? ' zoomed' : ''}${frozen ? ' hidden' : ''}`}
            src={fileUrl(shown)}
            alt={basename(shown)}
            style={{
              transform: `rotate(${rotation}deg)`,
              width: zoom && natural ? `${natural.w * zoom}px` : undefined,
              height: zoom ? 'auto' : undefined,
            }}
            draggable={false}
          />
          {/* The held frame of an animation, drawn once and shown in its place. */}
          <canvas ref={freezeRef} className={`pv-image${frozen ? '' : ' hidden'}`} style={{ transform: `rotate(${rotation}deg)` }} />
        </>
      ) : null}
    </div>
  );
}

const CSS = `
/* ── the folder ── */
.pv-browse { flex: 1; min-height: 0; display: flex; flex-direction: column; background: var(--window); }
.pv-browse-head {
  display: flex; align-items: center; gap: 10px; padding: 8px 12px 6px; flex-wrap: wrap;
  border-bottom: 1px solid var(--line-soft); background: var(--chrome);
}
.pv-browse-name { font-weight: 600; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 260px; }
.pv-browse-head .pv-tools { flex: 1; }
.pv-library { height: 100%; display: flex; flex-direction: column; min-height: 0; }
.pv-library .pv-tools { padding: 8px 10px 6px; }

.pv-tools { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; min-width: 0; }
.pv-tools .rw-search { flex: 1; min-width: 140px; max-width: 320px; }
.pv-tools.stacked { flex-direction: column; align-items: stretch; gap: 6px; }
.pv-tools.stacked .rw-search { max-width: none; }
.pv-kinds { display: flex; gap: 4px; flex-wrap: wrap; }
.pv-kind {
  border: 1px solid var(--line); background: var(--surface); color: var(--ink-2); font: inherit; font-size: 11.5px;
  padding: 2px 8px 2px 9px; border-radius: 999px; display: inline-flex; align-items: center; gap: 5px; cursor: pointer;
  transition: background var(--fast), color var(--fast), border-color var(--fast);
}
.pv-kind:hover { background: var(--hover); color: var(--ink); }
.pv-kind[aria-pressed='true'] { background: var(--selected); color: var(--accent); border-color: var(--accent-line); }
.pv-kind-n { font-size: 10.5px; color: var(--ink-3); font-variant-numeric: tabular-nums; }
.pv-kind[aria-pressed='true'] .pv-kind-n { color: var(--accent); opacity: 0.8; }
.pv-order { display: flex; align-items: center; gap: 6px; margin-left: auto; }
.pv-tools.stacked .pv-order { margin-left: 0; justify-content: space-between; }

.pv-sizes { display: inline-flex; gap: 1px; background: var(--surface-2); border: 1px solid var(--line); border-radius: var(--r-2); padding: 1px; }
.pv-size { border: 0; background: transparent; width: 24px; height: 20px; border-radius: 5px; display: grid; place-items: center; cursor: pointer; color: var(--ink-3); }
.pv-size i { display: block; background: currentColor; border-radius: 2px; }
.pv-size[data-size='s'] i { width: 7px; height: 7px; }
.pv-size[data-size='m'] i { width: 10px; height: 10px; }
.pv-size[data-size='l'] i { width: 13px; height: 13px; }
.pv-size:hover { background: var(--hover); color: var(--ink); }
.pv-size[aria-pressed='true'] { background: var(--surface); color: var(--accent); box-shadow: var(--shadow-1); }

/* ── the grid ── */
.pv-grid-scroll { flex: 1; min-height: 0; overflow-y: auto; overflow-x: hidden; position: relative; padding-top: 8px; scrollbar-width: thin; }
.pv-grid-scroll.browse { padding-top: 12px; }
.pv-grid { display: grid; justify-content: start; }
.pv-busy {
  position: sticky; top: 0; height: 2px; margin: -8px 0 6px; z-index: 2;
  background: linear-gradient(90deg, transparent, var(--accent), transparent); background-size: 50% 100%;
  animation: pv-busy 1.1s linear infinite;
}
@keyframes pv-busy { from { background-position: -50% 0; } to { background-position: 150% 0; } }
.pv-tile {
  border: 0; border-radius: var(--r-2); background: var(--sunken); padding: 0; cursor: pointer;
  overflow: hidden; display: flex; flex-direction: column; position: relative; box-sizing: border-box;
  box-shadow: inset 0 0 0 1px rgba(15, 20, 30, 0.06);
  transition: box-shadow var(--fast), transform var(--fast);
}
.pv-tile:hover { box-shadow: inset 0 0 0 1px var(--line-strong), var(--shadow-1); }
.pv-tile:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.pv-tile.active { box-shadow: inset 0 0 0 2px var(--accent), 0 0 0 2px var(--accent-soft); }
.pv-tile.folder { background: var(--surface-2); }
.pv-tile.folder .pv-tile-icon { color: var(--accent); opacity: 0.85; }
.pv-tile.folder .pv-tile-name { color: var(--ink); background: none; font-weight: 600; text-align: center; }
.pv-thumb {
  width: 100%; height: 100%; object-fit: cover; display: block; background: var(--sunken);
  opacity: 0; transition: opacity 180ms ease-out;
}
.pv-thumb.loaded { opacity: 1; }
.pv-tile-icon { flex: 1; display: grid; place-items: center; color: var(--ink-3); }
.pv-tile-badge {
  position: absolute; top: 5px; right: 5px; height: 18px; min-width: 18px; padding: 0 4px; border-radius: 9px;
  background: rgba(0, 0, 0, 0.58); color: #fff; display: grid; place-items: center; font-size: 9px; font-weight: 700; letter-spacing: 0.04em;
  backdrop-filter: blur(3px);
}
.pv-tile-name {
  position: absolute; left: 0; right: 0; bottom: 0; padding: 14px 6px 4px; font-size: 10.5px; line-height: 1.2;
  color: #fff; background: linear-gradient(transparent, rgba(0, 0, 0, 0.7)); text-align: left;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap; pointer-events: none;
}
.pv-grid-scroll.browse .pv-tile-name { font-size: 11.5px; padding: 18px 8px 6px; }

/* ── the stage ── */
/* One bounded track each way: an auto track would grow to the picture's own
   size and the stage would scroll by the few pixels it did not fit by. */
.pv-stage {
  flex: 1; min-height: 0; overflow: auto; display: grid; place-items: center;
  grid-template-columns: minmax(0, 1fr); grid-template-rows: minmax(0, 1fr);
  background: var(--sunken); padding: 16px; position: relative; scrollbar-width: thin;
}
.pv-stage > * { min-width: 0; min-height: 0; }
/* With a picture open, the stage keeps almost nothing back from it — and is
   dark, whatever the theme, because a photograph reads best against near-black. */
.pv-stage.viewing { padding: 4px; background: #15171b; }
.pv-stage.viewing .rw-empty { color: #c8ccd3; }
.pv-stage.zoomed { cursor: grab; }
.pv-stage.zoomed:active { cursor: grabbing; }
.pv-image { max-width: 100%; max-height: 100%; object-fit: contain; display: block; image-orientation: from-image; }
.pv-image.zoomed { max-width: none; max-height: none; }
.pv-image.hidden { display: none; }
.pv-loading {
  position: absolute; left: 0; right: 0; top: 0; height: 2px; z-index: 2; opacity: 0;
  background: linear-gradient(90deg, transparent, var(--accent), transparent); background-size: 50% 100%;
  animation: pv-busy 1.1s linear infinite, pv-appear 0.1s 0.25s forwards;
}
@keyframes pv-appear { to { opacity: 1; } }
.pv-video { max-width: 100%; max-height: 100%; display: block; background: #000; }
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

/* ── the filmstrip ── */
.pv-strip {
  display: flex; gap: ${STRIP_GAP}px; padding: 6px 10px; overflow-x: auto; overflow-y: hidden;
  background: var(--surface); border-top: 1px solid var(--line-soft); scrollbar-width: thin;
}
.pv-strip-item {
  flex: none; width: ${STRIP_ITEM}px; height: 54px; border: 0; border-radius: var(--r-2);
  overflow: hidden; padding: 0; background: var(--sunken); cursor: pointer; position: relative;
  box-shadow: inset 0 0 0 1px rgba(15, 20, 30, 0.06); transition: box-shadow var(--fast);
}
.pv-strip-item:hover { box-shadow: inset 0 0 0 1px var(--line-strong); }
.pv-strip-item.active { box-shadow: inset 0 0 0 2px var(--accent), 0 0 0 2px var(--accent-soft); }
.pv-strip-item .pv-tile-icon { height: 100%; }

/* ── the details ── */
.pv-details { padding: 4px 12px 14px; display: flex; flex-direction: column; gap: 12px; }
.pv-detail-name { font-weight: 600; font-size: 12.5px; word-break: break-all; }
.pv-detail-wait { font-size: 11.5px; color: var(--ink-3); }
.pv-details dl { display: grid; grid-template-columns: auto 1fr; gap: 4px 10px; margin: 0; font-size: 11.5px; }
.pv-details dt { color: var(--ink-3); }
.pv-details dd { margin: 0; word-break: break-word; }
`;
