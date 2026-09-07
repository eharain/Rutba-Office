// Rutba Pictures.
//
// A viewer whose whole job is to be instant. Two things make it so: the grid
// asks the backend for each file's dimensions rather than decoding it, and the
// images themselves are fetched by the window over rutba://file, so a 40 MP
// photograph is decoded by Chromium straight from disk and never passes through
// JavaScript as an array of bytes.
//
// It reads EXIF, which matters for one reason above the details panel: a phone
// records which way up the picture is, and a viewer that ignores that shows a
// third of everyone's holiday sideways.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Ribbon, Group, Button, Icon, Spacer, Chip, Empty, Spinner, Panel, Content, Search, useToast, useMenu, useCommands, menuItems, formatBytes, formatWhen, basename, dirname } from '@rutba/office-ui';
import { megapixels, aspectName } from '@rutba/imaging/probe';
import { AppFrame, useAppMenu, pickOpen, useFileDrop, openInApp } from '../shell.js';

const IMAGE_EXT = ['.jpg', '.jpeg', '.png', '.gif', '.webp', '.avif', '.bmp', '.tif', '.tiff', '.ico', '.svg', '.heic', '.pdf'];

const fileUrl = (p) => `rutba://file/${btoa(unescape(encodeURIComponent(p))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')}`;

export default function Pictures({ app, shell, boot }) {
  const toast = useToast();
  const [folder, setFolder] = useState(null);
  const [entries, setEntries] = useState([]);
  const [current, setCurrent] = useState(null);
  const [details, setDetails] = useState(null);
  const [query, setQuery] = useState('');
  const [zoom, setZoom] = useState(0); // 0 = fit
  const [busy, setBusy] = useState(false);
  const [showDetails, setShowDetails] = useState(true);
  const [tab, setTab] = useState('home');
  const [slideshow, setSlideshow] = useState(false);
  const menu = useMenu();
  const openFileRef = useRef(null);
  const appMenu = useAppMenu({ shell, appKey: 'pictures', onOpen: () => openFileRef.current?.() });

  const list = useCallback(
    async (dir) => {
      setBusy(true);
      try {
        const rows = await shell.fs.list({ path: dir, filter: IMAGE_EXT });
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

  // Details come from the backend because reading EXIF means reading bytes,
  // and the window should not be holding a 40 MB file to learn the shutter speed.
  useEffect(() => {
    if (!current) return;
    let alive = true;
    (async () => {
      try {
        const { bytes, stat } = await shell.fs.read({ path: current });
        const { probeImage } = await import('@rutba/imaging/probe');
        const { readExif, describeExif, orientationOf } = await import('@rutba/imaging/exif');
        const probe = probeImage(bytes.slice(0, 262144));
        const exif = readExif(bytes.slice(0, 262144));
        if (alive) setDetails({ stat, probe, exif, rows: describeExif(exif), orientation: orientationOf(exif) });
      } catch {
        if (alive) setDetails(null);
      }
    })();
    return () => {
      alive = false;
    };
  }, [current, shell]);

  const files = useMemo(() => {
    const rows = entries.filter((e) => !e.dir);
    if (!query.trim()) return rows;
    const q = query.toLowerCase();
    return rows.filter((r) => r.name.toLowerCase().includes(q));
  }, [entries, query]);

  const at = files.findIndex((f) => f.path === current);

  const step = useCallback(
    (delta) => {
      if (!files.length) return;
      const next = files[(at + delta + files.length) % files.length];
      if (next) {
        setCurrent(next.path);
        setZoom(0);
      }
    },
    [files, at]
  );

  useEffect(() => {
    if (!slideshow) return undefined;
    const timer = setInterval(() => step(1), 3500);
    return () => clearInterval(timer);
  }, [slideshow, step]);

  const openFolder = useCallback(async () => {
    const dirs = await shell.dialog.open({ title: 'Open a folder of pictures', directory: true });
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

  const commands = useMemo(
    () => ({
      'file.open': { label: 'Open…', icon: 'open', key: 'Mod+O', run: openFile },
      'file.folder': { label: 'Open folder…', icon: 'folderOpen', run: openFolder },
      'view.next': { label: 'Next', icon: 'chevronRight', key: 'arrowright', run: () => step(1) },
      'view.prev': { label: 'Previous', icon: 'chevronLeft', key: 'arrowleft', run: () => step(-1) },
      'view.fit': { label: 'Fit to window', icon: 'maximize', key: 'Mod+0', run: () => setZoom(0) },
      'view.in': { label: 'Zoom in', icon: 'zoomIn', key: 'Mod+Plus', run: () => setZoom((z) => Math.min(8, (z || 1) * 1.25)) },
      'view.out': { label: 'Zoom out', icon: 'zoomOut', key: 'Mod+-', run: () => setZoom((z) => Math.max(0.05, (z || 1) / 1.25)) },
      'view.details': { label: 'Details', icon: 'info', run: () => setShowDetails((s) => !s) },
      'view.slideshow': { label: 'Slideshow', icon: 'play', key: 'F5', run: () => setSlideshow((s) => !s) },
      'file.edit': { label: 'Edit this picture', icon: 'wand', run: () => current && shell.win.create({ app: 'image', file: current }) },
      'file.reveal': { label: 'Show in folder', icon: 'folderOpen', run: () => current && shell.shell.showInFolder({ path: current }) },
      'file.trash': {
        label: 'Move to trash',
        icon: 'trash',
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
    [openFile, openFolder, step, current, folder, list, shell, toast]
  );

  useCommands(commands, [files, at, current, folder]);

  const isPdf = current?.toLowerCase().endsWith('.pdf');
  const rotation = details?.orientation?.rotate || 0;
  const mirrored = details?.orientation?.flip;

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={current ? basename(current) : folder ? basename(folder) : 'Pictures'}
      menu={appMenu}
      ribbon={
        <Ribbon
          tabs={[{ id: 'home', label: 'Home' }, { id: 'view', label: 'View' }]}
          active={tab}
          onTab={setTab}
          quick={
            <>
              <Button icon="chevronLeft" title="Previous" onClick={() => step(-1)} disabled={!files.length} />
              <Button icon="chevronRight" title="Next" onClick={() => step(1)} disabled={!files.length} />
            </>
          }
        >
          {tab === 'home' ? (
            <>
              <Group label="Open">
                <Button tall icon="open" label="Picture" onClick={openFile} />
                <Button tall icon="folderOpen" label="Folder" onClick={openFolder} />
              </Group>
              <Group label="This picture">
                <Button tall icon="wand" label="Edit" disabled={!current} onClick={() => commands['file.edit'].run()} />
                <Button tall icon="folderOpen" label="Reveal" disabled={!current} onClick={() => commands['file.reveal'].run()} />
                <Button tall icon="trash" label="Trash" disabled={!current} onClick={() => commands['file.trash'].run()} />
              </Group>
              <Group label="Show">
                <Button tall icon="play" label={slideshow ? 'Stop' : 'Slideshow'} onClick={() => setSlideshow((s) => !s)} />
              </Group>
            </>
          ) : (
            <>
              <Group label="Zoom">
                <Button icon="zoomOut" label="Out" onClick={() => commands['view.out'].run()} />
                <Button icon="zoomIn" label="In" onClick={() => commands['view.in'].run()} />
                <Button icon="maximize" label="Fit" onClick={() => setZoom(0)} />
                <Button icon="check" label="100%" onClick={() => setZoom(1)} />
              </Group>
              <Group label="Panels">
                <Button icon="info" label="Details" pressed={showDetails} onClick={() => setShowDetails((s) => !s)} />
              </Group>
            </>
          )}
        </Ribbon>
      }
      status={
        <>
          <span>{folder || ''}</span>
          <Spacer />
          {details?.probe ? (
            <>
              <Chip>{details.probe.width} × {details.probe.height}</Chip>
              <Chip>{megapixels(details.probe.width, details.probe.height)} MP</Chip>
              <Chip>{aspectName(details.probe.width, details.probe.height)}</Chip>
            </>
          ) : null}
          {details?.stat ? <Chip>{formatBytes(details.stat.size)}</Chip> : null}
          <Chip>{files.length ? `${at + 1} of ${files.length}` : '0 pictures'}</Chip>
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
              <button
                key={f.path}
                type="button"
                className={`pv-tile${f.path === current ? ' active' : ''}`}
                onClick={() => {
                  setCurrent(f.path);
                  setZoom(0);
                }}
                onContextMenu={(e) => menu.open(e, menuItems(commands, ['file.edit', 'file.reveal', '-', 'file.trash']))}
                title={f.name}
              >
                {f.ext === '.pdf' ? (
                  <span className="pv-tile-pdf"><Icon name="pdf" size={22} /></span>
                ) : (
                  <img src={fileUrl(f.path)} alt="" loading="lazy" decoding="async" />
                )}
                <span className="pv-tile-name">{f.name}</span>
              </button>
            ))}
          </div>
        )}
      </Panel>

      <Content>
        {current ? (
          <div
            className="pv-stage"
            onWheel={(e) => {
              if (!e.ctrlKey) return;
              e.preventDefault();
              setZoom((z) => Math.max(0.05, Math.min(8, (z || 1) * (e.deltaY < 0 ? 1.1 : 0.9))));
            }}
            onDoubleClick={() => setZoom((z) => (z ? 0 : 1))}
            onContextMenu={(e) => menu.open(e, menuItems(commands, ['view.fit', 'view.in', 'view.out', '-', 'file.edit', 'file.reveal', '-', 'file.trash']))}
          >
            {isPdf ? (
              <embed src={fileUrl(current)} type="application/pdf" className="pv-pdf" />
            ) : (
              <img
                className={zoom ? 'pv-image zoomed' : 'pv-image'}
                src={fileUrl(current)}
                alt={basename(current)}
                style={{
                  transform: `rotate(${rotation}deg)${mirrored ? ' scaleX(-1)' : ''}`,
                  width: zoom ? `${(details?.probe?.width || 0) * zoom}px` : undefined,
                  height: zoom ? 'auto' : undefined,
                }}
                draggable={false}
              />
            )}
          </div>
        ) : (
          <Empty icon="pictures" title={files.length ? 'Choose a picture' : 'No pictures here'}>
            {files.length ? 'Pick one from the list to see it full size.' : 'Open a folder, or drop pictures onto this window.'}
          </Empty>
        )}
      </Content>

      {showDetails && details ? (
        <Panel right width={228} title="Details" resizable>
          <div className="pv-details">
            <div className="pv-detail-name">{basename(current)}</div>
            <dl>
              {details.probe ? (
                <>
                  <dt>Format</dt>
                  <dd>{details.probe.format.toUpperCase()}</dd>
                  <dt>Dimensions</dt>
                  <dd>{details.probe.width} × {details.probe.height}</dd>
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
            <Button icon="wand" label="Edit this picture" onClick={() => commands['file.edit'].run()} />
          </div>
        </Panel>
      ) : null}

      {menu.node}
      <style>{CSS}</style>
    </AppFrame>
  );
}

const CSS = `
.pv-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(96px, 1fr)); gap: 6px; padding: 4px 8px 12px; }
.pv-tile {
  border: 2px solid transparent; border-radius: var(--r-2); background: var(--sunken); padding: 0;
  overflow: hidden; display: flex; flex-direction: column; aspect-ratio: 1; position: relative;
}
.pv-tile img { width: 100%; height: 100%; object-fit: cover; display: block; background: var(--sunken); }
.pv-tile-pdf { flex: 1; display: grid; place-items: center; color: var(--ink-3); }
.pv-tile-name {
  position: absolute; left: 0; right: 0; bottom: 0; padding: 3px 5px; font-size: 10px;
  color: #fff; background: linear-gradient(transparent, rgba(0,0,0,0.72)); text-align: left;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.pv-tile:hover { border-color: var(--line-strong); }
.pv-tile.active { border-color: var(--accent); }

.pv-stage {
  flex: 1; min-height: 0; overflow: auto; display: grid; place-items: center;
  background: var(--sunken); padding: 16px;
}
.pv-image { max-width: 100%; max-height: 100%; object-fit: contain; display: block; box-shadow: var(--shadow-2); }
.pv-image.zoomed { max-width: none; max-height: none; }
.pv-pdf { width: 100%; height: 100%; border: 0; }

.pv-details { padding: 4px 12px 14px; display: flex; flex-direction: column; gap: 12px; }
.pv-detail-name { font-weight: 600; font-size: 12.5px; word-break: break-all; }
.pv-details dl { display: grid; grid-template-columns: auto 1fr; gap: 4px 10px; margin: 0; font-size: 11.5px; }
.pv-details dt { color: var(--ink-3); }
.pv-details dd { margin: 0; word-break: break-word; }
`;
