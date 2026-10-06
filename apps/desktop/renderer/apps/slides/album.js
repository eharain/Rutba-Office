// Insert → Photo Album: the pictures, in the order they will appear, how
// many go on a slide, whether each is captioned with its name, and the
// album's title. Create makes a new presentation and opens it in a window
// of its own, as PowerPoint's New Photo Album does.

import React, { useState } from 'react';
import { Button, Dialog, Field } from '@rutba/office-ui';

const baseName = (file) => String(file).split(/[\\/]/).pop();

export function PhotoAlbumDialog({ shell, onCreate, onClose }) {
  const [files, setFiles] = useState([]);
  const [picked, setPicked] = useState(0);
  const [perSlide, setPerSlide] = useState(1);
  const [captions, setCaptions] = useState(false);
  const [title, setTitle] = useState('Photo Album');
  const [busy, setBusy] = useState(false);

  const add = async () => {
    const more = await shell.dialog.open({
      title: 'Insert New Pictures',
      multiple: true,
      filters: [{ name: 'Pictures', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp'] }],
    });
    if (more?.length) setFiles((list) => [...list, ...more.filter((f) => !list.includes(f))]);
  };
  const move = (delta) => setFiles((list) => {
    const to = picked + delta;
    if (picked < 0 || to < 0 || to >= list.length) return list;
    const next = [...list];
    [next[picked], next[to]] = [next[to], next[picked]];
    setPicked(to);
    return next;
  });
  const remove = () => setFiles((list) => {
    const next = list.filter((_, i) => i !== picked);
    setPicked(Math.min(picked, next.length - 1));
    return next;
  });
  const create = async () => {
    if (!files.length || busy) return;
    setBusy(true);
    try {
      await onCreate({ files, perSlide, captions, title: title.trim() || 'Photo Album' });
    } finally {
      setBusy(false);
    }
  };

  return (
    <Dialog
      title="Photo Album"
      width={500}
      onClose={onClose}
      actions={<><Button label="Cancel" onClick={onClose} /><Button primary label={busy ? 'Creating…' : 'Create'} className="sl-album-create" disabled={!files.length || busy} onClick={create} /></>}
    >
      <div className="sl-album">
        <div className="sl-album-head">
          <span>Pictures in album</span>
          <Button icon="picture" label="File/Disk…" className="sl-album-add" title="Add pictures from a folder" onClick={add} />
        </div>
        <div className="sl-album-list" role="listbox" aria-label="Pictures in album">
          {files.length ? files.map((f, i) => (
            <button key={f} type="button" role="option" aria-selected={i === picked} className={`sl-album-row${i === picked ? ' on' : ''}`} title={f} onClick={() => setPicked(i)}>
              <span className="sl-album-n">{i + 1}</span>{baseName(f)}
            </button>
          )) : <div className="sl-album-empty">No pictures yet. File/Disk… adds them.</div>}
        </div>
        <div className="sl-album-tools">
          <Button icon="chevronUp" title="Move the picture earlier" disabled={picked <= 0} onClick={() => move(-1)} />
          <Button icon="chevronDown" title="Move the picture later" disabled={picked < 0 || picked >= files.length - 1} onClick={() => move(1)} />
          <Button icon="trash" label="Remove" disabled={!files.length} onClick={remove} />
          <span className="sl-album-count">{files.length} {files.length === 1 ? 'picture' : 'pictures'}</span>
        </div>
        <Field label="Picture layout">
          <select className="rw-input sl-album-layout" value={perSlide} onChange={(e) => setPerSlide(Number(e.target.value))}>
            <option value={1}>1 picture</option>
            <option value={2}>2 pictures</option>
            <option value={4}>4 pictures</option>
          </select>
        </Field>
        <label className="sl-album-check">
          <input type="checkbox" className="sl-album-captions" checked={captions} onChange={(e) => setCaptions(e.target.checked)} />
          Captions below all pictures
        </label>
        <Field label="Title">
          <input className="rw-input sl-album-title" value={title} onChange={(e) => setTitle(e.target.value)} />
        </Field>
      </div>
    </Dialog>
  );
}

export const ALBUM_CSS = `
.sl-album { display: flex; flex-direction: column; gap: 8px; font-size: 12.5px; }
.sl-album-head { display: flex; align-items: center; justify-content: space-between; color: var(--ink-2); }
.sl-album-head .rw-btn { border: 1px solid var(--line); background: var(--surface); }
.sl-album-list { height: 170px; overflow: auto; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); }
.sl-album-row { display: flex; align-items: center; gap: 8px; width: 100%; border: 0; border-bottom: 1px solid var(--line-soft); background: none; font: inherit; color: var(--ink); text-align: left; padding: 4px 8px; cursor: pointer; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sl-album-row:hover { background: var(--hover); }
.sl-album-row.on { background: var(--selected); color: var(--accent); }
.sl-album-n { width: 20px; color: var(--ink-3); font-variant-numeric: tabular-nums; flex: none; }
.sl-album-empty { padding: 10px; color: var(--ink-3); }
.sl-album-tools { display: flex; align-items: center; gap: 4px; }
.sl-album-count { margin-left: auto; color: var(--ink-3); }
.sl-album-check { display: flex; align-items: center; gap: 6px; }
`;
