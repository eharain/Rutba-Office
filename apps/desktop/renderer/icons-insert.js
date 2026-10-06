// Insert → Icons, shared by Documents, Worksheets and Presentations: the
// suite's own icons, drawn here rather than fetched from a library, in the
// colour chosen, put in as a picture. The set is the one the ribbon draws
// with — one grid, one stroke — so the icons look like a set on the page too.

import React, { useMemo, useState } from 'react';
import { Button, Dialog, Icon } from '@rutba/office-ui';
import { iconSvg } from '@rutba/office-ui/icons';

export const ICON_GROUPS = [
  ['Communication', ['mail', 'send', 'reply', 'forward', 'inbox', 'attach', 'contacts', 'calendar', 'globe', 'link']],
  ['Office', ['file', 'folder', 'folderOpen', 'word', 'sheets', 'slides', 'print', 'pdf', 'save', 'archive', 'table', 'chart', 'formula', 'sum', 'clock']],
  ['Media', ['pictures', 'image', 'video', 'play', 'pause', 'stop', 'volume', 'scissors', 'crop', 'rotate', 'zoomIn', 'zoomOut']],
  ['Signs and symbols', ['check', 'close', 'plus', 'minus', 'info', 'star', 'heart', 'flag', 'lock', 'shield', 'eye', 'sun', 'moon', 'home', 'settings', 'wand', 'refresh', 'download', 'export', 'import', 'trash', 'find', 'mouse']],
];

const COLOURS = [['#000000', 'Black'], ['#1F4E79', 'Dark blue'], ['#2B5FD9', 'Blue'], ['#00796B', 'Teal'], ['#2E7D32', 'Green'], ['#C62828', 'Red'], ['#EF6C00', 'Orange'], ['#6A1B9A', 'Purple'], ['#7F7F7F', 'Grey']];

/** A readable name from an icon's key: folderOpen → "folder open". */
const words = (name) => name.replace(/([A-Z])/g, ' $1').toLowerCase();

/** An icon drawn to a square PNG, `px` on a side, in `colour`. */
export async function iconPng(name, colour, px = 512) {
  const svg = iconSvg(name, { colour, size: px });
  const img = new Image();
  img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  await img.decode();
  const canvas = document.createElement('canvas');
  canvas.width = px;
  canvas.height = px;
  canvas.getContext('2d').drawImage(img, 0, 0, px, px);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
  return new Uint8Array(await blob.arrayBuffer());
}

/**
 * @param {{ onInsert, onClose }} props
 *   onInsert({ name, bytes, contentType }) — the picture, drawn
 */
export function IconsDialog({ onInsert, onClose }) {
  const [find, setFind] = useState('');
  const [picked, setPicked] = useState(null);
  const [colour, setColour] = useState('#2B5FD9');
  const [busy, setBusy] = useState(false);
  const groups = useMemo(() => {
    const f = find.trim().toLowerCase();
    return ICON_GROUPS.map(([title, names]) => [title, f ? names.filter((n) => words(n).includes(f) || title.toLowerCase().includes(f)) : names]).filter(([, names]) => names.length);
  }, [find]);
  const insert = async (name = picked) => {
    if (!name || busy) return;
    setBusy(true);
    try {
      await onInsert({ name, bytes: await iconPng(name, colour), contentType: 'image/png' });
    } finally {
      setBusy(false);
    }
  };

  return (
    <Dialog
      title="Insert Icons"
      width={560}
      onClose={onClose}
      actions={<><Button label="Cancel" onClick={onClose} /><Button primary label="Insert" className="ic-insert" disabled={!picked || busy} onClick={() => insert()} /></>}
    >
      <style>{ICONS_CSS}</style>
      <div className="ic">
        <input className="rw-input ic-find" placeholder="Find an icon" value={find} spellCheck={false} autoFocus onChange={(e) => setFind(e.target.value)} />
        <div className="ic-scroll">
          {groups.length ? groups.map(([title, names]) => (
            <section key={title}>
              <div className="ic-title">{title}</div>
              <div className="ic-grid">
                {names.map((n) => (
                  <button key={n} type="button" className={`ic-cell${picked === n ? ' on' : ''}`} data-icon={n} title={words(n)} aria-pressed={picked === n} onClick={() => setPicked(n)} onDoubleClick={() => { setPicked(n); insert(n); }} style={{ color: colour }}>
                    <Icon name={n} size={30} />
                  </button>
                ))}
              </div>
            </section>
          )) : <div className="ic-none">No icon matches.</div>}
        </div>
        <div className="ic-colours" role="radiogroup" aria-label="Colour">
          <span>Colour</span>
          {COLOURS.map(([c, label]) => (
            <button key={c} type="button" role="radio" aria-checked={colour === c} className={`ic-swatch${colour === c ? ' on' : ''}`} data-colour={c} title={label} style={{ background: c }} onClick={() => setColour(c)} />
          ))}
        </div>
      </div>
    </Dialog>
  );
}

const ICONS_CSS = `
.ic { display: flex; flex-direction: column; gap: 8px; }
.ic-find { width: 100%; box-sizing: border-box; }
.ic-scroll { height: 300px; overflow: auto; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); padding: 6px 8px; }
.ic-title { font-size: 11px; text-transform: uppercase; letter-spacing: .06em; color: var(--ink-3); font-weight: 600; margin: 6px 0 4px; }
.ic-grid { display: grid; grid-template-columns: repeat(auto-fill, 52px); gap: 6px; }
.ic-cell { width: 52px; height: 52px; display: grid; place-items: center; border: 1px solid transparent; border-radius: var(--r-1); background: #fff; cursor: pointer; }
.ic-cell:hover { border-color: var(--line); }
.ic-cell.on { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }
.ic-none { padding: 12px; color: var(--ink-3); }
.ic-colours { display: flex; align-items: center; gap: 6px; font-size: 12px; color: var(--ink-2); }
.ic-colours span { margin-right: 4px; }
.ic-swatch { width: 20px; height: 20px; border-radius: 50%; border: 2px solid var(--surface); box-shadow: 0 0 0 1px var(--line); cursor: pointer; padding: 0; }
.ic-swatch.on { box-shadow: 0 0 0 2px var(--accent); }
`;
