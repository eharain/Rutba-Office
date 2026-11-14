// Insert → Object, for Presentations (and Documents): PowerPoint's Insert
// Object box — Create new (a blank Word document, Excel worksheet or
// PowerPoint presentation) or Create from file (an Office document from
// this computer) — embedded as an icon with its name, and opened in its
// own app by a double-click.

import React, { useState } from 'react';
import { Button, Dialog, t } from '@rutba/office-ui';
import { iconSvg } from '@rutba/office-ui/icons';

export const OBJECT_TYPES = [
  { ext: 'docx', label: t('object-insert.microsoft-word'), glyph: 'word', colour: '#2B5FD9', app: 'word' },
  { ext: 'xlsx', label: t('object-insert.microsoft-excel'), glyph: 'sheets', colour: '#1E8A4C', app: 'sheets' },
  { ext: 'pptx', label: t('object-insert.microsoft-power'), glyph: 'slides', colour: '#C4422F', app: 'slides' },
];
export const typeOfExt = (ext) => OBJECT_TYPES.find((t) => t.ext === String(ext || '').toLowerCase()) || null;
export const appOfProgId = (progId) => (/^Word\./.test(progId || '') ? 'word' : /^Excel\./.test(progId || '') ? 'sheets' : /^PowerPoint\./.test(progId || '') ? 'slides' : null);

/** An object's icon, as Office draws one: its program's mark over its name, as PNG bytes. */
export async function objectIcon(type, name, { width = 128, height = 120 } = {}) {
  const img = new Image();
  img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(iconSvg(type.glyph, { colour: type.colour, size: 64 }))}`;
  await img.decode();
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const g = canvas.getContext('2d');
  g.drawImage(img, (width - 64) / 2, 8, 64, 64);
  g.fillStyle = '#333';
  g.font = '13px "Segoe UI", sans-serif';
  g.textAlign = 'center';
  let label = String(name || type.label);
  while (label.length > 4 && g.measureText(label).width > width - 8) label = label.slice(0, -2);
  if (label !== String(name || type.label)) label += '…';
  g.fillText(label, width / 2, 96);
  return new Uint8Array(await (await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'))).arrayBuffer());
}

/**
 * The Insert Object box. `onInsert({ ext, name, file })` — `file` a path for
 * Create from file, none for Create new.
 */
export function ObjectDialog({ shell, onInsert, onClose }) {
  const [mode, setMode] = useState('new');
  const [type, setType] = useState('docx');
  const [file, setFile] = useState(null);
  const browse = async () => {
    const [picked] = await shell.dialog.open({ title: t('object-insert.browse'), filters: [{ name: t('object-insert.office-documents'), extensions: ['docx', 'xlsx', 'pptx'] }] });
    if (picked) setFile(picked);
  };
  const ok = () => {
    if (mode === 'new') onInsert({ ext: type, name: typeOfExt(type).label, file: null });
    else if (file) onInsert({ ext: file.split('.').pop().toLowerCase(), name: file.split(/[\\/]/).pop(), file });
  };
  return (
    <Dialog
      title={t('object-insert.insert-object')}
      width={480}
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.ok')} className="obj-ok" disabled={mode === 'file' && !file} onClick={ok} /></>}
    >
      <div className="obj">
        <div className="obj-modes">
          <label><input type="radio" name="obj-mode" checked={mode === 'new'} onChange={() => setMode('new')} /> {t('object-insert.create-new')}</label>
          <label><input type="radio" name="obj-mode" className="obj-mode-file" checked={mode === 'file'} onChange={() => setMode('file')} /> {t('object-insert.create-from-file')}</label>
        </div>
        {mode === 'new' ? (
          <div className="obj-list" role="listbox" aria-label={t('object-insert.object-type')}>
            {OBJECT_TYPES.map((t) => (
              <button key={t.ext} type="button" role="option" aria-selected={type === t.ext} className={`obj-item${type === t.ext ? ' on' : ''}`} data-ext={t.ext} onClick={() => setType(t.ext)}>{t.label}</button>
            ))}
          </div>
        ) : (
          <div className="obj-file">
            <div className="obj-path">{file || t('object-insert.no-file-chosen')}</div>
            <Button label={t('common.browse')} className="obj-browse" onClick={browse} />
          </div>
        )}
        <p className="obj-lead">{t('object-insert.document-goes-whole')}</p>
      </div>
    </Dialog>
  );
}

export const OBJECT_CSS = `
.obj { display: flex; flex-direction: column; gap: 10px; font-size: 13px; }
.obj-modes { display: flex; gap: 18px; }
.obj-list { border: 1px solid var(--line-soft); border-radius: var(--r-2); }
.obj-item { display: block; width: 100%; text-align: left; border: 0; border-bottom: 1px solid var(--line-soft); background: none; padding: 6px 10px; font: inherit; color: var(--ink); cursor: pointer; }
.obj-item.on { background: var(--selected); color: var(--accent-text); }
.obj-file { display: flex; gap: 8px; align-items: center; }
.obj-path { flex: 1; padding: 6px 8px; border: 1px solid var(--line-soft); border-radius: var(--r-2); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.obj-lead { margin: 0; font-size: 11.5px; color: var(--ink-3); }
`;
