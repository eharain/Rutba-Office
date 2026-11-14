// Review → Compare → Compare Documents: Word's box — the original document,
// the revised one (this one, when it is saved), and the name its changes
// are marked with — then the compared document in a window of its own.

import React, { useState } from 'react';
import { Button, Dialog, t } from '@rutba/office-ui';

const nameOf = (p) => (p ? String(p).split(/[\\/]/).pop() : '');

export function CompareDialog({ shell, current = null, author = '', onCompare, onClose }) {
  const [original, setOriginal] = useState(null);
  const [revised, setRevised] = useState(current);
  const [label, setLabel] = useState(author);
  const browse = async (set) => {
    const [file] = await shell.dialog.open({ title: t('common.open'), filters: [{ name: t('word.compare.word-documents'), extensions: ['docx', 'docm', 'dotx'] }] });
    if (file) set(file);
  };
  return (
    <Dialog
      title={t('word.compare.compare-documents')}
      width={520}
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.ok')} className="wd-compare-ok" disabled={!original || !revised} onClick={() => onCompare({ original, revised, author: label.trim() || null })} /></>}
    >
      <div className="wd-compare">
        <div className="wd-compare-col">
          <div className="wd-compare-head">{t('word.compare.original-document')}</div>
          <div className="wd-compare-file wd-compare-original" title={original || ''}>{nameOf(original) || t('word.compare.none-chosen')}</div>
          <Button label={t('common.browse')} className="wd-compare-browse-original" onClick={() => browse(setOriginal)} />
        </div>
        <div className="wd-compare-col">
          <div className="wd-compare-head">{t('word.compare.revised-document')}</div>
          <div className="wd-compare-file wd-compare-revised" title={revised || ''}>{nameOf(revised) || t('word.compare.none-chosen')}</div>
          <Button label={t('common.browse')} className="wd-compare-browse-revised" onClick={() => browse(setRevised)} />
        </div>
        <label className="wd-compare-label">{t('word.compare.label-changes-with')} <input className="rw-input" value={label} onChange={(e) => setLabel(e.target.value)} /></label>
        <p className="wd-compare-lead">{t('word.compare.compared-document')}</p>
      </div>
    </Dialog>
  );
}

export const COMPARE_CSS = `
.wd-compare { display: grid; grid-template-columns: 1fr 1fr; gap: 12px 16px; font-size: 13px; }
.wd-compare-col { display: flex; flex-direction: column; gap: 6px; }
.wd-compare-head { font-size: 11.5px; color: var(--ink-2); }
.wd-compare-file { padding: 6px 8px; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); min-height: 18px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.wd-compare-label { grid-column: 1 / span 2; display: flex; align-items: center; gap: 10px; }
.wd-compare-label input { flex: 1; }
.wd-compare-lead { grid-column: 1 / span 2; margin: 0; font-size: 11.5px; color: var(--ink-3); }
`;
