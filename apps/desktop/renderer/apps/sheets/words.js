// Insert → WordArt on a sheet, and a double-click on a text box or a
// shape: the words it says, typed in — a sheet's drawings have no caret of
// their own, so the words are asked for in a box, with the look they will
// take shown as they are typed.

import React, { useState } from 'react';
import { Button, Dialog, t } from '@rutba/office-ui';
import { wordArtCss } from '../../wordart.js';

export function ShapeWordsDialog({ title = t('Edit Text'), initial = '', look = null, onOk, onClose }) {
  const [text, setText] = useState(initial);
  const ok = () => { if (text.trim()) onOk(text.replace(/\s+$/, '')); };
  return (
    <Dialog
      title={title}
      width={460}
      onClose={onClose}
      actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary label={t('OK')} className="sh-words-ok" disabled={!text.trim()} onClick={ok} /></>}
    >
      <div className="sh-words">
        <textarea className="rw-input sh-words-text" rows={3} value={text} autoFocus onFocus={(e) => e.target.select()} onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); ok(); } }} />
        {look ? (
          <div className="sh-words-preview" aria-label={t('Preview')}>
            <span style={{ fontWeight: 700, fontSize: 34, lineHeight: 1.15, whiteSpace: 'pre-wrap', fontFamily: 'Calibri, Carlito, sans-serif', ...wordArtCss(look, 1.4) }}>{text || ' '}</span>
          </div>
        ) : null}
        <p className="sh-words-lead">{t('Each line is a paragraph of its own. Ctrl+Enter puts the words in.')}</p>
      </div>
    </Dialog>
  );
}

export const WORDS_CSS = `
.sh-words { display: flex; flex-direction: column; gap: 10px; }
.sh-words-text { width: 100%; box-sizing: border-box; resize: vertical; font: inherit; font-size: 14px; }
.sh-words-preview { min-height: 70px; display: flex; align-items: center; justify-content: center; text-align: center; padding: 12px; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: #fff; overflow: hidden; }
.sh-words-lead { margin: 0; font-size: 11.5px; color: var(--ink-3); }
`;
