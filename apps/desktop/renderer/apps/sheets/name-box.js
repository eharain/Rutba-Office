// The Name Box, at the left of the formula bar.
//
// It shows the address of what is selected. Clicked, it takes typing: an
// address goes there, a name the workbook has goes to its range, and any
// other word names what is selected, as Excel's does. Enter does it, Escape
// puts the address back.

import React, { useEffect, useRef, useState } from 'react';
import { t } from '@rutba/office-ui';
import { nameBoxEntry } from '@rutba/sheet-view/name-box';

/**
 * `address` is the selection's address, `names` the workbook's names
 * (`[{ name }]`). `onGo(entry)` is given what was typed, read by
 * `nameBoxEntry`; `onDone()` is called when the box is finished with, so the
 * grid can take the keys back.
 */
export function NameBox({ address, names = [], onGo, onDone }) {
  // null while the box only shows the address; the typed text while it is in use.
  const [text, setText] = useState(null);
  const inputRef = useRef(null);
  const typing = text != null;
  useEffect(() => {
    if (!typing) return;
    inputRef.current?.focus();
    inputRef.current?.select();
  }, [typing]);

  if (!typing) {
    return (
      // Its name holds the address it shows, so a person who says what they see (A1) finds it (WCAG 2.5.3).
      <div className="sh-namebox" role="button" title={t('sheets.name-box.name-box')} aria-label={address ? t('sheets.name-box.name-box') + ' ' + address : t('sheets.name-box.name-box')}
        onMouseDown={(e) => { e.preventDefault(); setText(address || ''); }}>{address}</div>
    );
  }
  const finish = () => { setText(null); onDone?.(); };
  return (
    <div className="sh-namebox typing">
      <input
        ref={inputRef}
        className="sh-namebox-input"
        value={text}
        spellCheck={false}
        aria-label={t('sheets.name-box.name-box')}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            const entry = nameBoxEntry(text, names);
            finish();
            if (entry) onGo(entry);
          } else if (e.key === 'Escape') {
            e.preventDefault();
            finish();
          }
        }}
        onBlur={() => setText(null)}
      />
    </div>
  );
}

export const NAME_BOX_CSS = `
.sh-namebox { cursor: text; }
.sh-namebox.typing { padding: 0; border-color: var(--accent); }
.sh-namebox-input {
  width: 112px; height: 100%; padding: 0 6px; border: 0; outline: 0; background: transparent;
  font: inherit; color: var(--ink); text-align: center;
}
`;
