// Insert → SmartArt's box, for all three apps: Office's Choose a SmartArt
// Graphic — the kinds down the side, the layouts of the kind picked, and a
// picture of the one chosen — with Office's text pane beside it, where the
// diagram's words are typed a line each (a Tab in for a line under the one
// above). OK hands back the layout and the items; the app lays them out
// (smartart.js) and puts the shapes in.

import React, { useMemo, useState } from 'react';
import { Button, Dialog, t } from '@rutba/office-ui';
import { SMARTART_LAYOUTS, parseItems, smartArtSvg } from './smartart.js';

const KINDS = ['All', 'List', 'Process', 'Cycle', 'Hierarchy'];
const KIND_NAMES = { All: t('All'), List: t('List'), Process: t('Process'), Cycle: t('Cycle'), Hierarchy: t('Hierarchy') };

const ABOUT = {
  blockList: t('Shows ideas that need not follow one another, each in a box of its own, the boxes in rows.'),
  verticalList: t('Shows a list of ideas one under the other, each in a wide rounded box.'),
  process: t('Shows the steps of a process in order, left to right, an arrow from each to the next.'),
  chevron: t('Shows the stages of a process moving forward, each a chevron pointing to the next.'),
  cycle: t('Shows stages that come round again, in circles on a ring, clockwise from the top.'),
  hierarchy: t('Shows each top item over the items under it — a line to each. Tab a line in to put it under the one above.'),
};

const STARTER = {
  hierarchy: `${t('Lead')}\n\t${t('First')}\n\t${t('Second')}\n\t${t('Third')}`,
  default: `${t('First')}\n${t('Second')}\n${t('Third')}`,
};

/**
 * `initial` — the text it starts with (a box's words, for Convert to
 * SmartArt); `title` the box's own title. `onInsert({ layout, name, items })`.
 */
export function SmartArtDialog({ initial = '', title = t('Choose a SmartArt Graphic'), onInsert, onClose }) {
  const [kind, setKind] = useState('All');
  const [layout, setLayout] = useState('blockList');
  const [text, setText] = useState(initial || null);
  const words = text ?? (STARTER[layout] || STARTER.default);
  const items = useMemo(() => parseItems(words), [words]);
  const chosen = SMARTART_LAYOUTS.find((l) => l.id === layout);
  const shown = SMARTART_LAYOUTS.filter((l) => kind === 'All' || l.group === kind);
  const ok = () => { if (items.length) onInsert({ layout, name: chosen.name, items }); };
  return (
    <Dialog
      title={title}
      width={780}
      onClose={onClose}
      actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary label={t('OK')} className="sa-ok" disabled={!items.length} onClick={ok} /></>}
    >
      <div className="sa">
        <div className="sa-kinds" role="tablist" aria-label={t('Kind')}>
          {KINDS.map((k) => (
            <button key={k} type="button" role="tab" aria-selected={kind === k} className={`sa-kind${kind === k ? ' on' : ''}`} onClick={() => setKind(k)}>{KIND_NAMES[k]}</button>
          ))}
        </div>
        <div className="sa-gallery" role="listbox" aria-label={t('Layout')}>
          {shown.map((l) => (
            <button
              key={l.id}
              type="button"
              role="option"
              aria-selected={layout === l.id}
              className={`sa-item${layout === l.id ? ' on' : ''}`}
              data-layout={l.id}
              title={l.name}
              onClick={() => setLayout(l.id)}
            >
              <span className="sa-thumb" dangerouslySetInnerHTML={{ __html: smartArtSvg(l.id, [], { width: 96, height: 64, labels: false }) }} />
              <span className="sa-name">{l.name}</span>
            </button>
          ))}
        </div>
        <div className="sa-side">
          <div className="sa-preview" dangerouslySetInnerHTML={{ __html: smartArtSvg(layout, items, { width: 260, height: 160 }) }} />
          <div className="sa-title">{chosen.name}</div>
          <p className="sa-about">{ABOUT[layout]}</p>
          <label className="sa-label" htmlFor="sa-text">{t('Type your text here')}</label>
          <textarea
            id="sa-text"
            className="sa-text"
            value={words}
            spellCheck={false}
            onChange={(e) => setText(e.target.value)}
            onKeyDown={(e) => {
              // Tab puts the line under the one above, Shift+Tab brings it back, as in Office's text pane.
              if (e.key !== 'Tab') return;
              e.preventDefault();
              const el = e.currentTarget;
              const start = el.value.lastIndexOf('\n', el.selectionStart - 1) + 1;
              const line = el.value.slice(start);
              const next = e.shiftKey
                ? (line.startsWith('\t') ? el.value.slice(0, start) + line.slice(1) : el.value)
                : el.value.slice(0, start) + '\t' + line;
              const moved = next.length - el.value.length;
              const caret = Math.max(start, el.selectionStart + moved);
              setText(next);
              requestAnimationFrame(() => { el.selectionStart = el.selectionEnd = caret; });
            }}
          />
        </div>
      </div>
    </Dialog>
  );
}

export const SMARTART_CSS = `
.sa { display: grid; grid-template-columns: 120px 1fr 276px; gap: 12px; min-height: 360px; font-size: 13px; }
.sa-kinds { display: flex; flex-direction: column; gap: 2px; border-right: 1px solid var(--line-soft); padding-right: 8px; }
.sa-kind { text-align: left; border: 0; background: none; padding: 6px 8px; border-radius: var(--r-2); font: inherit; color: var(--ink); cursor: pointer; }
.sa-kind.on { background: var(--selected); color: var(--accent); }
.sa-gallery { display: grid; grid-template-columns: repeat(auto-fill, minmax(112px, 1fr)); gap: 8px; align-content: start; overflow: auto; max-height: 380px; }
.sa-item { display: flex; flex-direction: column; align-items: center; gap: 4px; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--window); padding: 8px 4px 6px; font: inherit; color: var(--ink-2); cursor: pointer; }
.sa-item.on { border-color: var(--accent); box-shadow: 0 0 0 1px var(--accent); color: var(--ink); }
.sa-thumb { display: block; line-height: 0; }
.sa-name { font-size: 11.5px; text-align: center; }
.sa-side { display: flex; flex-direction: column; gap: 6px; border-left: 1px solid var(--line-soft); padding-left: 12px; }
.sa-preview { line-height: 0; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: #fff; padding: 4px; align-self: flex-start; }
.sa-title { font-weight: 600; }
.sa-about { margin: 0; font-size: 11.5px; color: var(--ink-3); }
.sa-label { font-size: 11.5px; color: var(--ink-2); margin-top: 4px; }
.sa-text { flex: 1; min-height: 110px; resize: none; font: 13px/1.45 var(--font-ui, "Segoe UI", sans-serif); tab-size: 3; padding: 6px 8px; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--window); color: var(--ink); }
`;
