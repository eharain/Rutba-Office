// What sits around the page: the navigation pane, the ruler, and the styles
// the ribbon's own controls need.
//
// None of it is document. The navigation pane is a view of the headings; the
// ruler is a view of the margins; the view modes are ways of looking. That is
// why they live beside the page and not inside the engine, and why a document
// opened in two windows can show a different one in each.

import React from 'react';
import { Icon } from '@rutba/office-ui';

/**
 * Headings, in order, indented by level — the way Word's Navigation Pane draws
 * them, and the fastest way through a two-hundred-page specification.
 */
export function NavigationPane({ blocks, at, onGo, onClose }) {
  const headings = (blocks || [])
    .map((b) => ({ index: b.index, level: Number((/^Heading(\d)$/.exec(b.style || '') || [])[1] || (b.style === 'Title' ? 0 : 9)), text: b.text || (b.runs || []).map((r) => r.text).join('') }))
    .filter((h) => h.level <= 3 && h.text.trim());

  // The heading the caret is under: the last one at or before the caret.
  let current = -1;
  for (const h of headings) if (h.index <= at) current = h.index;

  return (
    <aside className="wd-nav">
      <div className="wd-nav-head">
        <span>Navigation</span>
        <button type="button" className="wd-nav-close" onClick={onClose} title="Close the navigation pane" aria-label="Close">
          <Icon name="close" size={12} />
        </button>
      </div>
      {headings.length ? (
        <div className="wd-nav-list">
          {headings.map((h) => (
            <button
              key={h.index}
              type="button"
              className={`wd-nav-item${h.index === current ? ' on' : ''}`}
              style={{ paddingLeft: 10 + Math.max(0, h.level) * 12 }}
              onClick={() => onGo(h.index)}
              title={h.text}
            >
              {h.text}
            </button>
          ))}
        </div>
      ) : (
        <p className="wd-nav-empty">
          No headings yet. Give the paragraphs you want listed a Heading style, and they appear here — and in a table of contents.
        </p>
      )}
    </aside>
  );
}

/**
 * A ruler in centimetres, with the margins shaded. Read-only: dragging the
 * indent markers is a nicety the paragraph dialog does not need yet.
 */
export function Ruler({ section }) {
  const widthPx = section ? Math.round(section.widthPx) : 794;
  const left = section?.margins?.left ?? 96;
  const right = section?.margins?.right ?? 96;
  const pxPerCm = 96 / 2.54;
  const cms = Math.floor(widthPx / pxPerCm);
  return (
    <div className="wd-ruler" style={{ width: widthPx }} aria-hidden="true">
      <div className="wd-ruler-margin" style={{ left: 0, width: left }} />
      <div className="wd-ruler-margin" style={{ right: 0, width: right }} />
      {Array.from({ length: cms + 1 }, (_, i) => (
        <span key={i} className={`wd-ruler-tick${i % 5 === 0 ? ' major' : ''}`} style={{ left: i * pxPerCm }}>
          {i % 5 === 0 && i > 0 ? i : ''}
        </span>
      ))}
    </div>
  );
}

const CSS = `
/* the ribbon's own controls ------------------------------------------------ */
.wd-styles { display: flex; gap: 3px; align-items: stretch; }
.wd-style {
  border: 1px solid var(--line); background: var(--surface); color: var(--ink);
  border-radius: var(--r-1); padding: 4px 9px; font: inherit; font-size: 12px; cursor: pointer;
  white-space: nowrap; min-width: 58px; text-align: center; line-height: 1.2; height: 44px;
}
.wd-style:hover { border-color: var(--accent-line); background: var(--hover); }
.wd-style.on { border-color: var(--accent); background: var(--selected); }
.wd-fields { display: flex; align-items: center; gap: 5px; font-size: 11.5px; color: var(--ink-2); }
.wd-fields label { min-width: 44px; }
.wd-field-value { min-width: 46px; text-align: center; font-variant-numeric: tabular-nums; }
.wd-symbols { display: grid; grid-template-columns: repeat(12, 1fr); gap: 3px; }
.wd-symbol {
  border: 1px solid var(--line); background: var(--surface); color: var(--ink); border-radius: var(--r-1);
  font-size: 18px; height: 34px; cursor: pointer; font-family: "Segoe UI Symbol", "Segoe UI", system-ui, sans-serif;
}
.wd-symbol:hover { border-color: var(--accent-line); background: var(--hover); }

/* the navigation pane ------------------------------------------------------ */
/* The pane sits beside the scroll area; the scroll area takes the rest, or
   the page drifts to the right of centre the moment .wd becomes a flex row. */
.wd { display: flex; }
.wd > .wd-scroll { flex: 1; min-width: 0; }
.wd-nav {
  width: 240px; flex: none; border-right: 1px solid var(--line); background: var(--chrome);
  display: flex; flex-direction: column; min-height: 0;
}
.wd-nav-head {
  display: flex; align-items: center; justify-content: space-between; padding: 9px 12px;
  font-size: 11px; text-transform: uppercase; letter-spacing: .06em; color: var(--ink-3);
  border-bottom: 1px solid var(--line-soft);
}
.wd-nav-close { border: 0; background: none; color: var(--ink-3); cursor: pointer; padding: 2px; display: grid; border-radius: var(--r-1); }
.wd-nav-close:hover { background: var(--hover); color: var(--ink); }
.wd-nav-list { overflow: auto; padding: 4px 0; }
.wd-nav-item {
  display: block; width: 100%; text-align: left; border: 0; background: none; color: var(--ink-2);
  font: inherit; font-size: 12.5px; padding: 5px 10px; cursor: pointer; white-space: nowrap;
  overflow: hidden; text-overflow: ellipsis;
}
.wd-nav-item:hover { background: var(--hover); color: var(--ink); }
.wd-nav-item.on { color: var(--accent); font-weight: 600; box-shadow: inset 3px 0 0 var(--accent); }
.wd-nav-empty { padding: 12px 14px; font-size: 12px; color: var(--ink-3); }

/* the ruler ---------------------------------------------------------------- */
.wd-ruler {
  position: relative; height: 18px; margin: 10px auto 0; background: var(--surface);
  border: 1px solid var(--line); border-radius: 3px 3px 0 0; overflow: hidden; font-size: 9px; color: var(--ink-3);
}
.wd-ruler-margin { position: absolute; top: 0; bottom: 0; background: var(--sunken); }
.wd-ruler-tick { position: absolute; top: 12px; width: 1px; height: 5px; background: var(--ink-4, var(--ink-3)); text-align: center; text-indent: -4px; line-height: 0; }
.wd-ruler-tick.major { top: 8px; height: 9px; }
.wd-ruler-tick.major::after { content: attr(data-n); }

/* formatting marks --------------------------------------------------------- */
.wd-page.marks .wd-block::after { content: '¶'; color: var(--accent); opacity: .55; margin-left: 2px; font-weight: 400; }
.wd-page.marks .wd-block { white-space: pre-wrap; }

/* view modes --------------------------------------------------------------- */
.wd.mode-web .wd-page, .wd.mode-draft .wd-page {
  width: auto !important; min-height: 0 !important; box-shadow: none; border: 0; margin: 0;
  padding: 24px 40px !important;
}
.wd.mode-web .wd-scroll, .wd.mode-draft .wd-scroll { background: var(--surface); }
.wd.mode-draft .wd-page img { display: none; }
.wd.mode-read .wd-page {
  width: min(760px, 92vw) !important; min-height: 0 !important; padding: 48px 56px !important;
  font-size: 1.15em; line-height: 1.7; box-shadow: none; border-radius: var(--r-2);
}
.wd.mode-outline .wd-page { width: auto !important; min-height: 0 !important; padding: 20px 32px !important; box-shadow: none; }
.wd.mode-outline .wd-block:not([data-style^="Heading"]):not([data-style="Title"]) { display: none; }
.wd.mode-outline .wd-table, .wd.mode-outline .wd-page img { display: none; }
.wd.mode-outline .wd-block[data-style="Heading2"] { margin-left: 24px; }
.wd.mode-outline .wd-block[data-style="Heading3"] { margin-left: 48px; }
.wd.focus .wd-scroll { background: #2b2d31; }
`;

let installed = false;
/** Injected once; the ribbon and the panes share these rules. */
export function installWordStyles() {
  if (installed || typeof document === 'undefined') return;
  installed = true;
  const style = document.createElement('style');
  style.id = 'rutba-word-panes-css';
  style.textContent = CSS;
  document.head.appendChild(style);
}
