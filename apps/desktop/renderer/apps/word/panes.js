// What sits around the page: the navigation pane and the styles the ribbon's
// own controls need. (The ruler, which moves things, lives in ruler.js.)
//
// None of it is document. The navigation pane is a view of the headings; the
// view modes are ways of looking. That is why they live beside the page and
// not inside the engine, and why a document opened in two windows can show a
// different one in each.

import React from 'react';
import { Icon, t } from '@rutba/office-ui';

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
        <span>{t('Navigation')}</span>
        <button type="button" className="wd-nav-close" onClick={onClose} title={t('Close the navigation pane')} aria-label={t('Close')}>
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
          {t('No headings yet. Give the paragraphs you want listed a Heading style, and they appear here — and in a table of contents.')}
        </p>
      )}
    </aside>
  );
}

const CSS = `
/* the ribbon's own controls ------------------------------------------------ */
/*
  The style gallery as Word's: each style a small sheet of paper with a
  sample in its look and its name under it. Paper, so dark or light the
  sample is the colour it will be on the page.
*/
.wd-styles { display: flex; align-items: stretch; height: 100%; padding: 3px 0 3px 3px; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface-2); }
.wd-styles > .rw-gallery-items { gap: 2px; height: 100%; padding: 1px; box-sizing: border-box; align-items: stretch; }
.wd-style {
  border: 1px solid transparent; background: #ffffff; color: #1b1b1b;
  border-radius: var(--r-1); padding: 3px 6px 2px; font: inherit; font-size: 13px; cursor: pointer;
  white-space: nowrap; min-width: 64px; max-width: 86px; height: 100%; text-align: center; line-height: 1.15;
  display: flex; flex-direction: column; align-items: center; justify-content: space-between; overflow: hidden;
  box-shadow: 0 0 0 1px rgba(0, 0, 0, 0.06); transition: border-color var(--fast), box-shadow var(--fast);
}
.wd-style::before { content: 'AaBbCc'; display: block; padding-top: 6px; max-width: 100%; overflow: hidden; }
.wd-style-name { font: 400 11px/1.3 var(--font); font-style: normal; color: #5b5f66; text-transform: capitalize; max-width: 100%; overflow: hidden; text-overflow: ellipsis; }
.wd-style:hover { border-color: var(--accent-line); }
.wd-style.on { border-color: var(--accent); box-shadow: 0 0 0 1px var(--accent); }
.wd-style:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.wd-fields { display: flex; align-items: center; gap: 5px; font-size: 11.5px; color: var(--ink-2); }
.wd-fields label { min-width: 44px; }
.wd-field-value { min-width: 46px; text-align: center; font-variant-numeric: tabular-nums; }

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
.wd-nav-item.on { color: var(--accent-text); font-weight: 600; box-shadow: inset 3px 0 0 var(--accent); }
.wd-nav-empty { padding: 12px 14px; font-size: 12px; color: var(--ink-3); }

/* formatting marks --------------------------------------------------------- */
/* The ¶ takes no room on its line — a box no wider than nothing, the mark drawn past it — so it may hang past the margin, as Word's does, rather than push a right-aligned page number onto a line of its own. */
.wd-page.marks .wd-block::after { content: '¶'; color: var(--accent-text); opacity: .55; display: inline-block; width: 0; text-indent: 0; transform: translateX(2px); font-weight: 400; }
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

/* Layout → Hyphenation → Manual: the word, its break points a click each. */
.wd-hyph-word {
  display: flex; align-items: baseline; flex-wrap: wrap; font-size: 20px; letter-spacing: .01em; color: var(--ink);
  padding: 8px 12px; border: 1px solid var(--line); border-radius: var(--r-1); background: var(--surface); min-height: 44px; box-sizing: border-box;
}
.wd-hyph-point {
  border: 0; background: none; color: var(--ink-3); font: inherit; font-size: 18px; padding: 0 2px; margin: 0 1px;
  cursor: pointer; border-radius: 3px; line-height: 1.2;
}
.wd-hyph-point:hover { background: var(--hover); color: var(--ink); }
.wd-hyph-point.on { background: var(--accent-fill); color: #fff; }
.wd-hyph-done { padding: 10px 2px; color: var(--ink-2); font-size: 13px; }
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
