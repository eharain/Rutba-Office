// View → Split: the document in two panes, one over the other — the page
// you work on above, and below it the same document again, scrolled on its
// own, so one part can be read while another is written. The lower pane is
// the page as it stands, brought up to date as you type; a click in it puts
// the caret there in the pane above. The bar between them is dragged to
// share the window out.

import React, { useEffect, useRef } from 'react';
import { t } from '@rutba/office-ui';

/** A node made safe to show as a copy: nothing in it editable, no id twice in the window. */
function inert(node) {
  if (node.nodeType !== 1) return node;
  node.removeAttribute('contenteditable');
  node.removeAttribute('id');
  for (const n of node.querySelectorAll('[contenteditable], [id]')) { n.removeAttribute('contenteditable'); n.removeAttribute('id'); }
  return node;
}

export function SplitPane({ pageRef, version, onGo, onResize }) {
  const host = useRef(null);
  // What changed on the page since the copy was made: the paragraphs typed
  // in, by their block, or `all` when the pages themselves moved. A long
  // document was copied whole a moment after every change.
  const dirty = useRef({ all: true, blocks: new Set() });
  useEffect(() => {
    const page = pageRef.current;
    if (!page) return undefined;
    const seen = new MutationObserver((records) => {
      const d = dirty.current;
      if (d.all) return;
      for (const r of records) {
        const at = r.target.nodeType === 1 ? r.target : r.target.parentElement;
        const block = at?.closest?.('[data-block]');
        // Inside one paragraph: that paragraph. Anything else — a paragraph
        // added or taken away, a page laid out again — the whole copy.
        if (block) d.blocks.add(block.dataset.block);
        else { d.all = true; return; }
      }
    });
    seen.observe(page, { subtree: true, childList: true, characterData: true, attributes: true });
    return () => seen.disconnect();
  }, [pageRef]);
  // The page as it stands, again: brought up to date a moment after each change, read-only.
  useEffect(() => {
    const t = setTimeout(() => {
      const page = pageRef.current;
      const el = host.current;
      if (!page || !el) return;
      const d = dirty.current;
      const copy = el.firstElementChild;
      let whole = d.all || !copy;
      if (!whole) {
        for (const i of d.blocks) {
          const now = page.querySelectorAll(`[data-block="${i}"]`);
          const was = copy.querySelectorAll(`[data-block="${i}"]`);
          // A paragraph now laid over a different number of pages: copy it all.
          if (now.length !== was.length) { whole = true; break; }
          now.forEach((n, k) => was[k].replaceWith(inert(n.cloneNode(true))));
        }
      }
      if (whole) {
        const fresh = inert(page.cloneNode(true));
        fresh.setAttribute('contenteditable', 'false');
        fresh.classList.add('wd-split-copy');
        el.replaceChildren(fresh);
      }
      dirty.current = { all: false, blocks: new Set() };
    }, 180);
    return () => clearTimeout(t);
  }, [pageRef, version]);
  const drag = (e) => {
    e.preventDefault();
    const box = e.currentTarget.parentElement.getBoundingClientRect();
    const move = (ev) => onResize(Math.max(0.15, Math.min(0.85, (ev.clientY - box.top) / box.height)));
    const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };
  return (
    <>
      <div className="wd-split-bar" role="separator" aria-orientation="horizontal" title={t('Drag to share the window between the panes')} onMouseDown={drag} />
      <div
        className="wd-split-pane"
        ref={host}
        onClick={(e) => {
          const b = e.target.closest?.('[data-block]');
          const i = b ? Number(b.dataset.block) : NaN;
          if (Number.isFinite(i)) onGo(i);
        }}
      />
    </>
  );
}

export const SPLIT_CSS = `
.wd > .wd-splits { flex: 1; min-width: 0; min-height: 0; display: flex; flex-direction: column; }
.wd-splits > .wd-scroll { min-height: 0; }
.wd-scroll.split-top { flex: none; height: var(--split, 50%); }
.wd-split-bar { flex: none; height: 6px; cursor: row-resize; background: var(--line); border-top: 1px solid var(--line-strong); border-bottom: 1px solid var(--line-strong); }
.wd-split-bar:hover { background: var(--accent-soft, var(--line-strong)); }
.wd-split-pane { flex: 1; min-height: 0; overflow: auto; padding: 26px 0 40px; display: flex; flex-direction: column; align-items: center; background: var(--window); cursor: default; }
.wd-split-copy { caret-color: transparent; user-select: text; }
`;
