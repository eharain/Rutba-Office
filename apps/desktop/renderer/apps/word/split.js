// View → Split: the document in two panes, one over the other — the page
// you work on above, and below it the same document again, scrolled on its
// own, so one part can be read while another is written. The lower pane is
// the page as it stands, brought up to date as you type; a click in it puts
// the caret there in the pane above. The bar between them is dragged to
// share the window out.

import React, { useEffect, useRef } from 'react';

export function SplitPane({ pageRef, version, onGo, onResize }) {
  const host = useRef(null);
  // The page as it stands, again: copied a moment after each change, read-only.
  useEffect(() => {
    const t = setTimeout(() => {
      const page = pageRef.current;
      const el = host.current;
      if (!page || !el) return;
      const copy = page.cloneNode(true);
      copy.removeAttribute('contenteditable');
      copy.setAttribute('contenteditable', 'false');
      copy.removeAttribute('id');
      copy.classList.add('wd-split-copy');
      for (const n of copy.querySelectorAll('[contenteditable], [id]')) { n.removeAttribute('contenteditable'); n.removeAttribute('id'); }
      el.replaceChildren(copy);
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
      <div className="wd-split-bar" role="separator" aria-orientation="horizontal" title="Drag to share the window between the panes" onMouseDown={drag} />
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
