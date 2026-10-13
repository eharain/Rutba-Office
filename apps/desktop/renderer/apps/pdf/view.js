// The PDF window's two views of the pages: the strip of thumbnails down the
// side, where pages are picked, dragged into a new order and right-clicked,
// and the pages themselves at the zoom chosen, one at a time, two side by
// side or in one column, each with its words laid over it to select, copy
// and find.
//
// A page is drawn only when it comes near the screen, and drawn again only
// when its size or turn changes.

import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import { t } from '@rutba/office-ui';
import { drawPage, drawText } from './pdfjs.js';
import { shownSize } from './pages.js';

/** CSS pixels a point at 100%: a PDF's 72 to the inch on a 96 to the inch screen. */
export const PX_PER_PT = 96 / 72;

/** Whether an element is near enough the screen to draw. */
function useNear(ref, root, margin = '600px') {
  const [near, setNear] = useState(false);
  useEffect(() => {
    const el = ref.current;
    if (!el) return undefined;
    const io = new IntersectionObserver((entries) => {
      for (const e of entries) setNear(e.isIntersecting);
    }, { root: root?.current || null, rootMargin: margin });
    io.observe(el);
    return () => io.disconnect();
  }, [ref, root, margin]);
  return near;
}

/** One page drawn on a canvas at `scale` CSS pixels a point. */
function PageCanvas({ doc, page, rotate, scale, priority, className, onDrawn }) {
  const ref = useRef(null);
  const [drawn, setDrawn] = useState(0);
  useEffect(() => {
    const canvas = ref.current;
    if (!canvas || !doc) return undefined;
    const job = drawPage(doc, page, canvas, { scale, rotate, priority });
    let live = true;
    job.done.then(() => {
      if (!live) return;
      setDrawn((n) => n + 1);
      onDrawn?.();
    }).catch(() => {});
    return () => { live = false; job.cancel(); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [doc, page, rotate, scale]);
  return <canvas ref={ref} className={className} data-drawn={drawn ? 'yes' : undefined} />;
}

/**
 * The thumbnail strip. A press picks a page (Ctrl adds or takes away one,
 * Shift takes the run from the last picked); dragging a picked page moves
 * every picked page to where the line shows, as Acrobat's and PowerPoint's
 * strips do.
 */
export function ThumbStrip({ pages, sources, selected, current, onPick, onMove, onMenu, onKey, width = 132 }) {
  const scroller = useRef(null);
  const [drop, setDrop] = useState(null);
  const drag = useRef(null);

  const dropIndexAt = useCallback((y) => {
    const items = [...(scroller.current?.querySelectorAll('.pdf-thumb') || [])];
    for (let i = 0; i < items.length; i++) {
      const r = items[i].getBoundingClientRect();
      if (y < r.top + r.height / 2) return i;
    }
    return items.length;
  }, []);

  useEffect(() => {
    const move = (e) => {
      const d = drag.current;
      if (!d) return;
      if (!d.moving && Math.hypot(e.clientX - d.x, e.clientY - d.y) < 5) return;
      d.moving = true;
      document.body.classList.add('pdf-dragging');
      setDrop(dropIndexAt(e.clientY));
      // Near an edge of the strip, the strip scrolls.
      const box = scroller.current?.getBoundingClientRect();
      if (box) {
        if (e.clientY < box.top + 30) scroller.current.scrollTop -= 14;
        else if (e.clientY > box.bottom - 30) scroller.current.scrollTop += 14;
      }
    };
    const up = (e) => {
      const d = drag.current;
      drag.current = null;
      document.body.classList.remove('pdf-dragging');
      if (!d) return;
      if (d.moving) {
        const at = dropIndexAt(e.clientY);
        setDrop(null);
        onMove(at, d.key);
      } else {
        setDrop(null);
        onPick(d.key, { ctrl: d.ctrl, shift: d.shift, release: true });
      }
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
    return () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
    };
  }, [dropIndexAt, onMove, onPick]);

  // The page in view stays in view in the strip.
  useEffect(() => {
    const el = scroller.current?.querySelector(`.pdf-thumb[data-key="${current}"]`);
    if (!el) return;
    const box = scroller.current.getBoundingClientRect();
    const r = el.getBoundingClientRect();
    // The strip alone scrolls: scrollIntoView would move the window's frame too.
    if (r.top < box.top) scroller.current.scrollTop += r.top - box.top - 8;
    else if (r.bottom > box.bottom) scroller.current.scrollTop += r.bottom - box.bottom + 8;
  }, [current]);

  const set = new Set(selected);
  return (
    <div className="pdf-strip" ref={scroller} tabIndex={0} onKeyDown={onKey} role="listbox" aria-multiselectable="true" aria-label={t('Pages')}>
      {pages.map((p, i) => {
        const source = sources[p.source];
        const info = source?.info.pages[p.page];
        const size = shownSize(info, p.rotate);
        // Room for the strip's scrollbar and the selection's outline on both sides.
        const scale = (width - 46) / Math.max(size.width, size.height * 0.75);
        return (
          <React.Fragment key={p.key}>
            {drop === i ? <div className="pdf-drop-line" /> : null}
            <Thumb
              entry={p}
              index={i}
              source={source}
              scale={scale}
              size={size}
              root={scroller}
              selected={set.has(p.key)}
              current={current === p.key}
              onDown={(e) => {
                if (e.button !== 0) return;
                e.preventDefault();
                scroller.current?.focus({ preventScroll: true });
                drag.current = { key: p.key, x: e.clientX, y: e.clientY, moving: false, ctrl: e.ctrlKey || e.metaKey, shift: e.shiftKey };
                onPick(p.key, { ctrl: e.ctrlKey || e.metaKey, shift: e.shiftKey, press: true });
              }}
              onMenu={(e) => onMenu(e, p.key)}
            />
          </React.Fragment>
        );
      })}
      {drop === pages.length ? <div className="pdf-drop-line" /> : null}
    </div>
  );
}

function Thumb({ entry, index, source, scale, size, root, selected, current, onDown, onMenu }) {
  const ref = useRef(null);
  const near = useNear(ref, root, '300px');
  return (
    <div
      ref={ref}
      className={`pdf-thumb${selected ? ' selected' : ''}${current ? ' current' : ''}`}
      data-key={entry.key}
      data-index={index}
      data-rotate={entry.rotate}
      data-page={entry.page}
      data-from={source?.name || ''}
      role="option"
      aria-selected={selected}
      onMouseDown={onDown}
      onContextMenu={onMenu}
    >
      <div className="pdf-thumb-paper" style={{ width: Math.round(size.width * scale), height: Math.round(size.height * scale) }}>
        {near && source?.doc ? <PageCanvas doc={source.doc} page={entry.page} rotate={entry.rotate} scale={scale} priority={2} className="pdf-thumb-canvas" /> : null}
      </div>
      <span className="pdf-thumb-label">{index + 1}</span>
    </div>
  );
}

/** One page of the page view: the drawing, and its words over it. */
function PageBox({ entry, index, source, zoom, root, onText }) {
  const ref = useRef(null);
  const textRef = useRef(null);
  const near = useNear(ref, root);
  const info = source?.info.pages[entry.page];
  const size = shownSize(info, entry.rotate);
  const scale = zoom * PX_PER_PT;
  const w = Math.round(size.width * scale);
  const h = Math.round(size.height * scale);
  const [textReady, setTextReady] = useState(false);

  useEffect(() => {
    const el = textRef.current;
    if (!near || !el || !source?.doc) return undefined;
    let live = true;
    setTextReady(false);
    drawText(source.doc, entry.page, el, { scale, rotate: entry.rotate }).then(() => {
      if (!live) return;
      setTextReady(true);
      onText?.();
    }).catch(() => {});
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [near, source, entry.page, entry.rotate, scale]);

  return (
    <div ref={ref} className="pdf-page" data-key={entry.key} data-index={index} data-rotate={entry.rotate} style={{ width: w, height: h }}>
      {near && source?.doc ? <PageCanvas doc={source.doc} page={entry.page} rotate={entry.rotate} scale={scale} priority={0} className="pdf-page-canvas" /> : null}
      <div ref={textRef} className="textLayer pdf-text" data-ready={textReady ? 'yes' : undefined} />
    </div>
  );
}

/**
 * The pages at the zoom chosen. `layout`: 'single' shows the page in view
 * alone, 'two' pairs them side by side, 'continuous' runs them down in one
 * column. Scrolling tells the window which page is in view.
 */
export function PageView({ pages, sources, zoom, layout, current, onCurrent, onText, scrollerRef, onFocusText }) {
  const shown = layout === 'single' ? pages.filter((p) => p.key === current).slice(0, 1) : pages;
  const rows = [];
  if (layout === 'two') {
    for (let i = 0; i < shown.length; i += 2) rows.push(shown.slice(i, i + 2));
  } else {
    for (const p of shown) rows.push([p]);
  }
  const ticking = useRef(false);

  const onScroll = useCallback(() => {
    if (layout === 'single' || ticking.current) return;
    ticking.current = true;
    requestAnimationFrame(() => {
      ticking.current = false;
      const box = scrollerRef.current?.getBoundingClientRect();
      if (!box) return;
      const mid = box.top + Math.min(box.height / 3, 240);
      let best = null;
      for (const el of scrollerRef.current.querySelectorAll('.pdf-page')) {
        const r = el.getBoundingClientRect();
        if (r.bottom < box.top || r.top > box.bottom) continue;
        if (r.top <= mid && r.bottom >= mid) { best = el; break; }
        if (!best) best = el;
      }
      if (best) onCurrent(best.dataset.key, { fromScroll: true });
    });
  }, [layout, onCurrent, scrollerRef]);

  useLayoutEffect(() => { onScroll(); }, [zoom]); // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <div className={`pdf-view layout-${layout}`} ref={scrollerRef} onScroll={onScroll} dir="ltr" onMouseDown={onFocusText}>
      <div className="pdf-view-inner">
        {rows.map((row) => (
          <div className="pdf-row" key={row.map((p) => p.key).join('-')}>
            {row.map((p) => (
              <PageBox
                key={p.key}
                entry={p}
                index={pages.indexOf(p)}
                source={sources[p.source]}
                zoom={zoom}
                root={scrollerRef}
                onText={onText}
              />
            ))}
          </div>
        ))}
      </div>
    </div>
  );
}
