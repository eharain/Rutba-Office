// The pages of a picture that has more than one: a fax, a scan of a
// letter, a multi-page TIFF. The stage shows one page at a time and a strip
// under it shows them all, as the Windows Photos and the old Windows Fax and
// Picture Viewer did; Page Up and Page Down step through them.
//
// The pages arrive over rutba://file, which decodes a TIFF on the main
// process's picture thread (packages/office-shell/src/electron/protocol.js):
// `?info` says how many there are, `?page=N` is a PNG of one, and
// `&fit=` a small one for the strip.

import React, { useEffect, useRef, useState } from 'react';
import { Button, t } from '@rutba/office-ui';
import { fileUrl } from './library.js';

/** Files whose pages the strip shows. */
export const PAGED = /\.(tiff?)$/i;

/** The address of one page of a file: page 0 is the file's own address, which a window may already hold. */
export function pageUrl(url, page, { fit = 0 } = {}) {
  if (!page && !fit) return url;
  const join = url.includes('?') ? '&' : '?';
  return `${url}${join}page=${page}${fit ? `&fit=${fit}` : ''}`;
}

/** What rutba://file says of a file it decodes: `{ pages, width, height }`, or `{ error }`. */
export async function pictureInfo(path) {
  try {
    const res = await fetch(`${fileUrl(path)}?info=1`);
    return await res.json();
  } catch (err) {
    return { error: err?.message || String(err) };
  }
}

/**
 * The page shown of the file on the stage, and how many it has. Back to
 * the first page whenever the file changes.
 */
export function usePages(current) {
  const [state, setState] = useState({ path: null, pages: 1, page: 0 });
  useEffect(() => {
    if (!current || !PAGED.test(current)) {
      setState({ path: current, pages: 1, page: 0 });
      return undefined;
    }
    let alive = true;
    setState({ path: current, pages: 1, page: 0 });
    pictureInfo(current).then((info) => {
      if (alive && info?.pages > 1) setState((s) => (s.path === current ? { ...s, pages: info.pages } : s));
    });
    return () => {
      alive = false;
    };
  }, [current]);
  const mine = state.path === current;
  const pages = mine ? state.pages : 1;
  const page = mine ? Math.min(state.page, pages - 1) : 0;
  const setPage = (next) => setState((s) => (s.path === current ? { ...s, page: Math.max(0, Math.min(s.pages - 1, typeof next === 'function' ? next(s.page) : next)) } : s));
  return { pages, page, setPage };
}

const THUMB_FIT = 128;

/** One page in the strip: drawn small, its number under it. */
function PageThumb({ path, page, active, onPick }) {
  const ref = useRef(null);
  useEffect(() => {
    if (active) ref.current?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
  }, [active]);
  return (
    <button
      ref={ref}
      type="button"
      className={`pv-page${active ? ' active' : ''}`}
      data-page={page}
      aria-label={t('Page {n}', { n: page + 1 })}
      aria-pressed={active}
      onClick={() => onPick(page)}
    >
      <span className="pv-page-pic">
        <img src={pageUrl(fileUrl(path), page, { fit: THUMB_FIT })} alt="" loading="lazy" decoding="async" draggable={false} />
      </span>
      <span className="pv-page-n">{page + 1}</span>
    </button>
  );
}

/** The strip of pages under the stage, with the page number and steps either way. */
export function PageStrip({ path, pages, page, setPage }) {
  if (!path || pages < 2) return null;
  return (
    <div className="pv-pages" role="toolbar" aria-label={t('Pages')}>
      <div className="pv-pages-step">
        <Button icon="chevronLeft" title={t('Previous page (Page Up)')} disabled={page <= 0} onClick={() => setPage((p) => p - 1)} />
        <span className="pv-pages-count" data-role="page-count">{t('Page {n} of {count}', { n: page + 1, count: pages })}</span>
        <Button icon="chevronRight" title={t('Next page (Page Down)')} disabled={page >= pages - 1} onClick={() => setPage((p) => p + 1)} />
      </div>
      <div className="pv-pages-list">
        {Array.from({ length: pages }, (_, i) => (
          <PageThumb key={i} path={path} page={i} active={i === page} onPick={setPage} />
        ))}
      </div>
    </div>
  );
}

export const PAGES_CSS = `
.pv-pages {
  display: flex; align-items: center; gap: 12px; padding: 6px 10px 7px;
  background: #1d2026; border-top: 1px solid #2a2e36; color: #d6dae0; min-width: 0;
}
.pv-pages-step { display: flex; align-items: center; gap: 2px; flex: none; }
.pv-pages-step .rw-btn { color: #d6dae0; }
.pv-pages-step .rw-btn:hover:not(:disabled) { background: rgba(255, 255, 255, 0.08); }
.pv-pages-count { font-size: 12px; font-variant-numeric: tabular-nums; min-width: 92px; text-align: center; white-space: nowrap; }
.pv-pages-list { display: flex; gap: 8px; overflow-x: auto; overflow-y: hidden; min-width: 0; flex: 1; padding: 2px; scrollbar-width: thin; }
.pv-page {
  flex: none; display: flex; flex-direction: column; align-items: center; gap: 3px;
  border: 0; background: none; padding: 0; cursor: pointer; color: inherit;
}
.pv-page-pic {
  width: 56px; height: 64px; display: grid; place-items: center; border-radius: 3px;
  background: #2a2e36; box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.06); overflow: hidden;
  transition: box-shadow var(--fast);
}
.pv-page-pic img { max-width: 52px; max-height: 60px; display: block; background: #fff; box-shadow: 0 1px 2px rgba(0, 0, 0, 0.5); }
.pv-page:hover .pv-page-pic { box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.28); }
.pv-page.active .pv-page-pic { box-shadow: inset 0 0 0 2px var(--accent), 0 0 0 2px var(--accent-soft); }
.pv-page-n { font-size: 10.5px; font-variant-numeric: tabular-nums; opacity: 0.8; }
.pv-page.active .pv-page-n { opacity: 1; font-weight: 600; }
`;
