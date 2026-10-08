// Pictures: ratings and tags, and the folders' counts.
//
// A rating (one to five stars) and tags are kept on this computer, by the
// file's path — never written into the picture, which a viewer has no
// business changing — and the grid filters and sorts by them; the name
// filter finds a tag too. A folder's tile says how many pictures, clips and
// documents it holds before it is entered, read once and kept for the
// window's life.

import { useCallback, useEffect, useRef, useState } from 'react';

const KEY = 'pictures.marks';

/** The ratings and tags, and the means to change them: `{ marks, rate(path, n), tag(path, tags) }`. */
export function useMarks(shell) {
  const [marks, setMarks] = useState({});
  const saveTimer = useRef(null);
  useEffect(() => {
    let live = true;
    shell.store.get({ key: KEY, fallback: {} }).then((m) => { if (live && m && typeof m === 'object') setMarks(m); }).catch(() => {});
    return () => { live = false; };
  }, [shell]);
  const change = useCallback((path, patch) => {
    setMarks((all) => {
      const was = all[path] || {};
      const next = { ...was, ...patch };
      if (!next.rating) delete next.rating;
      if (!next.tags?.length) delete next.tags;
      const out = { ...all };
      if (Object.keys(next).length) out[path] = next;
      else delete out[path];
      clearTimeout(saveTimer.current);
      saveTimer.current = setTimeout(() => shell.store.set({ key: KEY, value: out }).catch(() => {}), 300);
      return out;
    });
  }, [shell]);
  const rate = useCallback((path, rating) => change(path, { rating: Math.max(0, Math.min(5, Number(rating) || 0)) }), [change]);
  const tag = useCallback((path, tags) => change(path, { tags: [...new Set((tags || []).map((t) => String(t).trim()).filter(Boolean))] }), [change]);
  return { marks, rate, tag };
}

const counted = new Map();

/** How many viewable files a folder holds, read once: a number, or null while it is being read. */
export function useFolderCount(shell, path, filter) {
  const [n, setN] = useState(() => (counted.has(path) ? counted.get(path) : null));
  useEffect(() => {
    if (counted.has(path)) { setN(counted.get(path)); return undefined; }
    let live = true;
    shell.fs.list({ path, filter }).then((rows) => {
      const count = (rows || []).filter((r) => !r.dir).length;
      counted.set(path, count);
      if (live) setN(count);
    }).catch(() => { counted.set(path, 0); if (live) setN(0); });
    return () => { live = false; };
  }, [shell, path, filter]);
  return n;
}
