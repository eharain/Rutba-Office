// Pictures: ratings and tags, and the folders' counts.
//
// A rating (one to five stars) and tags are kept on this computer, by the
// file's path (never written into the picture, which a viewer has no
// business changing), and follow the file when it is renamed or moved
// (its size and time kept with them); the grid filters and sorts by them; the name
// filter finds a tag too. A folder's tile says how many pictures, clips and
// documents it holds before it is entered, read once and kept for the
// window's life.

import { useCallback, useEffect, useRef, useState } from 'react';
import { fingerprintOf, followMarks } from './library.js';

const KEY = 'pictures.marks';

/**
 * The ratings and tags, and the means to change them: `{ marks, rate(path,
 * n), tag(path, tags), follow(rows) }`. Each is kept with the file's
 * fingerprint (library.js), so `follow`, called with every folder listed,
 * finds a marked file renamed or moved outside the window and moves its
 * marks to the new path.
 */
export function useMarks(shell) {
  const [marks, setMarks] = useState({});
  const saveTimer = useRef(null);
  const latest = useRef({});
  // A listing that came before the kept marks were read waits for them.
  const loaded = useRef(null);
  if (!loaded.current) { let done; loaded.current = { promise: new Promise((r) => { done = r; }), done }; }
  useEffect(() => {
    let live = true;
    shell.store.get({ key: KEY, fallback: {} })
      .then((m) => { if (live && m && typeof m === 'object') { latest.current = m; setMarks(m); } })
      .catch(() => {})
      .finally(() => loaded.current.done());
    return () => { live = false; };
  }, [shell]);
  const update = useCallback((fn) => {
    setMarks((all) => {
      const out = fn(all);
      if (out === all) return all;
      latest.current = out;
      clearTimeout(saveTimer.current);
      saveTimer.current = setTimeout(() => shell.store.set({ key: KEY, value: out }).catch(() => {}), 300);
      return out;
    });
  }, [shell]);
  const change = useCallback((path, patch) => {
    update((all) => {
      const was = all[path] || {};
      const next = { ...was, ...patch };
      if (!next.rating) delete next.rating;
      if (!next.tags?.length) delete next.tags;
      const out = { ...all };
      if (next.rating || next.tags) out[path] = next;
      else delete out[path];
      return out;
    });
    // The fingerprint the marks will know the file by, should it be renamed or moved.
    shell.fs.stat({ path }).then((s) => {
      const fp = fingerprintOf(s);
      if (fp) update((all) => (all[path] && all[path].fp !== fp ? { ...all, [path]: { ...all[path], fp } } : all));
    }).catch(() => {});
  }, [shell, update]);
  const rate = useCallback((path, rating) => change(path, { rating: Math.max(0, Math.min(5, Number(rating) || 0)) }), [change]);
  const tag = useCallback((path, tags) => change(path, { tags: [...new Set((tags || []).map((t) => String(t).trim()).filter(Boolean))] }), [change]);
  const follow = useCallback(async (rows) => {
    await loaded.current.promise;
    const { fill, moves } = followMarks(latest.current, rows);
    const taken = [];
    for (const m of moves) {
      // Still there: a copy, not a move; the original keeps its marks.
      if (!(await shell.fs.exists({ path: m.from }).catch(() => true))) taken.push(m);
    }
    if (!taken.length && !Object.keys(fill).length) return;
    update((all) => {
      const out = { ...all };
      for (const [p, fp] of Object.entries(fill)) if (out[p]) out[p] = { ...out[p], fp };
      for (const { from, to } of taken) {
        if (!out[from] || out[to]) continue;
        out[to] = out[from];
        delete out[from];
      }
      return out;
    });
  }, [shell, update]);
  return { marks, rate, tag, follow };
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
