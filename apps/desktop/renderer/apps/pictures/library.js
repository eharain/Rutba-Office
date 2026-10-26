// The viewer's arithmetic, with no window in it.
//
// What a folder holds, in what order, and which of it is on screen: the grid
// and the filmstrip draw only the tiles a person can see, so a folder of five
// thousand files costs the same to draw as a folder of fifty. Everything here
// is a pure function of numbers, which is what lets the engine suite hold it
// to account without opening a window.

import { t } from '@rutba/office-ui/messages';

/** Everything the viewer will open. */
export const STILL = ['.jpg', '.jpeg', '.jpe', '.png', '.bmp', '.tif', '.tiff', '.ico', '.svg', '.avif', '.heic', '.heif'];
export const ANIMATED = ['.gif', '.webp', '.apng'];
export const VIDEO = ['.mp4', '.m4v', '.mov', '.webm', '.mkv', '.ogv', '.avi'];
export const AUDIO = ['.mp3', '.m4a', '.wav', '.flac', '.ogg', '.opus'];
export const VIEWABLE = [...STILL, ...ANIMATED, ...VIDEO, ...AUDIO, '.pdf'];

const extOf = (p) => {
  const name = String(p || '').split(/[\\/]/).pop() || '';
  const dot = name.lastIndexOf('.');
  return dot > 0 ? name.slice(dot).toLowerCase() : '';
};

/** 'still' | 'maybe-animated' | 'video' | 'audio' | 'pdf' */
export function kindOf(file) {
  const ext = extOf(file);
  if (VIDEO.includes(ext)) return 'video';
  if (AUDIO.includes(ext)) return 'audio';
  if (ext === '.pdf') return 'pdf';
  if (ANIMATED.includes(ext)) return 'maybe-animated';
  return 'still';
}

/** The filter chips: a family each, with what it covers. */
export const FAMILIES = [
  { id: 'all', label: t('All') },
  { id: 'pictures', label: t('Pictures'), kinds: ['still', 'maybe-animated'] },
  { id: 'video', label: t('Videos'), kinds: ['video'] },
  { id: 'audio', label: t('Audio'), kinds: ['audio'] },
  { id: 'pdf', label: 'PDF', kinds: ['pdf'] }, // words-ok: a format's name
];

export function familyOf(file) {
  const kind = kindOf(file);
  return FAMILIES.find((f) => f.kinds?.includes(kind))?.id || 'pictures';
}

/** How many of each family a folder holds, `all` included. */
export function countByFamily(entries) {
  const counts = { all: 0 };
  for (const f of FAMILIES) if (f.id !== 'all') counts[f.id] = 0;
  for (const e of entries) {
    if (e.dir) continue;
    counts.all += 1;
    counts[familyOf(e.path)] += 1;
  }
  return counts;
}

export const SORTS = [
  { id: 'name', label: t('Name') },
  { id: 'newest', label: t('Newest first') },
  { id: 'oldest', label: t('Oldest first') },
  { id: 'largest', label: t('Largest first') },
  { id: 'kind', label: t('Kind') },
  { id: 'rating', label: t('Highest rated') },
];

/** The rating filter: any, or at least so many stars. */
export const RATED = [[0, t('Any rating')], [1, t('★ and up')], [2, t('★★ and up')], [3, t('★★★ and up')], [4, t('★★★★ and up')], [5, '★★★★★']];

const byName = (a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
const KIND_ORDER = ['still', 'maybe-animated', 'video', 'audio', 'pdf'];

/**
 * The files of a folder as the grid shows them: no directories, the family
 * asked for, the name filter, in the order chosen. Ties fall back to the
 * name, so an order is the same order on every visit.
 */
export function arrange(entries, { query = '', family = 'all', sort = 'name', marks = null, rated = 0 } = {}) {
  const q = query.trim().toLowerCase();
  const fam = FAMILIES.find((f) => f.id === family);
  const markOf = (e) => (marks && marks[e.path]) || null;
  let rows = entries.filter((e) => !e.dir);
  if (fam?.kinds) rows = rows.filter((e) => fam.kinds.includes(kindOf(e.path)));
  // The filter finds a name, or a tag given to the file here.
  if (q) rows = rows.filter((e) => e.name.toLowerCase().includes(q) || (markOf(e)?.tags || []).some((t) => t.toLowerCase().includes(q)));
  if (rated > 0) rows = rows.filter((e) => (markOf(e)?.rating || 0) >= rated);
  const cmp =
    sort === 'newest' ? (a, b) => (b.mtime || 0) - (a.mtime || 0) || byName(a, b)
    : sort === 'oldest' ? (a, b) => (a.mtime || 0) - (b.mtime || 0) || byName(a, b)
    : sort === 'largest' ? (a, b) => (b.size || 0) - (a.size || 0) || byName(a, b)
    : sort === 'kind' ? (a, b) => KIND_ORDER.indexOf(kindOf(a.path)) - KIND_ORDER.indexOf(kindOf(b.path)) || byName(a, b)
    : sort === 'rating' ? (a, b) => (markOf(b)?.rating || 0) - (markOf(a)?.rating || 0) || byName(a, b)
    : byName;
  return rows.slice().sort(cmp);
}

/**
 * Search the folders inside one: every file under `root` — through
 * `list(dir)`, a folder's entries — whose name or tag holds `query`, as
 * far down as `depth` and no more than `limit` of them, nearest first.
 */
export async function searchTree(list, root, { query = '', marks = null, depth = 6, limit = 2000, filter = null } = {}) {
  const q = query.trim().toLowerCase();
  if (!q) return [];
  const out = [];
  let level = [root];
  for (let d = 0; d <= depth && level.length && out.length < limit; d++) {
    const next = [];
    for (const dir of level) {
      let rows = [];
      try { rows = await list(dir); } catch { rows = []; }
      for (const e of rows) {
        if (e.dir) { next.push(e.path); continue; }
        if (filter && !filter(e)) continue;
        const tags = (marks && marks[e.path]?.tags) || [];
        if (e.name.toLowerCase().includes(q) || tags.some((t) => t.toLowerCase().includes(q))) out.push(e);
        if (out.length >= limit) break;
      }
      if (out.length >= limit) break;
    }
    level = next;
  }
  return out;
}

/**
 * What a file's ratings and tags know it by besides its path: its size and
 * the time it was last changed, which renaming or moving it in Explorer
 * leaves as they were. Null for a folder, or a row without them.
 */
export function fingerprintOf(row) {
  if (!row || row.dir || !Number.isFinite(Number(row.size)) || !Number.isFinite(Number(row.mtime)) || row.size == null || row.mtime == null) return null;
  return `${Number(row.size)}:${Math.round(Number(row.mtime))}`;
}

const kept = (m) => Boolean(m && (m.rating || m.tags?.length));

/**
 * Ratings and tags that follow a file renamed or moved outside the window.
 * Marks are kept by path, so each also keeps the file's fingerprint; a
 * folder listed answers:
 *   `fill`:  `{ path: fingerprint }` for each marked file listed whose
 *             fingerprint is not kept yet, or has changed (the file edited);
 *   `moves`: `[{ from, to }]`: a file listed with no marks of its own whose
 *             fingerprint is that of exactly one marked file that is not
 *             listed here, and no other file listed shares it.
 * A move is only a candidate: the caller takes it when `from` is no longer
 * there, so a copy leaves its original's marks where they are.
 */
export function followMarks(marks, rows) {
  const fill = {};
  const moves = [];
  if (!marks || !rows?.length) return { fill, moves };
  const listed = new Set(rows.map((r) => r.path));
  const here = new Map();
  for (const r of rows) {
    const fp = fingerprintOf(r);
    if (!fp) continue;
    here.set(fp, (here.get(fp) || 0) + 1);
    if (kept(marks[r.path]) && marks[r.path].fp !== fp) fill[r.path] = fp;
  }
  const away = new Map();
  for (const [p, m] of Object.entries(marks)) {
    if (!kept(m) || !m.fp || listed.has(p)) continue;
    away.set(m.fp, away.has(m.fp) ? null : p);
  }
  for (const r of rows) {
    const fp = fingerprintOf(r);
    if (!fp || kept(marks[r.path]) || here.get(fp) !== 1) continue;
    const from = away.get(fp);
    if (from) moves.push({ from, to: r.path });
  }
  return { fill, moves };
}

/** The folders inside a folder, by name, for the tiles that lead into them. */
export function foldersOf(entries) {
  return entries.filter((e) => e.dir).slice().sort(byName);
}

/**
 * The tile sizes on offer, in CSS pixels: `px` beside an open picture,
 * `browsePx` when the grid has the whole window. The label is what the
 * button says.
 */
export const TILE_SIZES = [
  { id: 's', label: t('Small'), px: 84, browsePx: 124 },
  { id: 'm', label: t('Medium'), px: 116, browsePx: 172 },
  { id: 'l', label: t('Large'), px: 168, browsePx: 248 },
];

/**
 * How many columns of `tile`-wide tiles fit `width`, to the nearest — a
 * tile is then stretched or squeezed a little to fill the row exactly,
 * rather than leaving a gutter down one side.
 */
export function fitColumns({ width, tile, gap = 6, padding = 0 }) {
  const inner = Math.max(0, width - padding * 2);
  const columns = Math.max(1, Math.round((inner + gap) / (tile + gap)));
  return { columns, tile: Math.max(24, Math.floor((inner - (columns - 1) * gap) / columns)) };
}

/**
 * Which tiles of a grid are in view, and how much room to leave for the rest.
 *
 * The grid is rows of `columns` square tiles of `tile` px with `gap` between
 * them, inside a scroller `height` px tall scrolled to `scrollTop`. Only the
 * rows in view and `overscan` rows either side are drawn; the space of the
 * rows above and below is held by two empty blocks, so the scrollbar is the
 * scrollbar of the whole folder.
 */
export function gridWindow({ count, width, tile, gap = 6, scrollTop = 0, height = 0, overscan = 2, padding = 0, columns: fixed = 0 }) {
  const inner = Math.max(0, width - padding * 2);
  const columns = fixed > 0 ? fixed : Math.max(1, Math.floor((inner + gap) / (tile + gap)));
  const rowHeight = tile + gap;
  const rows = Math.ceil(count / columns);
  if (!count) return { columns, rowHeight, rows: 0, first: 0, last: 0, top: 0, bottom: 0 };
  const firstRow = Math.max(0, Math.floor(scrollTop / rowHeight) - overscan);
  const lastRow = Math.min(rows, Math.ceil((scrollTop + Math.max(height, rowHeight)) / rowHeight) + overscan);
  return {
    columns,
    rowHeight,
    rows,
    first: firstRow * columns,
    last: Math.min(count, lastRow * columns),
    top: firstRow * rowHeight,
    bottom: Math.max(0, (rows - lastRow) * rowHeight),
  };
}

/** The same for a strip of `count` items `itemWidth` wide, scrolled sideways. */
export function stripWindow({ count, itemWidth, gap = 5, scrollLeft = 0, width = 0, overscan = 6 }) {
  const step = itemWidth + gap;
  if (!count) return { first: 0, last: 0, left: 0, right: 0, step };
  const first = Math.max(0, Math.floor(scrollLeft / step) - overscan);
  const last = Math.min(count, Math.ceil((scrollLeft + Math.max(width, step)) / step) + overscan);
  return { first, last, left: first * step, right: Math.max(0, (count - last) * step), step };
}

/** Where a strip must scroll to put item `index` in the middle. */
export function stripCentre({ index, itemWidth, gap = 5, width }) {
  return Math.max(0, index * (itemWidth + gap) + itemWidth / 2 - width / 2);
}

/**
 * The files worth fetching before they are asked for: the next ones first,
 * then the previous, `radius` deep, wrapping when the folder loops. The one
 * on screen is never in the list.
 */
export function neighbours(index, count, { radius = 2, loop = true } = {}) {
  if (count < 2 || index < 0) return [];
  const out = [];
  for (let d = 1; d <= radius; d++) {
    for (const i of [index + d, index - d]) {
      const j = loop ? ((i % count) + count) % count : i;
      if (j < 0 || j >= count || j === index || out.includes(j)) continue;
      out.push(j);
    }
  }
  return out;
}

// The addresses of a file and of its thumbnail are the core's (file-urls.js): Worksheets uses them too.
export { fileUrl, thumbUrl } from '../../file-urls.js';
