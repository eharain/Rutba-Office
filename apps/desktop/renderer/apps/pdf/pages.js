// The PDF window's document: a list of pages, each naming the file it comes
// from, its page there and how it is turned. Every page operation is a new
// list made from the old one, so undo and redo are the lists before and
// after, and nothing is drawn again to rearrange anything.
//
// Pure functions, tested in tests/pdf-window.test.js.

let counter = 0;
/** A key for a page of the list, unique in this window. */
export const pageKey = () => `pg${(counter += 1).toString(36)}`;

/** The pages of one file as list entries. */
export function pagesOf(source, info) {
  return info.pages.map((p, page) => ({ key: pageKey(), source, page, rotate: p.rotate || 0 }));
}

const turn = (deg) => (((Math.round(deg / 90) * 90) % 360) + 360) % 360;

/** The pages named turned by `by` degrees (90 is a right turn). */
export function rotatePages(pages, keys, by) {
  const set = new Set(keys);
  return pages.map((p) => (set.has(p.key) ? { ...p, rotate: turn(p.rotate + by) } : p));
}

/** The list without the pages named. */
export function deletePages(pages, keys) {
  const set = new Set(keys);
  return pages.filter((p) => !set.has(p.key));
}

/** `items` put in before index `at` (0 is the front, pages.length the end). */
export function insertPages(pages, at, items) {
  const where = Math.max(0, Math.min(pages.length, at));
  return [...pages.slice(0, where), ...items, ...pages.slice(where)];
}

/**
 * The pages named moved together, in their order, to land before the page
 * that stood at index `to` in the list as it was (pages.length: the end).
 */
export function movePages(pages, keys, to) {
  const set = new Set(keys);
  const moving = pages.filter((p) => set.has(p.key));
  if (!moving.length) return pages;
  const before = pages.slice(0, Math.max(0, Math.min(pages.length, to))).filter((p) => !set.has(p.key)).length;
  const rest = pages.filter((p) => !set.has(p.key));
  const next = [...rest.slice(0, before), ...moving, ...rest.slice(before)];
  return next.every((p, i) => p === pages[i]) ? pages : next;
}

/** The keys from `anchor` to `key` inclusive, in list order (Shift and a click). */
export function rangeKeys(pages, anchor, key) {
  const a = pages.findIndex((p) => p.key === anchor);
  const b = pages.findIndex((p) => p.key === key);
  if (a < 0 || b < 0) return b < 0 ? [] : [key];
  const [from, to] = a <= b ? [a, b] : [b, a];
  return pages.slice(from, to + 1).map((p) => p.key);
}

/** Page numbers (1-based, sorted) written short: "1-3, 5, 8-9". */
export function rangeText(numbers) {
  const sorted = [...new Set(numbers)].sort((a, b) => a - b);
  const out = [];
  for (let i = 0; i < sorted.length;) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] + 1) j += 1;
    out.push(j > i ? `${sorted[i]}-${sorted[j]}` : String(sorted[i]));
    i = j + 1;
  }
  return out.join(', ');
}

/** The selection's pages in list order; the page in view when nothing is selected. */
export function chosenKeys(pages, selected, current) {
  const set = new Set(selected);
  const keys = pages.filter((p) => set.has(p.key)).map((p) => p.key);
  if (keys.length) return keys;
  return current && pages.some((p) => p.key === current) ? [current] : [];
}

/** The size a page is shown at, in points, once turned. */
export function shownSize(info, rotate) {
  const w = info?.width || 612;
  const h = info?.height || 792;
  return rotate % 180 ? { width: h, height: w } : { width: w, height: h };
}

/**
 * Where a split puts each page: files of `every` pages, or a new file
 * starting at each page whose index is in `starts`.
 */
export function splitGroups(count, { every = null, starts = null } = {}) {
  if (starts && starts.length) {
    const cuts = [...new Set(starts.filter((n) => n > 0 && n < count))].sort((a, b) => a - b);
    const out = [];
    let from = 0;
    for (const cut of [...cuts, count]) {
      if (cut > from) out.push(Array.from({ length: cut - from }, (_, k) => from + k));
      from = cut;
    }
    return out;
  }
  const size = Math.max(1, Math.floor(Number(every) || 1));
  const out = [];
  for (let i = 0; i < count; i += size) out.push(Array.from({ length: Math.min(size, count - i) }, (_, k) => i + k));
  return out;
}
