// Layout of our own from word boxes (whatever read them) and ruling lines:
// lines of words, the columns of a page, and a table's rows, columns and
// cells. Deterministic geometry; the engine tests prove each case.
//
// A word is { text, box: { x, y, width, height } }; rulings are what
// analyse.js `rulings` finds.

const centreY = (b) => b.y + b.height / 2;
const centreX = (b) => b.x + b.width / 2;
const median = (xs) => {
  if (!xs.length) return 0;
  const s = [...xs].sort((a, b) => a - b);
  return s[s.length >> 1];
};

/**
 * Words gathered into lines: a word joins the line whose vertical extent
 * it overlaps by more than half of the smaller height. Lines run top to
 * bottom; the words of each left to right, or right to left when `rtl`.
 * @returns {{ box, words: object[] }[]}
 */
export function groupLines(words, { rtl = false } = {}) {
  const lines = [];
  for (const w of [...words].sort((a, b) => centreY(a.box) - centreY(b.box))) {
    const b = w.box;
    let best = null, bestOverlap = 0;
    for (const line of lines) {
      const top = Math.max(line.y0, b.y), bottom = Math.min(line.y1, b.y + b.height);
      const overlap = (bottom - top) / Math.max(1, Math.min(line.y1 - line.y0, b.height));
      if (overlap > 0.5 && overlap > bestOverlap) { best = line; bestOverlap = overlap; }
    }
    if (best) {
      best.words.push(w);
      best.y0 = Math.min(best.y0, b.y); best.y1 = Math.max(best.y1, b.y + b.height);
    } else lines.push({ y0: b.y, y1: b.y + b.height, words: [w] });
  }
  return lines
    .sort((a, b) => a.y0 - b.y0)
    .map((line) => {
      const ws = line.words.sort((a, b) => (rtl ? b.box.x - a.box.x : a.box.x - b.box.x));
      const x0 = Math.min(...ws.map((w) => w.box.x)), x1 = Math.max(...ws.map((w) => w.box.x + w.box.width));
      return { box: { x: x0, y: line.y0, width: x1 - x0, height: line.y1 - line.y0 }, words: ws };
    });
}

/**
 * Coverage of x by word boxes: for each pixel column of [x0, x1), how many
 * lines have a word over it.
 */
function coverage(lines, x0, x1) {
  const cov = new Uint16Array(Math.max(0, x1 - x0));
  for (const line of lines) {
    const seen = new Uint8Array(cov.length);
    for (const w of line.words) {
      for (let x = Math.max(x0, w.box.x); x < Math.min(x1, w.box.x + w.box.width); x++) seen[x - x0] = 1;
    }
    for (let i = 0; i < cov.length; i++) cov[i] += seen[i];
  }
  return cov;
}

/** The runs of x where coverage is at most `most`, at least `wide` long, inside [x0, x1). */
function gaps(cov, x0, most, wide) {
  const out = [];
  let i = 0;
  while (i < cov.length) {
    if (cov[i] > most) { i++; continue; }
    const s = i;
    while (i < cov.length && cov[i] <= most) i++;
    if (i - s >= wide) out.push({ x0: x0 + s, x1: x0 + i });
  }
  return out;
}

/**
 * The columns of a page of text: vertical strips of white that no line
 * crosses (a word or two may, as a heading does), at least two text
 * heights wide, running the height of the text.
 * @returns {{ x0, x1 }[]} the columns, left to right
 */
export function columns(words) {
  if (!words.length) return [];
  const lines = groupLines(words);
  const x0 = Math.min(...words.map((w) => w.box.x)), x1 = Math.max(...words.map((w) => w.box.x + w.box.width));
  const h = median(words.map((w) => w.box.height)) || 10;
  const cov = coverage(lines, x0, x1);
  const allowed = Math.floor(lines.length * 0.08);
  // Gutters lie inside the body of the text: white beyond its edges, under a heading wider than the body, is margin.
  let first = 0, last = cov.length - 1;
  while (first < cov.length && cov[first] <= allowed) first++;
  while (last > first && cov[last] <= allowed) last--;
  const splits = gaps(cov.subarray(first, last + 1), x0 + first, allowed, Math.max(4, Math.round(h * 2)));
  const out = [];
  let from = x0;
  for (const g of splits) {
    out.push({ x0: from, x1: g.x0 });
    from = g.x1;
  }
  out.push({ x0: from, x1 });
  return out;
}

/** Values that lie within `near` of each other, merged into their mean. */
function cluster(values, near) {
  const s = [...values].sort((a, b) => a - b);
  const out = [];
  for (const v of s) {
    const last = out[out.length - 1];
    if (last && v - last.max <= near) { last.sum += v; last.n++; last.max = v; }
    else out.push({ sum: v, n: 1, max: v });
  }
  return out.map((c) => c.sum / c.n);
}

/**
 * A table from the words of a picture of one and the ruling lines drawn in
 * it. Where the rules make a grid (two or more of each way), its cells are
 * the boxes between them; where they do not (a table set with space alone),
 * rows are the lines of words and columns the white strips the rows leave,
 * as `columns` finds for a page but row by row. Rules one way only are rows
 * or columns when they hold every word and (for rows) one line between
 * each two, as a table ruled under each row is; rules under the heading and
 * the last row alone are not, and the lines are the rows. A line or two
 * across the gutters (a title, a note, a long name) do not join the columns
 * either side; a column with a word in only a row or two is kept. Each word
 * lands in the cell its centre falls in; a cell's words are joined in
 * reading order.
 *
 * @param {object[]} words
 * @param {{ horizontal: object[], vertical: object[] }} [rules]
 * @param {{ rtl?: boolean }} [options] right-to-left: the first column is the rightmost
 * @returns {{ rows: { y0, y1 }[], columns: { x0, x1 }[], cells: string[][], confidence: (number|null)[][], ruled: boolean }}
 *   `confidence`: each cell's least sure word (0 to 100), null for an empty cell
 */
export function tableCells(words, rules = { horizontal: [], vertical: [] }, { rtl = false } = {}) {
  if (!words.length) return { rows: [], columns: [], cells: [], confidence: [], ruled: false };
  const h = median(words.map((w) => w.box.height)) || 10;
  const left = Math.min(...words.map((w) => w.box.x)), right = Math.max(...words.map((w) => w.box.x + w.box.width));
  const top = Math.min(...words.map((w) => w.box.y)), bottom = Math.max(...words.map((w) => w.box.y + w.box.height));

  let rowEdges = null, colEdges = null;
  const hy = cluster((rules.horizontal || []).map((r) => r.y), Math.max(3, h / 3));
  const vx = cluster((rules.vertical || []).map((r) => r.x), Math.max(3, h / 3));
  const ruled = hy.length >= 2 && vx.length >= 2;

  // A ruled table's rules can be read as strokes of their own (a bar, a dash): they are not words.
  if (ruled) words = words.filter((w) => !/^[|│¦_\-\u2014\u2013=]+$/.test(w.text));
  if (!words.length) return { rows: [], columns: [], cells: [], confidence: [], ruled };
  const lines = groupLines(words, { rtl });
  if (ruled) {
    rowEdges = hy;
    colEdges = vx;
  } else {
    // Rules one way only: the rows (or columns) they make when every word is
    // between the first and the last, and (rows) no two lines share a band.
    const ys = lines.map((l) => centreY(l.box));
    if (hy.length >= 2 && ys.every((y) => y >= hy[0] && y < hy[hy.length - 1])
      && hy.slice(1).every((y1, i) => ys.filter((y) => y >= hy[i] && y < y1).length <= 1)) rowEdges = hy;
    if (vx.length >= 2 && words.every((w) => centreX(w.box) >= vx[0] && centreX(w.box) < vx[vx.length - 1])) colEdges = vx;
  }
  if (!rowEdges) {
    // Rows from the lines of words: a boundary halfway between each line and the next.
    const bands = lines.map((l) => ({ y0: l.box.y, y1: l.box.y + l.box.height }));
    rowEdges = [Math.min(top, bands[0].y0) - 1];
    for (let i = 1; i < bands.length; i++) rowEdges.push((bands[i - 1].y1 + bands[i].y0) / 2);
    rowEdges.push(bottom + 1);
  }
  if (!colEdges) {
    // Columns from the white the rows leave. A quarter of the lines may cross
    // a gutter (from four lines up), as a title or a long name does. Inside a
    // strip that few lines cross, the gutter is the white of its own least
    // crossed runs: two such runs either side of a column with a word in
    // only a row or two leave that column standing. White that reaches the
    // table's edge (under a title wider than the table) is margin.
    const cov = coverage(lines, left, right);
    const most = lines.length >= 4 ? Math.floor(lines.length / 4) : 0;
    const wide = Math.max(3, Math.round(h * 0.8));
    const splits = [];
    for (const g of gaps(cov, left, most, wide)) {
      const strip = cov.subarray(g.x0 - left, g.x1 - left);
      const low = Math.min(...strip);
      const runs = gaps(strip, g.x0, low, Math.max(2, Math.round(wide / 2))).filter((r) => r.x0 > left && r.x1 < right);
      for (const r of runs) splits.push((r.x0 + r.x1) / 2);
    }
    colEdges = [left - 1, ...splits, right + 1];
  }

  const rows = [];
  for (let i = 1; i < rowEdges.length; i++) rows.push({ y0: rowEdges[i - 1], y1: rowEdges[i] });
  let cols = [];
  for (let i = 1; i < colEdges.length; i++) cols.push({ x0: colEdges[i - 1], x1: colEdges[i] });
  // Ruled rows and columns with nothing between their rules at all (a double rule) are not cells.
  const holds = (r, c) => words.some((w) => {
    const cx = centreX(w.box), cy = centreY(w.box);
    return (!r || (cy >= r.y0 && cy < r.y1)) && (!c || (cx >= c.x0 && cx < c.x1));
  });
  const keepRows = rows.filter((r) => r.y1 - r.y0 > h * 0.6 || holds(r, null));
  cols = cols.filter((c) => c.x1 - c.x0 > h * 0.6 || holds(null, c));
  if (rtl) cols.reverse();

  const cells = keepRows.map(() => cols.map(() => []));
  for (const line of lines) {
    for (const w of line.words) {
      const cx = centreX(w.box), cy = centreY(w.box);
      const r = keepRows.findIndex((row) => cy >= row.y0 && cy < row.y1);
      const c = cols.findIndex((col) => cx >= col.x0 && cx < col.x1);
      if (r >= 0 && c >= 0) cells[r][c].push(w);
    }
  }
  const text = cells.map((row) => row.map((ws) => groupLines(ws, { rtl }).map((l) => l.words.map((w) => w.text).join(' ')).join(' ')));
  const sure = cells.map((row) => row.map((ws) => (ws.length && ws.every((w) => typeof w.confidence === 'number') ? Math.min(...ws.map((w) => w.confidence)) : ws.length ? 100 : null)));
  // Rows that came out empty in every cell (the space between two ruled tables) are dropped.
  const used = text.map((row) => row.some((v) => v !== ''));
  return {
    rows: keepRows.filter((_, i) => used[i]),
    columns: cols,
    cells: text.filter((_, i) => used[i]),
    confidence: sure.filter((_, i) => used[i]),
    ruled,
  };
}
