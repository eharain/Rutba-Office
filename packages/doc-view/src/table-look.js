// A table cell's look as Word draws it — on the page and on paper alike.
//
// A cell's sides are its own border where it gives one; else the last of its
// table style's conditional parts that reaches it and says (a row's part
// gives its top and bottom, and its inside line between the cells across it;
// a column's part its left and right, and its inside line between the cells
// down it); else the table's outer side at an edge of the table and its
// inside line between cells. The parts reach a cell by where it stands and as
// the table's tblLook turns them on, in Word's order: banded columns, banded
// rows, first and last column, header and total row, the corners. Its shading
// and its words' weight, slant and colour come the same way, its own first.

/** The style's parts that reach a cell, in the order they are laid on. */
export function partsAt(style, at) {
  if (!style?.parts) return [];
  const { look = {}, parts, rowBand = 1, colBand = 1 } = style;
  const out = [];
  const add = (type, rowScope, colScope) => { if (parts[type]) out.push({ ...parts[type], rowScope, colScope }); };
  const first = at.row === 0;
  const last = at.row + (at.rowSpan || 1) >= at.rows;
  const start = at.column === 0;
  const end = at.column + (at.span || 1) >= at.columns;
  const header = look.firstRow && first;
  const total = look.lastRow && last;
  const firstCol = look.firstColumn && start;
  const lastCol = look.lastColumn && end;
  const kc = at.column - (look.firstColumn ? 1 : 0);
  if (!look.noVBand && kc >= 0 && !lastCol) add(Math.floor(kc / colBand) % 2 ? 'band2Vert' : 'band1Vert', false, true);
  const kr = at.row - (look.firstRow ? 1 : 0);
  if (!look.noHBand && kr >= 0 && !total) add(Math.floor(kr / rowBand) % 2 ? 'band2Horz' : 'band1Horz', true, false);
  if (firstCol) add('firstCol', false, true);
  if (lastCol) add('lastCol', false, true);
  if (header) add('firstRow', true, false);
  if (total) add('lastRow', true, false);
  if (header && firstCol) add('nwCell', true, true);
  if (header && lastCol) add('neCell', true, true);
  if (total && firstCol) add('swCell', true, true);
  if (total && lastCol) add('seCell', true, true);
  return out;
}

/** Whether a table draws lines at all: its own or its style's sides, a cell's own, a part's. */
export function tableRuled({ borders, style, anyCellBorders }) {
  return Boolean(borders || anyCellBorders || (style && Object.values(style.parts || {}).some((p) => p.borders)));
}

/**
 * A cell's look.
 *   style   the table style: { fill, parts, rowBand, colBand, look } or null
 *   borders the table's sides (its own over its style's) or null
 *   own     the cell's own { borders, fill } or null
 *   at      { row, column, span, rowSpan, rows, columns } — its place in the table
 *   ruled   whether the table draws lines at all (tableRuled)
 * Returns { sides: { top, bottom, left, right } of { style, widthPx, colour }
 * (undefined for none) or null when the table draws none, fill, text: { bold,
 * italic, colour } as far as anything says }.
 */
export function cellLook({ style = null, borders = null, own = null, at, ruled = true }) {
  const parts = partsAt(style, at);
  const first = at.row === 0;
  const last = at.row + (at.rowSpan || 1) >= at.rows;
  const start = at.column === 0;
  const end = at.column + (at.span || 1) >= at.columns;
  let sides = null;
  if (ruled) {
    const t = borders || {};
    const side = { top: first ? t.top : t.insideH, bottom: last ? t.bottom : t.insideH, left: start ? t.left : t.insideV, right: end ? t.right : t.insideV };
    for (const p of parts) {
      const b = p.borders;
      if (!b) continue;
      const set = (name, v) => { if (v !== undefined) side[name] = v; };
      set('top', p.rowScope || first ? b.top : b.insideH);
      set('bottom', p.rowScope || last ? b.bottom : b.insideH);
      set('left', p.colScope || start ? b.left : b.insideV);
      set('right', p.colScope || end ? b.right : b.insideV);
    }
    const mine = own?.borders || {};
    sides = { top: mine.top ?? side.top, bottom: mine.bottom ?? side.bottom, left: mine.left ?? side.left, right: mine.right ?? side.right };
  }
  const lastSaid = (key) => { for (let k = parts.length - 1; k >= 0; k--) if (parts[k][key] !== undefined) return parts[k][key]; return undefined; };
  return {
    sides,
    fill: own?.fill ?? lastSaid('fill') ?? style?.fill ?? null,
    text: { bold: lastSaid('bold'), italic: lastSaid('italic'), colour: lastSaid('colour') },
  };
}

/** A side drawn at all: not set to none, nil or nothing wide. */
export const drawnSide = (b) => Boolean(b) && b.widthPx !== 0 && b.style !== 'none' && b.style !== 'nil';
