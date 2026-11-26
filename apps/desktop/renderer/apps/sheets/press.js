/**
 * Which cell a press on the grid is for.
 *
 * A cell whose words are longer than its column is drawn over the empty
 * cells beside it, as Excel draws it. Its box on the screen is then wider
 * than the cell, and a press on the words where they lie over a neighbour
 * lands in that box. The press is for the neighbour: Excel selects the cell
 * under the pointer, whoever's words are drawn there.
 */

/** The column and the row of the frame under a point of the grid, or null off the grid. */
export function cellUnder(model, x, y) {
  const col = (model?.columns || []).find((c) => x >= c.x && x < c.x + c.width);
  const row = (model?.rows || []).find((r) => y >= r.y && y < r.y + r.height);
  return col && row ? { row: row.index, col: col.index, x: col.x } : null;
}

/**
 * The cell a press inside a drawn cell's box is for: the cell itself while
 * the pointer is over its own columns (a merged cell's are all its own),
 * and the cell under the pointer, in the same row, where the box is only
 * the cell's words lying over a neighbour.
 */
export function pressedCell(cell, under) {
  const own = { row: cell.row, col: cell.col };
  if (!under || under.col === cell.col) return own;
  if (under.x >= cell.x && under.x < cell.x + cell.width) return own;
  return { row: cell.row, col: under.col };
}
