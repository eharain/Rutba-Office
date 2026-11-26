/**
 * Fill Down, Right, Up and Left (Ctrl+D, Ctrl+R, and Home → Fill): the
 * operations the window sends.
 *
 * One cell fills from its neighbour above, left, below or right; a range
 * from its own first row or column, or its last for up and left. These copy,
 * as Excel's do: a date stays that date, Mon stays Mon and Item 1 stays
 * Item 1, with a formula's references still moving. Continuing a series is
 * the fill handle's work, not this command's.
 *
 * `range` is the selection, `at` its active cell. Returns the operations in
 * order, or null when one cell has no neighbour to fill from.
 */
export function fillOps(direction, range, at) {
  const single = range.top === range.bottom && range.left === range.right;
  if (single) {
    const source = direction === 'right' ? { row: at.row, col: at.col - 1 }
      : direction === 'left' ? { row: at.row, col: at.col + 1 }
        : direction === 'up' ? { row: at.row + 1, col: at.col }
          : { row: at.row - 1, col: at.col };
    if (source.row < 0 || source.col < 0) return null;
    const target = { top: Math.min(source.row, at.row), left: Math.min(source.col, at.col), bottom: Math.max(source.row, at.row), right: Math.max(source.col, at.col) };
    return [
      { op: 'select', row: source.row, col: source.col },
      { op: 'fill', target, mode: 'copy' },
      { op: 'select', row: at.row, col: at.col },
    ];
  }
  const source = direction === 'right'
    ? { top: range.top, bottom: range.bottom, left: range.left, right: range.left }
    : direction === 'left' ? { top: range.top, bottom: range.bottom, left: range.right, right: range.right }
      : direction === 'up' ? { top: range.bottom, bottom: range.bottom, left: range.left, right: range.right }
        : { top: range.top, bottom: range.top, left: range.left, right: range.right };
  return [
    { op: 'select', row: source.top, col: source.left },
    { op: 'select', row: source.bottom, col: source.right, extend: true },
    { op: 'fill', target: { top: range.top, left: range.left, bottom: range.bottom, right: range.right }, mode: 'copy' },
    { op: 'select', row: range.top, col: range.left },
    { op: 'select', row: range.bottom, col: range.right, extend: true },
  ];
}
