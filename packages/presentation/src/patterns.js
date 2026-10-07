// DrawingML's pattern fills — `<a:pattFill prst="…">` — as the eight-by-
// eight tiles Office draws them: a foreground colour on a background, the
// fifty-four presets by their file names and the names PowerPoint's Format
// Shape pane gives them. Each tile is a rule for which of its 64 dots are
// the foreground; the renderer makes an SVG pattern of it.

/** An 8×8 ordered dither: a dot is on when its rank is under the share asked. */
const BAYER = [
  [0, 32, 8, 40, 2, 34, 10, 42], [48, 16, 56, 24, 50, 18, 58, 26],
  [12, 44, 4, 36, 14, 46, 6, 38], [60, 28, 52, 20, 62, 30, 54, 22],
  [3, 35, 11, 43, 1, 33, 9, 41], [51, 19, 59, 27, 49, 17, 57, 25],
  [15, 47, 7, 39, 13, 45, 5, 37], [63, 31, 55, 23, 61, 29, 53, 21],
];
const percent = (p) => (x, y) => BAYER[y][x] < Math.round((p / 100) * 64);
const down = (x, y) => (x - y + 64) % 8;
const up = (x, y) => (x + y) % 8;
const confetti = (dots) => { const set = new Set(dots.map(([x, y]) => x + ',' + y)); return (x, y) => set.has(x + ',' + y); };

/** Each preset: its file name, its name in PowerPoint, and its rule. */
export const PATTERNS = [
  ['pct5', '5%', percent(5)], ['pct10', '10%', percent(10)], ['pct20', '20%', percent(20)], ['pct25', '25%', percent(25)],
  ['pct30', '30%', percent(30)], ['pct40', '40%', percent(40)], ['pct50', '50%', percent(50)], ['pct60', '60%', percent(60)],
  ['pct70', '70%', percent(70)], ['pct75', '75%', percent(75)], ['pct80', '80%', percent(80)], ['pct90', '90%', percent(90)],
  ['horz', 'Horizontal', (x, y) => y % 8 === 0], ['vert', 'Vertical', (x) => x % 8 === 0],
  ['ltHorz', 'Light horizontal', (x, y) => y % 4 === 0], ['ltVert', 'Light vertical', (x) => x % 4 === 0],
  ['dkHorz', 'Dark horizontal', (x, y) => y % 4 < 2], ['dkVert', 'Dark vertical', (x) => x % 4 < 2],
  ['narHorz', 'Narrow horizontal', (x, y) => y % 2 === 0], ['narVert', 'Narrow vertical', (x) => x % 2 === 0],
  ['dashHorz', 'Dashed horizontal', (x, y) => y % 4 === 0 && x % 8 < 4], ['dashVert', 'Dashed vertical', (x, y) => x % 4 === 0 && y % 8 < 4],
  ['cross', 'Cross', (x, y) => x % 8 === 0 || y % 8 === 0],
  ['dnDiag', 'Downward diagonal', (x, y) => down(x, y) < 2], ['upDiag', 'Upward diagonal', (x, y) => up(x, y) < 2],
  ['ltDnDiag', 'Light downward diagonal', (x, y) => down(x, y) % 4 === 0], ['ltUpDiag', 'Light upward diagonal', (x, y) => up(x, y) % 4 === 0],
  ['dkDnDiag', 'Dark downward diagonal', (x, y) => down(x, y) % 4 < 2], ['dkUpDiag', 'Dark upward diagonal', (x, y) => up(x, y) % 4 < 2],
  ['wdDnDiag', 'Wide downward diagonal', (x, y) => down(x, y) < 3], ['wdUpDiag', 'Wide upward diagonal', (x, y) => up(x, y) < 3],
  ['dashDnDiag', 'Dashed downward diagonal', (x, y) => down(x, y) % 4 === 0 && y % 8 < 4], ['dashUpDiag', 'Dashed upward diagonal', (x, y) => up(x, y) % 4 === 0 && y % 8 < 4],
  ['diagCross', 'Diagonal cross', (x, y) => down(x, y) === 0 || up(x, y) === 0],
  ['smCheck', 'Small checker board', (x, y) => ((x >> 1) + (y >> 1)) % 2 === 0], ['lgCheck', 'Large checker board', (x, y) => ((x >> 2) + (y >> 2)) % 2 === 0],
  ['smGrid', 'Small grid', (x, y) => x % 4 === 0 || y % 4 === 0], ['lgGrid', 'Large grid', (x, y) => x % 8 === 0 || y % 8 === 0],
  ['dotGrid', 'Dotted grid', (x, y) => (x % 4 === 0 && y % 2 === 0) || (y % 4 === 0 && x % 2 === 0)],
  ['smConfetti', 'Small confetti', confetti([[0, 0], [4, 1], [2, 2], [7, 3], [1, 4], [5, 5], [3, 6], [6, 7]])],
  ['lgConfetti', 'Large confetti', confetti([[0, 0], [1, 0], [0, 1], [1, 1], [4, 2], [5, 2], [4, 3], [5, 3], [2, 5], [3, 5], [2, 6], [3, 6], [6, 6], [7, 6], [6, 7], [7, 7]])],
  ['horzBrick', 'Horizontal brick', (x, y) => y % 4 === 0 || x === (Math.floor(y / 4) % 2 ? 4 : 0)],
  ['diagBrick', 'Diagonal brick', (x, y) => up(x, y) === 0 || (down(x, y) === 0 && x % 8 < 4)],
  ['solidDmnd', 'Solid diamond', (x, y) => Math.abs(x - 3.5) + Math.abs(y - 3.5) <= 3.5],
  ['openDmnd', 'Outlined diamond', (x, y) => Math.abs(Math.abs(x - 3.5) + Math.abs(y - 3.5) - 3) < 0.6],
  ['dotDmnd', 'Dotted diamond', (x, y) => (x + y) % 4 === 0 && (x - y + 8) % 4 === 0],
  ['plaid', 'Plaid', (x, y) => (x % 8 < 4 && y % 2 === 0) || (y % 8 < 4 && x % 2 === 0)],
  ['sphere', 'Sphere', (x, y) => { const d = Math.hypot(x - 3.5, y - 3.5); return d > 2.4 && d < 3.8; }],
  ['weave', 'Weave', (x, y) => up(x, y) % 4 === 0 || (down(x, y) === 0 && x % 2 === 0)],
  ['divot', 'Divot', confetti([[2, 1], [3, 2], [2, 3], [6, 5], [7, 6], [6, 7]])],
  ['shingle', 'Shingle', (x, y) => (up(x, y) === 0 && y < 4) || (y === 4 && x < 4) || (x === 0 && y > 4)],
  ['wave', 'Wave', (x, y) => y === Math.round(3.5 + 2 * Math.sin((x / 8) * 2 * Math.PI))],
  ['trellis', 'Trellis', (x, y) => up(x, y) % 4 === 0 || down(x, y) % 4 === 0],
  ['zigZag', 'Zig zag', (x, y) => y % 4 === (x % 4 < 2 ? x % 4 : 3 - (x % 4))],
];

const BY_NAME = new Map(PATTERNS.map(([name, label, rule]) => [name, { name, label, rule }]));

/** Whether `name` is one of DrawingML's presets. */
export const isPattern = (name) => BY_NAME.has(String(name));

/**
 * An SVG `<pattern>` of the preset: the background, then each row's runs of
 * foreground dots as rectangles, one dot a unit, the tile `size` units wide.
 */
export function patternDef(id, name, fg = '#000000', bg = '#ffffff', { fgAlpha = 1, bgAlpha = 1, size = 8 } = {}) {
  const rule = BY_NAME.get(String(name))?.rule ?? percent(50);
  const unit = size / 8;
  const rects = [];
  for (let y = 0; y < 8; y++) {
    let x = 0;
    while (x < 8) {
      if (!rule(x, y)) { x += 1; continue; }
      let end = x;
      while (end < 8 && rule(end, y)) end += 1;
      rects.push(`<rect x="${x * unit}" y="${y * unit}" width="${(end - x) * unit}" height="${unit}"/>`);
      x = end;
    }
  }
  const bgRect = `<rect width="${size}" height="${size}" fill="${bg}"${bgAlpha < 1 ? ` fill-opacity="${bgAlpha}"` : ''}/>`;
  return `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${size}" height="${size}">${bgRect}<g fill="${fg}"${fgAlpha < 1 ? ` fill-opacity="${fgAlpha}"` : ''}>${rects.join('')}</g></pattern>`;
}
