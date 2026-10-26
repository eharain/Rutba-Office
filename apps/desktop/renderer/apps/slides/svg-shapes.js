// A slide's drawing as the window has it: each shape in a <g data-shape="id">
// of its own (the engine's tagShapes), so one shape can be taken out of it.

/**
 * The slide's drawing without the shapes named: each shape is drawn in a
 * `<g data-shape="id">` of its own, taken out whole, groups inside it too.
 */
export function withoutShapes(svg, ids) {
  let out = svg;
  for (const id of ids) {
    const open = out.indexOf(`<g data-shape="${id}"`);
    if (open < 0) continue;
    const tag = /<g[\s>/]|<\/g>/g;
    tag.lastIndex = open + 2;
    let depth = 1;
    let end = -1;
    for (let m = tag.exec(out); m; m = tag.exec(out)) {
      if (m[0] === '</g>') depth -= 1;
      else if (!out.slice(m.index, out.indexOf('>', m.index) + 1).endsWith('/>')) depth += 1;
      if (depth === 0) { end = m.index + 4; break; }
    }
    if (end > open) out = out.slice(0, open) + out.slice(end);
  }
  return out;
}
