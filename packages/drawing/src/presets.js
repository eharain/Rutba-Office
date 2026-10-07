// Preset geometries a renderer draws no shape of its own for, as SVG path
// data in a box: the outlines DrawingML's preset definitions give them (the
// guides and paths of ECMA-376's presetShapeDefinitions), at their default
// adjustments. A heart, a smiley, callouts, the block arrows, a can, a cube,
// a donut, a frame, stars, a pentagon arrow, a cloud and the flowchart shapes.
// Several outlines are subpaths of one path — a smiley's eyes and mouth, a
// cube's faces — drawn in the shape's fill and outline. Pure.

const f = (n) => Math.round(n * 100) / 100;
const pts = (list) => 'M' + list.map(([x, y]) => `${f(x)} ${f(y)}`).join(' L') + ' Z';

/**
 * An elliptical arc as DrawingML's arcTo draws one from where the pen is:
 * its radii, and the angles it starts at and sweeps through, in degrees,
 * clockwise from three o'clock as a page turns them. Returns the SVG arc and
 * where it ends.
 */
function arcTo(at, wR, hR, stAng, swAng) {
  // The angle on the ellipse a visual angle is.
  const param = (deg) => {
    const a = (deg * Math.PI) / 180;
    return Math.atan2(wR * Math.sin(a), hR * Math.cos(a));
  };
  const t1 = param(stAng);
  const t2 = param(stAng + swAng);
  const cx = at[0] - wR * Math.cos(t1);
  const cy = at[1] - hR * Math.sin(t1);
  const end = [cx + wR * Math.cos(t2), cy + hR * Math.sin(t2)];
  const large = Math.abs(swAng) > 180 ? 1 : 0;
  const sweep = swAng > 0 ? 1 : 0;
  return { d: ` A${f(wR)} ${f(hR)} 0 ${large} ${sweep} ${f(end[0])} ${f(end[1])}`, end };
}

/** The flowchart shapes that are other shapes by another name. */
const SAME_AS = {
  flowChartProcess: 'rect', flowChartAlternateProcess: 'roundRect', flowChartDecision: 'diamond', flowChartConnector: 'ellipse',
  flowChartInputOutput: 'flowChartIO', flowChartPreparation: 'hexagon', flowChartMerge: 'flowChartMergeShape', flowChartExtract: 'triangle',
};

/**
 * A preset's outline as SVG path data in the box (x, y, w, h), or null for
 * one not drawn here.
 */
export function presetPath(preset, x, y, w, h) {
  const name = SAME_AS[preset] || preset;
  const r = x + w;
  const b = y + h;
  const hc = x + w / 2;
  const vc = y + h / 2;
  const ss = Math.min(w, h);
  switch (name) {
    case 'rect':
      return pts([[x, y], [r, y], [r, b], [x, b]]);
    case 'roundRect': {
      const rr = ss * 0.16667;
      return `M${f(x + rr)} ${f(y)} L${f(r - rr)} ${f(y)} Q${f(r)} ${f(y)} ${f(r)} ${f(y + rr)} L${f(r)} ${f(b - rr)} Q${f(r)} ${f(b)} ${f(r - rr)} ${f(b)} L${f(x + rr)} ${f(b)} Q${f(x)} ${f(b)} ${f(x)} ${f(b - rr)} L${f(x)} ${f(y + rr)} Q${f(x)} ${f(y)} ${f(x + rr)} ${f(y)} Z`;
    }
    case 'ellipse':
      return `M${f(x)} ${f(vc)} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(r)} ${f(vc)} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(x)} ${f(vc)} Z`;
    case 'diamond':
      return pts([[hc, y], [r, vc], [hc, b], [x, vc]]);
    case 'triangle':
      return pts([[hc, y], [r, b], [x, b]]);
    case 'hexagon': {
      const dx = ss * 0.25;
      return pts([[x, vc], [x + dx, y], [r - dx, y], [r, vc], [r - dx, b], [x + dx, b]]);
    }
    case 'heart': {
      // Two curves from the cleft at the top to the point at the bottom.
      const dx1 = (w * 49) / 48;
      const dx2 = (w * 10) / 48;
      const y1 = y - h / 3;
      const hd4 = y + h / 4;
      return `M${f(hc)} ${f(hd4)} C${f(hc + dx2)} ${f(y1)} ${f(hc + dx1)} ${f(hd4)} ${f(hc)} ${f(b)} C${f(hc - dx1)} ${f(hd4)} ${f(hc - dx2)} ${f(y1)} ${f(hc)} ${f(hd4)} Z`;
    }
    case 'smileyFace': {
      // The face, its eyes, and its mouth (a curve the adjustment bends).
      const face = presetPath('ellipse', x, y, w, h);
      const ex1 = x + (w * 6215) / 21600;
      const ex2 = x + (w * 13135) / 21600;
      const ey = y + (h * 7570) / 21600;
      const ew = (w * 1125) / 21600;
      const eh = (h * 1125) / 21600;
      const eye = (cx) => ` M${f(cx - ew)} ${f(ey)} A${f(ew)} ${f(eh)} 0 1 1 ${f(cx + ew)} ${f(ey)} A${f(ew)} ${f(eh)} 0 1 1 ${f(cx - ew)} ${f(ey)} Z`;
      const mx1 = x + (w * 4969) / 21699;
      const mx2 = x + (w * 16640) / 21600;
      const my3 = y + (h * 16515) / 21600;
      const dy2 = (h * 4653) / 100000;
      const my2 = my3 - dy2;
      const my5 = my3 + dy2 + (h * 4653) / 50000;
      // The mouth turning the way the face does, so it is a line on it and not a hole in it.
      return `${face}${eye(ex1)}${eye(ex2)} M${f(mx2)} ${f(my2)} Q${f(hc)} ${f(my5)} ${f(mx1)} ${f(my2)}`;
    }
    case 'wedgeRectCallout':
    case 'wedgeRoundRectCallout': {
      // A box with a tail to its tip, set out from its middle by the adjustments.
      const dxPos = (w * -20833) / 100000;
      const dyPos = (h * 62500) / 100000;
      const xPos = hc + dxPos;
      const yPos = vc + dyPos;
      const dq = (dxPos * h) / w;
      const vertical = Math.abs(dyPos) - Math.abs(dq) > 0;
      const xg1 = x + (w * (dxPos > 0 ? 7 : 2)) / 12;
      const xg2 = x + (w * (dxPos > 0 ? 10 : 5)) / 12;
      const yg1 = y + (h * (dyPos > 0 ? 7 : 2)) / 12;
      const yg2 = y + (h * (dyPos > 0 ? 10 : 5)) / 12;
      const rr = name === 'wedgeRoundRectCallout' ? ss * 0.16667 : 0;
      const corner = (cx, cy, tx, ty) => (rr ? ` Q${f(cx)} ${f(cy)} ${f(tx)} ${f(ty)}` : '');
      let d = `M${f(x + rr)} ${f(y)}`;
      if (vertical && dyPos < 0) d += ` L${f(xg1)} ${f(y)} L${f(xPos)} ${f(yPos)} L${f(xg2)} ${f(y)}`;
      d += ` L${f(r - rr)} ${f(y)}${corner(r, y, r, y + rr)}`;
      if (!vertical && dxPos > 0) d += ` L${f(r)} ${f(yg1)} L${f(xPos)} ${f(yPos)} L${f(r)} ${f(yg2)}`;
      d += ` L${f(r)} ${f(b - rr)}${corner(r, b, r - rr, b)}`;
      if (vertical && dyPos > 0) d += ` L${f(xg2)} ${f(b)} L${f(xPos)} ${f(yPos)} L${f(xg1)} ${f(b)}`;
      d += ` L${f(x + rr)} ${f(b)}${corner(x, b, x, b - rr)}`;
      if (!vertical && dxPos < 0) d += ` L${f(x)} ${f(yg2)} L${f(xPos)} ${f(yPos)} L${f(x)} ${f(yg1)}`;
      d += ` L${f(x)} ${f(y + rr)}${corner(x, y, x + rr, y)} Z`;
      return d;
    }
    case 'wedgeEllipseCallout': {
      // An ellipse whose edge opens to a tail, either side of the line to its tip.
      const xPos = hc + (w * -20833) / 100000;
      const yPos = vc + (h * 62500) / 100000;
      const to = Math.atan2((yPos - vc) / (h / 2), (xPos - hc) / (w / 2));
      const at = (t) => [hc + (w / 2) * Math.cos(t), vc + (h / 2) * Math.sin(t)];
      const [p1, p2] = [at(to + 0.18), at(to - 0.18)];
      return `M${f(p1[0])} ${f(p1[1])} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(p2[0])} ${f(p2[1])} L${f(xPos)} ${f(yPos)} Z`;
    }
    case 'upArrow':
    case 'downArrow':
    case 'leftArrow':
    case 'rightArrow': {
      // A shaft half the width across, a head as long as half the shorter side.
      const head = (ss * 50000) / 100000;
      if (name === 'upArrow' || name === 'downArrow') {
        const x1 = hc - w / 4;
        const x2 = hc + w / 4;
        return name === 'upArrow'
          ? pts([[x, y + head], [hc, y], [r, y + head], [x2, y + head], [x2, b], [x1, b], [x1, y + head]])
          : pts([[x1, y], [x2, y], [x2, b - head], [r, b - head], [hc, b], [x, b - head], [x1, b - head]]);
      }
      const y1 = vc - h / 4;
      const y2 = vc + h / 4;
      return name === 'rightArrow'
        ? pts([[x, y1], [r - head, y1], [r - head, y], [r, vc], [r - head, b], [r - head, y2], [x, y2]])
        : pts([[x + head, y], [x + head, y1], [r, y1], [r, y2], [x + head, y2], [x + head, b], [x, vc]]);
    }
    case 'leftRightArrow': {
      const head = (ss * 50000) / 100000;
      const y1 = vc - h / 4;
      const y2 = vc + h / 4;
      return pts([[x, vc], [x + head, y], [x + head, y1], [r - head, y1], [r - head, y], [r, vc], [r - head, b], [r - head, y2], [x + head, y2], [x + head, b]]);
    }
    case 'upDownArrow': {
      const head = (ss * 50000) / 100000;
      const x1 = hc - w / 4;
      const x2 = hc + w / 4;
      return pts([[hc, y], [r, y + head], [x2, y + head], [x2, b - head], [r, b - head], [hc, b], [x, b - head], [x1, b - head], [x1, y + head], [x, y + head]]);
    }
    case 'homePlate': {
      const x1 = r - (ss * 50000) / 100000;
      return pts([[x, y], [x1, y], [r, vc], [x1, b], [x, b]]);
    }
    case 'can': {
      // A body between two half ellipses, and the top, a whole one.
      const ry = (ss * 25000) / 200000;
      const y3 = b - ry;
      return `M${f(x)} ${f(y + ry)} A${f(w / 2)} ${f(ry)} 0 0 0 ${f(r)} ${f(y + ry)} L${f(r)} ${f(y3)} A${f(w / 2)} ${f(ry)} 0 0 1 ${f(x)} ${f(y3)} Z`
        + ` M${f(x)} ${f(y + ry)} A${f(w / 2)} ${f(ry)} 0 1 1 ${f(r)} ${f(y + ry)} A${f(w / 2)} ${f(ry)} 0 1 1 ${f(x)} ${f(y + ry)} Z`;
    }
    case 'cube': {
      const d = (ss * 25000) / 100000;
      return pts([[x, y + d], [r - d, y + d], [r - d, b], [x, b]])
        + ' ' + pts([[x, y + d], [x + d, y], [r, y], [r - d, y + d]])
        + ' ' + pts([[r - d, y + d], [r, y], [r, b - d], [r - d, b]]);
    }
    case 'donut': {
      // A ring: the inner ellipse turning the other way, so it is a hole.
      const dr = (ss * 25000) / 100000;
      const iw = w / 2 - dr;
      const ih = h / 2 - dr;
      return `${presetPath('ellipse', x, y, w, h)} M${f(hc - iw)} ${f(vc)} A${f(iw)} ${f(ih)} 0 1 0 ${f(hc + iw)} ${f(vc)} A${f(iw)} ${f(ih)} 0 1 0 ${f(hc - iw)} ${f(vc)} Z`;
    }
    case 'frame': {
      const t = (ss * 12500) / 100000;
      return `${pts([[x, y], [r, y], [r, b], [x, b]])} M${f(x + t)} ${f(y + t)} L${f(x + t)} ${f(b - t)} L${f(r - t)} ${f(b - t)} L${f(r - t)} ${f(y + t)} Z`;
    }
    case 'plus': {
      const d = (ss * 25000) / 100000;
      return pts([[x, y + d], [x + d, y + d], [x + d, y], [r - d, y], [r - d, y + d], [r, y + d], [r, b - d], [r - d, b - d], [r - d, b], [x + d, b], [x + d, b - d], [x, b - d]]);
    }
    case 'octagon': {
      const d = (ss * 29289) / 100000;
      return pts([[x, y + d], [x + d, y], [r - d, y], [r, y + d], [r, b - d], [r - d, b], [x + d, b], [x, b - d]]);
    }
    case 'star4':
    case 'star6':
    case 'star7':
    case 'star8':
    case 'star10':
    case 'star12': {
      // Points round the box's ellipse, the inner ones as far in as the preset's adjustment says.
      const n = Number(name.slice(4));
      const inner = { 4: 0.25, 6: 0.57736, 7: 0.69202, 8: 0.75, 10: 0.85066, 12: 0.75 }[n];
      const list = [];
      for (let i = 0; i < n * 2; i++) {
        const a = (Math.PI * i) / n - Math.PI / 2;
        const k = i % 2 ? inner : 1;
        list.push([hc + (w / 2) * k * Math.cos(a), vc + (h / 2) * k * Math.sin(a)]);
      }
      return pts(list);
    }
    case 'flowChartTerminator': {
      const rx = Math.min(w / 2, h / 2);
      return `M${f(x + rx)} ${f(y)} L${f(r - rx)} ${f(y)} A${f(rx)} ${f(h / 2)} 0 0 1 ${f(r - rx)} ${f(b)} L${f(x + rx)} ${f(b)} A${f(rx)} ${f(h / 2)} 0 0 1 ${f(x + rx)} ${f(y)} Z`;
    }
    case 'flowChartIO':
      return pts([[x + w / 5, y], [r, y], [r - w / 5, b], [x, b]]);
    case 'flowChartMergeShape':
      return pts([[x, y], [r, y], [hc, b]]);
    case 'flowChartPredefinedProcess': {
      const d = w / 8;
      return `${pts([[x, y], [r, y], [r, b], [x, b]])} M${f(x + d)} ${f(y)} L${f(x + d)} ${f(b)} M${f(r - d)} ${f(y)} L${f(r - d)} ${f(b)}`;
    }
    case 'flowChartDocument': {
      // A page whose foot is a wave.
      const y1 = y + (h * 17322) / 21600;
      const y2 = y + (h * 20172) / 21600;
      const y3 = y + (h * 23922) / 21600;
      return `M${f(x)} ${f(y)} L${f(r)} ${f(y)} L${f(r)} ${f(y1)} C${f(x + (w * 10800) / 21600)} ${f(y1)} ${f(x + (w * 10800) / 21600)} ${f(y3)} ${f(x)} ${f(y2)} Z`;
    }
    case 'flowChartManualInput':
      return pts([[x, y + h / 5], [r, y], [r, b], [x, b]]);
    case 'flowChartManualOperation':
      return pts([[x, y], [r, y], [r - w / 5, b], [x + w / 5, b]]);
    case 'flowChartOffpageConnector':
      return pts([[x, y], [r, y], [r, y + (h * 4) / 5], [hc, b], [x, y + (h * 4) / 5]]);
    case 'flowChartMagneticDisk':
      return presetPath('can', x, y, w, h);
    case 'cloud': {
      // Eleven arcs round a 43200-unit box, as DrawingML draws a cloud.
      const sx = w / 43200;
      const sy = h / 43200;
      const arcs = [[6753, 9190, -11429249, 7426832], [5333, 7267, -8646143, 5396714], [4365, 5945, -8748475, 5983381], [4857, 6595, -7859164, 7034504],
        [5333, 7273, -4722533, 6541615], [6775, 9220, -2776035, 7816140], [5785, 7867, 37501, 6842000], [6752, 9215, 1347096, 6910353],
        [7720, 10543, 3974558, 4542661], [4360, 5918, -16496525, 8804134], [4345, 5945, -14809710, 9151131]];
      let at = [x + 3900 * sx, y + 14370 * sy];
      let d = `M${f(at[0])} ${f(at[1])}`;
      for (const [wR, hR, st, sw] of arcs) {
        const arc = arcTo(at, wR * sx, hR * sy, st / 60000, sw / 60000);
        d += arc.d;
        at = arc.end;
      }
      return d + ' Z';
    }
    default:
      return null;
  }
}

/** The presets presetPath draws. */
export const PRESET_PATHS = [
  'rect', 'roundRect', 'ellipse', 'diamond', 'triangle', 'hexagon', 'heart', 'smileyFace', 'wedgeRectCallout', 'wedgeRoundRectCallout',
  'wedgeEllipseCallout', 'upArrow', 'downArrow', 'leftArrow', 'rightArrow', 'leftRightArrow', 'upDownArrow', 'homePlate', 'can', 'cube',
  'donut', 'frame', 'plus', 'octagon', 'star4', 'star6', 'star7', 'star8', 'star10', 'star12', 'flowChartTerminator', 'flowChartPredefinedProcess',
  'flowChartDocument', 'flowChartManualInput', 'flowChartManualOperation', 'flowChartOffpageConnector', 'flowChartMagneticDisk', 'cloud',
  ...Object.keys(SAME_AS),
];
