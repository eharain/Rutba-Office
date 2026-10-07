// Format Shape → 3-D Format and 3-D Rotation, as DrawingML keeps them: a
// top bevel (`a:sp3d/a:bevelT`, its preset, width and height), a depth
// (`a:sp3d@extrusionH`) and its colour, and the camera the shape is seen
// through (`a:scene3d/a:camera@prst`). PowerPoint draws them in true 3-D;
// this suite draws a bevel lit from the top left, and a rotation as the
// flat projection nearest it with the depth stepped out behind — the file
// keeps exactly what was chosen, for PowerPoint to draw.

/** The top-bevel presets, in PowerPoint's gallery order. */
export const BEVELS = [
  ['circle', 'Circle'], ['relaxedInset', 'Relaxed Inset'], ['cross', 'Cross'], ['coolSlant', 'Cool Slant'],
  ['angle', 'Angle'], ['softRound', 'Soft Round'], ['convex', 'Convex'], ['slope', 'Slope'],
  ['divot', 'Divot'], ['riblet', 'Riblet'], ['hardEdge', 'Hard Edge'], ['artDeco', 'Art Deco'],
];

/**
 * The 3-D rotation presets the gallery offers, each with the flat matrix
 * (`a b c d`, about the shape's centre) nearest what PowerPoint shows, and
 * the way its depth runs on the page (a unit step).
 */
export const CAMERAS = [
  ['orthographicFront', 'No Rotation', [1, 0, 0, 1], [0, 0]],
  ['isometricLeftDown', 'Isometric: Left Down', [0.866, 0.5, 0, 1], [0.7, -0.4]],
  ['isometricRightUp', 'Isometric: Right Up', [0.866, -0.5, 0, 1], [-0.7, -0.4]],
  ['isometricTopUp', 'Isometric: Top Up', [0.866, 0.5, -0.866, 0.5], [0, 0.85]],
  ['isometricBottomDown', 'Isometric: Bottom Down', [0.866, -0.5, 0.866, 0.5], [0, -0.85]],
  ['isometricOffAxis1Left', 'Off Axis 1: Left', [0.94, 0.17, 0, 1], [0.75, -0.2]],
  ['isometricOffAxis1Right', 'Off Axis 1: Right', [0.94, -0.17, 0, 1], [-0.75, -0.2]],
  ['obliqueTopLeft', 'Oblique: Top Left', [1, 0, 0, 1], [-0.55, -0.55]],
  ['obliqueTopRight', 'Oblique: Top Right', [1, 0, 0, 1], [0.55, -0.55]],
  ['perspectiveFront', 'Perspective: Front', [1, 0, 0, 1], [0, 0]],
  ['perspectiveLeft', 'Perspective: Left', [0.9, 0.12, 0, 1], [0.65, 0]],
  ['perspectiveRight', 'Perspective: Right', [0.9, -0.12, 0, 1], [-0.65, 0]],
  ['perspectiveAbove', 'Perspective: Above', [1, 0, 0, 0.85], [0, 0.65]],
  ['perspectiveBelow', 'Perspective: Below', [1, 0, 0, 0.85], [0, -0.65]],
];

/** PowerPoint's own depths, in points, as its 3-D Format pane offers them. */
export const DEPTHS = [0, 4.5, 9, 18, 36, 72];

const CAMERA = new Map(CAMERAS.map(([prst, label, matrix, step]) => [prst, { prst, label, matrix, step }]));

/** A camera preset's flat look: its matrix and its depth step — the front view for one this suite does not name. */
export const cameraLook = (prst) => CAMERA.get(prst) ?? { prst, label: prst, matrix: [1, 0, 0, 1], step: [0, 0] };

const EMU_PT = 12700;

/**
 * `<a:scene3d>` and `<a:sp3d>` for a shape's 3-D: `bevel` `{ prst, w, h }`
 * (points), `depth` in points, `depthColor` a hex, `camera` a preset name.
 * Empty when there is nothing to say — a shape seen from the front with no
 * bevel and no depth.
 */
export function shape3dXml({ bevel = null, depth = 0, depthColor = null, camera = 'orthographicFront' } = {}) {
  const flat = !bevel && !(depth > 0) && (!camera || camera === 'orthographicFront');
  if (flat) return '';
  const scene = `<a:scene3d><a:camera prst="${camera || 'orthographicFront'}"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>`;
  const attrs = depth > 0 ? ` extrusionH="${Math.round(depth * EMU_PT)}"` : '';
  const bevelXml = bevel ? `<a:bevelT w="${Math.round((bevel.w ?? 6) * EMU_PT)}" h="${Math.round((bevel.h ?? 6) * EMU_PT)}"${bevel.prst && bevel.prst !== 'circle' ? ` prst="${bevel.prst}"` : ''}/>` : '';
  const colour = depth > 0 && depthColor ? `<a:extrusionClr><a:srgbClr val="${String(depthColor).replace('#', '').toUpperCase()}"/></a:extrusionClr>` : '';
  return scene + `<a:sp3d${attrs}>${bevelXml}${colour}</a:sp3d>`;
}
