/**
 * @rutba/drawing — charts, shapes and drawings, shared.
 *
 * Built for three consumers from the start, because retrofitting shareability is
 * how you end up with three chart libraries that disagree:
 *
 *   Workspace — charts in a sheet, drawings in a document
 *   Studio    — the design surface itself
 *   Mail      — a chart in an email (see the raster caveat in svg.js)
 *
 * Nothing here imports a product. There is no DOM, no filesystem, no package
 * format — a scene is JSON and a render is a string, so the same code runs in a
 * browser, on a server and in a test.
 *
 * The layers, weakest coupling first:
 *   palette.js  the validated colour system — the one thing all three must share
 *   scene.js    the JSON scene graph
 *   measure.js  text metrics without a DOM
 *   svg.js      scene -> SVG string
 *   chart.js    chart spec -> scene, with the design rules enforced
 *   ooxml.js    DrawingML adapter, at the edge and optional
 */
export {
  CATEGORICAL, CATEGORICAL_HUES, RELIEF_REQUIRED, SCATTER_SERIES_CAP,
  SEQUENTIAL_BLUE, ORDINAL_START, DIVERGING, STATUS, INK, MARKS, FONT,
  theme, seriesColour, needsRelief,
} from './palette.js';

export {
  scene, group, rect, ellipse, line, polyline, polygon, path, text, image,
  roundedBarPath, bounds, walk, markCount,
  EMU_PER_PIXEL, EMU_PER_POINT, emuToPx, pxToEmu,
} from './scene.js';

export { measureText, widestText, wrapText, truncateText, lineHeight, capHeight } from './measure.js';
export { renderSvg, svgDataUri, RASTER_REQUIRED_NOTE } from './svg.js';
export {
  buildChart, normaliseSpec, niceScale, labelPolicy, describe, chartTable, CHART_TYPES,
} from './chart.js';
export {
  parseChartXml, parseDrawingAnchors, parseShapeXml, parsePictureXml, readSheetDrawings,
} from './ooxml.js';
export {
  buildShape, buildPicture, resolveColour, toDataUri,
  SUPPORTED_GEOMETRY, MEDIA_TYPES,
} from './shapes.js';
