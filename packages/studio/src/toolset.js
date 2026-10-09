/**
 * The Studio toolset — one import that installs every extension layer type.
 *
 * The three modules register themselves on import, which makes the import
 * itself load-bearing: a host that imports `renderVideo` but not this file gets
 * a renderer that silently skips every word-art and shape layer, because
 * paintFrame's `if (!painter) continue;` is a deliberate forward-compatibility
 * rule, not an error path. That failure looks like "the headline vanished in
 * the render but shows in the preview", which is a miserable thing to debug.
 *
 * So both hosts import THIS, not the individual files:
 *   - the editor page, when it mounts the canvas
 *   - the render worker, before it builds its first plan
 *
 * Re-exported here as well: the vocabularies an editor needs to build its
 * panels (presets, warps, geometries) without reaching past this barrel.
 */

export { WORDART, PRESETS as WORDART_PRESETS, WARPS, wordArtResize } from './wordart.js';
export { SHAPE, GEOMETRIES, SHAPE_GROUPS, GEOMETRY_LABELS, paintShapeTile } from './shapes.js';
export { ADJUST, ADJUST_PRESETS, applyAdjustment } from './adjust.js';
export { ICON, ICONS, ICON_NAMES, ICON_GROUPS, ICON_LABELS } from './icons.js';
export { TEXTBLOCK, LIST_KINDS, LIST_LABELS } from './textblock.js';
export { CHART, CHART_KINDS, CHART_LABELS, CHART_PALETTE, niceMax, readData } from './charts.js';
export { registerLayerType, registeredLayerTypes } from './index.js';

/**
 * What the editor's "Add" row offers, in the order it offers it.
 *
 * Ordered by how often it is reached for, not alphabetically: a headline is the
 * first thing anyone adds to a creative, and an adjustment is usually the last.
 * Each entry carries the patch a click should append, so the Add row stays a
 * list of data rather than a switch statement in the UI.
 */
export const ADD_MENU = [
    {
        type: 'wordart',
        label: 'Word art',
        icon: 'fa-wand-magic-sparkles',
        patch: { type: 'wordart', text: 'HEADLINE', preset: 'sticker', fx: 0.5, fy: 0.32, sizeFrac: 0.11 },
    },
    {
        type: 'shape',
        label: 'Shape',
        icon: 'fa-shapes',
        patch: { type: 'shape', geometry: 'rect', fx: 0.5, fy: 0.5, fw: 0.36, fh: 0.18, radius: 0.12 },
    },
    {
        type: 'shape',
        label: 'Badge',
        icon: 'fa-certificate',
        patch: { type: 'shape', geometry: 'burst', fx: 0.76, fy: 0.22, fw: 0.3, fh: 0.3, innerRatio: 0.82 },
    },
    {
        type: 'shape',
        label: 'Speech bubble',
        icon: 'fa-comment',
        patch: { type: 'shape', geometry: 'bubble', fx: 0.5, fy: 0.4, fw: 0.5, fh: 0.28, radius: 0.14 },
    },
    {
        type: 'adjust',
        label: 'Adjustment',
        icon: 'fa-sliders',
        patch: { type: 'adjust', preset: 'punch' },
    },
];
