/**
 * Typing on the canvas (docs/STUDIO.md 7.2, I8): what the Image window's
 * in-place text box needs to look like the text layer it stands over. The box
 * is drawn by the window in the layer's own face, size, weight, colour and
 * alignment; these say which, from the layer as stored or as compiled.
 *
 * Pure: no DOM, no canvas. Runs in a page and in Node.
 */

/** The face named first in a font stack: `"Georgia", Georgia, serif` is Georgia. */
export function faceOf(family) {
    const first = String(family || '').split(',')[0].trim().replace(/^["']|["']$/g, '');
    return first || 'Segoe UI';
}

/** A face as a stack the canvas and the text box both draw with, falling back by kind (serif or sans). */
export function stackOf(face) {
    const name = String(face || 'Segoe UI').replace(/["\\]/g, '');
    const serif = /georgia|times|palatino|cambria|constantia|garamond|book antiqua/i.test(name);
    return `"${name}", ${serif ? 'Georgia, serif' : '"Segoe UI", system-ui, sans-serif'}`;
}

/** The colour a layer's words are drawn in: its solid fill, or a gradient's first stop; white when it says none. */
export function inkOf(layer) {
    const f = layer?.fill;
    if (!f) return '#ffffff';
    if (f.kind === 'solid') return f.color || '#ffffff';
    if (Array.isArray(f.stops) && f.stops.length) {
        const s = f.stops[0];
        return (Array.isArray(s) ? s[1] : s?.color) || '#ffffff';
    }
    return f.color || '#ffffff';
}

/** The size in canvas pixels a layer's words are set at (at least 8, as the painter sets them). */
export const sizeOf = (layer, canvasWidth) => Math.max(8, Math.round(canvasWidth * (Number(layer?.sizeFrac) > 0 ? Number(layer.sizeFrac) : 0.035)));

/** The stored size for a size in canvas pixels: a fraction of the canvas's width. */
export const sizeFracOf = (px, canvasWidth) => +(Math.max(8, Number(px) || 8) / Math.max(1, canvasWidth)).toFixed(5);

/** Whether a layer's words are bold, as the type bar shows it: 700 and heavier (a semibold 600 is not, as in Word). */
export const isBold = (layer) => (Number(layer?.weight) || 500) >= 700;
