/**
 * Zoom and pan on the stage — the arithmetic, apart from `Stage.js` so it can
 * be tested without a browser.
 *
 * The canvas is always the full frame (see the note in Stage.js); zoom is only
 * how large CSS draws it. "Fit" is the scale that shows the whole frame in the
 * stage and is what the editor opens on. Past fit the stage scrolls, and
 * zooming keeps the point under the pointer where it was — the thing that
 * makes Ctrl+wheel usable rather than a scroll back to where you were.
 */

/** The steps the buttons and keys move between, as design tools offer them. */
export const ZOOM_STEPS = [0.1, 0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4];
export const MIN_ZOOM = ZOOM_STEPS[0];
export const MAX_ZOOM = ZOOM_STEPS[ZOOM_STEPS.length - 1];

export const clampZoom = (z) => Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, Number(z) || 1));

/** The scale that shows the whole frame inside a box, never above 100%. */
export function fitScale(box, frame) {
    if (!box?.width || !box?.height || !frame?.width || !frame?.height) return 1;
    return Math.min(1, box.width / frame.width, box.height / frame.height);
}

/**
 * The next step in or out from where the stage is now. From an in-between
 * scale (fit is rarely a step) the first press lands on the nearest step in
 * that direction, never further.
 */
export function stepZoom(current, direction) {
    const z = clampZoom(current);
    if (direction > 0) return ZOOM_STEPS.find((s) => s > z + 1e-6) ?? MAX_ZOOM;
    return [...ZOOM_STEPS].reverse().find((s) => s < z - 1e-6) ?? MIN_ZOOM;
}

/** "67%" — what the zoom button says. */
export const zoomLabel = (scale) => `${Math.round(scale * 100)}%`;

/**
 * Where the scroller must scroll to so the content point under `pointer`
 * (relative to the scroller's visible box) stays under it after the scale
 * changes from `from` to `to`. `margin` is the space around the frame, which
 * does not scale.
 */
export function scrollToKeep({ scrollLeft, scrollTop }, pointer, from, to, margin = 0) {
    const k = to / from;
    const along = (scroll, p) => Math.max(0, (scroll + p - margin) * k + margin - p);
    return { left: along(scrollLeft, pointer.x), top: along(scrollTop, pointer.y) };
}

/** What a key press asks of the zoom: Ctrl/Cmd with +, - or 0. Null for anything else. */
export function zoomKey(e) {
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return null;
    if (e.key === "=" || e.key === "+") return "in";
    if (e.key === "-" || e.key === "_") return "out";
    if (e.key === "0") return "fit";
    return null;
}
