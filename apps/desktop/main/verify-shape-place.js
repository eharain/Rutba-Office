// Presentations: Insert → Shapes waits for the chosen shape to be drawn on
// the slide (a drag draws it, a click puts it down at its usual size). The
// checks that only need a shape on the slide put it down with a click in
// the middle, where it used to land by itself.

/**
 * A click on the drawing surface, at a share of the slide's width and
 * height; or a drag from one share to another when `to` is given.
 * @param {Electron.WebContents} wc
 * @returns {Promise<string>} 'put', or 'no surface' when nothing waited
 *
 * The surface is waited for inside the page, where verify-waits.js cannot
 * give the wait its floor: so it has the same 15 seconds here (it was 3, and
 * 01.0109 found no surface once, beside a full run of 2.26.0). It ends as
 * soon as the surface is there.
 */
export function putShapeDown(wc, { x = 0.5, y = 0.5, to = null, shift = false } = {}) {
  return wc.executeJavaScript(`(async () => {
    for (let i = 0; i < 300 && !document.querySelector('[data-role="shape-draw"]'); i++) await new Promise((r) => setTimeout(r, 50));
    const s = document.querySelector('[data-role="shape-draw"]');
    if (!s) return 'no surface';
    const r = s.getBoundingClientRect();
    const at = (px, py) => ({ bubbles: true, cancelable: true, button: 0, shiftKey: ${shift ? 'true' : 'false'}, clientX: r.left + r.width * px, clientY: r.top + r.height * py });
    s.dispatchEvent(new MouseEvent('mousedown', at(${x}, ${y})));
    ${to ? `window.dispatchEvent(new MouseEvent('mousemove', at(${(x + to.x) / 2}, ${(y + to.y) / 2})));
    window.dispatchEvent(new MouseEvent('mouseup', at(${to.x}, ${to.y})));` : `window.dispatchEvent(new MouseEvent('mouseup', at(${x}, ${y})));`}
    return 'put';
  })()`);
}
