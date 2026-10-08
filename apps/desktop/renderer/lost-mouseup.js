// A drag that loses its mouseup still ends.
//
// Every drag in the suite — selecting cells, moving a picture, the ruler's
// markers, a split pane, a slide's points — listens on the window for the
// mouseup that ends it, and only then lets go of its listeners and its
// auto-scroll timer. The window losing focus mid-drag (Alt+Tab, a dialog or
// a menu of the system's) sends no mouseup, nor does a button let go after
// the pointer has left the window and come back; and the drag ran on: the
// sheet kept scrolling, the picture kept following the pointer. This watches
// for those and sends the mouseup that never came, so each drag ends the way
// it always does.

export function installLostMouseupGuard(win = window) {
  let pressed = false;
  let away = false;
  let last = { x: 0, y: 0 };
  const release = () => {
    if (!pressed) return;
    pressed = false;
    away = false;
    win.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, clientX: last.x, clientY: last.y, button: 0, buttons: 0 }));
  };
  win.addEventListener('mousedown', (e) => { pressed = true; away = false; last = { x: e.clientX, y: e.clientY }; }, true);
  win.addEventListener('mouseup', () => { pressed = false; away = false; }, true);
  // The pointer leaving the window altogether (no element to go to).
  win.addEventListener('mouseout', (e) => { if (pressed && !e.relatedTarget) away = true; }, true);
  win.addEventListener('mousemove', (e) => {
    last = { x: e.clientX, y: e.clientY };
    // Back in the window with no button held: it was let go out there.
    if (pressed && away && e.buttons === 0) release();
    else if (away) away = false;
  }, true);
  win.addEventListener('blur', release);
  return () => release();
}
