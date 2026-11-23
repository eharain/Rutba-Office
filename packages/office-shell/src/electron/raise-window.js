// The window that comes forward when Office is started again while it runs.
//
// Not every window is one a person can be shown. The ring for a call that
// comes with no Talk window open (notify-electron.js) and a web page read as
// a source (apps/desktop web-source.js) live in windows that are never shown
// and cannot take focus. With Office kept by its icon beside the clock and
// only such a window alive, a second start has to open Home: bringing that
// window "forward" shows the person nothing.

/** The first window that can be shown to the person, or null when there is none. */
export function windowToRaise(list = []) {
  for (const w of list) {
    try {
      if (w && !w.isDestroyed() && w.isFocusable()) return w;
    } catch {
      // A window on its way out is not one to bring forward.
    }
  }
  return null;
}
