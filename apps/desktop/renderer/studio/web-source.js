// The window's half of a web page source (main/web-source.js renders the page
// out of sight): a canvas per page, painted from the frames the main process
// sends, which the program draws from as it draws a camera.
//
// Only the part of a page that changed comes across, so the canvas is kept and
// painted over. The page's pixels arrive BGRA and premultiplied, as Chromium
// keeps them, and a canvas wants RGBA and straight: a page with a transparent
// background (?bg=none) keys over the camera properly because of that.
//
// A page is laid out at 1920 × 1080, or 1080 × 1920 for the tall picture, as the
// store's show pages expect.

const WIDE = { width: 1920, height: 1080 };
const TALL = { width: 1080, height: 1920 };

/** The page's own loop option, so the strip knows whether it greys at the ends. */
export function pageLoops(url) {
  try {
    return /^(1|true|yes|on)$/i.test(new URL(url).searchParams.get('loop') || '');
  } catch {
    return false;
  }
}

/** BGRA premultiplied, in place, to RGBA straight. */
export function toStraightRgba(px) {
  for (let i = 0; i < px.length; i += 4) {
    const b = px[i];
    const a = px[i + 3];
    if (a === 255 || a === 0) {
      px[i] = px[i + 2];
      px[i + 2] = b;
      continue;
    }
    px[i] = Math.min(255, Math.round((px[i + 2] * 255) / a));
    px[i + 1] = Math.min(255, Math.round((px[i + 1] * 255) / a));
    px[i + 2] = Math.min(255, Math.round((b * 255) / a));
  }
  return px;
}

const idFor = (url, tall) => `${url}${tall ? '#tall' : ''}`;

export class WebSources {
  /** @param {{ shell, onChange?: () => void }} o */
  constructor({ shell, onChange = null }) {
    this.shell = shell;
    this.onChange = onChange;
    this.tall = false;
    this.tallUsers = 0;
    this.items = [];
    // address#shape -> { canvas, ctx, state }
    this.pages = new Map();
    this.offFrame = shell.on('web:frame', (f) => this.frame(f));
    this.offState = shell.on('web:state', (s) => this.state(s));
  }

  changed() { this.onChange?.(); }

  /** Whether the tall picture is being made: a page is then laid out tall as well. */
  setTall(on) {
    this.tallUsers = Math.max(0, this.tallUsers + (on ? 1 : -1));
    const want = this.tallUsers > 0;
    if (this.tall === want) return;
    this.tall = want;
    this.sync(this.items);
  }

  /** Opens the pages the started items need and lets go of the rest. */
  sync(items) {
    this.items = items;
    const want = new Map();
    for (const it of items) {
      if (it.kind !== 'web' || !it.live || !it.ref) continue;
      want.set(idFor(it.ref, false), { url: it.ref, sound: it.sound, ...WIDE });
      if (this.tall) want.set(idFor(it.ref, true), { url: it.ref, sound: false, ...TALL });
    }
    for (const [id] of this.pages) {
      if (want.has(id)) continue;
      this.pages.delete(id);
      this.shell.capture.webClose({ id }).catch(() => {});
    }
    for (const [id, p] of want) {
      if (this.pages.has(id)) continue;
      const canvas = document.createElement('canvas');
      canvas.width = p.width;
      canvas.height = p.height;
      const page = { canvas, ctx: canvas.getContext('2d'), url: p.url, painted: false, state: { status: 'loading', reported: false, index: 0, count: 0, title: '', loop: pageLoops(p.url), sound: p.sound } };
      this.pages.set(id, page);
      this.shell.capture.webOpen({ id, url: p.url, width: p.width, height: p.height, sound: p.sound, loop: page.state.loop }).catch((err) => {
        page.state = { ...page.state, status: 'unavailable', error: err?.message || String(err) };
        this.changed();
      });
    }
    this.changed();
  }

  frame({ id, x, y, w, h, bytes }) {
    const page = this.pages.get(id);
    if (!page) return;
    const px = new Uint8ClampedArray(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    toStraightRgba(px);
    page.ctx.putImageData(new ImageData(px, w, h), x, y);
    page.canvas.dataUrl = null;
    // The list's small picture is made when the window draws: once more after a flip's frame has come.
    if (!page.painted || page.stale) { page.painted = true; page.stale = false; this.changed(); }
  }

  state({ id, ...s }) {
    const page = this.pages.get(id);
    if (!page) return;
    page.state = { ...page.state, ...s };
    page.stale = true;
    this.changed();
  }

  page(item) {
    return this.pages.get(idFor(item.ref, Boolean(item.tall))) || this.pages.get(idFor(item.ref, false)) || null;
  }

  /** The canvas an item is drawn from, or null before its first frame or while the page does not answer. */
  drawable(item) {
    const page = this.page(item);
    return page?.painted ? page.canvas : null;
  }

  /** ready, loading, unavailable, or off for an item stopped. */
  statusOf(item) {
    const page = this.page(item);
    return page ? page.state.status : 'off';
  }

  /** What the page last reported: { reported, index, count, title, loop, status }, or null. */
  stateOf(item) {
    return this.page(item)?.state || null;
  }

  /** A flip: the page and its tall layout together, since one page is on air in both. */
  async flip(item, verb, index) {
    const ids = [idFor(item.ref, false), ...(this.tall ? [idFor(item.ref, true)] : [])];
    for (const id of ids) if (this.pages.has(id)) await this.shell.capture.webFlip({ id, verb, index }).catch(() => {});
  }

  /** The page's own sound on or off. */
  sound(item, on) {
    const page = this.page(item);
    if (page) page.state = { ...page.state, sound: Boolean(on) };
    this.shell.capture.webSound({ id: idFor(item.ref, false), on }).catch(() => {});
  }

  dispose() {
    this.offFrame?.();
    this.offState?.();
    for (const id of this.pages.keys()) this.shell.capture.webClose({ id }).catch(() => {});
    this.pages.clear();
  }
}
