// A web page as a picture source (owner, 5 Nov: "a screen from Rutba consumer
// commerce website that can be docked in live streaming as shared screen showing
// the products by group and user able to flip next and previous").
//
// Office opens the page in a window of its own, never shown and with no access
// to Office: its own session in memory, no Node, context isolation and a
// sandbox. It may only be https, it cannot navigate away to anything else, open
// windows, download, or ask for any permission. A small bridge is loaded and
// nothing else: window.rutbaShow.state, as the store's show pages expect
// (consumer's show-deck.ts).
//
// The frames go to the studio window as they are painted. Only the part that
// changed is sent, so a page that sits still costs nothing and a flip costs one
// picture. A page that does not answer (a 404, or 503 while the store is down)
// is "Not available" and tried again; its error page is never shown on air.
//
// Flipping: a page that reports its state is sent the contract's messages. One
// that has never reported is sent the arrow keys, and the strip then has only
// Previous and Next.

import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import electron from 'electron';
import { sendEvent } from '@rutba/office-shell/electron/main';

const { BrowserWindow, session } = electron;
const here = path.dirname(fileURLToPath(import.meta.url));

const VERBS = new Set(['next', 'previous', 'first', 'last', 'go']);
/** How long to wait before trying a page again that did not answer, in seconds. */
const RETRY_S = [5, 10, 20, 30];

/** Why an address cannot be a source, or null. `allowLoopback` is for the checks' local page. */
export function refuseAddress(raw, { allowLoopback = false } = {}) {
  let u;
  try { u = new URL(String(raw)); } catch { return 'not-an-address'; }
  if (u.protocol === 'https:') return null;
  if (allowLoopback && u.protocol === 'http:' && u.hostname === '127.0.0.1') return null;
  return 'not-https';
}

/**
 * The address a page opens at, with its options kept. Tall pages use the same
 * address: the store's show pages lay themselves out for 1080 × 1920.
 */
export function createWebSources({ allowLoopback = false, frameRate = 30 } = {}) {
  /** id -> the page's state */
  const pages = new Map();
  /** The windows whose pages are let go when they close or reload. */
  const watched = new WeakSet();

  const tell = (page, event, payload) => {
    if (!page.win.isDestroyed()) sendEvent(page.owner, event, { id: page.id, ...payload });
  };

  const stateOf = (page) => ({ status: page.status, reported: page.reported, index: page.index, count: page.count, title: page.title, loop: page.loop, width: page.width, height: page.height, sound: page.sound });
  const tellState = (page) => tell(page, 'web:state', stateOf(page));

  /** The page did not answer: say so, and try again after a while. */
  function unavailable(page) {
    if (page.closed) return;
    page.status = 'unavailable';
    page.reported = false;
    tellState(page);
    clearTimeout(page.retry);
    const wait = RETRY_S[Math.min(page.tries, RETRY_S.length - 1)];
    page.tries += 1;
    page.retry = setTimeout(() => { if (!page.closed && !page.win.isDestroyed()) page.win.webContents.loadURL(page.url).catch(() => {}); }, wait * 1000);
  }

  function open({ id, url, width = 1920, height = 1080, sound = false, loop = false }, owner) {
    const refused = refuseAddress(url, { allowLoopback });
    if (refused) throw new Error(refused === 'not-https' ? 'A web page source must start https://.' : 'That is not a web address.');
    if (owner && !watched.has(owner)) {
      watched.add(owner);
      owner.once('closed', () => closeAll(owner));
      owner.webContents.on('did-navigate', () => closeAll(owner));
    }
    const was = pages.get(id);
    if (was && was.url === url && was.width === width && was.height === height) {
      was.owner = owner;
      if (!was.win.isDestroyed()) { was.win.webContents.invalidate(); tellState(was); }
      return stateOf(was);
    }
    if (was) close({ id });

    // Nothing shared with Office or any other page: an in-memory session of its own.
    const partition = `web-source-${crypto.randomBytes(8).toString('hex')}`;
    const ses = session.fromPartition(partition, { cache: false });
    ses.setPermissionRequestHandler((_wc, _permission, callback) => callback(false));
    ses.setPermissionCheckHandler(() => false);
    ses.on('will-download', (e) => e.preventDefault());
    ses.webRequest.onBeforeRequest((details, callback) => {
      const ok = /^(https:|data:|blob:)/i.test(details.url) || (allowLoopback && details.url.startsWith('http://127.0.0.1'));
      callback({ cancel: !ok });
    });

    const win = new BrowserWindow({
      show: false,
      width,
      height,
      useContentSize: true,
      frame: false,
      transparent: true,
      resizable: false,
      skipTaskbar: true,
      focusable: false,
      paintWhenInitiallyHidden: true,
      backgroundColor: '#00000000',
      webPreferences: {
        offscreen: true,
        sandbox: true,
        contextIsolation: true,
        nodeIntegration: false,
        webSecurity: true,
        session: ses,
        preload: path.join(here, 'web-source-preload.cjs'),
        backgroundThrottling: false,
        devTools: false,
        spellcheck: false,
      },
    });
    win.setMenu?.(null);
    const wc = win.webContents;
    wc.setAudioMuted(!sound);
    wc.setFrameRate(frameRate);

    const page = { id, url, width, height, sound, loop, win, owner, closed: false, status: 'loading', reported: false, index: 0, count: 0, title: '', tries: 0, retry: null, painted: false, partition };
    pages.set(id, page);

    wc.setWindowOpenHandler(() => ({ action: 'deny' }));
    const stays = (next) => !refuseAddress(next, { allowLoopback });
    wc.on('will-navigate', (e, next) => { if (!stays(next)) e.preventDefault(); });
    wc.on('will-redirect', (e, next) => { if (!stays(next)) e.preventDefault(); });

    // The page's own report, through the bridge: this page's and nobody else's.
    wc.ipc.on('rutba-show:state', (_e, s) => {
      if (page.closed) return;
      page.reported = true;
      page.index = Number.isInteger(s?.index) ? s.index : 0;
      page.count = Number.isInteger(s?.count) ? s.count : 0;
      page.title = String(s?.title ?? '').slice(0, 200);
      tellState(page);
    });

    wc.on('did-navigate', (_e, _url, code) => {
      if (code >= 400) { page.badResponse = true; unavailable(page); } else page.badResponse = false;
    });
    wc.on('did-fail-load', (_e, code, _desc, _url, mainFrame) => {
      if (mainFrame && code !== -3) unavailable(page);
    });
    wc.on('render-process-gone', () => unavailable(page));
    wc.on('did-finish-load', () => {
      if (page.badResponse) return;
      page.status = 'ready';
      page.tries = 0;
      wc.invalidate();
      tellState(page);
      // The page has its state by now, or says so when asked.
      wc.send('rutba-show:command', { rutbaShow: 'report' });
    });

    // Only the rectangle that changed; while the page is an error page, nothing.
    wc.on('paint', (_e, dirty, image) => {
      if (page.closed || page.status === 'unavailable' || page.badResponse) return;
      if (!dirty.width || !dirty.height) return;
      page.painted = true;
      tell(page, 'web:frame', { x: dirty.x, y: dirty.y, w: dirty.width, h: dirty.height, width: page.width, height: page.height, bytes: new Uint8Array(image.crop(dirty).toBitmap()) });
    });

    wc.loadURL(url).catch(() => {});
    return stateOf(page);
  }

  function close({ id }) {
    const page = pages.get(id);
    if (!page) return { closed: false };
    page.closed = true;
    clearTimeout(page.retry);
    pages.delete(id);
    if (!page.win.isDestroyed()) page.win.destroy();
    session.fromPartition(page.partition, { cache: false }).clearStorageData().catch(() => {});
    return { closed: true };
  }

  /** One flip: the page reports, so a message; a page that never has gets the arrow keys. */
  function flip({ id, verb, index }) {
    const page = pages.get(id);
    if (!page || page.closed || page.win.isDestroyed()) return { sent: false };
    if (!VERBS.has(verb)) throw new Error('Not a way to flip a page.');
    const wc = page.win.webContents;
    if (page.reported) {
      wc.send('rutba-show:command', verb === 'go' ? { rutbaShow: 'go', index: Number(index) } : { rutbaShow: verb });
      return { sent: true, via: 'message' };
    }
    if (verb === 'next' || verb === 'previous') {
      const keyCode = verb === 'next' ? 'Right' : 'Left';
      wc.sendInputEvent({ type: 'keyDown', keyCode });
      wc.sendInputEvent({ type: 'keyUp', keyCode });
      return { sent: true, via: 'key' };
    }
    return { sent: false };
  }

  /** The page's own sound, off until it is turned on. */
  function sound({ id, on }) {
    const page = pages.get(id);
    if (!page || page.win.isDestroyed()) return { on: false };
    page.sound = Boolean(on);
    page.win.webContents.setAudioMuted(!page.sound);
    tellState(page);
    return { on: page.sound };
  }

  function closeAll(owner) {
    for (const page of [...pages.values()]) if (!owner || page.owner === owner) close({ id: page.id });
  }

  return {
    /** ({ id, url, width, height, sound, loop }, win) -> state: opened (or, the same page, asked to send its picture again) */
    webOpen: open,
    /** ({ id }) -> { closed } */
    webClose: close,
    /** ({ id, verb: 'next' | 'previous' | 'first' | 'last' | 'go', index? }) -> { sent, via } */
    webFlip: flip,
    /** ({ id, on }) -> { on } */
    webSound: sound,
    /** () -> [state] for the checks */
    webPages: () => [...pages.values()].map((p) => ({ id: p.id, ...stateOf(p) })),
    closeAll,
  };
}
