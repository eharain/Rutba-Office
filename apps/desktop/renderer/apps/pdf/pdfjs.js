// Drawing PDF pages in the window: pdf.js (Mozilla, Apache-2.0), vendored in
// apps/desktop/vendor/pdfjs and served from the app's own files, loaded the
// first time a PDF is shown. It parses and draws in a worker of its own, so
// a large file never holds the window; nothing here reaches the network.
//
// One pdf.js document per source file, pages fetched once, and every draw
// through one queue that does the pages on screen first and drops a draw
// nobody is waiting for any more.

const base = () => new URL('pdfjs/', window.location.href).href;

let loading = null;
/** The pdf.js library, loaded once. */
export function pdfjs() {
  if (!loading) {
    loading = (async () => {
      const url = `${base()}pdf.min.mjs`;
      const lib = await import(/* the app's own file, never bundled */ url);
      lib.GlobalWorkerOptions.workerSrc = `${base()}pdf.worker.min.mjs`;
      return lib;
    })();
    loading.catch(() => { loading = null; });
  }
  return loading;
}

/**
 * Open bytes as a pdf.js document. A file that needs a password to be read
 * at all rejects with `{ password: true }`.
 */
export async function openDocument(bytes) {
  const lib = await pdfjs();
  const task = lib.getDocument({
    // pdf.js takes the buffer over to its worker: it is given a copy.
    data: bytes.slice(),
    standardFontDataUrl: `${base()}standard_fonts/`,
    wasmUrl: `${base()}wasm/`,
    iccUrl: `${base()}iccs/`,
    isEvalSupported: false,
    enableXfa: false,
    verbosity: 0,
  });
  task.onPassword = () => {
    task.destroy();
  };
  try {
    return await task.promise;
  } catch (err) {
    if (err?.name === 'PasswordException' || /password/i.test(err?.message || '')) {
      const e = new Error('password');
      e.password = true;
      throw e;
    }
    throw err;
  }
}

/** Page proxies by document and index, fetched once. */
const pageCache = new WeakMap();
export function getPage(doc, index) {
  let pages = pageCache.get(doc);
  if (!pages) { pages = new Map(); pageCache.set(doc, pages); }
  if (!pages.has(index)) pages.set(index, doc.getPage(index + 1));
  return pages.get(index);
}

/** A page's words as pdf.js reads them: the string find searches, and where each text item starts in it. */
const textCache = new WeakMap();
export function pageText(doc, index) {
  let texts = textCache.get(doc);
  if (!texts) { texts = new Map(); textCache.set(doc, texts); }
  if (!texts.has(index)) {
    texts.set(index, (async () => {
      const page = await getPage(doc, index);
      const content = await page.getTextContent();
      let text = '';
      for (const item of content.items) {
        if (typeof item.str !== 'string') continue;
        text += item.str;
        if (item.hasEOL) text += '\n';
      }
      return text;
    })());
  }
  return texts.get(index);
}

// ── the queue ─────────────────────────────────────────────────────────────

const queue = [];
let running = 0;
const MAX_AT_ONCE = 2;

function pump() {
  while (running < MAX_AT_ONCE && queue.length) {
    queue.sort((a, b) => a.priority - b.priority);
    const job = queue.shift();
    if (job.cancelled) { job.reject(Object.assign(new Error('cancelled'), { cancelled: true })); continue; }
    running += 1;
    job.run().then(job.resolve, job.reject).finally(() => { running -= 1; pump(); });
  }
}

/**
 * Draw a page onto a canvas, `scale` CSS pixels a point, turned `rotate`
 * degrees, sharp for the screen's pixel ratio. Returns a handle whose
 * `cancel()` drops a draw not yet begun and stops one under way.
 */
export function drawPage(doc, index, canvas, { scale, rotate = 0, priority = 1 }) {
  let task = null;
  const job = { priority, cancelled: false };
  job.run = async () => {
    const page = await getPage(doc, index);
    if (job.cancelled) throw Object.assign(new Error('cancelled'), { cancelled: true });
    const ratio = Math.min(3, window.devicePixelRatio || 1);
    const viewport = page.getViewport({ scale: scale * ratio, rotation: rotate });
    const off = document.createElement('canvas');
    off.width = Math.max(1, Math.floor(viewport.width));
    off.height = Math.max(1, Math.floor(viewport.height));
    const ctx = off.getContext('2d', { alpha: false });
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, off.width, off.height);
    task = page.render({ canvasContext: ctx, canvas: off, viewport, background: '#ffffff' });
    await task.promise;
    if (job.cancelled) throw Object.assign(new Error('cancelled'), { cancelled: true });
    // Drawn off screen and copied over at once, so a page never shows half drawn.
    canvas.width = off.width;
    canvas.height = off.height;
    canvas.getContext('2d').drawImage(off, 0, 0);
    return { width: off.width, height: off.height };
  };
  const done = new Promise((resolve, reject) => { job.resolve = resolve; job.reject = reject; });
  queue.push(job);
  pump();
  return {
    done,
    cancel() {
      job.cancelled = true;
      try { task?.cancel(); } catch { /* finished */ }
    },
  };
}

/**
 * The page's words laid over it as selectable, findable text (pdf.js's text
 * layer): positioned by the same viewport the canvas was drawn with.
 */
export async function drawText(doc, index, container, { scale, rotate = 0 }) {
  // Drawn again before the last draw ended (the zoom changed as a pane opened): the last one is
  // stopped, or both would add their words to the one layer and every word would be there twice.
  const token = (container.rutbaTextDraw = (container.rutbaTextDraw || 0) + 1);
  const lib = await pdfjs();
  const page = await getPage(doc, index);
  if (container.rutbaTextDraw !== token) return null;
  container.rutbaTextLayer?.cancel?.();
  const viewport = page.getViewport({ scale, rotation: rotate });
  container.replaceChildren();
  container.style.setProperty('--total-scale-factor', String(scale));
  container.style.setProperty('--scale-factor', String(scale));
  const layer = new lib.TextLayer({ textContentSource: page.streamTextContent(), container, viewport });
  container.rutbaTextLayer = layer;
  try {
    await layer.render();
  } catch (err) {
    if (container.rutbaTextDraw !== token) return null;
    throw err;
  }
  if (container.rutbaTextDraw !== token) return null;
  const end = document.createElement('div');
  end.className = 'endOfContent';
  container.append(end);
  return layer;
}

/** A page drawn to a JPEG at `dpi`, for printing. */
export async function pageImage(doc, index, { rotate = 0, dpi = 150 }) {
  const page = await getPage(doc, index);
  const viewport = page.getViewport({ scale: dpi / 72, rotation: rotate });
  const canvas = document.createElement('canvas');
  canvas.width = Math.floor(viewport.width);
  canvas.height = Math.floor(viewport.height);
  const ctx = canvas.getContext('2d', { alpha: false });
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, canvas, viewport, background: '#ffffff' }).promise;
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', 0.92));
  return { url: URL.createObjectURL(blob), width: viewport.width / (dpi / 72), height: viewport.height / (dpi / 72) };
}
