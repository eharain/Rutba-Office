// Equations for paper: the same MathML the page draws, laid out by the same
// Chromium, measured and pictured for the PDF writer.
//
// The document engine prints through its own PDF writer, which draws text in
// the base-14 fonts and boxes — it cannot set a radical or a stacked limit,
// and it has no glyph for √ or ∑ at all. Rather than teach it a second,
// lesser math layout, a printout asks Chromium: each equation goes into a
// hidden, off-screen page in the Cambria Math face the Word window uses, its
// box and baseline are read, and the page is captured at three times its
// size. The writer then places that picture where the paginator left room
// for it, so paper shows exactly what the screen shows.
//
// One hidden window per printout, closed when the equations are measured.
// Nothing it loads runs anything but the few lines that measure.

const REF_PX = 16; // the font size the measurements are taken at, so they come back in ems
const SCALE = 3; // the picture is drawn this many times larger than it prints, for a crisp page

const PAGE = `<!doctype html><html><head><meta charset="utf-8"><style>
html, body { margin: 0; padding: 0; background: #ffffff; color: #000000; }
.eq { position: absolute; left: 0; display: inline-block; padding: 0.15em 0.18em; font-size: ${REF_PX * SCALE}px; line-height: normal; white-space: nowrap; }
math { font-family: 'Cambria Math', 'STIX Two Math', 'Latin Modern Math', math; font-size: 1.04em; display: inline math; }
.eq.display math { math-style: normal; }
.eq.inline math { math-style: compact; }
.base { display: inline-block; width: 0; height: 0; }
</style></head><body></body></html>`;

/**
 * @returns {{ measure: (list: Array<{key:string, mathml:string, display:boolean}>) => Promise<Map<string, object>>, close: () => void }}
 */
export function createMathMeasurer() {
  let win = null;
  let chain = Promise.resolve();

  async function ensure() {
    if (win && !win.isDestroyed()) return win;
    const { BrowserWindow } = await import('electron');
    win = new BrowserWindow({
      show: false,
      width: 2400,
      height: 1600,
      webPreferences: { offscreen: true, sandbox: true, contextIsolation: true, nodeIntegration: false, backgroundThrottling: false },
    });
    await win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(PAGE));
    return win;
  }

  const close = () => {
    if (win && !win.isDestroyed()) win.destroy();
    win = null;
  };

  /** Lay out a batch, stacked down the page, and read each one's box and baseline. */
  async function layBatch(w, batch) {
    const html = batch.map((item, i) =>
      `<div class="eq ${item.display ? 'display' : 'inline'}" id="eq${i}">${item.mathml}<span class="base"></span></div>`
    ).join('');
    return w.webContents.executeJavaScript(`(async () => {
      document.body.innerHTML = ${JSON.stringify(html)};
      let y = 4;
      const boxes = [];
      for (const el of document.querySelectorAll('.eq')) {
        el.style.top = y + 'px';
        const r = el.getBoundingClientRect();
        y += Math.ceil(r.height) + 8;
      }
      await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
      for (const el of document.querySelectorAll('.eq')) {
        const r = el.getBoundingClientRect();
        const b = el.querySelector('.base').getBoundingClientRect();
        boxes.push({ x: r.left, y: r.top, w: r.width, h: r.height, base: b.top - r.top });
      }
      return { boxes, height: y };
    })()`);
  }

  async function measureNow(list) {
    const out = new Map();
    if (!list.length) return out;
    const w = await ensure();
    // Batches whose stacked height the window can hold in one frame.
    let at = 0;
    while (at < list.length) {
      const batch = list.slice(at, at + 12);
      at += batch.length;
      const { boxes, height } = await layBatch(w, batch);
      if (height > 1600) w.setContentSize(2400, Math.min(16000, Math.ceil(height) + 20));
      w.webContents.invalidate();
      await new Promise((r) => setTimeout(r, 120));
      for (let i = 0; i < batch.length; i++) {
        const box = boxes[i];
        if (!box || box.w < 1 || box.h < 1) continue;
        const rect = { x: Math.floor(box.x), y: Math.floor(box.y), width: Math.ceil(box.w), height: Math.ceil(box.h) };
        const image = await w.webContents.capturePage(rect);
        const size = image.getSize();
        const bitmap = image.toBitmap();
        // The bitmap is in device pixels, which on a scaled display are more
        // than the rect's: its own length says how many.
        const factor = size.width && size.height ? Math.sqrt(bitmap.length / (4 * size.width * size.height)) : 1;
        const pw = Math.round(size.width * factor);
        const ph = Math.round(size.height * factor);
        if (!pw || !ph || bitmap.length < pw * ph * 4) continue;
        const rgb = Buffer.alloc(pw * ph * 3);
        for (let p = 0, q = 0; p < pw * ph; p++, q += 4) {
          // BGRA, as Chromium hands it back on Windows and Linux.
          rgb[p * 3] = bitmap[q + 2];
          rgb[p * 3 + 1] = bitmap[q + 1];
          rgb[p * 3 + 2] = bitmap[q];
        }
        const em = REF_PX * SCALE;
        out.set(batch[i].key, {
          widthEm: box.w / em,
          heightEm: box.h / em,
          baselineEm: box.base / em,
          raster: { width: pw, height: ph, data: rgb },
        });
      }
    }
    return out;
  }

  return {
    /** Measure and picture a list of equations; answers by key. One request at a time. */
    measure(list) {
      // The window goes as soon as the job is done: a hidden window left
      // standing would keep the app running after its last real window closed.
      const run = chain.then(() => measureNow(list)).finally(close);
      chain = run.catch(() => {});
      return run;
    },
    close,
  };
}
