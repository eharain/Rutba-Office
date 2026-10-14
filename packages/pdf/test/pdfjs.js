'use strict';

// pdf.js, the reader the window draws with (vendored in apps/desktop/vendor),
// loaded in Node for the tests: a second reader's word on what the writer
// and the updates make. Its browser build wants a few things this Node lacks
// (DOMMatrix, Path2D and ImageData at load time; Uint8Array's toHex,
// Promise.try and Map's getOrInsertComputed as it reads), none of which
// reading text, annotations, fields or outlines draws with: they are given
// as stand-ins. Resolves to null when the vendored copy is not there.

const path = require('path');
const fs = require('fs');
const { pathToFileURL } = require('url');

const ROOT = path.join(__dirname, '..', '..', '..', 'apps', 'desktop', 'vendor', 'pdfjs');

let loading = null;
function loadPdfjs() {
  if (!loading) {
    loading = (async () => {
      if (!fs.existsSync(path.join(ROOT, 'pdf.min.mjs'))) return null;
      globalThis.DOMMatrix ||= class DOMMatrix {
        constructor(m) { [this.a, this.b, this.c, this.d, this.e, this.f] = Array.isArray(m) ? m : [1, 0, 0, 1, 0, 0]; }
      };
      globalThis.Path2D ||= class Path2D {};
      globalThis.ImageData ||= class ImageData {};
      if (!Uint8Array.prototype.toHex) Object.defineProperty(Uint8Array.prototype, 'toHex', { value() { return Buffer.from(this.buffer, this.byteOffset, this.byteLength).toString('hex'); }, configurable: true });
      for (const M of [Map, WeakMap]) {
        if (!M.prototype.getOrInsertComputed) Object.defineProperty(M.prototype, 'getOrInsertComputed', { value(key, make) { if (!this.has(key)) this.set(key, make(key)); return this.get(key); }, configurable: true });
        if (!M.prototype.getOrInsert) Object.defineProperty(M.prototype, 'getOrInsert', { value(key, v) { if (!this.has(key)) this.set(key, v); return this.get(key); }, configurable: true });
      }
      if (!Promise.try) Promise.try = (fn, ...args) => new Promise((resolve) => resolve(fn(...args)));
      const [log, warn] = [console.log, console.warn];
      const quiet = (to) => (...a) => { if (!/legacy/.test(String(a[0]))) to(...a); };
      console.log = quiet(log);
      console.warn = quiet(warn);
      try {
        const lib = await import(pathToFileURL(path.join(ROOT, 'pdf.min.mjs')).href);
        lib.GlobalWorkerOptions.workerSrc = pathToFileURL(path.join(ROOT, 'pdf.worker.min.mjs')).href;
        return lib;
      } finally {
        console.log = log;
        console.warn = warn;
      }
    })();
  }
  return loading;
}

/** Every page's words and annotations as pdf.js reads them. */
async function readWithPdfjs(bytes) {
  const lib = await loadPdfjs();
  if (!lib) return null;
  const task = lib.getDocument({ data: new Uint8Array(bytes), verbosity: 0, standardFontDataUrl: `${ROOT.replace(/\\/g, '/')}/standard_fonts/` });
  const doc = await task.promise;
  const pages = [];
  try {
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i);
      const tc = await page.getTextContent();
      const annotations = await page.getAnnotations();
      pages.push({ text: tc.items.map((it) => it.str).join(' '), annotations });
    }
    const fields = await doc.getFieldObjects();
    const outline = await doc.getOutline();
    const labels = await doc.getPageLabels();
    const meta = await doc.getMetadata();
    const metadata = meta.metadata ? JSON.stringify(meta.metadata.getAll ? meta.metadata.getAll() : meta.metadata) : '';
    return { pages, fields, outline, labels, info: meta.info, metadata, numPages: doc.numPages };
  } finally {
    await task.destroy();
  }
}

module.exports = { loadPdfjs, readWithPdfjs };
