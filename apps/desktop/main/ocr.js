// Text recognition for the windows: the PDF app's Recognise Text and its
// fields, Copy Text from Picture in Pictures and Image, Data from Picture and
// Receipt or Invoice in Worksheets, an attached invoice's details in Mail.
//
// A window hands over a picture's pixels and gets back the text, its lines
// and its words with their boxes. The reading is done by @rutba/ocr's pool
// of threads, never on this process's thread and never on the window's, so
// a three-hundred-page scan holds up neither the other windows nor the one
// that asked. A window names each job, so it can cancel one; a window that
// closes cancels whatever it was still waiting for.

import { ocrHost, LANGUAGES } from '@rutba/ocr';

export function createOcrService({ host = ocrHost() } = {}) {
  /** job name -> { controller, owner } */
  const jobs = new Map();
  const watched = new WeakSet();

  const watch = (win) => {
    if (!win || watched.has(win)) return;
    watched.add(win);
    win.once('closed', () => {
      for (const [name, job] of jobs) if (job.owner === win) { job.controller.abort(); jobs.delete(name); }
    });
  };

  return {
    /** The languages there are models for: [{ code, name, script, direction, beta? }]. */
    languages: () => LANGUAGES.map((l) => ({ ...l })),

    /**
     * ({ image: { width, height, data } | PNG bytes, languages, dpi, mode, table, document, job })
     * -> { text, lines, words, width, height, languages, angle, ms, table?, document? }
     * `document`: the picture read as a bill, an invoice, a receipt, a letter
     * or a form; its kind, fields, items, sums and tick boxes come back too.
     */
    async recognize({ image, languages, dpi, mode, table, document, job } = {}, win) {
      const controller = new AbortController();
      const name = job ? `${win?.id ?? 0}:${job}` : null;
      if (name) {
        jobs.get(name)?.controller.abort();
        jobs.set(name, { controller, owner: win });
      }
      watch(win);
      try {
        const result = await host.recognize(image, { languages, dpi, mode, table: Boolean(table), document: Boolean(document), signal: controller.signal });
        // The second readings the fields were made from stay here: the window has no use for them.
        if (result.document) delete result.document.again;
        return result;
      } finally {
        if (name && jobs.get(name)?.controller === controller) jobs.delete(name);
      }
    },

    /** ({ job }) -> { cancelled }: stop a job this window named, waiting or reading. */
    cancel({ job } = {}, win) {
      const name = `${win?.id ?? 0}:${job}`;
      const found = jobs.get(name);
      found?.controller.abort();
      jobs.delete(name);
      return { cancelled: Boolean(found) };
    },

    /** The pool's threads and counts, for the checks. */
    stats: () => host.stats,
  };
}
