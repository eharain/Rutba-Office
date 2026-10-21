// The features the website shows, each as `{ name, app, file, setup }`.
//
// `file` is a sample the screens build wrote (build/make-screens.js and
// build/feature-samples.js), by its name in the samples folder; `setup(h)`
// drives the opened window into the state that shows the feature with the
// harness of screen-run.js (a tab, a button by its tip or words, a field, a
// key, a drag) and may answer another window to photograph instead.
// Optional: `before(b)` for what has to be in place before the window opens
// (a message delivered, ratings stored), `query` for the window's query,
// `boot` ({ theme, language, group }) for what windows read as they start
// (a run of its own per boot), `size` for a window size of its own,
// `settle` for the pause before the setup, `done()` for what to stop after
// the picture (the AI app connected for AI Tools' picture).
//
// Every person, firm and figure in them is invented.
//
// Kept free of Electron imports, so make-screens.js can read the names and
// the boot settings from plain Node.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

/** The Talk picture's stand-in introducer, kept until its picture is taken. */
let talkStandIn = null;

/** Words selected in a PDF page's text layer, from one word to another. */
const pdfSelect = (h, index, from, to) => h.js(`(() => {
  const layer = document.querySelector('.pdf-page[data-index="${index}"] .textLayer[data-ready]');
  if (!layer) return 'no text';
  const nodes = [];
  const walker = document.createTreeWalker(layer, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) nodes.push(n);
  const a = nodes.find((n) => n.data.includes(${JSON.stringify(from)}));
  const b = nodes.find((n) => n.data.includes(${JSON.stringify(to)}));
  if (!a || !b) return 'no words';
  const r = document.createRange();
  r.setStart(a, a.data.indexOf(${JSON.stringify(from)}));
  r.setEnd(b, b.data.indexOf(${JSON.stringify(to)}) + ${JSON.stringify(to)}.length);
  getSelection().removeAllRanges();
  getSelection().addRange(r);
  return getSelection().toString();
})()`);

/** A drag on a PDF page's comment layer, through points given as fractions of the page. */
const pdfDrag = (h, index, points) => h.js(`(async () => {
  const svg = document.querySelector('.pdf-page[data-index="${index}"] .pdf-annots');
  if (!svg) return 'no layer';
  const r = svg.getBoundingClientRect();
  const at = ([fx, fy]) => ({ clientX: r.left + r.width * fx, clientY: r.top + r.height * fy, bubbles: true, cancelable: true, button: 0, buttons: 1 });
  const pts = ${JSON.stringify(points)};
  svg.dispatchEvent(new MouseEvent('mousedown', at(pts[0])));
  for (const p of pts.slice(1)) { window.dispatchEvent(new MouseEvent('mousemove', at(p))); await new Promise((res) => setTimeout(res, 20)); }
  window.dispatchEvent(new MouseEvent('mouseup', at(pts[pts.length - 1])));
  return 'dragged';
})()`);

/** A control by its data-role, once it is there and enabled. */
const role = async (h, r) => {
  const sel = `[data-role="${r}"]`;
  await h.until(`Boolean(document.querySelector(${JSON.stringify(sel)}) && !document.querySelector(${JSON.stringify(sel)}).disabled)`, `the ${r} control`, 6000);
  await h.js(`document.querySelector(${JSON.stringify(sel)}).click(), 1`);
  await h.wait(250);
};

/** A PDF tool put down again, so the page shows no tool's cursor or pressed button. */
const putDown = (h, r) => h.js(`(() => { const b = document.querySelector('[data-role="${r}"]'); if (b && (b.getAttribute('aria-pressed') === 'true' || b.classList.contains('pressed') || b.classList.contains('on'))) b.click(); return 1; })()`);

const pdfReady = (h) => h.until(`document.querySelectorAll('.pdf-page canvas[data-drawn]').length >= 1 && Boolean(document.querySelector('.pdf-page[data-index="0"] .pdf-annots[data-ready]')) && Boolean(document.querySelector('.pdf-page[data-index="0"] .textLayer[data-ready]'))`, 'the first page drawn', 20000);

/** Sets a field the way typing does, then leaves it, for the PDF's own inputs. */
const pdfField = (h, sel, value) => h.js(`(() => {
  const el = document.querySelector(${JSON.stringify(sel)});
  if (!el) return false;
  el.focus();
  const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
  Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)});
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.blur();
  return true;
})()`);

/** Text recognition in English only, as the checks pin it. */
const englishOnly = (h) => h.js(`localStorage.setItem('rutba-ocr-languages', '["eng"]'), 1`);

/** Presses an element the way the forms check does: down, then a click. */
const press = (h, selector) => h.js(`(() => {
  const el = document.querySelector(${JSON.stringify(selector)});
  if (!el) return false;
  el.scrollIntoView({ block: 'center' });
  el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 }));
  el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, button: 0 }));
  return true;
})()`);

/** An image tool taken up, as the image checks take one. */
const tool = async (h, id) => {
  await h.until(`Boolean(document.querySelector('.im-tool[data-tool-id="${id}"]:not(:disabled)'))`, `the ${id} tool`, 5000);
  await h.js(`document.querySelector('.im-tool[data-tool-id="${id}"]').click(), 1`);
  await h.until(`document.querySelector('.im-tool[data-tool-id="${id}"]')?.getAttribute('aria-pressed') === 'true'`, `the ${id} tool in hand`, 5000);
};

/** A press, moves and a release on the image's tool stage, at picture pixels. */
const gesture = (h, points, width) => h.js(`(() => {
  const el = document.querySelector('.im-tool-stage');
  if (!el) return 'no stage';
  const r = el.getBoundingClientRect();
  const k = r.width / ${width};
  const pts = ${JSON.stringify(points)};
  const at = (p) => ({ clientX: r.left + p[0] * k, clientY: r.top + p[1] * k, bubbles: true, cancelable: true, pointerId: 7, pointerType: 'mouse', pressure: 0.5, button: 0, buttons: 1, isPrimary: true });
  el.dispatchEvent(new PointerEvent('pointerdown', at(pts[0])));
  for (const p of pts.slice(1)) el.dispatchEvent(new PointerEvent('pointermove', at(p)));
  el.dispatchEvent(new PointerEvent('pointerup', at(pts[pts.length - 1])));
  return 'done';
})()`);

const imageReady = (h) => h.until(`Number(document.querySelector('.im-canvas')?.dataset.drawn || 0) > 0`, 'the picture on the canvas', 15000);

/** Waits for the image canvas to stop redrawing. */
const imageSettled = async (h) => {
  let last = -1;
  await h.untilHere(async () => {
    const n = await h.js(`Number(document.querySelector('.im-canvas')?.dataset.drawn || 0)`);
    const still = n === last;
    last = n;
    await h.wait(250);
    return still;
  }, 'the canvas settled', 10000).catch(() => {});
};

const LANDSCAPE = path.join('Sample Pictures', 'Lakeside at dawn.png');

/**
 * Recorder recording a document: the quarterly review opened in Documents,
 * Window chosen and that window picked, and the program drawn with it.
 */
async function recordWindowIn(h) {
  const doc = h.windows.create({ app: 'word', file: h.sample('quarterly-review.docx') });
  await new Promise((resolve) => doc.webContents.once('did-finish-load', resolve));
  await h.untilHere(() => doc.webContents.executeJavaScript(`Boolean(document.querySelector('.wd-block'))`), 'the document to record', 12000);
  await h.wait(1200);
  await h.until(`Boolean(document.querySelector('.st-kinds button[data-value="window"]'))`, 'the source kinds', 8000);
  await h.click('.st-kinds button[data-value="window"]', 'Window');
  // The list was read when the window opened, before the document was: read again.
  await h.button('Refresh the list');
  await h.until(`[...document.querySelectorAll('.st-sources .rec-source')].some((n) => /quarterly-review/.test(n.dataset.tip || ''))`, 'the document in the list', 10000);
  await h.js(`(() => { [...document.querySelectorAll('.st-sources .rec-source')].find((n) => /quarterly-review/.test(n.dataset.tip || ''))?.click(); return 1; })()`);
  await h.until(`(document.querySelector('.st-canvas')?.dataset.layers || '').includes('screen')`, 'the window in the program', 20000);
}

/** Video's playhead moved, as dragging the slider under the preview moves it. */
const scrubVideo = async (h, t) => {
  await h.js(`(() => { const s = document.querySelector('.vd-scrub'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, ${JSON.stringify(String(t))}); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
  await h.until(`Math.abs(Number(document.querySelector('.vd-scrub').value) - ${t}) < 0.05`, 'the playhead to move', 3000).catch(() => {});
  await h.wait(800);
};

/**
 * An AI app connected to the AI tools, for their picture: it starts the
 * tools as AI Tools says to, and asks them for a document, a workbook and a
 * presentation in the folder they were given, so the window has a
 * connection and a log to show. Closed by the feature's `done`.
 */
let aiApp = null;
async function connectAiApp({ profile, folder }) {
  const { spawn } = await import('node:child_process');
  const node = path.join(path.dirname(fileURLToPath(import.meta.url)), 'mcp', 'node.js');
  aiApp = spawn(process.execPath, [node], { stdio: ['pipe', 'pipe', 'pipe'], env: { ...process.env, ELECTRON_RUN_AS_NODE: '1', RUTBA_OFFICE_PROFILE: profile } });
  let out = '';
  const waiting = new Map();
  aiApp.stdout.on('data', (d) => {
    out += d;
    let at;
    while ((at = out.indexOf('\n')) >= 0) {
      const line = out.slice(0, at);
      out = out.slice(at + 1);
      try {
        const m = JSON.parse(line);
        waiting.get(m.id)?.(m);
      } catch {
        /* not an answer */
      }
    }
  });
  let id = 0;
  const ask = (method, params) => new Promise((resolve) => {
    const my = ++id;
    const timer = setTimeout(() => resolve(null), 30000);
    waiting.set(my, (m) => { clearTimeout(timer); resolve(m); });
    aiApp.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: my, method, params })}\n`);
  });
  const call = (name, args) => ask('tools/call', { name, arguments: args });
  await ask('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'Desktop assistant', version: '2.1' } });
  await call('list_folders', {});
  await call('create_workbook', {
    path: 'Freight figures.xlsx',
    sheets: [{ name: 'By region', rows: [['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Year'], ['North', 142000, 161000, 158500, 174200, '=SUM(B2:E2)'], ['South', 86000, 91000, 94750, 98300, '=SUM(B3:E3)'], ['East', 51200, 57400, 60100, 66600, '=SUM(B4:E4)'], ['West', 73900, 70200, 76800, 81500, '=SUM(B5:E5)']] }],
  });
  await call('create_document', {
    path: 'Board summary.docx',
    title: 'Board summary: the year in freight',
    theme: 'harbour',
    blocks: [
      { type: 'paragraph', text: 'Prepared from Freight figures.xlsx for the November board. Figures are illustrative.' },
      { type: 'heading', text: 'Highlights', level: 1 },
      { type: 'bullets', items: ['The north grew fastest, adding £32,200 between the first and fourth quarters.', 'The east finished 30% ahead of where it started, without a down quarter.', 'The Leeds depot opened in September, on time and under budget.'] },
      { type: 'heading', text: 'Revenue by region', level: 1 },
      { type: 'table', rows: [['Region', 'Q1', 'Q4', 'Change'], ['North', '£142,000', '£174,200', '+22.7%'], ['South', '£86,000', '£98,300', '+14.3%'], ['East', '£51,200', '£66,600', '+30.1%'], ['West', '£73,900', '£81,500', '+10.3%']] },
      { type: 'heading', text: 'Decisions for the board', level: 1 },
      { type: 'numbered', items: ['Approve the second northern depot for spring.', 'Move the east region to the new routing system.', 'Review fleet leasing terms before the year end.'] },
    ],
  });
  await call('create_presentation', {
    path: 'Board briefing.pptx',
    title: 'The year in freight',
    subtitle: 'Northbank Logistics, November board',
    slides: [{ title: 'Highlights', points: ['North grew fastest', 'East up 30%', 'Leeds depot open'] }, { title: 'Decisions', points: ['Second northern depot', 'New routing in the east', 'Fleet leasing review'] }],
  });
  await call('read_file', { path: 'Board summary.docx' });
  return folder;
}


/**
 * The status bar's zoom stepped down until the whole chart is in the
 * window: scrolled sideways instead, the grid cut its first column in half.
 */
const scrollToChart = async (h) => {
  const fits = `(() => { const c = document.querySelector('.sh-drawing[data-kind="chart"]')?.getBoundingClientRect(); const g = document.querySelector('.sh-grid')?.getBoundingClientRect(); return Boolean(c && g && c.right <= g.right - 8); })()`;
  for (let i = 0; i < 3 && !(await h.js(fits)); i++) {
    await h.js(`document.querySelector('.rw-zoom-step[data-tip^="Zoom out"]')?.click(), 1`);
    await h.wait(900);
  }
};

export const FEATURES = [
  /* ── Documents ─────────────────────────────────────────────────────── */
  {
    name: 'word-track-changes',
    app: 'word',
    file: 'Tenancy agreement.docx',
    async setup(h) {
      await h.until(`Boolean(document.querySelector('.wd-page .wd-ins'))`, 'the tracked changes', 10000);
      await h.tab('Review');
      await h.button('Display for Review');
      await h.menu('All Markup');
      // The page alone: the Comments box is modal and would blur it, and
      // Show (2) on the ribbon says the comments are there.
      await h.until(`Boolean(document.querySelector('.wd-page .wd-del'))`, 'the deletions shown', 5000);
      await h.wait(500);
    },
  },
  {
    name: 'word-forms',
    app: 'word',
    file: 'Volunteer sign-up.docx',
    async setup(h) {
      await h.until(`document.querySelectorAll('.wd-page .wd-cc').length >= 4`, 'the form controls', 10000);
      await h.tab('Developer');
      await h.until(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].some((n) => (n.dataset.tip || '').startsWith('Rich Text Content Control'))`, 'the Developer tab', 5000);
      // The list of the drop-down open, as a person choosing a team sees it.
      // The first press puts the control in hand, the second opens its list.
      await press(h, '.wd-page .wd-cc[data-kind="dropDownList"]');
      const open = await h.until(`Boolean(document.querySelector('.wd-cc-list'))`, 'the list open', 2000).catch(() => false);
      if (!open) {
        await h.wait(300);
        await press(h, '.wd-page .wd-cc[data-kind="dropDownList"]');
        await h.until(`Boolean(document.querySelector('.wd-cc-list'))`, 'the list open', 4000);
      }
      await h.wait(500);
    },
    keepFocus: true,
  },
  {
    name: 'word-redact',
    app: 'word',
    file: 'Offer letter.docx',
    async setup(h) {
      await h.until(`Boolean(document.querySelector('.wd-page img.wd-image'))`, 'the letter and its picture', 10000);
      await h.tab('Review');
      await h.button('Redact');
      await h.until(`Boolean(document.querySelector('.rd-pane .rd-find'))`, 'the Redact pane', 5000);
      await h.field('.rd-find', 'Ayesha Khan');
      await h.click('.rd-find-mark');
      await h.click('.rd-pattern-email');
      await h.click('.rd-pattern-phone');
      await h.until(`document.querySelectorAll('.rd-mark').length >= 3 && document.querySelectorAll('.rd-hit').length >= 3`, 'the marks on the page', 6000);
      await h.until(`document.querySelectorAll('.rd-where-parts span[data-kind]').length >= 4`, 'the parts counted', 6000).catch(() => {});
      await h.wait(600);
    },
  },
  {
    name: 'word-styles-navigation',
    app: 'word',
    file: 'quarterly-review.docx',
    async setup(h) {
      await h.tab('View');
      await h.button('Navigation Pane');
      await h.until(`document.querySelectorAll('.wd-nav-item').length > 2`, 'the headings in the pane', 6000);
      await h.tab('Home');
      // The caret in the second heading, so the gallery marks its style.
      const at = await h.js(`(() => { const b = [...document.querySelectorAll('.wd-page .wd-block')].find((n) => n.textContent.trim() === 'Summary'); if (!b) return null; const r = b.getBoundingClientRect(); return { x: Math.round(r.left + Math.min(40, r.width / 2)), y: Math.round(r.top + r.height / 2) }; })()`);
      if (at) await h.clickAt(at.x, at.y);
      await h.until(`Boolean(document.querySelector('.rw-ribbon .wd-style.on'))`, 'the style marked in the gallery', 4000).catch(() => {});
      await h.wait(500);
    },
  },

  /* ── Worksheets ────────────────────────────────────────────────────── */
  {
    name: 'sheets-pivot',
    app: 'sheets',
    file: 'Orders 2026.xlsx',
    async setup(h) {
      await h.until(`Boolean(document.querySelector('.sh-cell[data-ref="G4"]'))`, 'the PivotTable', 10000);
      await h.js(`document.querySelector('.sh-cell[data-ref="G4"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })), 1`);
      await h.js(`window.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, button: 0 })), 1`);
      await h.until(`Boolean(document.querySelector('.sh-pf'))`, 'the PivotTable Fields pane', 6000);
      await h.wait(600);
    },
  },
  {
    name: 'sheets-data-bars',
    app: 'sheets',
    file: 'Team results.xlsx',
    async setup(h) {
      await h.until(`document.querySelectorAll('.sh-cf-bar').length >= 6`, 'the data bars', 10000);
      await h.tab('Home');
      await h.wait(400);
    },
  },
  {
    name: 'sheets-data-from-picture',
    app: 'sheets',
    file: 'Stock count.xlsx',
    async setup(h) {
      await englishOnly(h);
      await h.until(`document.querySelectorAll('.sh-colheads .sh-head').length > 2`, 'the grid', 8000);
      const cell = await h.js(`(() => {
        const col = [...document.querySelectorAll('.sh-colheads .sh-head')].find((n) => n.textContent.trim() === 'A').getBoundingClientRect();
        const row = [...document.querySelectorAll('.sh-rowheads .sh-head')].find((n) => n.textContent.trim() === '3').getBoundingClientRect();
        return { x: Math.round(col.left + col.width / 2), y: Math.round(row.top + row.height / 2) };
      })()`);
      await h.clickAt(cell.x, cell.y);
      h.answer('open', [h.sample('Printed stock sheet.png')]);
      await h.tab('Data');
      await h.click('[data-role="data-from-picture"]');
      await h.menu('Picture From File');
      await h.until(`Boolean(document.querySelector('[data-role="pd-grid"]'))`, 'the preview read', 60000);
      await h.wait(800);
    },
  },
  {
    name: 'sheets-chart',
    app: 'sheets',
    file: 'regional-sales.xlsx',
    async setup(h) {
      await h.until(`Boolean(document.querySelector('.sh-drawing[data-kind="chart"] svg'))`, 'the chart', 10000);
      await scrollToChart(h);
      const at = await h.js(`(() => { const r = document.querySelector('.sh-drawing[data-kind="chart"]').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 3) }; })()`);
      await h.clickAt(at.x, at.y);
      await h.until(`Boolean(document.querySelector('.sh-drawing[data-kind="chart"].picked'))`, 'the chart picked', 4000).catch(() => {});
      // The pointer off the chart, so no tip covers it.
      h.wc.sendInputEvent({ type: 'mouseMove', x: 5, y: 600 });
      await h.wait(600);
    },
  },

  /* ── Presentations ─────────────────────────────────────────────────── */
  {
    name: 'slides-master',
    app: 'slides',
    file: 'product-review.pptx',
    async setup(h) {
      await h.until(`document.querySelectorAll('.sl-thumb').length >= 3`, 'the slides');
      await h.js(`document.querySelectorAll('.sl-thumb')[1].click(), 1`);
      await h.wait(300);
      await h.tab('View');
      await h.button('Slide Master');
      await h.until(`document.querySelectorAll('.sl-mthumb svg').length >= 3 && (document.querySelector('.sl-svg')?.innerHTML || '').includes('Click to edit Master title style')`, 'the master and its layouts', 10000);
      await h.js(`document.querySelector('.sl-mthumb-master')?.click(), 1`);
      await h.until(`(document.querySelector('.sl-svg')?.innerHTML || '').includes('Fifth level')`, 'the master itself', 6000).catch(() => {});
      await h.wait(600);
    },
  },
  {
    name: 'slides-compare',
    app: 'slides',
    file: 'Launch plan.pptx',
    async setup(h) {
      await h.until(`Boolean(document.querySelector('.sl-slide'))`, 'the slide');
      await h.tab('Review');
      h.answer('open', [h.sample('Launch plan (reviewed).pptx')]);
      await h.button('Compare');
      await h.until(`Boolean(document.querySelector('.sl-rev'))`, 'the Revisions pane', 15000);
      await h.until(`Boolean(document.querySelector('.sl-rev-jump .sl-rev-goto'))`, 'the changed slide', 6000).catch(() => {});
      await h.js(`document.querySelector('.sl-rev-jump .sl-rev-goto')?.click(), 1`);
      await h.until(`document.querySelectorAll('.sl-slide .sl-rev-marker').length >= 2`, 'the changes marked on the slide', 6000).catch(() => {});
      await h.wait(800);
    },
  },
  {
    name: 'slides-design-ideas',
    app: 'slides',
    file: 'Design ideas.pptx',
    async setup(h) {
      await h.tab('Design');
      await h.button('Design Ideas');
      await h.until(`document.querySelectorAll('.sl-idea svg').length >= 3`, 'the ideas', 10000);
      await h.wait(800);
    },
  },
  {
    name: 'slides-presenter',
    app: 'slides',
    file: 'Opening week.pptx',
    async setup(h) {
      await h.until(`document.querySelectorAll('.sl-thumb').length >= 3`, 'the slides');
      await h.tab('Slide Show');
      const before = h.allWindows();
      await h.button('Use Presenter View');
      const presenter = await h.newWindow(before, 'the presenter view');
      await h.untilHere(() => presenter.webContents.executeJavaScript(`Boolean(document.querySelector('.pv-stage svg') && document.querySelector('.pv-next .pv-thumb'))`), 'the presenter view drawn', 12000);
      await h.wait(1500);
      return presenter;
    },
  },

  /* ── PDF ───────────────────────────────────────────────────────────── */
  {
    name: 'pdf-comments',
    app: 'pdf',
    file: 'Annual Review 2026.pdf',
    async setup(h) {
      await pdfReady(h);
      await h.tab('Comment');
      await h.until(`Boolean(document.querySelector('[data-role="comments-pane"]'))`, 'the Comments pane');
      await pdfSelect(h, 0, 'every region', 'last year');
      await role(h, 'tool-highlight');
      await h.until(`document.querySelectorAll('.pdf-page .pdf-annot[data-type="Highlight"]').length >= 1`, 'the highlight', 6000);
      await role(h, 'tool-note');
      await pdfDrag(h, 0, [[0.86, 0.22], [0.86, 0.22]]);
      await h.until(`Boolean(document.querySelector('[data-role="note-text"]'))`, 'the note editor', 5000);
      await pdfField(h, '[data-role="note-text"]', 'Can we show the South East depot on its own here?');
      await role(h, 'note-post');
      await h.until(`document.querySelectorAll('.pdf-comments .pdf-comment').length >= 2`, 'two comments in the pane', 6000);
      await putDown(h, 'tool-note');
      await h.wait(400);
      await pdfSelect(h, 0, 'renewable', 'diesel');
      await role(h, 'tool-underline');
      await h.js(`document.querySelector('.pdf-view')?.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })), 1`);
      await h.until(`document.querySelectorAll('.pdf-comments .pdf-comment').length >= 3`, 'three comments in the pane', 6000).catch(() => {});
      await putDown(h, 'tool-underline');
      await h.js(`getSelection().removeAllRanges(), 1`);
      // The last comment let go, so the pane shows three alike.
      await h.js(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
      await h.wait(3600);
    },
  },
  {
    name: 'pdf-fill-sign',
    app: 'pdf',
    file: 'Application form.pdf',
    async setup(h) {
      await h.until(`document.querySelectorAll('.pdf-field').length >= 10`, 'the form fields', 15000);
      await h.tab('Fill & Sign');
      await pdfField(h, '.pdf-field[data-field="name"]', 'Imogen Hartley');
      await pdfField(h, '.pdf-field[data-field="address"]', '7 Quayside Walk\nFalmouth TR11 3XA');
      await pdfField(h, '.pdf-field[data-field="amount"]', '1,250.00');
      await h.js(`document.querySelector('.pdf-field[data-field="agree"]')?.click(), 1`);
      await h.js(`document.querySelector('.pdf-field[data-field="size"][data-state="M"]')?.click(), 1`);
      await h.js(`(() => { const el = document.querySelector('.pdf-field[data-field="country"]'); if (!el) return 0; const o = [...el.options].find((x) => /United Kingdom|^GB$|^UK$/.test(x.textContent.trim() + ' ' + x.value)) || el.options[1]; if (o) { o.selected = true; el.value = o.value; } el.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
      await h.js(`(() => { const el = document.querySelector('.pdf-field[data-field="colours"]'); if (!el) return 0; for (const o of el.options) o.selected = ['Green', 'Blue'].includes(o.value) || ['Green', 'Blue'].includes(o.textContent.trim()); el.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
      await h.wait(400);
      await role(h, 'sign-signature');
      await h.until(`Boolean(document.querySelector('[data-role="sign-dialog"]'))`, 'the signature dialog', 5000);
      await pdfField(h, '[data-role="sign-text"]', 'Imogen Hartley');
      await role(h, 'sign-place');
      await h.until(`document.querySelector('.pdf-page[data-index="0"] .pdf-annots')?.dataset.tool === 'stamp'`, 'the signature to place', 5000);
      // Signed just under the form's last field.
      await pdfDrag(h, 0, [[0.42, 0.665], [0.42, 0.665]]);
      await h.until(`document.querySelectorAll('.pdf-page[data-index="0"] .pdf-annot[data-type="Stamp"]').length >= 1`, 'the signature on the page', 5000);
      // From the amount down in view: the tick, the choices, the lists and
      // the signature, the signature's foot at the bottom of the view. Two
      // steps of zoom out first: at 100% they are taller than the view.
      for (let i = 0; i < 2; i++) {
        await h.js(`document.querySelector('.rw-zoom-step[data-tip^="Zoom out"]')?.click(), 1`);
        await h.wait(700);
      }
      await h.until(`document.querySelectorAll('.pdf-page canvas[data-drawn]').length >= 1`, 'the page drawn again', 8000).catch(() => {});
      await h.wait(800);
      await h.js(`(() => { const v = document.querySelector('.pdf-view'); const s = document.querySelector('.pdf-page[data-index="0"] .pdf-annot[data-type="Stamp"]'); const f = document.querySelector('.pdf-field[data-field="amount"]'); if (!v || !s || !f) return 0; const vr = v.getBoundingClientRect(); v.scrollTop += Math.min(f.getBoundingClientRect().top - vr.top - 20, s.getBoundingClientRect().bottom - vr.bottom + 40); return v.scrollTop; })()`);
      await h.wait(3600);
    },
  },
  {
    name: 'pdf-redact',
    app: 'pdf',
    file: 'Annual Review 2026.pdf',
    async setup(h) {
      await pdfReady(h);
      await h.tab('Redact');
      await pdfSelect(h, 0, '£48.6', 'million');
      await role(h, 'redact-text');
      await h.until(`document.querySelectorAll('.pdf-page .pdf-annot[data-type="Redact"]').length >= 1`, 'the first mark', 5000);
      await pdfSelect(h, 0, '97.4%', '97.4%');
      await role(h, 'redact-text');
      await role(h, 'redact-search');
      await h.until(`Boolean(document.querySelector('[data-role="find-input"]'))`, 'Find', 5000);
      await h.field('[data-role="find-input"]', 'Northbank');
      await h.until(`Number(document.querySelector('.pdf-find-count')?.dataset.count || 0) >= 1`, 'the matches', 8000).catch(() => {});
      await role(h, 'redact-search');
      await h.wait(500);
      await h.js(`document.querySelector('[data-tip^="Close Find"]')?.click(), 1`);
      await h.js(`getSelection().removeAllRanges(), 1`);
      // Back to the first page, where most of the marks are.
      await h.js(`(() => { const v = document.querySelector('.pdf-view'); if (v) v.scrollTop = 0; return 1; })()`);
      await h.wait(3600);
    },
  },
  {
    name: 'pdf-recognise-text',
    app: 'pdf',
    file: 'Scanned letter.pdf',
    async setup(h) {
      await h.until(`Boolean(document.querySelector('.pdf-page canvas[data-drawn]'))`, 'the scan drawn', 15000);
      await h.button('Recognise Text');
      await h.until(`Boolean(document.querySelector('[data-role="ocr-dialog"]'))`, 'the Recognise Text dialog', 6000);
      await h.js(`(() => { for (const row of document.querySelectorAll('.ocr-lang')) { const box = row.querySelector('input'); const want = row.dataset.code === 'eng'; if (box && box.checked !== want) box.click(); } return 1; })()`);
      await role(h, 'ocr-go');
      await h.until(`!document.querySelector('[data-role="ocr-progress"]') && !document.querySelector('[data-role="ocr-dialog"]')`, 'the words recognised', 120000);
      await h.wait(800);
      await h.button('Find');
      await h.until(`Boolean(document.querySelector('[data-role="find-input"]'))`, 'Find', 5000);
      await h.field('[data-role="find-input"]', 'signed copy');
      await h.until(`Number(document.querySelector('.pdf-find-count')?.dataset.count || 0) >= 1`, 'the words found on the scan', 10000);
      await h.wait(5200);
    },
    keepFocus: true,
  },

  /* ── Mail ──────────────────────────────────────────────────────────── */
  {
    name: 'mail-signed',
    app: 'mail',
    async before({ mail }) {
      const { protect } = await import('@rutba/mailbox/smime');
      const { makeIdentities } = await import('../../../tests/fixtures/smime-keys.js');
      const people = makeIdentities({ org: 'Northwind Certificates' });
      const made = await mail.addAccount({ account: { email: people.owner.email, name: 'Harbour Lane', local: true } });
      await mail.trustForTest({ pem: people.ca.cert });
      const raw = Buffer.from(
        `From: ${people.amina.name} <${people.amina.email}>\r\nTo: Harbour Lane <${people.owner.email}>\r\nSubject: Harbour lease: the signed copy\r\nDate: ${new Date().toUTCString()}\r\nMessage-ID: <lease-1@northwind.example>\r\nMIME-Version: 1.0\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Transfer-Encoding: 7bit\r\n\r\nHello,\r\n\r\nHere is the lease as we agreed it on Friday, with the yard included and the rent at 1,380 a month.\r\n\r\nThis message is signed, so you can see it is from me and that nobody has changed it on the way.\r\n\r\nBest wishes,\r\nAmina\r\n`,
        'latin1',
      );
      const signed = protect(raw, { sign: { key: people.amina.key, cert: people.amina.cert, chain: [people.ca.cert] } });
      await mail.deliverTest({ accountId: made.id, folder: 'Inbox', raw: signed.toString('latin1') });
    },
    async setup(h) {
      await h.until(`document.querySelectorAll('.ml-accounts button').length > 0`, 'the sidebar', 10000);
      await h.js(`(() => { const b = [...document.querySelectorAll('.ml-accounts button')].find((n) => n.textContent.trim().startsWith('Harbour Lane')); b?.click(); return Boolean(b); })()`);
      await h.until(`[...document.querySelectorAll('.ml-row .ml-subject')].some((n) => n.textContent.trim() === 'Harbour lease: the signed copy')`, 'the message in the list', 8000);
      await h.js(`[...document.querySelectorAll('.ml-row .ml-subject')].find((n) => n.textContent.trim() === 'Harbour lease: the signed copy').click(), 1`);
      await h.until(`document.querySelector('.ml-smime')?.dataset.signed === 'valid'`, 'the signature strip', 8000);
      await h.wait(800);
    },
  },
  {
    name: 'mail-archive-search',
    app: 'mail',
    async before({ mail, sample, wait }) {
      const { rfc822 } = await import('../../../tests/fixtures/mail-archives.js');
      const { ACCOUNT_MAIL } = await import('../build/feature-samples.js');
      const made = await mail.addAccount({ account: { email: 'legal@harbourlane.example', name: 'Legal', local: true } });
      for (const m of ACCOUNT_MAIL) await mail.deliverTest({ accountId: made.id, folder: 'Inbox', raw: rfc822(m) });
      const dir = sample('Mail archives');
      for (const name of fs.readdirSync(dir).filter((n) => /\.(pst|olm)$/i.test(n))) await mail.openArchive({ path: path.join(dir, name) });
      const deadline = Date.now() + 60000;
      while (Date.now() < deadline) {
        const list = await mail.archives();
        if (list.length >= 2 && list.every((a) => a.index?.state === 'ready')) break;
        await wait(300);
      }
    },
    async setup(h) {
      await h.until(`document.querySelectorAll('.ml-archives button').length >= 2`, 'the archives in the sidebar', 10000);
      await h.tab('Search');
      await h.field('.rw-search input', 'halvorsen');
      await h.until(`Boolean(document.querySelector('.ml-scope'))`, 'the search scope', 6000);
      await h.js(`[...document.querySelectorAll('.ml-scope-seg button')].find((n) => n.textContent.trim().startsWith('All mail'))?.click(), 1`);
      await h.until(`Number(document.querySelector('.ml-scope-found')?.dataset.total || 0) >= 3`, 'the results from every source', 15000);
      await h.js(`[...document.querySelectorAll('.ml-row .ml-subject')].find((n) => /Halvorsen site visit/.test(n.textContent))?.click(), 1`);
      await h.wait(1200);
    },
  },
  {
    name: 'mail-add-accounts',
    app: 'mail',
    async setup(h) {
      await h.until(`document.querySelectorAll('.ml-accounts button').length > 0`, 'the sidebar', 10000);
      await h.js(`[...document.querySelectorAll('.ml-accounts button, button')].find((b) => /^Add account$/.test(b.textContent.trim()))?.click(), 1`);
      await h.click('.ml-several-accounts', 'Several accounts');
      await h.until(`Boolean(document.querySelector('textarea.ml-bulk-text'))`, 'the rows box', 5000);
      // Pasted, as a list from a spreadsheet would be: the box reads the
      // rows into the list and keeps no password on screen.
      const rows = ['imogen@example.com\tHarbour-Lane-26', 'orders@example.org,Spring-Planting-3', 'team@example.net:Quayside-77', 'bookings@example.co.uk'].join('\n');
      await h.js(`(() => { const el = document.querySelector('textarea.ml-bulk-text'); el.focus(); const dt = new DataTransfer(); dt.setData('text/plain', ${JSON.stringify(rows)}); el.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); return 1; })()`);
      await h.until(`document.querySelectorAll('.ml-bulk-row').length >= 4`, 'the rows listed', 5000);
      await h.wait(600);
    },
  },

  /* ── Calendar and Contacts ─────────────────────────────────────────── */
  {
    name: 'calendar-week',
    app: 'calendar',
    file: 'this-week.ics',
    settle: 3000,
    async setup(h) {
      // The opened file's events kept, so the banner offering them goes.
      await h.until(`Boolean(document.querySelector('.cal-invite .rw-btn'))`, 'the opened file', 8000).catch(() => {});
      await h.js(`(() => { const b = [...document.querySelectorAll('.cal-invite .rw-btn')].pop(); b?.click(); return Boolean(b); })()`);
      await h.until(`!document.querySelector('.cal-invite')`, 'the events kept', 6000).catch(() => {});
      await h.js(`[...document.querySelectorAll('.cal-view')].find((b) => b.textContent.trim() === 'Week')?.click(), 1`);
      await h.until(`document.querySelectorAll('.cal-block').length >= 4`, 'the week drawn', 8000);
      await h.wait(800);
    },
  },
  {
    name: 'contacts-card',
    app: 'contacts',
    file: 'address-book.vcf',
    settle: 2500,
    async setup(h) {
      await h.until(`Boolean(document.querySelector('.ct-offer .rw-btn'))`, 'the opened file', 8000).catch(() => {});
      await h.js(`(() => { const b = [...document.querySelectorAll('.ct-offer .rw-btn')].pop(); b?.click(); return Boolean(b); })()`);
      await h.until(`!document.querySelector('.ct-offer') && document.querySelectorAll('.ct-item').length >= 6`, 'the contacts kept', 6000).catch(() => {});
      await h.js(`[...document.querySelectorAll('.ct-item')].find((n) => n.querySelector('.name')?.textContent.trim() === 'Aoife Brennan')?.click(), 1`);
      await h.until(`document.querySelector('.ct-card h2')?.textContent.trim() === 'Aoife Brennan'`, 'the card', 5000);
      await h.wait(600);
    },
  },

  /* ── Pictures ──────────────────────────────────────────────────────── */
  {
    name: 'pictures-ratings',
    app: 'pictures',
    file: 'Sample Pictures',
    settle: 3000,
    before({ stores, sample }) {
      const at = (name) => path.join(sample('Sample Pictures'), name);
      stores.settings.set('pictures.marks', {
        [at('Coast at sunset.png')]: { rating: 5, tags: ['holiday'] },
        [at('Lakeside at dawn.png')]: { rating: 4 },
        [at('Hills at noon.png')]: { rating: 3, tags: ['walks'] },
        [at('Valley in mist.png')]: { rating: 2 },
      });
    },
    async setup(h) {
      await h.until(`document.querySelectorAll('.pv-tile:not(.folder)').length >= 4 && document.querySelectorAll('.pv-tile-stars[data-rating]').length >= 4`, 'the rated pictures', 10000);
      // The largest tiles, so four pictures fill the window rather than a third of it.
      await h.js(`[...document.querySelectorAll('.pv-size')].pop()?.click(), 1`);
      await h.wait(600);
      await h.until(`[...document.querySelectorAll('.pv-tile img')].every((i) => i.complete && i.naturalWidth > 0)`, 'the thumbnails', 10000).catch(() => {});
      await h.wait(1200);
    },
  },
  {
    name: 'pictures-copy-text',
    app: 'pictures',
    file: path.join('Scans', 'Printed letter.png'),
    async setup(h) {
      await englishOnly(h);
      await h.until(`(() => { const i = document.querySelector('.pv-image'); return Boolean(i && i.complete && i.naturalWidth); })()`, 'the picture', 10000);
      await h.click('[data-role="copy-text"]');
      await h.until(`document.querySelector('[data-role="ocr-bar"]')?.dataset.state === 'ready'`, 'the words read', 60000);
      // One line selected, as a person picking words out of it would.
      await h.js(`(() => { const line = document.querySelectorAll('.ocr-layer .ocr-line')[1]; if (!line) return 0; const r = document.createRange(); r.selectNodeContents(line); getSelection().removeAllRanges(); getSelection().addRange(r); return 1; })()`);
      await h.wait(5400);
    },
    keepFocus: true,
  },

  /* ── Image ─────────────────────────────────────────────────────────── */
  {
    name: 'image-layers',
    app: 'image',
    file: LANDSCAPE,
    async setup(h) {
      await imageReady(h);
      // A see-through panel, and words on a layer above it: a title card.
      await h.tab('Insert');
      await h.click('[data-insert="shape"]');
      await h.until(`document.querySelectorAll('.im-gallery-item').length > 20`, 'the shapes', 5000);
      await h.click('.im-gallery-item[data-item="rect"]');
      await h.until(`Number(document.querySelector('.im-canvas-wrap')?.dataset.layers || 0) === 2`, 'the shape layer', 6000);
      await imageSettled(h);
      await h.tab('Layers');
      await h.click('[data-role="show-layers"]');
      await h.until(`Boolean(document.querySelector('[data-role="opacity"]'))`, 'the shape picked in the panel', 5000);
      await h.field('[data-role="opacity"]', '45');
      await imageSettled(h);
      await h.tab('Insert');
      await h.click('[data-insert="text"]');
      await h.until(`Number(document.querySelector('.im-canvas-wrap')?.dataset.layers || 0) === 3`, 'the text layer', 6000);
      await imageSettled(h);
      await h.tab('Layers');
      await h.js(`document.querySelector('[data-side="layers"]')?.click(), 1`);
      await h.until(`Boolean(document.querySelector('[data-role="layer-text"]'))`, 'the words in the panel', 5000);
      await h.field('[data-role="layer-text"]', 'Lakeside at dawn');
      await imageSettled(h);
      await h.until(`document.querySelectorAll('.im-layer').length >= 3`, 'the Layers panel', 5000);
      await h.wait(800);
    },
  },
  {
    name: 'image-selection-mask',
    app: 'image',
    file: LANDSCAPE,
    async setup(h) {
      await imageReady(h);
      await h.tab('Layers');
      // The ribbon's Layers turns the picture into a Background layer first.
      await h.click('[data-role="show-layers"]');
      await h.until(`Boolean(document.querySelector('.im-layer[data-layer="background"]'))`, 'the Layers panel', 6000);
      // The right half kept and the rest masked away, then a new selection
      // around the sun with its marching ants.
      await tool(h, 'marquee');
      await gesture(h, [[600, 0], [900, 400], [1200, 800]], 1200);
      await h.until(`/^\\d+,\\d+,\\d+,\\d+$/.test(document.querySelector('.im-ants')?.dataset.selection || '')`, 'the selection', 4000);
      await h.js(`document.querySelector('.im-layer[data-layer="background"]')?.click(), 1`);
      await h.until(`Boolean(document.querySelector('[data-role="mask-selection"]:not(:disabled)'))`, 'From Selection', 4000);
      await h.click('[data-role="mask-selection"]');
      await imageSettled(h);
      await tool(h, 'marquee');
      await gesture(h, [[720, 300], [800, 360], [920, 500]], 1200);
      await h.until(`/^\\d+,\\d+,\\d+,\\d+$/.test(document.querySelector('.im-ants')?.dataset.selection || '')`, 'the new selection', 4000);
      await h.wait(800);
    },
  },
  {
    name: 'image-curves',
    app: 'image',
    file: LANDSCAPE,
    async setup(h) {
      await imageReady(h);
      await h.tab('Adjust');
      await h.button('Curves');
      await h.until(`Boolean(document.querySelector('[data-role="curve-preset"]'))`, 'the Curves section', 5000);
      await h.field('[data-role="curve-preset"]', 'lighter');
      await h.until(`(document.querySelector('.im-curve-line')?.getAttribute('d') || '').length > 100`, 'the curve', 5000).catch(() => {});
      await imageSettled(h);
      // The whole Curves section in view, its heading included.
      await h.js(`(document.querySelector('.im-curve-line')?.closest('svg') || document.querySelector('[data-role="curve-preset"]'))?.scrollIntoView({ block: 'center' }), 1`);
      await h.wait(800);
    },
  },

  /* ── Video ─────────────────────────────────────────────────────────── */
  {
    name: 'video-intro',
    app: 'video',
    query: { intro: 1 },
    async setup(h) {
      await h.until(`document.querySelectorAll('.vi-scene').length >= 4`, 'the scenes', 10000);
      await h.js(`document.querySelector('.vi-scene .vi-scene-head')?.click(), 1`);
      await h.until(`Boolean(document.querySelector('.vi-scene.open'))`, 'the first scene open', 4000).catch(() => {});
      await h.wait(1500);
    },
  },
  {
    name: 'video-recorder',
    app: 'video',
    file: 'Our year.rvid',
    async setup(h) {
      // A window of the suite to record: the recorder offers only the
      // suite's own windows in a screens run, never the desktop.
      const doc = h.windows.create({ app: 'word', file: h.sample('quarterly-review.docx') });
      await new Promise((resolve) => doc.webContents.once('did-finish-load', resolve));
      await h.untilHere(() => doc.webContents.executeJavaScript(`Boolean(document.querySelector('.wd-block'))`), 'the document to record', 12000);
      await h.wait(1500);
      await h.js(`window.rutbaOffice.store.set({ key: 'recorder.settings', value: { source: 'window', fps: 30, cap: 1080, quality: 'standard', countdown: 3, camera: false, mic: true, system: false, minimise: true } }), 1`);
      await h.tab('Home');
      await h.button('Record Screen');
      await h.until(`Boolean(document.querySelector('.rec-dialog'))`, 'the recorder', 6000);
      await h.until(`[...document.querySelectorAll('.rec-source')].some((n) => /quarterly-review/.test(n.dataset.tip || n.textContent))`, 'the window listed', 10000);
      await h.js(`[...document.querySelectorAll('.rec-source')].find((n) => /quarterly-review/.test(n.dataset.tip || n.textContent)).click(), 1`);
      await h.until(`(document.querySelector('.rec-screen')?.videoWidth || 0) > 0`, 'the window live in the recorder', 10000);
      await h.wait(1500);
    },
  },
  {
    name: 'video-export',
    app: 'video',
    file: 'Our year.rvid',
    settle: 3000,
    async setup(h) {
      await h.until(`document.querySelectorAll('.vd-clip').length >= 3`, 'the timeline', 10000);
      await h.tab('Home');
      await h.button('Export');
      await h.until(`Boolean(document.querySelector('.vd-export-preset'))`, 'the Export box', 6000);
      await h.field('.vd-export-preset', 'youtube');
      await h.wait(1000);
    },
  },
  {
    // The inset still picked on the timeline: on the preview its outline,
    // the grips at its corners and the one to turn it; in the panel its
    // place, size and angle as numbers.
    name: 'video-picture-in-picture',
    app: 'video',
    file: 'Our year.rvid',
    settle: 2500,
    async setup(h) {
      await h.until(`document.querySelectorAll('.vd-clip').length >= 6`, 'the timeline', 10000);
      await scrubVideo(h, 11.6);
      await h.tab('Clip');
      const inset = await h.centre('.vd-clip[data-kind="image"][data-start="9.5"]');
      if (!inset) throw new Error('no inset on the timeline');
      await h.clickAt(inset.x, inset.y);
      await h.until(`Boolean(document.querySelector('.vd-outline')) && Boolean(document.querySelector('.vd-place-props .vd-rot'))`, 'the inset picked, with its grips', 5000);
      await h.js(`document.querySelector('.vd-place-props')?.scrollIntoView({ block: 'start' }), 1`);
      h.wc.sendInputEvent({ type: 'mouseMove', x: 30, y: 600 });
      await h.wait(1200);
    },
  },
  {
    // The music bed picked, with its volume and fades, on a cut marked In
    // and Out, and the Edit tab's track tools.
    name: 'video-tracks-sound',
    app: 'video',
    file: 'Our year, final cut.rvid',
    settle: 2500,
    async setup(h) {
      await h.until(`document.querySelectorAll('.vd-clip').length >= 6 && Boolean(document.querySelector('.vd-range'))`, 'the timeline and its In and Out', 10000);
      await scrubVideo(h, 14.6);
      await h.tab('Edit');
      await h.js(`document.querySelector('.vd-clip[data-kind="sound"]')?.scrollIntoView({ block: 'end' }), 1`);
      await h.wait(400);
      const bed = await h.centre('.vd-clip[data-kind="sound"]');
      if (!bed) throw new Error('no music on the timeline');
      await h.clickAt(bed.x, bed.y);
      await h.until(`[...document.querySelectorAll('.vd-insp-section h3')].some((n) => /Sound/.test(n.textContent))`, 'the sound settings', 5000);
      await h.js(`(() => { const s = [...document.querySelectorAll('.vd-insp-section h3')].find((n) => /Sound/.test(n.textContent)); s?.scrollIntoView({ block: 'start' }); return 1; })()`);
      h.wc.sendInputEvent({ type: 'mouseMove', x: 30, y: 600 });
      await h.wait(1500);
    },
  },

  /* ── Recorder ──────────────────────────────────────────────────────── */
  {
    // Recorder's main picture: a document window being recorded with the
    // camera in its corner and a title across the foot, taken the feature
    // way because the window has to be picked first.
    name: 'recorder-studio',
    app: 'recorder',
    asApp: 'recorder',
    async before({ stores, samples }) {
      stores.settings.set('recorder.studio', { scene: 'screen-camera', size: 1080, fps: 30, countdown: 3, camera: true, mic: true, system: false, titleOn: true, title: 'Imogen Hartley', subtitle: 'Northbank Logistics, quarterly review', folder: samples, minimise: true, shape: 'circle', corner: 'br', insetSize: 'medium' });
    },
    async setup(h) {
      await recordWindowIn(h);
      await h.wait(1500);
    },
  },
  {
    // Go Live's box: YouTube with its key kept (a sample key, never used to
    // connect) and Facebook waiting for one, and the copy recorded meanwhile.
    name: 'recorder-go-live',
    app: 'recorder',
    async before({ stores, samples }) {
      stores.settings.set('recorder.studio', { scene: 'split', size: 1080, fps: 30, countdown: 3, camera: true, mic: true, system: false, titleOn: false, folder: samples, minimise: true, liveCopy: true });
    },
    async setup(h) {
      await recordWindowIn(h);
      for (const d of await h.js(`window.rutbaOffice.live.destinations()`)) await h.js(`window.rutbaOffice.live.removeDestination({ id: ${JSON.stringify(d.id)} })`);
      await h.js(`window.rutbaOffice.live.saveDestination({ platform: 'youtube', name: 'YouTube', on: true, key: 'sample-stream-key-2025' })`);
      await h.js(`window.rutbaOffice.live.saveDestination({ platform: 'facebook', name: 'Facebook page', on: true })`);
      await h.click('.st-dests-btn', 'Destinations');
      await h.until(`document.querySelectorAll('.lv-dest').length === 2`, 'the destinations', 6000);
      await h.wait(800);
    },
  },

  /* ── Talk ──────────────────────────────────────────────────────────── */
  {
    // Talk's main picture: signed in to a stand-in introducer (the window
    // checks' own, verify-talk.js) with six invented colleagues, some online,
    // and a conversation from this morning open. Nothing reaches a real server.
    name: 'talk-conversation',
    app: 'talk',
    asApp: 'talk',
    async setup(h) {
      const { standInIntroducer } = await import('./verify-talk.js');
      const at = (minutesAgo) => Date.now() - minutesAgo * 60_000;
      const me = 'u-amira';
      talkStandIn = standInIntroducer({
        people: {
          grace: { id: 'u-grace', name: 'Grace Okafor', email: 'grace@example.com' },
          tariq: { id: 'u-tariq', name: 'Tariq Mahmood', email: 'tariq@example.com' },
          sofia: { id: 'u-sofia', name: 'Sofia Marchetti', email: 'sofia@example.com' },
          hiroshi: { id: 'u-hiroshi', name: 'Hiroshi Tanaka', email: 'hiroshi@example.com' },
          eleanor: { id: 'u-eleanor', name: 'Eleanor Price', email: 'eleanor@example.com' },
        },
        contacts: ['u-ben', 'u-grace', 'u-tariq', 'u-sofia', 'u-hiroshi', 'u-eleanor'].map((id) => [me, id]),
        presence: { 'u-ben': 'online', 'u-grace': 'busy', 'u-tariq': 'away', 'u-hiroshi': 'online' },
        history: [
          { from: 'u-grace', to: me, text: 'Lunch at one on Thursday still good for you?', at: at(1440) },
          { from: me, to: 'u-grace', text: 'Yes, see you there.', at: at(1430) },
          { from: 'u-ben', to: me, text: 'Morning! Did the quarterly figures come through from finance?', at: at(96) },
          { from: me, to: 'u-ben', text: 'They did, an hour ago. Freight is up 12% on last quarter.', at: at(92) },
          { from: 'u-ben', to: me, text: 'Good news. Can you put the regional table into the board pack?', at: at(90) },
          { from: me, to: 'u-ben', text: 'Already in. I have sent you the draft to look over.', at: at(41) },
          { from: 'u-ben', to: me, text: 'Read it. Two small changes on page 3, otherwise ready to go.', at: at(18) },
          { from: me, to: 'u-ben', text: 'Thanks, fixing them now. A quick call in ten minutes to go through the slides?', at: at(6) },
          { from: 'u-ben', to: me, text: 'Perfect, call me then.', at: at(4) },
        ],
      });
      const port = await talkStandIn.listen();
      await h.until(`Boolean(document.querySelector('.tk-signin'))`, 'the sign-in', 8000);
      await h.js(`document.querySelector('.tk-test')?.setAttribute('open', ''), 1`);
      await h.field('.tk-server', `ws://127.0.0.1:${port}/ws`);
      await h.field('.tk-token', 'amira');
      await h.click('.tk-connect', 'Connect');
      await h.until(`document.querySelectorAll('.tk-contact').length >= 6`, 'the contacts', 8000);
      await h.click('.tk-contact[data-contact="u-ben"]', 'Ben');
      await h.until(`document.querySelectorAll('.tk-bubble').length >= 7`, 'the conversation', 6000);
      h.wc.sendInputEvent({ type: 'mouseMove', x: 1300, y: 860 });
      await h.wait(1200);
    },
    async done() {
      const going = talkStandIn;
      talkStandIn = null;
      await going?.close();
    },
  },

  /* ── The launcher ──────────────────────────────────────────────────── */
  {
    name: 'home-templates',
    app: 'home',
    size: '1180x780',
    async setup(h) {
      await h.click('.home-more');
      await h.until(`(() => { const c = [...document.querySelectorAll('.tg-card[data-template] img')]; return c.length > 8 && c.slice(0, 12).every((i) => i.complete && i.naturalWidth > 100); })()`, 'the templates drawn', 10000);
      await h.wait(800);
    },
  },

  /* ── Across the suite ──────────────────────────────────────────────── */
  {
    // Mail's main picture: the imported mailbox with a conversation open in
    // the reading pane. Taken as a feature because it needs the click, and
    // written over the app capture's mail.png, whose reading pane said
    // "No message selected".
    name: 'mail-reading',
    app: 'mail',
    asApp: 'mail',
    async setup(h) {
      await h.until(`document.querySelectorAll('.ml-row').length >= 3`, 'the messages', 10000);
      await h.js(`(() => { const r = [...document.querySelectorAll('.ml-row .ml-subject')].find((n) => /Q3 numbers/.test(n.textContent)) || document.querySelector('.ml-row .ml-subject'); r?.click(); return 1; })()`);
      await h.until(`Boolean(document.querySelector('.ml-head h2'))`, 'the conversation open', 6000);
      h.wc.sendInputEvent({ type: 'mouseMove', x: 1300, y: 860 });
      await h.wait(1200);
    },
  },
  {
    // Mail, whose window is all the suite's own: a workbook's light cell
    // fills stay light in the dark theme, as Excel keeps them.
    name: 'suite-dark',
    app: 'mail',
    boot: { theme: 'dark' },
    async setup(h) {
      await h.until(`document.documentElement.dataset.theme === 'dark'`, 'the dark theme', 5000);
      await h.until(`document.querySelectorAll('.ml-row').length >= 3`, 'the messages', 10000);
      await h.js(`(() => { const r = [...document.querySelectorAll('.ml-row .ml-subject')].find((n) => /Q3 numbers/.test(n.textContent)) || document.querySelector('.ml-row .ml-subject'); r?.click(); return 1; })()`);
      await h.until(`Boolean(document.querySelector('.ml-head h2'))`, 'the message open', 6000).catch(() => {});
      await h.wait(1000);
    },
  },
  {
    name: 'suite-urdu',
    app: 'word',
    file: 'quarterly-review.docx',
    boot: { language: 'ur' },
    async setup(h) {
      await h.until(`document.documentElement.lang === 'ur' && document.documentElement.dir === 'rtl'`, 'the Urdu interface', 6000);
      await h.wait(800);
    },
  },
  {
    name: 'suite-icons-only',
    app: 'slides',
    file: 'product-review.pptx',
    // A run of its own: the choice is kept for every window after it.
    boot: { group: 'icons' },
    async setup(h) {
      // A figure card picked, so the ribbon's buttons have something to act on.
      await h.until(`document.querySelectorAll('.sl-hit').length >= 3`, 'the slide\'s shapes', 8000);
      const card = await h.js(`(() => { const r = [...document.querySelectorAll('.sl-hit')].map((n) => n.getBoundingClientRect()).filter((b) => b.width > 20).sort((a, b) => b.left - a.left || a.top - b.top)[0]; return r ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) } : null; })()`);
      if (card) await h.clickAt(card.x, card.y);
      await h.wait(500);
      await h.click('.rw-ribbon-options', 'the ribbon options');
      await h.menu('Icons only');
      await h.until(`Boolean(document.querySelector('.rw-ribbon.icons-only'))`, 'the ribbon in icons', 4000);
      h.wc.sendInputEvent({ type: 'mouseMove', x: 700, y: 600 });
      await h.wait(800);
    },
  },
  {
    name: 'suite-keytips',
    app: 'word',
    file: 'quarterly-review.docx',
    async setup(h) {
      h.win.webContents.focus();
      await h.wait(300);
      h.wc.sendInputEvent({ type: 'keyDown', keyCode: 'Alt' });
      h.wc.sendInputEvent({ type: 'keyUp', keyCode: 'Alt' });
      await h.until(`[...document.querySelectorAll('.rw-keytip')].some((n) => n.dataset.tab)`, 'the KeyTips', 4000);
      await h.wait(400);
    },
    keepFocus: true,
  },
  {
    // AI Tools on, always, for a folder the person chose, with an AI app
    // connected that has just made the document the window is showing:
    // the badge in the title bar, who is connected, and what was done.
    name: 'suite-ai-tools',
    app: 'word',
    file: path.join('Northbank work', 'Board summary.docx'),
    async before({ stores, samples, profile }) {
      const folder = path.join(samples, 'Northbank work');
      fs.mkdirSync(folder, { recursive: true });
      stores.settings.set('ai.tools.folders', [folder]);
      stores.settings.set('ai.tools.mode', 'always');
      await connectAiApp({ profile, folder });
    },
    async setup(h) {
      await h.until(`document.querySelector('.ai-badge')?.dataset.connections === '1'`, 'the badge to count the AI app', 15000);
      await h.click('.ai-badge', 'the AI Tools badge');
      await h.until(`document.querySelectorAll('.ai-connections li').length === 1 && document.querySelectorAll('.ai-log li').length >= 4`, 'who is connected and the log', 8000);
      // Down to what the tools have done, with who is connected still above it.
      await h.js(`(() => { const h3 = [...document.querySelectorAll('.ai-tools-dialog .ai-h')][0]; h3?.scrollIntoView({ block: 'start' }); return 1; })()`);
      h.wc.sendInputEvent({ type: 'mouseMove', x: 30, y: 700 });
      await h.wait(800);
    },
    async done() {
      if (!aiApp) return;
      const going = aiApp;
      aiApp = null;
      going.stdin.end();
      await new Promise((resolve) => { going.once('exit', resolve); setTimeout(resolve, 5000); });
    },
  },
];
