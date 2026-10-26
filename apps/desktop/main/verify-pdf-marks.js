// PDF: page moves and the marks put on a page (2.18.0, A7.9). Date and
// Checkmark go back once placed, so a second click adds nothing; a placed
// tick is moved by dragging it and a date resized by its corner, and the
// file the window keeps has them there; pressing a comment's card with a
// tool in hand says the tool was put down; dropping selected pages among
// themselves says why nothing moved; Split names its files and offers Show
// in Folder. Run alone with RUTBA_VERIFY_ONLY=pdf-marks.

import fs from 'node:fs';
import path from 'node:path';
import pdfWriter from '@rutba/pdf';

const { PdfDocument } = pdfWriter;

function makePdf(file) {
  const doc = new PdfDocument({ title: 'Form to sign', created: '2026-10-26T09:00:00Z' });
  for (const [i, words] of ['Application form', 'Second page', 'Third page', 'Fourth page'].entries()) {
    const page = doc.addPage();
    page.text(words, 72, 110, { font: 'Helvetica-Bold', size: 28 });
    page.text(`Page ${i + 1}: sign and date below.`, 72, 160, { font: 'Helvetica', size: 14 });
  }
  fs.writeFileSync(file, doc.toBuffer());
}

const answer = (kind, value) => {
  (globalThis.__rutbaCheckDialogAnswers ||= []).push({ kind, answer: value });
};

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made
 */
export async function verifyPdfMarks({ open, check, until, wait, errorsIn }, { dir }) {
  const root = path.join(dir, 'pdf-marks');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'form-to-sign.pdf');
  makePdf(file);
  const win = await open('pdf', file);
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };
  const tab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 1`);
  const role = async (r) => {
    const sel = `[data-role="${r}"]`;
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(sel)}) && !document.querySelector(${JSON.stringify(sel)}).disabled)`), `the ${r} button`, 5000).catch(() => {});
    return js(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'no button ${r}'; if (b.disabled) return 'disabled ${r}'; b.click(); return 'clicked'; })()`);
  };
  const pressed = (r) => js(`document.querySelector('[data-role="${r}"]')?.getAttribute('aria-pressed') === 'true'`);
  /** A click on page `index`'s layer, at fractions of the page. */
  const clickOn = (index, [fx, fy]) => js(`(() => {
    const svg = document.querySelector('.pdf-page[data-index="${index}"] .pdf-annots');
    if (!svg) return 'no layer';
    const r = svg.getBoundingClientRect();
    const at = { clientX: r.left + r.width * ${fx}, clientY: r.top + r.height * ${fy}, bubbles: true, cancelable: true, button: 0, buttons: 1 };
    svg.dispatchEvent(new MouseEvent('mousedown', at));
    window.dispatchEvent(new MouseEvent('mouseup', at));
    return 'clicked';
  })()`);
  const marks = (type) => js(`[...document.querySelectorAll('.pdf-page[data-index="0"] .pdf-annot[data-type="${type}"]')].map((g) => { const r = g.querySelector('polygon').getBoundingClientRect(); return { id: g.dataset.id, x: Math.round(r.left), y: Math.round(r.top), w: Math.round(r.width), h: Math.round(r.height) }; })`);
  /** A drag from the middle of a mark (or one of its handles) by dx, dy pixels. */
  const dragMark = (sel, dx, dy) => js(`(async () => {
    const el = document.querySelector(${JSON.stringify(sel)});
    if (!el) return 'nothing at ${sel}';
    const r = el.getBoundingClientRect();
    const x = r.left + r.width / 2;
    const y = r.top + r.height / 2;
    const ev = (px, py) => ({ clientX: px, clientY: py, bubbles: true, cancelable: true, button: 0, buttons: 1 });
    el.dispatchEvent(new MouseEvent('mousedown', ev(x, y)));
    for (let i = 1; i <= 8; i++) {
      window.dispatchEvent(new MouseEvent('mousemove', ev(x + (${dx} * i) / 8, y + (${dy} * i) / 8)));
      await new Promise((res) => setTimeout(res, 25));
    }
    window.dispatchEvent(new MouseEvent('mouseup', ev(x + ${dx}, y + ${dy})));
    return 'dragged';
  })()`);
  const toasts = () => js(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent.trim()).join(' | ')`);

  try {
    await until(() => js(`document.querySelectorAll('.pdf-page canvas[data-drawn]').length >= 1 && Boolean(document.querySelector('.pdf-page[data-index="0"] .pdf-annots[data-ready]'))`), 'the first page drawn', 20000).catch(() => {});

    // Checkmark: one tick a click, the tool given back.
    await tab('Fill & Sign');
    await role('fill-check');
    const inHand = await pressed('fill-check');
    await clickOn(0, [0.3, 0.35]);
    await until(async () => (await marks('Stamp')).length === 1, 'the tick', 10000).catch(() => {});
    const released = await until(async () => !(await pressed('fill-check')), 'Checkmark given back', 3000).then(() => true).catch(() => false);
    await clickOn(0, [0.6, 0.35]);
    await wait(900);
    const ticks = await marks('Stamp');
    check('pdf-marks: Checkmark puts one tick down and goes back, so a second click adds no other', inHand && released && ticks.length === 1, JSON.stringify({ inHand, released, ticks }));

    // Date: the same.
    await role('fill-date');
    await clickOn(0, [0.3, 0.5]);
    await until(async () => (await marks('FreeText')).length === 1, 'the date', 10000).catch(() => {});
    const dateBack = await until(async () => !(await pressed('fill-date')), 'Date given back', 3000).then(() => true).catch(() => false);
    await clickOn(0, [0.6, 0.6]);
    await wait(900);
    const dates = await marks('FreeText');
    check('pdf-marks: Date writes one date and goes back, so a second click adds no other', dateBack && dates.length === 1, JSON.stringify({ dateBack, dates }));

    // The tick dragged 80 pixels right and 40 down: moved in the file the window keeps.
    const tickBefore = ticks[0];
    await dragMark(`.pdf-page[data-index="0"] .pdf-annot[data-id="${tickBefore?.id}"] polygon`, 80, 40);
    const tickMoved = await until(async () => {
      const [m] = await marks('Stamp');
      return m && Math.abs(m.x - tickBefore.x - 80) <= 3 && Math.abs(m.y - tickBefore.y - 40) <= 3 && m.w === tickBefore.w ? m : null;
    }, 'the tick moved', 10000).catch(() => null);
    const drawnThere = await js(`document.querySelectorAll('.pdf-page[data-index="0"] canvas[data-drawn]').length >= 1`);
    await capture('pdf-marks-01-moved.png');
    check('pdf-marks: a placed tick is moved by dragging it, its size kept', Boolean(tickMoved) && drawnThere, JSON.stringify({ tickBefore, after: await marks('Stamp') }));

    // The date picked shows its corners; the top right one dragged 60 pixels right makes it wider.
    const dateBefore = dates[0];
    await dragMark(`.pdf-page[data-index="0"] .pdf-annot[data-id="${dateBefore?.id}"] polygon`, 0, 0);
    const corners = () => js(`document.querySelectorAll('.pdf-page[data-index="0"] .pdf-annot[data-id="${dateBefore?.id}"] .pdf-handle').length`);
    await until(async () => (await corners()) === 4, 'the corners', 4000).catch(() => {});
    const handles = await corners();
    await capture('pdf-marks-02-corners.png');
    await dragMark(`.pdf-page[data-index="0"] .pdf-annot[data-id="${dateBefore?.id}"] .pdf-handle[data-handle="2"]`, 60, 0);
    const wider = await until(async () => {
      const [m] = await marks('FreeText');
      return m && m.w >= dateBefore.w + 50 && Math.abs(m.x - dateBefore.x) <= 3 ? m : null;
    }, 'the date wider', 10000).catch(() => null);
    check('pdf-marks: a placed date shows its corners when picked, and a corner dragged resizes it, the other side staying', handles === 4 && Boolean(wider), JSON.stringify({ handles, dateBefore, after: await marks('FreeText') }));

    // Rectangle in hand, a card pressed in the Comments pane: the tool is put down, and it is said.
    await tab('Comment');
    await until(() => js(`document.querySelectorAll('.pdf-comments .pdf-comment').length >= 2`), 'the cards', 6000).catch(() => {});
    await role('tool-rect');
    const rectInHand = await pressed('tool-rect');
    await js(`(() => { const c = document.querySelector('.pdf-comments .pdf-comment'); c?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); return 1; })()`);
    const said = await until(async () => /Rectangle put down/.test(await toasts()), 'the word', 4000).then(() => true).catch(() => false);
    const rectBack = !(await pressed('tool-rect'));
    check('pdf-marks: pressing a comment\'s card with a tool in hand puts the tool down and says so', rectInHand && said && rectBack, JSON.stringify({ rectInHand, rectBack, toasts: await toasts() }));

    // Pages 1 and 2 selected, page 2 dropped between them: nothing moves, and the window says why.
    await tab('Pages');
    const pick = (i, { ctrl = false, shift = false } = {}) => js(`(() => {
      const el = document.querySelectorAll('.pdf-thumb')[${i}];
      if (!el) return 'no thumbnail';
      const r = el.getBoundingClientRect();
      const at = { clientX: r.left + r.width / 2, clientY: r.top + r.height / 2, bubbles: true, cancelable: true, button: 0, ctrlKey: ${ctrl}, shiftKey: ${shift} };
      el.dispatchEvent(new MouseEvent('mousedown', at));
      el.dispatchEvent(new MouseEvent('mouseup', at));
      return 'picked';
    })()`);
    await until(() => js(`document.querySelectorAll('.pdf-thumb').length === 4`), 'the thumbnails', 6000).catch(() => {});
    await pick(0);
    await pick(1, { shift: true });
    await until(() => js(`document.querySelectorAll('.pdf-thumb.selected').length === 2`), 'two pages selected', 3000).catch(() => {});
    const order = () => js(`[...document.querySelectorAll('.pdf-thumb')].map((n) => n.dataset.page).join(',')`);
    const before = await order();
    await js(`(async () => {
      const all = document.querySelectorAll('.pdf-thumb');
      const a = all[1].getBoundingClientRect();
      const b = all[0].getBoundingClientRect();
      const x = a.left + a.width / 2;
      const ev = (y) => ({ clientX: x, clientY: y, bubbles: true, cancelable: true, button: 0, buttons: 1 });
      all[1].dispatchEvent(new MouseEvent('mousedown', ev(a.top + a.height / 2)));
      const to = b.top + b.height * 0.75;
      for (let i = 1; i <= 8; i++) {
        window.dispatchEvent(new MouseEvent('mousemove', ev(a.top + a.height / 2 + (to - (a.top + a.height / 2)) * (i / 8))));
        await new Promise((r) => setTimeout(r, 30));
      }
      window.dispatchEvent(new MouseEvent('mouseup', ev(to)));
      return 1;
    })()`);
    const told = await until(async () => /already there/.test(await toasts()), 'the word', 4000).then(() => true).catch(() => false);
    const after = await order();
    check('pdf-marks: selected pages dropped among themselves stay where they are, and the window says why', told && before === after, JSON.stringify({ before, after, toasts: await toasts() }));

    // Split: the files named, and Show in Folder offered.
    const splitDir = path.join(root, 'split');
    fs.mkdirSync(splitDir, { recursive: true });
    await role('split');
    await until(() => js(`Boolean(document.querySelector('[data-role="split-every"]'))`), 'the Split dialog', 4000).catch(() => {});
    await js(`(() => {
      const el = document.querySelector('[data-role="split-every"]');
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, '2');
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return 1;
    })()`);
    answer('open', [splitDir]);
    await role('split-go');
    const named = await until(async () => /form-to-sign_Part1\.pdf to form-to-sign_Part2\.pdf/.test(await toasts()), 'the files named', 10000).then(() => true).catch(() => false);
    const show = await js(`[...document.querySelectorAll('.rw-toast button')].some((b) => b.textContent.trim() === 'Show in Folder')`);
    await capture('pdf-marks-03-split.png');
    check('pdf-marks: Split names the files it wrote and offers Show in Folder', named && show && fs.existsSync(path.join(splitDir, 'form-to-sign_Part2.pdf')), JSON.stringify({ named, show, toasts: await toasts() }));

    const complaints = await errorsIn(win);
    check('pdf-marks: the window reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('pdf-marks: the checks ran', false, err?.message || String(err));
  }
}
