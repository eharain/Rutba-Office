// PDF: Find Fields, pressed in the window. A PDF made here holds two scanned
// pages: an invoice and an application form with six tick boxes, two of them
// ticked. Recognise Text, with Find fields ticked as it is by default, reads
// both and opens the Fields pane: the invoice's supplier, number, dates and
// figures with its sums agreeing, and the form's tick boxes by the words
// beside them. A field pressed is marked on its page where it was read; Copy
// as a Row puts the invoice on the clipboard as one row a sheet takes cell by
// cell; Make Check Box Fields turns the tick boxes into the form's own check
// boxes, ticked as read, as one step Undo takes back and Redo puts back; Save
// As writes them into the file, where the form's reader finds six check boxes
// named as the page names them. The clipboard is put back.
// Run alone with RUTBA_VERIFY_ONLY=pdf-fields.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { clipboard } from 'electron';
import pdfWriter from '@rutba/pdf';
import pdfReader from '@rutba/pdf/reader';
import pdfForms from '@rutba/pdf/forms';
import { decodePng } from '@rutba/ocr/png';

const { PdfDocument } = pdfWriter;
const { openPdf } = pdfReader;
const { readFields } = pdfForms;
const fixture = (name) => fileURLToPath(new URL(`../test/fixtures/${name}`, import.meta.resolve('@rutba/ocr')));

const answer = (kind, value) => {
  (globalThis.__rutbaCheckDialogAnswers ||= []).push({ kind, answer: value });
};

/** The fixture's pictures as scanned pages, drawn at 150 dots per inch. */
function scannedPdf(file, names) {
  const doc = new PdfDocument({ title: 'Scanned post', created: '2026-10-14T09:00:00Z' });
  for (const name of names) {
    const { width, height, data } = decodePng(fs.readFileSync(fixture(name)));
    const channels = data.length / (width * height);
    const rgb = Buffer.alloc(width * height * 3);
    for (let i = 0, o = 0; o < rgb.length; i += channels, o += 3) { rgb[o] = data[i]; rgb[o + 1] = data[i + 1]; rgb[o + 2] = data[i + 2]; }
    const ref = doc.addImage({ width, height, data: rgb, colorSpace: 'DeviceRGB' });
    doc.addPage().image(ref, 0, 0, (width / 150) * 72, (height / 150) * 72);
  }
  fs.writeFileSync(file, doc.toBuffer());
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made (this block's own folder in it)
 */
export async function verifyPdfFields(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'pdf-fields');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'scanned-post.pdf');
  scannedPdf(file, ['invoice-eng.png', 'form-eng.png']);
  const truth = JSON.parse(fs.readFileSync(fixture('invoice-eng.json'), 'utf8')).fields;
  const formTruth = JSON.parse(fs.readFileSync(fixture('form-eng.json'), 'utf8')).checkboxes;
  const keptText = clipboard.readText();

  const win = await open('pdf', file);
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };
  const settle = async (read, label, ms) => {
    let got = null;
    await until(async () => (got = await read()), label, ms).catch(() => {});
    return got || null;
  };
  const tip = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith(${JSON.stringify(label)})); if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
  const role = (r) => js(`(() => { const b = document.querySelector('[data-role="${r}"]'); if (!b) return 'no ${r}'; if (b.disabled) return 'disabled ${r}'; b.click(); return 'clicked'; })()`);
  const pane = () => js(`(() => {
    const p = document.querySelector('[data-role="fields-pane"]');
    if (!p) return null;
    return [...p.querySelectorAll('[data-role="found-card"]')].map((c) => ({
      kind: c.dataset.kind, head: c.querySelector('.pdf-found-head')?.textContent || '', sums: c.querySelector('[data-role="found-sums"]')?.className || '',
      fields: Object.fromEntries([...c.querySelectorAll('.pdf-found-field')].map((b) => [b.dataset.field, b.querySelector('.v').textContent])),
      boxes: [...c.querySelectorAll('.pdf-found-box')].map((b) => ({ label: b.textContent.trim(), checked: b.dataset.checked === 'yes' })),
      make: Boolean(c.querySelector('[data-role="found-make"]')), made: Boolean(c.querySelector('[data-role="found-made"]')),
    }));
  })()`);
  const boxesOnPage = () => js(`(() => { const boxes = [...document.querySelectorAll('.pdf-page[data-index="1"] .pdf-fields [role="checkbox"]')]; return { count: boxes.length, checked: boxes.filter((b) => b.getAttribute('aria-checked') === 'true').length }; })()`);

  try {
    await until(() => js(`document.querySelectorAll('.pdf-thumb').length === 2 && document.querySelectorAll('.pdf-page canvas[data-drawn]').length > 0`), 'two pages drawn', 20000).catch(() => {});
    await js(`localStorage.setItem('rutba-ocr-languages', '["eng"]'); localStorage.removeItem('rutba-ocr-find-fields'); 1`);

    // Recognise Text, Find fields ticked by default.
    await tip('Recognise Text');
    const offered = await settle(() => js(`(() => { const c = document.querySelector('[data-role="ocr-find-fields"]'); return c ? { checked: c.checked, text: c.parentElement.textContent } : null; })()`), 'the Recognise Text dialog', 5000);
    await role('ocr-go');
    const cards = await settle(async () => { const p = await pane(); return p && p.length === 2 ? p : null; }, 'the Fields pane with both pages', 150000);
    const [inv, form] = cards || [];
    const invoiceRight = inv && inv.kind === 'invoice' && /good/.test(inv.sums) && inv.fields.supplier?.startsWith('Northgate Supplies') && inv.fields.number === truth.number && inv.fields.date === truth.date
      && inv.fields.dueDate === truth.dueDate && inv.fields.subtotal === truth.subtotal.toFixed(2) && inv.fields.tax === truth.tax.toFixed(2) && inv.fields.total === truth.total.toFixed(2) && inv.fields.currency === truth.currency;
    check('pdf-fields: Recognise Text, Find fields ticked, opens the Fields pane with the scanned invoice\'s supplier, number, dates and figures, its sums agreeing',
      offered?.checked === true && invoiceRight, JSON.stringify({ offered, inv }));
    check('pdf-fields: the scanned form\'s six tick boxes are listed by the words beside them, ticked as they are on the page',
      form?.kind === 'form' && JSON.stringify(form.boxes) === JSON.stringify(formTruth.map((c) => ({ label: c.label, checked: c.checked }))) && form.make, JSON.stringify(form));

    // A field pressed: marked on its page where it was read.
    await js(`document.querySelector('[data-role="found-card"][data-kind="invoice"] .pdf-found-field[data-field="total"]')?.click(), 1`);
    const mark = await settle(() => js(`(() => {
      const m = document.querySelector('.pdf-page[data-index="0"] [data-role="found-mark"]');
      const page = document.querySelector('.pdf-page[data-index="0"]');
      if (!m || !page) return null;
      const a = m.getBoundingClientRect(), p = page.getBoundingClientRect();
      return { x: (a.left - p.left) / p.width, y: (a.top - p.top) / p.height, w: a.width / p.width };
    })()`), 'the total marked on its page', 6000);
    await capture('pdf-fields-01-pane.png');
    check('pdf-fields: a field pressed in the pane is marked on its page where it was read',
      Boolean(mark) && mark.x > 0.7 && mark.y > 0.38 && mark.y < 0.46 && mark.w < 0.2, JSON.stringify(mark));

    // Copy as a Row: the sheet's columns, tab apart.
    clipboard.writeText('');
    await js(`document.querySelector('[data-role="found-card"][data-kind="invoice"] [data-role="found-copy-row"]')?.click(), 1`);
    await until(() => clipboard.readText().length > 0, 'the row copied', 4000).catch(() => {});
    const row = clipboard.readText().split('\t');
    check('pdf-fields: Copy as a Row puts the invoice on the clipboard as one row of date, supplier, number, subtotal, tax, total and currency',
      row.length === 7 && row[0] === truth.date && row[2] === truth.number && row[5] === truth.total.toFixed(2) && row[6] === truth.currency, JSON.stringify(row));

    // Make Check Box Fields: the form's own check boxes, ticked as read.
    // A tick box pressed brings its page into view and marks it there.
    await js(`document.querySelector('[data-role="found-card"][data-kind="form"] .pdf-found-box[data-box="1"]')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.pdf-page[data-index="1"] [data-role="found-mark"]'))`), 'the tick box marked', 6000).catch(() => {});
    await js(`document.querySelector('[data-role="found-card"][data-kind="form"] [data-role="found-make"]')?.click(), 1`);
    const madeBoxes = await settle(async () => { const b = await boxesOnPage(); return b.count === 6 ? b : null; }, 'six check boxes on the page', 10000);
    const afterMake = await settle(async () => { const p = await pane(); return p?.[1]?.made ? p[1] : null; }, 'the pane saying they are made', 5000);
    await wait(500);
    await capture('pdf-fields-02-checkboxes.png');
    check('pdf-fields: Make Check Box Fields turns the tick boxes into six check boxes of the form on the page, two ticked as read',
      madeBoxes?.count === 6 && madeBoxes.checked === 2 && Boolean(afterMake) && !afterMake.make, JSON.stringify({ madeBoxes, made: Boolean(afterMake) }));

    // One step: Undo takes them away, Redo puts them back.
    await tip('Undo');
    const undone = await settle(async () => { const b = await boxesOnPage(); const p = await pane(); return b.count === 0 && p?.[1]?.make ? b : null; }, 'the check boxes gone after Undo', 8000);
    await wait(300);
    await tip('Redo');
    const redone = await settle(async () => { const b = await boxesOnPage(); return b.count === 6 ? b : null; }, 'the check boxes back after Redo', 8000);
    check('pdf-fields: the check boxes made are one step that Undo takes back and Redo puts back', Boolean(undone) && Boolean(redone), JSON.stringify({ undone, redone }));

    // Save As: the form's reader finds them in the file, named and ticked.
    const saved = path.join(root, 'scanned-post-fields.pdf');
    answer('save', saved);
    await tip('Save As');
    const wrote = await until(() => fs.existsSync(saved) && fs.statSync(saved).size > fs.statSync(file).size, 'the saved file', 10000).then(() => true).catch(() => false);
    let fields = [];
    if (wrote) {
      await wait(300);
      fields = readFields(openPdf(fs.readFileSync(saved)));
    }
    const boxes = fields.filter((f) => f.type === 'checkbox');
    check('pdf-fields: Save As writes them into the file as check box fields on the form\'s page, named by the words beside them and ticked as read',
      wrote && boxes.length === 6 && JSON.stringify(boxes.map((f) => [f.name, f.value])) === JSON.stringify(formTruth.map((c) => [c.label, c.checked ? 'Yes' : 'Off'])) && boxes.every((f) => f.widgets[0].page === 1),
      JSON.stringify({ wrote, boxes: boxes.map((f) => [f.name, f.value, f.widgets[0]?.page]) }));
    const errors = await errorsIn(win);
    check('pdf-fields: no error was shown', !errors.length, errors.join(' | '));
  } finally {
    clipboard.writeText(keptText || '');
    await js(`localStorage.removeItem('rutba-ocr-languages'), 1`).catch(() => {});
  }
}
