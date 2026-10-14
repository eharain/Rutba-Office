// Worksheets: Data → From Picture → Receipt or Invoice, pressed in the
// window. Three files are chosen at once: a till receipt as a picture, an
// invoice as a page of a PDF, and the same invoice with its total misprinted.
// Each is read as a document on the main process's threads and shown as a
// row of a preview beside its picture: the first two with their sums agreeing
// and nothing marked, the third with its figures marked to check. A value
// focused shows the box it was read from on the picture; Arabic and Urdu
// chosen say what is not read yet; corrected cells go in as corrected; Insert
// puts a row of headings and one row a file in at the selected cell, the date
// a date and the amounts numbers, as one step Undo takes back and Redo puts
// back. Run alone with RUTBA_VERIFY_ONLY=receipt-row.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import pdfWriter from '@rutba/pdf';
import { decodePng } from '@rutba/ocr/png';
import { buildXlsx } from '@rutba/ooxml/build';

const { PdfDocument } = pdfWriter;
const fixture = (name) => fileURLToPath(new URL(`../test/fixtures/${name}`, import.meta.resolve('@rutba/ocr')));

const answer = (kind, value) => {
  (globalThis.__rutbaCheckDialogAnswers ||= []).push({ kind, answer: value });
};

/** A one-page PDF holding a picture drawn at 150 dots per inch, as a scanner writes one. */
function scannedPdf(png, file) {
  const { width, height, data } = decodePng(fs.readFileSync(png));
  const channels = data.length / (width * height);
  const rgb = Buffer.alloc(width * height * 3);
  for (let i = 0, o = 0; o < rgb.length; i += channels, o += 3) { rgb[o] = data[i]; rgb[o + 1] = data[i + 1]; rgb[o + 2] = data[i + 2]; }
  const doc = new PdfDocument({ title: 'Scanned invoice', created: '2026-10-14T09:00:00Z' });
  const ref = doc.addImage({ width, height, data: rgb, colorSpace: 'DeviceRGB' });
  doc.addPage().image(ref, 0, 0, (width / 150) * 72, (height / 150) * 72);
  fs.writeFileSync(file, doc.toBuffer());
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the files are made (this block's own folder in it)
 */
export async function verifyReceiptRow(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc, sessionFor } = h;
  const root = path.join(dir, 'receipt-row');
  fs.mkdirSync(root, { recursive: true });
  const book = path.join(root, 'expenses.xlsx');
  fs.writeFileSync(book, buildXlsx({ sheets: [{ name: 'Expenses', rows: [[]] }] }));
  const receipt = path.join(root, 'till-receipt.png');
  fs.copyFileSync(fixture('receipt-eng.png'), receipt);
  const invoice = path.join(root, 'invoice-2675.pdf');
  scannedPdf(fixture('invoice-eng.png'), invoice);
  const wrong = path.join(root, 'invoice-misprinted.png');
  fs.copyFileSync(fixture('invoice-eng-wrong.png'), wrong);
  const truth = JSON.parse(fs.readFileSync(fixture('invoice-eng.json'), 'utf8')).fields;

  const win = await open('sheets', book);
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };
  const model = () => doc.model({ id: sessionFor('sheet').id });
  const cell = (ref) => (model().cells || []).find((c) => c.ref === ref) || null;
  const text = (ref) => String(cell(ref)?.text ?? '');
  const grid = () => js(`[...document.querySelectorAll('[data-role="rr-grid"] tbody tr')].map((tr) => ({
    state: tr.dataset.state,
    values: [...tr.querySelectorAll('td[data-field] input')].map((i) => i.value),
    marked: [...tr.querySelectorAll('td[data-check="yes"]')].map((td) => td.dataset.field),
  }))`);
  const setCell = (row, field, value) => js(`(() => {
    const input = document.querySelectorAll('[data-role="rr-grid"] tbody tr')[${row}]?.querySelector('td[data-field="${field}"] input');
    if (!input) return false;
    input.focus();
    input.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, ${JSON.stringify(value)});
    input.dispatchEvent(new Event('input', { bubbles: true }));
    return true;
  })()`);
  // A window the checks drive is not the focused one, so focus fires no events of itself: they are sent too.
  const focusCell = (row, field) => js(`(() => {
    const i = document.querySelectorAll('[data-role="rr-grid"] tbody tr')[${row}]?.querySelector('td[data-field="${field}"] input');
    if (!i) return false;
    i.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    i.focus();
    i.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    return true;
  })()`);
  // What the window shows once it shows it: waited for, then read (the harness's until answers only whether it came).
  const settle = async (read, label, ms) => {
    let got = null;
    await until(async () => (got = await read()), label, ms).catch(() => {});
    return got || null;
  };
  const tipButton = (label) => `[...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('${label}'))`;

  try {
    await until(() => js(`Boolean(document.querySelector('.sh-colheads .sh-head'))`), 'the grid', 10000).catch(() => {});
    await js(`localStorage.setItem('rutba-ocr-languages', '["eng"]'), 1`);

    // Data → From Picture → Receipt or Invoice…, three files chosen at once.
    answer('open', [receipt, invoice, wrong]);
    await js(`[...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'Data')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('[data-role="data-from-picture"]'))`), 'the From Picture button', 5000).catch(() => {});
    await js(`document.querySelector('[data-role="data-from-picture"]')?.click(), 1`);
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.includes('Receipt or Invoice'))`), 'the menu', 4000).catch(() => {});
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.includes('Receipt or Invoice')); if (!b) return 'no item'; b.click(); return 'clicked'; })()`);
    const shown = await until(() => js(`document.querySelectorAll('[data-role="rr-grid"] tbody tr').length === 3`), 'three rows', 15000).then(() => true).catch(() => false);
    const read = await until(async () => (await grid()).every((r) => r.state === 'done'), 'the three read', 120000).then(() => true).catch(() => false);
    const rows = await grid();
    const [r1, r2, r3] = rows;
    const receiptRight = JSON.stringify(r1?.values) === JSON.stringify(['2025-09-04', 'Harbour Print Co.', '342457', '181.08', '9.05', '190.13', 'USD']);
    const money = (v) => v.toFixed(2);
    const invoiceRight = r2 && r2.values[0] === truth.date && r2.values[2] === truth.number && r2.values[3] === money(truth.subtotal) && r2.values[4] === money(truth.tax) && r2.values[5] === money(truth.total) && r2.values[6] === truth.currency && r2.values[1].startsWith('Northgate Supplies');
    check('receipt-row: Data → From Picture → Receipt or Invoice reads a picture and a PDF page into a row each: date, supplier, number, subtotal, tax, total and currency',
      pressed === 'clicked' && shown && read && receiptRight && invoiceRight, JSON.stringify({ pressed, shown, read, rows }));
    check('receipt-row: figures go in as read on the pages whose sums agree, and are marked to check on the page whose total is misprinted',
      r1?.marked.length === 0 && r2?.marked.length === 0 && JSON.stringify(r3?.marked) === JSON.stringify(['subtotal', 'tax', 'total']) && r3?.values[5] === money(truth.total + 10),
      JSON.stringify(rows.map((r) => r.marked)));

    // The misprinted page: its note says why, its total's box marked on its picture.
    await focusCell(2, 'total');
    const note3 = await settle(async () => {
      const n = await js(`(() => { const n = document.querySelector('[data-role="rr-note"]'); return n ? { tone: n.dataset.tone, text: n.textContent, mark: document.querySelector('[data-role="ocr-boxview"]')?.dataset.marked, loupe: document.querySelector('[data-role="rr-where"]')?.dataset.field || '' } : null; })()`);
      return n?.tone === 'warn' && n.mark === 'yes' && n.loupe === 'total' ? n : null;
    }, 'the third row chosen', 5000);
    const look = await js(`(() => {
      const pic = document.querySelector('.rr-picture')?.getBoundingClientRect();
      const grid = document.querySelector('.rr-grid-wrap')?.getBoundingClientRect();
      const body = document.querySelector('.rr-body')?.getBoundingClientRect();
      const inputs = [...document.querySelectorAll('[data-role="rr-grid"] tbody tr:first-child input')].map((i) => i.scrollWidth <= i.clientWidth + 1);
      return { beside: Boolean(pic && grid && pic.right <= grid.left + 2), fits: Boolean(body && grid && grid.right <= body.right + 1), unclipped: inputs.filter(Boolean).length, inputs: inputs.length };
    })()`);
    await capture('receipt-row-01-preview.png');
    check('receipt-row: the page whose sums do not agree says so, and its total is marked on its picture beside the preview',
      Boolean(note3) && /do not agree/.test(note3.text) && look.beside && look.fits && look.unclipped === look.inputs, JSON.stringify({ note3, look }));

    // A value focused on the invoice: the box it was read from, on the page drawn from the PDF.
    await focusCell(1, 'total');
    const mark = await settle(async () => {
      const m = await js(`(() => {
        const n = document.querySelector('[data-role="rr-note"]');
        const view = document.querySelector('[data-role="ocr-boxview"]')?.getBoundingClientRect();
        const box = document.querySelector('[data-role="ocr-boxmark"]')?.getBoundingClientRect();
        return n && view && box ? { tone: n.dataset.tone, text: n.textContent, x: (box.left - view.left) / view.width, y: (box.top - view.top) / view.height, w: box.width / view.width } : null;
      })()`);
      return m?.tone === 'good' ? m : null;
    }, 'the invoice chosen', 5000);
    await capture('receipt-row-02-box.png');
    // The invoice's total is printed at the right, 42 percent of the way down the page.
    check('receipt-row: a value focused shows where it was read on the page, and an agreeing page says its figures go in as read',
      Boolean(mark) && /agree/.test(mark.text) && mark.x > 0.7 && mark.y > 0.38 && mark.y < 0.46 && mark.w < 0.2, JSON.stringify(mark));

    // Arabic and Urdu chosen: what is not read yet is said.
    await js(`document.querySelector('[data-role="rr-languages"]')?.click(), 1`);
    await until(() => js(`document.querySelectorAll('.rr-body .ocr-lang').length === 6`), 'the languages', 4000).catch(() => {});
    const tick = (codes) => js(`(() => { for (const n of document.querySelectorAll('.rr-body .ocr-lang')) { const box = n.querySelector('input'); if (box.checked !== ${JSON.stringify(codes)}.includes(n.dataset.code)) box.click(); } return 1; })()`);
    await tick(['eng', 'ara']);
    const arabic = await settle(() => js(`(() => { const d = document.querySelector('.rr-body [data-role="ocr-digits-note"]'); return d && !document.querySelector('.rr-body [data-role="ocr-urdu-note"]') ? d.textContent : null; })()`), 'the digits note', 4000);
    await tick(['eng', 'ara', 'urd']);
    const urdu = await settle(() => js(`(() => { const u = document.querySelector('.rr-body [data-role="ocr-urdu-note"]'); return u && document.querySelector('.rr-body [data-role="ocr-digits-note"]') ? u.textContent : null; })()`), 'the Nastaliq note', 4000);
    await capture('receipt-row-03-limits.png');
    await tick(['eng']);
    const gone = await until(() => js(`!document.querySelector('.rr-body [data-role="ocr-limits"]')`), 'the notes gone', 4000).then(() => true).catch(() => false);
    check('receipt-row: choosing Arabic says Arabic-Indic and Urdu digits are not read yet, Urdu adds that Nastaliq is beta, English alone says neither',
      /Arabic-Indic and Urdu digits are not read yet/.test(arabic || '') && /beta/.test(urdu || '') && /Nastaliq/.test(urdu || '') && gone, JSON.stringify({ arabic, urdu, gone }));

    // Corrections, then Insert: headings and three rows, typed.
    await setCell(1, 'supplier', 'Northgate Supplies Ltd');
    await setCell(2, 'total', money(truth.total));
    const cleared = await settle(async () => {
      const g = await grid();
      return g[1]?.values[1] === 'Northgate Supplies Ltd' && g[2]?.values[5] === money(truth.total) && !g[2].marked.includes('total') ? g : null;
    }, 'the corrections', 4000);
    const headed = await js(`document.querySelector('[data-role="rr-headings"]')?.checked`);
    await wait(300);
    const insertLabel = await js(`document.querySelector('[data-role="rr-insert"]')?.textContent || ''`);
    await js(`document.querySelector('[data-role="rr-insert"]')?.click(), 1`);
    const landed = await until(() => text('A1') === 'Date' && text('A2') === '2025-09-04' && text('B3') === 'Northgate Supplies Ltd' && text('F4') === '1,601.12', 'the rows in the sheet', 10000).then(() => true).catch(() => false);
    const got = {
      A1: text('A1'), G1: text('G1'), A2: text('A2'), C2: text('C2'), F2: text('F2'), A3: text('A3'), C3: text('C3'), D3: text('D3'), F4: text('F4'),
      dateRight: cell('A2')?.align === 'right', numberRight: cell('F2')?.align === 'right', idLeft: cell('C2')?.align !== 'right',
    };
    await capture('receipt-row-04-inserted.png');
    check('receipt-row: Insert puts headings and a row a file in at the selected cell, the date a date, the amounts numbers and the corrections as corrected',
      Boolean(cleared) && headed === true && /3/.test(insertLabel) && landed && got.G1 === 'Currency' && got.C2 === '342457' && got.F2 === '190.13' && got.A3 === truth.date && got.C3 === truth.number && got.D3 === '1,334.27' && got.dateRight && got.numberRight && got.idLeft,
      JSON.stringify({ cleared: Boolean(cleared), headed, insertLabel, got }));

    // One step: Undo takes it all back, Redo puts it back.
    await js(`${tipButton('Undo')}?.click(), 1`);
    const undone = await until(() => !text('A1') && !text('A2') && !text('F4'), 'the rows gone after Undo', 5000).then(() => true).catch(() => false);
    await until(() => js(`Boolean(${tipButton('Redo')} && !${tipButton('Redo')}.disabled)`), 'Redo ready', 4000).catch(() => {});
    await wait(300);
    await js(`${tipButton('Redo')}?.click(), 1`);
    const redone = await until(() => text('A1') === 'Date' && text('F4') === '1,601.12' && text('A2') === '2025-09-04', 'the rows back after Redo', 5000).then(() => true).catch(() => false);
    check('receipt-row: the rows are one step that Undo takes back and Redo puts back', undone && redone, JSON.stringify({ undone, redone }));
    const errors = await errorsIn(win);
    check('receipt-row: no error was shown', !errors.length, errors.join(' | '));
  } finally {
    await js(`localStorage.removeItem('rutba-ocr-languages'), 1`).catch(() => {});
  }
}
