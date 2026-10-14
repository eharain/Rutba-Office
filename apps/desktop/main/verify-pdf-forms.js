// PDF: filling a form and Fill & Sign, pressed in the window and read back
// from the file it saves. A form made here (text, multi-line and
// right-aligned fields, a check box, a radio group, a drop-down, a list and
// a read-only field) opens with a control over each of its fields; the Fill
// & Sign tab says plainly that a signature placed is a picture, not a
// digital signature; words are typed, a box ticked, a size chosen, a country
// picked and two colours chosen, and the read-only field cannot be changed;
// Undo takes the last value away and Redo puts it back; a typed signature
// and drawn initials are placed on the page; Save As writes a file that
// begins with every byte of the original, the values after it with the
// appearance streams other readers draw, and the stamps; a signature taken
// from a picture of one signed on paper, its paper made see-through, today's
// date and a checkmark are placed on the page too.
// Run alone with RUTBA_VERIFY_ONLY=pdf-forms.

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import zlib from 'node:zlib';
import pdfReader from '@rutba/pdf/reader';
import pdfForms from '@rutba/pdf/forms';
import pdfAnnotations from '@rutba/pdf/annotations';

const require = createRequire(import.meta.url);
const { formPdf } = require('../../../packages/pdf/test/form-fixture.js');
const { openPdf, Stream } = pdfReader;
const { readFields } = pdfForms;
const { readAnnotations } = pdfAnnotations;

/** A photo of a signature on white paper, as a PNG: dark strokes on a slightly grey white. */
function signaturePhoto(file) {
  const w = 240;
  const h = 90;
  const rows = [];
  for (let y = 0; y < h; y++) {
    const row = Buffer.alloc(1 + w * 3, 0);
    for (let x = 0; x < w; x++) {
      // A loop and a long stroke, in a dark blue ink.
      const loop = Math.abs(Math.hypot(x - 60, y - 45) - 22) < 2.6;
      const stroke = Math.abs(y - (60 - x * 0.12) - 8 * Math.sin(x / 14)) < 2.2 && x > 70 && x < 225;
      const ink = loop || stroke;
      row[1 + x * 3] = ink ? 20 : 246;
      row[2 + x * 3] = ink ? 30 : 244;
      row[3 + x * 3] = ink ? 90 : 240;
    }
    rows.push(row);
  }
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, 'latin1'), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(zlib.crc32(body) >>> 0);
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8; ihdr[9] = 2;
  fs.writeFileSync(file, Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(Buffer.concat(rows))), chunk('IEND', Buffer.alloc(0)),
  ]));
}

function driver(win, { until, wait }) {
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
  const tip = (label) => js(`(() => {
    const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith(${JSON.stringify(label)}));
    if (!b) return 'no button ${label}';
    if (b.disabled) return 'disabled ${label}';
    b.click();
    return 'clicked';
  })()`);
  const field = (name, extra = '') => `document.querySelector('.pdf-field[data-field="${name}"]${extra}')`;
  /** Words typed into a field and the field left, as a person tabs on. */
  const type = (name, text) => js(`(() => {
    const el = ${field(name)};
    if (!el) return 'no field';
    el.focus();
    const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(text)});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return 'typed';
  })()`);
  const leave = (name) => js(`(${field(name)}?.blur(), 1)`);
  /** Whether a text field's value has come back from the file (its draft no longer shown over the page). */
  const settled = (name, text) => js(`(() => { const el = ${field(name)}; return Boolean(el) && el.value === ${JSON.stringify(text)} && !el.classList.contains('shown') && document.activeElement !== el; })()`);
  const choose = (name, values) => js(`(() => {
    const el = ${field(name)};
    if (!el) return 'no field';
    const want = ${JSON.stringify(values)};
    for (const o of el.options) o.selected = want.includes(o.value);
    if (!el.multiple) el.value = want[0];
    el.dispatchEvent(new Event('change', { bubbles: true }));
    return 'chosen';
  })()`);
  const fields = () => js(`[...document.querySelectorAll('.pdf-field')].map((n) => ({ name: n.dataset.field, type: n.dataset.type, value: n.multiple ? [...n.selectedOptions].map((o) => o.value) : n.value, checked: n.getAttribute('aria-checked'), state: n.dataset.state, disabled: n.disabled }))`);
  /** A press on page 0's layer at fractions of the page, then a drag through the points given. */
  const dragOn = (points) => js(`(async () => {
    const svg = document.querySelector('.pdf-page[data-index="0"] .pdf-annots');
    if (!svg) return 'no layer';
    const r = svg.getBoundingClientRect();
    const at = ([fx, fy]) => ({ clientX: r.left + r.width * fx, clientY: r.top + r.height * fy, bubbles: true, cancelable: true, button: 0, buttons: 1 });
    const pts = ${JSON.stringify(points)};
    svg.dispatchEvent(new MouseEvent('mousedown', at(pts[0])));
    for (const p of pts.slice(1)) { window.dispatchEvent(new MouseEvent('mousemove', at(p))); await new Promise((res) => setTimeout(res, 20)); }
    window.dispatchEvent(new MouseEvent('mouseup', at(pts[pts.length - 1])));
    return 'pressed';
  })()`);
  const stamps = () => js(`document.querySelectorAll('.pdf-page[data-index="0"] .pdf-annot[data-type="Stamp"]').length`);
  return { wc, js, capture, tab, role, tip, type, leave, settled, choose, fields, dragOn, stamps };
}

const answer = (kind, value) => {
  (globalThis.__rutbaCheckDialogAnswers ||= []).push({ kind, answer: value });
};

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made (a folder of this block's own is made in it)
 */
export async function verifyPdfForms(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'pdf-forms');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'application.pdf');
  fs.writeFileSync(file, formPdf());
  const original = fs.readFileSync(file);

  const win = await open('pdf', file);
  const d = driver(win, { until, wait });
  const listed = await until(async () => (await d.fields()).length === 10, 'a control over each field', 20000).then(() => true).catch(() => false);
  const controls = await d.fields();
  check('pdf-forms: a control lies over each of the form\'s fields, of its kind, the read-only one disabled', listed
    && controls.map((c) => c.type).join() === 'text,text,text,checkbox,radio,radio,radio,combo,list,text' && controls.find((c) => c.name === 'locked')?.disabled === true,
  JSON.stringify(controls));

  await d.tab('Fill & Sign');
  await until(() => d.js(`Boolean(document.querySelector('[data-role="sign-notice"]'))`), 'the notice', 4000).catch(() => {});
  const notice = await d.js(`document.querySelector('[data-role="sign-notice"]')?.textContent || ''`);
  check('pdf-forms: the Fill & Sign tab says plainly that a signature placed is a picture, not a digital signature', /not a digital signature/.test(notice), notice);

  // Words typed and the field left; the next field typed while the first is written.
  await d.type('name', 'Ejaz Arain');
  await d.leave('name');
  const named = await until(() => d.settled('name', 'Ejaz Arain'), 'the name written', 10000).then(() => true).catch(() => false);
  await d.type('address', 'Flat 2, 10 Long Street, Lahore, a line long enough to wrap');
  await d.leave('address');
  await d.type('amount', '1,250.00');
  await d.leave('amount');
  const typed = await until(async () => (await d.settled('address', 'Flat 2, 10 Long Street, Lahore, a line long enough to wrap')) && (await d.settled('amount', '1,250.00')), 'the address and amount written', 10000).then(() => true).catch(() => false);
  check('pdf-forms: words typed into text fields, a multi-line one among them, are written when the field is left', named && typed, JSON.stringify(await d.fields()));

  await d.js(`(document.querySelector('.pdf-field[data-field="agree"]').click(), 1)`);
  const ticked = await until(() => d.js(`document.querySelector('.pdf-field[data-field="agree"]').getAttribute('aria-checked') === 'true'`), 'the box ticked', 8000).then(() => true).catch(() => false);
  await d.js(`(document.querySelector('.pdf-field[data-field="size"][data-state="M"]').click(), 1)`);
  const sized = await until(() => d.js(`[...document.querySelectorAll('.pdf-field[data-field="size"]')].map((n) => n.getAttribute('aria-checked')).join() === 'false,true,false'`), 'M chosen', 8000).then(() => true).catch(() => false);
  check('pdf-forms: a check box is ticked and one radio button of a group chosen, the others left off', ticked && sized, JSON.stringify(await d.fields()));

  await d.choose('country', ['IN']);
  const country = await until(() => d.js(`document.querySelector('.pdf-field[data-field="country"]').value === 'IN'`), 'the country', 8000).then(() => true).catch(() => false);
  await d.choose('colours', ['Green', 'Blue']);
  await wait(300);
  const colours = await until(async () => JSON.stringify((await d.fields()).find((f) => f.name === 'colours')?.value) === '["Green","Blue"]', 'two colours', 8000).then(() => true).catch(() => false);
  check('pdf-forms: an option is chosen from a drop-down and two from a list', country && colours, JSON.stringify(await d.fields()));
  await d.capture('pdf-forms-01-filled.png');

  // Undo takes the last value away; Redo puts it back.
  await d.tip('Undo');
  const undone = await until(async () => JSON.stringify((await d.fields()).find((f) => f.name === 'colours')?.value) === '[]', 'the colours taken away', 8000).then(() => true).catch(() => false);
  await d.tip('Redo');
  const redone = await until(async () => JSON.stringify((await d.fields()).find((f) => f.name === 'colours')?.value) === '["Green","Blue"]', 'the colours back', 8000).then(() => true).catch(() => false);
  check('pdf-forms: Undo takes the last value away and Redo puts it back', undone && redone, JSON.stringify({ undone, redone }));

  // A typed signature, placed where the page is clicked.
  await d.role('sign-signature');
  const dialog = await until(() => d.js(`Boolean(document.querySelector('[data-role="sign-dialog"]'))`), 'the signature dialog', 5000).then(() => true).catch(() => false);
  const said = await d.js(`document.querySelector('[data-role="sign-note"]')?.textContent || ''`);
  await d.js(`(() => { const el = document.querySelector('[data-role="sign-text"]'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, 'Ejaz Arain'); el.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
  await d.capture('pdf-forms-02-sign-type.png');
  await d.role('sign-place');
  await until(() => d.js(`document.querySelector('.pdf-annots')?.dataset.tool === 'stamp'`), 'the signature in hand', 4000).catch(() => {});
  await d.dragOn([[0.25, 0.8], [0.25, 0.8]]);
  const signed = await until(async () => (await d.stamps()) === 1, 'the signature on the page', 10000).then(() => true).catch(() => false);
  check('pdf-forms: a typed signature is made in a dialog that says it is not a digital signature, and placed where the page is clicked', dialog && /not a digital signature/.test(said) && signed,
    JSON.stringify({ dialog, said, signed }));

  // Drawn initials.
  await d.role('sign-initials');
  await until(() => d.js(`Boolean(document.querySelector('[data-role="sign-dialog"]'))`), 'the initials dialog', 5000).catch(() => {});
  await d.role('sign-draw');
  await until(() => d.js(`Boolean(document.querySelector('[data-role="sign-pad"]'))`), 'the pad', 3000).catch(() => {});
  await d.js(`(async () => {
    const pad = document.querySelector('[data-role="sign-pad"]');
    const r = pad.getBoundingClientRect();
    const at = (x, y) => ({ clientX: r.left + x, clientY: r.top + y, bubbles: true, button: 0, buttons: 1 });
    const stroke = async (pts) => {
      pad.dispatchEvent(new MouseEvent('mousedown', at(...pts[0])));
      for (const p of pts.slice(1)) { window.dispatchEvent(new MouseEvent('mousemove', at(...p))); await new Promise((res) => setTimeout(res, 15)); }
      window.dispatchEvent(new MouseEvent('mouseup', at(...pts[pts.length - 1])));
    };
    await stroke([[60, 110], [80, 40], [100, 110], [90, 80], [70, 80]]);
    await stroke([[140, 40], [140, 110], [180, 110]]);
    return 1;
  })()`);
  await d.capture('pdf-forms-03-sign-draw.png');
  await d.role('sign-place');
  await until(() => d.js(`document.querySelector('.pdf-annots')?.dataset.tool === 'stamp'`), 'the initials in hand', 4000).catch(() => {});
  await d.dragOn([[0.6, 0.8], [0.6, 0.8]]);
  const initialled = await until(async () => (await d.stamps()) === 2, 'the initials on the page', 10000).then(() => true).catch(() => false);
  check('pdf-forms: drawn initials are placed on the page as a second stamp', initialled, JSON.stringify({ stamps: await d.stamps() }));
  await d.capture('pdf-forms-04-signed.png');

  // A signature from a picture of one signed on paper, the paper made see-through.
  const photo = path.join(root, 'signature-photo.png');
  signaturePhoto(photo);
  await d.role('sign-change');
  await until(() => d.js(`Boolean(document.querySelector('[data-role="sign-dialog"]'))`), 'the signature dialog again', 5000).catch(() => {});
  await d.role('sign-image');
  answer('open', [photo]);
  await d.role('sign-choose');
  const pictured = await until(() => d.js(`Boolean(document.querySelector('[data-role="sign-picture"] img')?.naturalWidth)`), 'the picture in the dialog', 8000).then(() => true).catch(() => false);
  const clearPaper = await d.js(`(async () => {
    const img = document.querySelector('[data-role="sign-picture"] img');
    if (!img) return null;
    const c = document.createElement('canvas');
    c.width = img.naturalWidth; c.height = img.naturalHeight;
    const ctx = c.getContext('2d');
    ctx.drawImage(img, 0, 0);
    const d = ctx.getImageData(0, 0, c.width, c.height).data;
    let clear = 0, ink = 0;
    for (let i = 3; i < d.length; i += 4) { if (d[i] === 0) clear++; else if (d[i] > 200) ink++; }
    return { clear, ink, w: c.width, h: c.height };
  })()`);
  await d.capture('pdf-forms-05-sign-picture.png');
  await d.role('sign-place');
  await until(() => d.js(`document.querySelector('.pdf-annots')?.dataset.tool === 'stamp'`), 'the picture in hand', 4000).catch(() => {});
  await d.dragOn([[0.25, 0.9], [0.25, 0.9]]);
  const third = await until(async () => (await d.stamps()) === 3, 'the picture signature on the page', 10000).then(() => true).catch(() => false);
  check('pdf-forms: a signature taken from a picture of one signed on paper has its paper made see-through and is placed on the page', pictured && third && clearPaper && clearPaper.clear > clearPaper.ink && clearPaper.ink > 50,
    JSON.stringify({ pictured, third, clearPaper }));

  // Today's date and a checkmark, where the page is clicked.
  await d.role('fill-date');
  await d.dragOn([[0.6, 0.9], [0.6, 0.9]]);
  const dated = await until(() => d.js(`document.querySelectorAll('.pdf-page[data-index="0"] .pdf-annot[data-type="FreeText"]').length === 1`), 'the date', 8000).then(() => true).catch(() => false);
  await d.role('fill-check');
  await d.dragOn([[0.8, 0.9], [0.8, 0.9]]);
  const ticked2 = await until(async () => (await d.stamps()) === 4, 'the checkmark', 8000).then(() => true).catch(() => false);
  await d.js(`(document.querySelector('[data-role="fill-check"]').click(), 1)`);
  await d.js(`(() => { const v = document.querySelector('.pdf-view'); v.scrollTop = v.scrollHeight; return 1; })()`);
  await d.capture('pdf-forms-06-date-check.png');
  check('pdf-forms: today\'s date and a checkmark are placed where the page is clicked', dated && ticked2, JSON.stringify({ dated, ticked2 }));

  // Save As: the original, then the values and stamps, read back.
  await d.tab('Home');
  const saved = path.join(root, 'application-filled.pdf');
  answer('save', saved);
  await d.tip('Save As');
  const wrote = await until(() => fs.existsSync(saved) && fs.statSync(saved).size > original.length, 'the saved file', 10000).then(() => true).catch(() => false);
  let readBack = {};
  if (wrote) {
    const bytes = fs.readFileSync(saved);
    const pdf = openPdf(bytes);
    const values = Object.fromEntries(readFields(pdf).map((f) => [f.name, f.value]));
    const nameField = readFields(pdf).find((f) => f.name === 'name');
    const ap = pdf.resolve(pdf.resolve(pdf.getObject(Number(nameField.id)).get('AP'))?.get('N'));
    const stampList = readAnnotations(pdf).filter((a) => a.type === 'Stamp');
    const today = new Intl.DateTimeFormat('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }).format(new Date());
    readBack = {
      prefix: bytes.subarray(0, original.length).equals(original),
      values,
      appearance: ap instanceof Stream && /\(Ejaz Arain\) Tj/.test(pdf.decode(ap).toString('latin1')),
      stamps: stampList.map((a) => `${a.icon}:${a.hasAppearance}`),
      date: readAnnotations(pdf).find((a) => a.type === 'FreeText')?.contents === today,
    };
  }
  check('pdf-forms: Save As writes the original bytes unchanged with the values, their appearances, the three signatures, the date and the checkmark appended', wrote && readBack.prefix
    && readBack.values.name === 'Ejaz Arain' && /Long Street/.test(readBack.values.address) && readBack.values.amount === '1,250.00' && readBack.values.agree === 'Yes'
    && readBack.values.size === 'M' && readBack.values.country === 'IN' && JSON.stringify(readBack.values.colours) === '["Green","Blue"]' && readBack.values.locked === 'fixed'
    && readBack.appearance && readBack.stamps.join() === 'Signature:true,Initials:true,Signature:true,Check:true' && readBack.date, JSON.stringify({ wrote, ...readBack }));
  const errors = await errorsIn(win);
  check('pdf-forms: no error was shown', !errors.length, errors.join(' | '));
}
