// Mail: an attachment's details, in the window. Into an account of the
// check's own arrive an invoice as a PDF, a till receipt as a picture, a
// form as a picture and the receipt again; nothing is read as they arrive.
// Opening the invoice reads its PDF on the text recognition threads and a
// note under the attachments says who it is from, the total, the due date and
// the number; a detail pressed shows where on the page it was read. The
// receipt opens with its own note; the form is read and shows none. Opened
// again, the invoice shows its note at once, read from what was kept with the
// message, and nothing is read again. In Urdu the note reads right to left,
// and the reading in Urdu and English says that Arabic-Indic and Urdu digits
// are not read yet and that Nastaliq is beta. Everything the check made is
// taken away at the end. Run alone with RUTBA_VERIFY_ONLY=mail-invoice.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import pdfWriter from '@rutba/pdf';
import { decodePng } from '@rutba/ocr/png';

const { PdfDocument } = pdfWriter;
const fixture = (name) => fileURLToPath(new URL(`../test/fixtures/${name}`, import.meta.resolve('@rutba/ocr')));

/** A one-page PDF holding a picture drawn at 150 dots per inch, as a scanner writes one. */
function scannedPdf(png) {
  const { width, height, data } = decodePng(fs.readFileSync(png));
  const channels = data.length / (width * height);
  const rgb = Buffer.alloc(width * height * 3);
  for (let i = 0, o = 0; o < rgb.length; i += channels, o += 3) { rgb[o] = data[i]; rgb[o + 1] = data[i + 1]; rgb[o + 2] = data[i + 2]; }
  const doc = new PdfDocument({ title: 'Invoice INV-2675', created: '2026-08-03T09:00:00Z' });
  const ref = doc.addImage({ width, height, data: rgb, colorSpace: 'DeviceRGB' });
  doc.addPage().image(ref, 0, 0, (width / 150) * 72, (height / 150) * 72);
  return doc.toBuffer();
}

/** A message on the wire with one file attached. */
function withAttachment({ from, subject, id, body, filename, type, bytes }) {
  const boundary = `----rutba-${id.replace(/[^a-z0-9]/gi, '')}`;
  const b64 = Buffer.from(bytes).toString('base64').replace(/.{76}/g, '$&\r\n');
  return [
    `From: ${from}`, 'To: Check Owner <owner@check.example>', `Subject: ${subject}`, `Date: ${new Date().toUTCString()}`, `Message-ID: <${id}>`, 'MIME-Version: 1.0',
    `Content-Type: multipart/mixed; boundary="${boundary}"`, '', `--${boundary}`, 'Content-Type: text/plain; charset=utf-8', '', body, '',
    `--${boundary}`, `Content-Type: ${type}`, `Content-Disposition: attachment; filename="${filename}"`, 'Content-Transfer-Encoding: base64', '', b64, `--${boundary}--`, '',
  ].join('\r\n');
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made (this block's own folder in it)
 */
export async function verifyMailInvoice(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'mail-invoice');
  fs.mkdirSync(root, { recursive: true });
  const truth = JSON.parse(fs.readFileSync(fixture('invoice-eng.json'), 'utf8')).fields;
  let win = null;
  let js = null;
  let accountId = null;
  let language = null;
  const capture = async (w, name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    w.webContents.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await w.webContents.capturePage()).toPNG());
  };
  const settle = async (read, label, ms) => {
    let got = null;
    await until(async () => (got = await read()), label, ms).catch(() => {});
    return got || null;
  };
  try {
    win = await open('mail');
    js = (code) => win.webContents.executeJavaScript(code);
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args ?? {})})`);
    const reads = async () => { const s = await js('window.rutbaOffice.ocr.stats()'); return s.done + s.failed + s.cancelled + s.busy + s.queued; };
    language = await js(`window.rutbaOffice.store.get({ key: 'language', fallback: null })`);
    await js(`localStorage.setItem('rutba-ocr-languages', '["eng"]'), 1`);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the sidebar', 10000).catch(() => {});

    // The mail arrives: nothing is read.
    const made = await mail('addAccount', { account: { email: 'owner@check.example', name: 'Invoice Owner', local: true } });
    accountId = made.id;
    const before = await reads();
    const pdf = scannedPdf(fixture('invoice-eng.png'));
    fs.writeFileSync(path.join(root, 'invoice-2675.pdf'), pdf);
    const arrivals = [
      withAttachment({ from: 'Northgate Supplies <accounts@northgate.example>', subject: 'Invoice INV-2675', id: 'inv-2675@northgate.example', body: 'Please find our invoice attached.', filename: 'invoice-2675.pdf', type: 'application/pdf', bytes: pdf }),
      withAttachment({ from: 'Harbour Print Co. <till@harbourprint.example>', subject: 'Your receipt', id: 'receipt-342457@harbourprint.example', body: 'Thank you for shopping with us.', filename: 'receipt.png', type: 'image/png', bytes: fs.readFileSync(fixture('receipt-eng.png')) }),
      withAttachment({ from: 'Copperline Electrical <forms@copperline.example>', subject: 'Your application form', id: 'form-7104@copperline.example', body: 'The form you sent us, for your records.', filename: 'application.png', type: 'image/png', bytes: fs.readFileSync(fixture('form-eng.png')) }),
      withAttachment({ from: 'Harbour Print Co. <till@harbourprint.example>', subject: 'Receipt, Karachi branch', id: 'receipt-karachi@harbourprint.example', body: 'Thank you for shopping with us.', filename: 'receipt-karachi.png', type: 'image/png', bytes: fs.readFileSync(fixture('receipt-eng.png')) }),
    ];
    for (const raw of arrivals) await mail('deliverTest', { accountId, folder: 'Inbox', raw });
    await js(`(() => { const b = [...document.querySelectorAll('.ml-accounts button')].find((n) => n.textContent.trim().startsWith('Invoice Owner')); b?.click(); return Boolean(b); })()`);
    await until(() => js(`document.querySelectorAll('.ml-row').length >= 4`), 'the four messages', 10000).catch(() => {});
    await wait(1500);
    const afterArrival = await reads();
    check('mail-invoice: mail with invoices attached arrives and is listed without anything being read', afterArrival === before, JSON.stringify({ before, afterArrival }));

    const openSubject = (subject) => js(`(() => { const r = [...document.querySelectorAll('.ml-row .ml-subject')].find((n) => n.textContent.trim() === ${JSON.stringify(subject)}); r?.click(); return Boolean(r); })()`);
    const noteOf = () => js(`(() => {
      const n = document.querySelector('[data-role="ml-docnote"]');
      if (!n) return null;
      return { state: n.dataset.state, kind: n.dataset.kind || '', head: document.querySelector('.ml-head h2')?.textContent || '', title: n.querySelector('.ml-docnote-title')?.textContent || '',
        facts: [...n.querySelectorAll('.ml-docnote-fact')].map((b) => ({ field: b.dataset.field, sure: b.dataset.sure, text: b.querySelector('.v')?.textContent || '' })) };
    })()`);

    // The invoice opened: read now, its note under the attachment.
    await openSubject('Invoice INV-2675');
    const reading = await settle(async () => { const n = await noteOf(); return n && n.head.includes('INV-2675') ? n : null; }, 'reading begun', 10000);
    const invoice = await settle(async () => { const n = await noteOf(); return n?.state === 'done' && n.head.includes('INV-2675') ? n : null; }, 'the invoice note', 90000);
    const fact = (n, field) => n?.facts.find((f) => f.field === field)?.text || '';
    const due = await js(`new Intl.DateTimeFormat('en', { dateStyle: 'medium', timeZone: 'UTC' }).format(Date.parse('${truth.dueDate}T00:00:00Z'))`);
    check('mail-invoice: opening a message reads its PDF invoice, and a note says who it is from, the total, the due date and the number',
      reading?.state === 'reading' && invoice?.kind === 'invoice' && fact(invoice, 'from').startsWith('Northgate Supplies') && fact(invoice, 'total') === '£1,601.12' && fact(invoice, 'dueDate') === due && fact(invoice, 'number') === truth.number && invoice.facts.every((f) => f.sure === 'yes'),
      JSON.stringify({ reading: reading?.state, invoice, due }));

    // A detail pressed: the part of the page it was read from.
    await js(`document.querySelector('[data-role="ml-docnote"] [data-field="total"]')?.click(), 1`);
    const peek = await settle(() => js(`(() => {
      const p = document.querySelector('[data-role="ml-docnote-peek"]');
      const view = p?.querySelector('[data-role="ocr-boxview"]')?.getBoundingClientRect();
      const box = p?.querySelector('[data-role="ocr-boxmark"]')?.getBoundingClientRect();
      const loupe = p?.querySelector('[data-role="ocr-loupe"]')?.getBoundingClientRect();
      return p && view && box && loupe && loupe.height > 20 ? { field: p.dataset.field, head: p.textContent, x: (box.left - view.left) / view.width, y: (box.top - view.top) / view.height, loupe: Math.round(loupe.width) } : null;
    })()`), 'where the total was read', 15000);
    await capture(win, 'mail-invoice-01-total.png');
    check('mail-invoice: the total pressed shows the part of the invoice it was read from, marked on the page',
      peek?.field === 'total' && /where it was read in invoice-2675\.pdf/.test(peek.head) && peek.x > 0.7 && peek.y > 0.38 && peek.y < 0.46 && peek.loupe > 200, JSON.stringify(peek));

    // The receipt, a picture: its own note. The form: read, and no note.
    await openSubject('Your receipt');
    const receipt = await settle(async () => { const n = await noteOf(); return n?.state === 'done' && n.head.includes('Your receipt') ? n : null; }, 'the receipt note', 90000);
    await capture(win, 'mail-invoice-02-receipt.png');
    check('mail-invoice: a picture of a till receipt gives a receipt\'s note: from, total and number',
      receipt?.kind === 'receipt' && fact(receipt, 'from') === 'Harbour Print Co.' && fact(receipt, 'total') === '$190.13' && fact(receipt, 'number') === '342457', JSON.stringify(receipt));
    const readsBeforeForm = await reads();
    await openSubject('Your application form');
    await until(async () => (await reads()) > readsBeforeForm && (await js('window.rutbaOffice.ocr.stats()')).busy === 0, 'the form read', 90000).catch(() => {});
    await until(async () => (await js(`document.querySelector('.ml-head h2')?.textContent || ''`)).includes('application form'), 'the form open', 5000).catch(() => {});
    await wait(800);
    const formNote = await noteOf();
    const formKept = await mail('attachmentReading', { accountId, folder: 'Inbox', id: (await mail('messages', { accountId, folder: 'Inbox' })).rows.find((r) => r.subject === 'Your application form')?.id, index: 0 });
    check('mail-invoice: a form attached is read when opened and shows no note, and that is kept too',
      !formNote && formKept?.version === 1 && formKept.kind === 'form' && formKept.note === null, JSON.stringify({ formNote, kind: formKept?.kind, note: formKept?.note }));

    // The invoice again: its note at once, from what was kept; nothing read again.
    const readsBeforeAgain = await reads();
    await openSubject('Invoice INV-2675');
    const again = await settle(async () => { const n = await noteOf(); return n?.state === 'done' && n.head.includes('INV-2675') ? n : null; }, 'the invoice note again', 10000);
    await wait(1000);
    const readsAfterAgain = await reads();
    const kept = await mail('attachmentReading', { accountId, folder: 'Inbox', id: (await mail('messages', { accountId, folder: 'Inbox' })).rows.find((r) => r.subject === 'Invoice INV-2675')?.id, index: 0 });
    check('mail-invoice: opened again, the invoice\'s note comes from what was kept with the message, and nothing is read again',
      Boolean(again) && fact(again, 'total') === '£1,601.12' && readsAfterAgain === readsBeforeAgain && kept?.note?.total?.value === truth.total, JSON.stringify({ again: Boolean(again), readsBeforeAgain, readsAfterAgain, keptTotal: kept?.note?.total?.value }));
    const complaints = await errorsIn(win);
    check('mail-invoice: the window reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');

    // In Urdu: the note right to left, its words in Urdu, and the reading in Urdu and English says what is not read yet.
    await js(`localStorage.setItem('rutba-ocr-languages', '["urd","eng"]'), 1`);
    await js(`window.rutbaOffice.store.set({ key: 'language', value: 'ur' }).then(() => 1)`);
    const w = await open('mail');
    const wjs = (code) => w.webContents.executeJavaScript(code);
    await until(() => wjs(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the Urdu window', 10000).catch(() => {});
    await wjs(`(() => { const b = [...document.querySelectorAll('.ml-accounts button')].find((n) => n.textContent.trim().startsWith('Invoice Owner')); b?.click(); return Boolean(b); })()`);
    await until(() => wjs(`document.querySelectorAll('.ml-row').length >= 4`), 'the messages in Urdu', 10000).catch(() => {});
    await wjs(`(() => { const r = [...document.querySelectorAll('.ml-row .ml-subject')].find((n) => n.textContent.trim() === 'Receipt, Karachi branch'); r?.click(); return Boolean(r); })()`);
    await until(() => wjs(`document.querySelector('[data-role="ml-docnote"]')?.dataset.state === 'done'`), 'the note in Urdu', 90000).catch(() => {});
    await wjs(`document.querySelector('[data-role="ml-docnote"] [data-field="total"]')?.click(), 1`);
    const urdu = await settle(() => wjs(`(() => {
      const n = document.querySelector('[data-role="ml-docnote"]');
      const p = n?.querySelector('[data-role="ml-docnote-peek"]');
      if (!n || !p || !p.querySelector('[data-role="ocr-boxmark"]')) return null;
      const icon = n.querySelector('.ml-docnote-icon').getBoundingClientRect(), main = n.querySelector('.ml-docnote-main').getBoundingClientRect();
      return { dir: document.documentElement.dir, title: n.querySelector('.ml-docnote-title').textContent, total: n.querySelector('[data-field="total"] .v')?.textContent || '', iconRight: icon.left > main.left,
        digits: p.querySelector('[data-role="ocr-digits-note"]')?.textContent || '', nastaliq: p.querySelector('[data-role="ocr-urdu-note"]')?.textContent || '' };
    })()`), 'the Urdu note and its limits', 15000);
    await capture(w, 'mail-invoice-03-urdu.png');
    check('mail-invoice: in Urdu the note reads right to left in Urdu, and a reading in Urdu says Arabic-Indic and Urdu digits are not read yet and Nastaliq is beta',
      urdu?.dir === 'rtl' && urdu.iconRight && /[؀-ۿ]/.test(urdu.title) && /190\.13/.test(urdu.total) && /[؀-ۿ]/.test(urdu.digits) && /[؀-ۿ]/.test(urdu.nastaliq), JSON.stringify(urdu));
    w.destroy();
  } catch (err) {
    check('mail-invoice: the checks ran', false, err.stack || err.message);
  } finally {
    if (js) {
      await js(`localStorage.removeItem('rutba-ocr-languages'), 1`).catch(() => {});
      await js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(language ?? null)} }).then(() => 1)`).catch(() => {});
      if (accountId) await js(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)}, keepMail: false })`).catch(() => {});
    }
  }
}
