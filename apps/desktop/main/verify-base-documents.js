// Data Files' print templates, checked in the window: New Print Template
// makes an invoice for each record of Invoices with the business's details
// as its letterhead, its lines and their total; one record chosen prints
// alone and Save as PDF writes it; Design turns it into a till receipt and
// adds words, the pages following, kept in the file; and the Invoices form
// prints the invoice it shows.

import fs from 'node:fs';
import path from 'node:path';
import { Base, buildFromSpec } from '@rutba/database';

const HELPERS = `
  window.__d = window.__d || {
    set(el, value) {
      if (typeof el === 'string') el = document.querySelector(el);
      if (!el) return false;
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, value);
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return true;
    },
    click(sel) { const el = document.querySelector(sel); if (!el) return false; el.click(); return true; },
    pages() { return [...document.querySelectorAll('.bs-report-scroll .bs-doc-page')]; },
  };
  true;
`;

export async function verifyBaseDocuments(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'Repairs.rbase');
  const pdf = path.join(dir, 'Invoice 2.pdf');
  const formPdf = path.join(dir, 'Invoice from the form.pdf');
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    for (const f of [pdf, formPdf]) fs.rmSync(f, { force: true });
    const ref = (t) => ({ type: 'integer', points: t });
    const b = Base.open(file);
    buildFromSpec(b, { tables: [
      { name: 'Business Details', fields: [{ name: 'Name', type: 'text' }, { name: 'Address', type: 'text' }] },
      { name: 'Customers', fields: [{ name: 'Name', type: 'text' }] },
      { name: 'Invoices', fields: [{ name: 'Customer', ...ref('Customers') }, { name: 'Invoice Date', type: 'date' }] },
      { name: 'Invoice Lines', fields: [{ name: 'Invoice', ...ref('Invoices') }, { name: 'Description', type: 'text' }, { name: 'Quantity', type: 'integer' }, { name: 'Unit Price', type: 'currency' }] },
    ] });
    b.insert('Business Details', { Name: 'Larkspur Repairs', Address: '4 Mill Lane, Leeds' });
    b.insert('Customers', { Name: 'Amira Haddad' });
    b.insert('Customers', { Name: 'Ben Cole' });
    b.insert('Invoices', { Customer: 1, 'Invoice Date': '2026-10-21' });
    b.insert('Invoices', { Customer: 2, 'Invoice Date': '2026-10-22' });
    b.insert('Invoice Lines', { Invoice: 1, Description: 'Screen repair', Quantity: 1, 'Unit Price': 89 });
    b.insert('Invoice Lines', { Invoice: 1, Description: 'Phone case', Quantity: 2, 'Unit Price': 7.5 });
    b.insert('Invoice Lines', { Invoice: 2, Description: 'Battery', Quantity: 1, 'Unit Price': 45 });
    b.close();

    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await js(HELPERS);
    await until(() => js(`Boolean(document.querySelector('.bs-new-document'))`), 'the database', 8000);

    await js(`__d.click('.bs-new-document')`);
    await until(() => js(`Boolean(document.querySelector('.bs-document-table'))`), 'the New Print Template box', 4000).catch(() => {});
    await js(`__d.set('.bs-document-table', 'Invoices')`);
    await wait(100);
    await js(`__d.click('.bs-make-document')`);
    await until(() => js(`__d.pages().length === 2`), 'two invoices', 6000).catch(() => {});
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await wait(400);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-invoice.png'), (await win.webContents.capturePage()).toPNG());
    }
    const first = await js(`(() => { const p = __d.pages()[0]; return { head: p?.querySelector('.bs-doc-letterhead')?.innerText || '', heading: p?.querySelector('.bs-doc-heading')?.textContent || '', lines: [...(p?.querySelectorAll('.bs-doc-lines tbody tr') || [])].map((r) => r.cells[0].textContent), total: p?.querySelector('.bs-doc-totals tr.strong td')?.textContent || '', fields: p?.querySelector('.bs-doc-fields')?.innerText || '' }; })()`);
    const listed = await js(`Boolean(document.querySelector('.bs-documents [data-document="Invoice"]'))`);
    check('base: New Print Template makes an invoice for each record of Invoices: the business\'s details as the letterhead, the customer, the lines and their total',
      listed && /Larkspur Repairs/.test(first.head) && first.heading === 'Invoice' && /Amira Haddad/.test(first.fields)
        && JSON.stringify(first.lines) === JSON.stringify(['Screen repair', 'Phone case']) && /104\.00/.test(first.total),
      JSON.stringify({ listed, first }));

    // One record, then Save as PDF.
    await js(`__d.set('.bs-doc-record', '2')`);
    await until(() => js(`__d.pages().length === 1 && __d.pages()[0].dataset.key === '2'`), 'invoice 2 alone', 5000).catch(() => {});
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: pdf });
    await js(`__d.click('.bs-doc-pdf')`);
    const wrote = await until(() => (fs.existsSync(pdf) && fs.statSync(pdf).size > 1000) || null, 'the PDF', 10000).then(() => true, () => false);
    const head = wrote ? fs.readFileSync(pdf).subarray(0, 5).toString('latin1') : '';
    check('base: a print template prints one record when it is chosen, and Save as PDF writes it', wrote && head === '%PDF-' && await js(`__d.pages().length === 1`), JSON.stringify({ wrote, head }));

    // Design: a till roll, and words added at the foot.
    await js(`__d.click('.bs-doc-design')`);
    await until(() => js(`Boolean(document.querySelector('.bs-dd-paper'))`), 'the design panel', 4000).catch(() => {});
    await js(`__d.set('.bs-dd-paper', 'receipt-80')`);
    await wait(150);
    await js(`__d.set('.bs-dd-add-kind', 'text')`);
    await wait(100);
    await js(`__d.click('.bs-dd-add-block')`);
    await until(() => js(`Boolean([...document.querySelectorAll('.bs-dd-text')].pop())`), 'the text block', 4000).catch(() => {});
    await js(`__d.set([...document.querySelectorAll('.bs-dd-text')].pop(), 'Thank you, {Customer}!')`);
    const drawn = await until(() => js(`(() => { const p = __d.pages()[0]; return p && p.classList.contains('paper-receipt-80') && /Thank you, Ben Cole!/.test(p.textContent); })()`), 'the receipt', 6000).then(() => true, () => false);
    const kept = await until(() => {
      const d = Base.open(file);
      try { const l = d.documents().find((x) => x.name === 'Invoice')?.layout; return (l?.paper === 'receipt-80' && l.blocks.some((x) => x.type === 'text' && /Thank you/.test(x.text))) || null; } finally { d.close(); }
    }, 'the change kept', 6000).then(() => true, () => false);
    check('base: Design turns the template into a till receipt and adds words with a blank filled from the record; the pages follow, and the change is kept in the file', drawn && kept, JSON.stringify({ drawn, kept }));

    // The Invoices form prints the invoice it shows.
    await js(`__d.click('.bs-forms [data-form="Invoices"]')`);
    await until(() => js(`Boolean(document.querySelector('.bs-form-print-now'))`), 'the form\'s Print', 6000).catch(() => {});
    const label = await js(`document.querySelector('.bs-form-print-now')?.textContent || ''`);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: formPdf });
    await js(`__d.click('.bs-form-pdf')`);
    const fromForm = await until(() => (fs.existsSync(formPdf) && fs.statSync(formPdf).size > 1000) || null, 'the form\'s PDF', 10000).then(() => true, () => false);
    check('base: a form whose table has a print template offers Print Invoice, and its PDF is the record shown', /Print Invoice/.test(label) && fromForm, JSON.stringify({ label, fromForm }));
  } catch (err) {
    check('base: the print template checks ran', false, err.stack || err.message);
  }
}
