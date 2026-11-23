// The print dialog is the only one (Office 2.27.0): Print sends the job
// straight to the printer chosen in it, on the paper the pages were laid out
// for, and the system's dialog opens only from its own button.
//
// A check run never reaches a printer: the print service lays the job out and
// keeps what the printer would have been told in globalThis.__rutbaPrints,
// answering as sent, or as asked in globalThis.__rutbaPrintAnswers. So this
// reads, in the real windows, what a real printer would be handed.
//
// A document: the dialog names it, draws page one small, counts the sheets
// and opens on the system's default printer; Print sends every sheet there,
// silently, upright on A4. Pages typed are the only ones sent, and what is
// not pages of the print stops Print. A print the printer does not take
// leaves the dialog open and says so; "System dialog" sends the same job
// with the system's dialog asked for. A workbook too wide for upright paper
// goes sideways on Automatic, on the paper chosen, and upright when a person
// says so. On a computer with no printer the dialog says so and only Save as
// PDF can be pressed.
// Run alone with RUTBA_VERIFY_ONLY=print-dialog.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the files are written
 */
export async function verifyPrintDialog({ open, check, until, wait, press, errorsIn }, { dir }) {
  const ok = (p) => p.then(() => true).catch(() => false);
  const root = path.join(dir, 'print-dialog');
  const sentBefore = (globalThis.__rutbaPrints ||= []).length;
  const sent = () => globalThis.__rutbaPrints.slice(sentBefore);
  try {
    fs.mkdirSync(root, { recursive: true });
    const docName = 'Lease for the shop.docx';
    const docFile = path.join(root, docName);
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: Array.from({ length: 150 }, (_, i) => ({ text: `Clause ${i + 1}. The tenant keeps the shop in good order and the owner keeps the roof.` })) }));
    const rows = [];
    for (let r = 0; r < 150; r++) {
      const row = [];
      for (let c = 0; c < 13; c++) row.push(r === 0 ? `Month ${c + 1}` : r * 100 + c);
      rows.push(row);
    }
    const bookFile = path.join(root, 'Sales by month.xlsx');
    fs.writeFileSync(bookFile, buildXlsx({ sheets: [{ name: 'Sales', rows }] }));

    const word = await open('word', docFile);
    const tools = (win) => {
      const js = (code) => win.webContents.executeJavaScript(code);
      const q = (sel) => `document.querySelector(${JSON.stringify(sel)})`;
      const PRINT = '.rw-dialog-foot .rw-btn.primary';
      const count = async () => {
        const said = await js(`${q('.rw-print-what .chip')}?.textContent || ''`);
        return { said, numbers: (said.match(/\d+/g) || []).map(Number) };
      };
      return {
        js, q, PRINT, count,
        has: (sel) => js(`Boolean(${q(sel)})`),
        click: (sel) => js(`(${q(sel)}?.click(), 1)`),
        type: (sel, value) => js(`(() => { const n = ${q(sel)}; if (!n) return 0; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(String(value))}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`),
        /** A list set to one of its values, as a person choosing it does. `find` is the list, as an expression. */
        choose: (find, value) => js(`(() => { const n = ${find}; if (!n) return 0; Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`),
        /** Ctrl+P, and the dialog with its sheets counted. */
        dialog: async () => {
          win.focus();
          win.webContents.focus();
          await press(win.webContents, 'p', { modifiers: ['control'] });
          const opened = await ok(until(() => js(`Boolean(${q('.rw-print-what')})`), 'the print dialog', 8000));
          await until(async () => (await count()).numbers.length > 0, 'the sheets counted', 25000).catch(() => {});
          return opened;
        },
        closed: () => ok(until(() => js(`!${q('.rw-print-what')}`), 'the dialog to close', 8000)),
        disabled: (sel) => js(`Boolean(${q(sel)}?.disabled)`),
      };
    };
    const w = tools(word);
    const opened = await w.dialog();
    const listed = (await w.js(`window.rutbaOffice.print.printers()`)) || [];

    if (!listed.length) {
      // No printer on this computer: the dialog says so, and only a PDF can be had.
      const none = await w.js(`(() => {
        const s = ${w.q('.rw-print-printer')};
        return { greyed: Boolean(s?.disabled), says: (s?.closest('.rw-field')?.querySelector('.rw-hint')?.textContent || '').trim().length, print: Boolean(${w.q(w.PRINT)}?.disabled), system: Boolean(${w.q('.rw-print-system')}?.disabled),
          pdf: [...document.querySelectorAll('.rw-dialog-foot .rw-btn')].filter((b) => !b.disabled).length };
      })()`);
      check('print dialog: with no printer installed the dialog says so under a greyed list, Print and System dialog cannot be pressed, and Save as PDF and Cancel can',
        opened && none.greyed && none.says > 20 && none.print && none.system && none.pdf === 2, JSON.stringify(none));
      await press(word.webContents, 'Escape');
      return;
    }

    /* ── a document: named, pictured, counted, on the default printer ── */
    const drawn = await ok(until(() => w.js(`${w.q('.rw-print-first canvas')}?.dataset.drawn === 'yes'`), 'page one drawn', 20000));
    const ink = await w.js(`(() => { const c = ${w.q('.rw-print-first canvas')}; if (!c || !c.width) return 0; const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; let n = 0; for (let i = 0; i < d.length; i += 4) if (d[i] + d[i + 1] + d[i + 2] < 600) n += 1; return n; })()`);
    const named = await w.js(`${w.q('.rw-print-name')}?.textContent || ''`);
    const whole = (await w.count()).numbers[0] || 0;
    const expected = listed.find((p) => p.default)?.name || listed[0].name;
    const list = await w.js(`(() => { const s = ${w.q('.rw-print-printer')}; return { value: s?.value || '', first: s?.options[0]?.value || '', options: s?.options.length || 0, greyed: Boolean(s?.disabled) }; })()`);
    check('print dialog: a document is named at the top with page one drawn small and its sheets counted, and the list of printers opens on the system default, put first',
      opened && named.includes('Lease for the shop') && drawn && ink > 40 && whole >= 2 && list.options === listed.length && !list.greyed && list.value === expected && list.first === listed[0].name,
      JSON.stringify({ opened, named, drawn, ink, whole, list, expected, printers: listed.length }));

    // A printer that is not ready says so under the list once the system has answered; a ready one says nothing.
    const marked = (await w.js(`window.rutbaOffice.print.printers({ states: true })`)) || listed;
    const state = marked.find((p) => p.name === list.value)?.state || '';
    const hinted = () => w.js(`Boolean((${w.q('.rw-print-printer')}?.closest('.rw-field')?.querySelector('.rw-hint')?.textContent || '').trim())`);
    const agrees = await ok(until(async () => (await hinted()) === Boolean(state), 'the list to say what the system says of the printer', 30000));
    check('print dialog: the printer it opens on says under the list that it is not ready when the system says so, and says nothing when it is ready',
      agrees, JSON.stringify({ printer: list.value, state: state || 'ready', hinted: await hinted() }));

    /* ── Print: straight to the printer chosen, every sheet, upright on A4 ── */
    await w.click(w.PRINT);
    const went = await ok(until(() => sent().length === 1 || null, 'the job handed on', 30000));
    const shut = await w.closed();
    const first = sent()[0] || {};
    const saidSent = await w.js(`[...document.querySelectorAll('.rw-toast.good')].map((n) => n.textContent).pop() || ''`);
    check('print dialog: Print sends the document straight to the printer chosen, with no system dialog asked for, every sheet, one copy, upright on A4; the dialog closes and says how many sheets went',
      went && first.silent === true && first.deviceName === expected && first.pageSize === 'A4' && first.landscape === false && first.copies === 1 && first.sheets === whole && shut && saidSent.includes(String(whole)),
      JSON.stringify({ went, job: first, shut, saidSent }));

    /* ── pages typed, copies, a print the printer does not take, and the system's dialog ── */
    await w.dialog();
    await w.type('.rw-print-pages', '999');
    await wait(300);
    const wrong = await w.js(`({ marked: ${w.q('.rw-print-pages')}?.getAttribute('aria-invalid') === 'true', print: Boolean(${w.q(w.PRINT)}?.disabled) })`);
    await w.type('.rw-print-pages', '2');
    await w.type('.rw-dialog input[type="number"]', '3');
    await until(async () => { const n = (await w.count()).numbers; return n.length === 2 && n[0] === 1 && n[1] === whole; }, 'one of the sheets chosen', 6000).catch(() => {});
    const chosen = await w.count();
    globalThis.__rutbaPrintAnswers = [{ ok: false, said: 'Print job failed' }];
    await w.click(w.PRINT);
    const refusedSent = await ok(until(() => sent().length === 2 || null, 'the second job handed on', 30000));
    const saysWhy = await ok(until(() => w.js(`document.querySelectorAll('.rw-toast.bad').length > 0`), 'the dialog says the printer did not take it', 8000));
    const refusal = await w.js(`[...document.querySelectorAll('.rw-toast.bad')].map((n) => n.textContent).pop() || ''`);
    await wait(400);
    const stayed = await w.has('.rw-print-what');
    const second = sent()[1] || {};
    check('print dialog: what is typed in Pages that is not pages of the print is marked and stops Print; page 2 typed with three copies sends that one sheet three times; a print the printer does not take leaves the dialog open and says so in its own words',
      wrong.marked && wrong.print && chosen.numbers[0] === 1 && chosen.numbers[1] === whole && refusedSent && second.sheets === 1 && second.copies === 3 && second.silent === true && saysWhy && refusal.length > 20 && !refusal.includes('Print job failed') && stayed,
      JSON.stringify({ wrong, chosen: chosen.said, job: second, saysWhy, refusal, stayed }));

    await w.click('.rw-print-system');
    const thirdSent = await ok(until(() => sent().length === 3 || null, 'the third job handed on', 30000));
    const shutAgain = await w.closed();
    const third = sent()[2] || {};
    check('print dialog: System dialog sends the same job to the same printer with the system\'s own dialog asked for, and that is the only way it is asked for',
      thirdSent && third.silent === false && third.deviceName === expected && third.sheets === 1 && third.copies === 3 && shutAgain && sent().filter((j) => j.silent === false).length === 1,
      JSON.stringify({ job: third, shutAgain }));
    const complaints = (await errorsIn(word)).filter((s) => !refusal || !(s.includes(refusal) || refusal.includes(s)));

    /* ── a workbook too wide for upright paper ── */
    const book = await open('sheets', bookFile);
    const s = tools(book);
    const bookOpened = await s.dialog();
    const PAPER = `[...document.querySelectorAll('.rw-dialog select')].find((n) => [...n.options].some((o) => o.value === 'Letter') && [...n.options].some((o) => o.value === 'A4'))`;
    const way = await s.js(`${s.q('.rw-print-way')}?.value || ''`);
    await s.choose(PAPER, 'Letter');
    await wait(900);
    await until(async () => (await s.count()).numbers.length > 0, 'the sheets counted on Letter', 15000).catch(() => {});
    await s.click(s.PRINT);
    const wideSent = await ok(until(() => sent().length === 4 || null, 'the workbook handed on', 30000));
    await s.closed();
    const wide = sent()[3] || {};
    // Upright because a person says so: the choice stands over the sheet's shape.
    await s.dialog();
    await s.choose(s.q('.rw-print-way'), 'portrait');
    await wait(900);
    await until(async () => (await s.count()).numbers.length > 0, 'the sheets counted upright', 15000).catch(() => {});
    await s.click(s.PRINT);
    const tallSent = await ok(until(() => sent().length === 5 || null, 'the workbook handed on upright', 30000));
    await s.closed();
    const tall = sent()[4] || {};
    check('print dialog: a workbook too wide for upright paper is on Automatic and goes to the printer sideways, on the paper chosen in the dialog; upright chosen by a person goes upright',
      bookOpened && way === 'auto' && wideSent && wide.silent === true && wide.landscape === true && wide.pageSize === 'Letter' && wide.deviceName === expected && tallSent && tall.landscape === false,
      JSON.stringify({ way, wide, tall }));

    complaints.push(...(await errorsIn(book)));
    check('print dialog: none of that reported an error beyond the print the check had refused', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('print dialog: the check ran to the end', false, err?.message || String(err));
  } finally {
    delete globalThis.__rutbaPrintAnswers;
  }
}
