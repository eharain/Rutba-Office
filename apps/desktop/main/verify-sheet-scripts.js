// Worksheets: Automate → scripts — New Script opens the Code Editor with a
// sample; a script typed there runs in its worker on the workbook and its
// edits land as one undo step; Save keeps it under All Scripts; a script
// that tries to fetch from the network is refused by the worker's policy;
// Record Actions turns a typed value and Bold into a script. Run alone with
// RUTBA_VERIFY_ONLY=scripts.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

export async function verifySheetScripts({ open, check, until, wait, press, errorsIn, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'scripts.xlsx');
  try {
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sales', rows: [['Item', 'Qty', 'Price'], ['Ink', 4, 2.5], ['Paper', 6, 1.25]] }] }));
    const win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('sheet').id });
    const text = (ref) => String((model().cells || []).find((c) => c.ref === ref)?.text ?? '');
    const setCode = (code) => js(`(() => { const el = document.querySelector('.sh-script-code'); if (!el) return false; Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, ${JSON.stringify(code)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const output = () => js(`document.querySelector('.sh-script-output')?.textContent || ''`);
    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 8000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Automate')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.sh-new-script'))`), 'the Automate tab', 4000);

    // New Script: the Code Editor with a sample.
    await js(`document.querySelector('.sh-new-script').click(), 1`);
    await until(() => js(`/function main\\(workbook\\)/.test(document.querySelector('.sh-script-code')?.value || '')`), 'the Code Editor', 4000).catch(() => {});
    const sample = await js(`document.querySelector('.sh-script-code')?.value || ''`);
    check('sheets: Automate → New Script opens the Code Editor on a sample script, function main(workbook)', /function main\(workbook\)/.test(sample), sample.slice(0, 60));

    // A script of our own: totals in D, bold headings, a summary sheet.
    await setCode(`function main(workbook) {
  const sheet = workbook.getActiveWorksheet();
  const rows = sheet.getUsedRange().getValues();
  sheet.getRange("D1").setValue("Total");
  for (let r = 2; r <= rows.length; r++) sheet.getRange("D" + r).setFormula("=B" + r + "*C" + r);
  sheet.getRange("A1:D1").getFormat().getFont().setBold(true);
  console.log("lines", rows.length - 1);
}`);
    await wait(150);
    await js(`document.querySelector('.sh-script-run').click(), 1`);
    await until(async () => /Done:/.test(await output()), 'the run', 10000).catch(() => {});
    const ran = await output();
    await until(() => text('D3') === '7.5', 'the totals', 4000).catch(() => {});
    check('sheets: Run works the script out in its worker on the workbook: totals as formulas, a log line, the cells counted',
      /Done: 3 cells set/.test(ran) && /lines 2/.test(ran) && text('D1') === 'Total' && text('D2') === '10' && text('D3') === '7.5', `${ran} | ${['D1', 'D2', 'D3'].map(text).join(',')}`);
    await press(wc, 'Z', { modifiers: ['control'] });
    await until(() => text('D1') === '', 'the undo', 4000).catch(() => {});
    check('sheets: one Ctrl+Z takes the whole script back', text('D1') === '' && text('D3') === '', ['D1', 'D3'].map(text).join(','));

    // Save: kept under All Scripts.
    await js(`(() => { const n = document.querySelector('.sh-script-name'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, 'Totals'); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(100);
    await js(`document.querySelector('.sh-script-save').click(), 1`);
    await until(() => js(`[...document.querySelectorAll('.sh-script-title')].some((t) => t.textContent === 'Totals')`), 'the saved script', 4000).catch(() => {});
    const listed = await js(`[...document.querySelectorAll('.sh-script-title')].map((t) => t.textContent)`);
    check('sheets: Save keeps the script on this computer, listed under All Scripts', listed.includes('Totals'), JSON.stringify(listed));

    // The worker reaches nothing: even the app's own files, which the window may read, are refused to it.
    await setCode(`function main(workbook) { const x = new XMLHttpRequest(); x.open("GET", "scripts-worker.js", false); x.send(); console.log("reached", x.status); }`);
    await wait(100);
    await js(`document.querySelector('.sh-script-run').click(), 1`);
    await until(async () => /Done:|refused|Content Security|Failed|Error/i.test(await output()), 'the refusal', 8000).catch(() => {});
    const refused = await output();
    check('sheets: a script cannot reach the network from its worker: the request is refused by the worker\'s policy', /Failed to load|Content Security Policy|Refused to connect/i.test(refused) && !/reached/.test(refused), refused.slice(0, 200));

    // Record Actions: a typed value and Bold become a script.
    await js(`document.querySelector('.sh-record-actions').click(), 1`);
    await doc.apply({ id: sessionFor('sheet').id, ops: [{ op: 'select', row: 5, col: 0 }] });
    await js(`(async () => {
      await window.rutbaOffice.doc.apply({ id: ${JSON.stringify(sessionFor('sheet').id)}, ops: [{ op: 'select', row: 5, col: 0 }] });
      return 1;
    })()`);
    await wait(200);
    // Typed through the grid as a person types.
    await js(`document.querySelector('.sh-grid, [data-grid]')?.focus?.(), 1`);
    await press(wc, 'F2');
    await wait(150);
    for (const ch of 'Done') { wc.sendInputEvent({ type: 'char', keyCode: ch }); await wait(30); }
    await press(wc, 'Enter');
    await wait(300);
    await js(`document.querySelector('.sh-record-actions').click(), 1`);
    await until(() => js(`/Recorded script/.test(document.querySelector('.sh-script-name')?.value || '')`), 'the recorded script', 4000).catch(() => {});
    const recorded = await js(`document.querySelector('.sh-script-code')?.value || ''`);
    check('sheets: Record Actions turns what was typed into a script', /setValue\("Done"\)/.test(recorded), recorded.slice(0, 200));

    const complaints = await errorsIn(win);
    check('sheets: scripts report nothing', complaints.filter((c) => !/Content Security|Refused to connect/i.test(c)).length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the script checks ran', false, err?.message || JSON.stringify(err));
  }
}
