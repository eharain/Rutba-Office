// Worksheets: Review → Redact cells, and File → Info → Inspect Workbook.
// A name and the telephone pattern are marked and outlined on the grid;
// Apply saves a copy that, read back, holds none of them in any part; the
// inspector lists the note, the header, the hidden sheet, the hidden row
// and column, the hidden name and the authors, and Remove All takes each
// out of the open workbook. Run alone with RUTBA_VERIFY_ONLY=sheet-inspect.

import fs from 'node:fs';
import path from 'node:path';
import { partsHolding } from '@rutba/ooxml/redact';
import { OoxmlPackage } from '@rutba/ooxml';
import { secretWorkbook } from './verify-review-fixtures.js';

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifySheetInspect(h, { dir }) {
  const { open, check, wait, capture, doc } = h;
  // Waits as the harness's until does, and answers what the condition found once it held.
  const until = async (fn, what, ms) => { let found; await h.until(async () => { found = await fn(); return found; }, what, ms); return found; };
  const own = path.join(dir, 'sheet-inspect');
  const file = path.join(own, 'payroll.xlsx');
  const copy = path.join(own, 'payroll (redacted).xlsx');
  try {
    fs.mkdirSync(own, { recursive: true });
    fs.writeFileSync(file, await secretWorkbook());
  } catch (err) {
    return check('sheet-inspect: the fixture was made', false, err.message);
  }
  try {
    const win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'sheet' && s.path === file).pop();
    const fresh = async (name) => { wc.invalidate(); await wait(700); await capture(win, name); };
    const setInput = (sel, value) => js(`(() => { const el = document.querySelector(${JSON.stringify(sel)}); if (!el) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const click = (sel) => js(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'none'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);

    await until(() => js(`[...document.querySelectorAll('.sh-cells .sh-cell')].some((n) => n.textContent === 'Ayesha Khan')`), 'the grid', 10000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Review')?.click(), 1`);
    await wait(250);
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Redact')); if (!b) return 'no button'; b.click(); return 'clicked'; })()`);
    await until(() => js(`Boolean(document.querySelector('.rd-pane .rd-find'))`), 'the Redact pane', 5000).catch(() => {});
    await setInput('.rd-find', 'Ayesha Khan');
    await click('.rd-find-mark');
    await js(`document.querySelector('.rd-pattern-phone')?.click(), 1`);
    const lit = await until(() => js(`(() => { const t = [...document.querySelectorAll('.rd-hit')].map((n) => n.dataset.text); return t.includes('Ayesha Khan') && t.includes('+92 300 1234567') ? t : null; })()`), 'the marks on the grid', 6000).catch(() => null);
    const where = await until(() => js(`(() => { const k = [...document.querySelectorAll('.rd-where-parts span')].map((s) => s.dataset.kind); return ['cells', 'comments', 'properties'].every((x) => k.includes(x)) ? k : null; })()`), 'the parts counted', 6000).catch(() => null);
    await fresh('sheet-inspect-redact.png');
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: copy });
    await click('.rd-apply-btn');
    const done = await until(() => js(`Boolean(document.querySelector('.rd-done'))`), 'the copy saved', 20000).catch(() => false);
    const left = fs.existsSync(copy) ? partsHolding(fs.readFileSync(copy), ['Ayesha Khan', '300 1234567', '321 7654321']) : ['no copy'];
    check(
      'sheet-inspect: Redact marks a name and telephone numbers on the grid as they are made, counts them in cells, notes and properties, and Apply saves a copy that holds none of them in any part',
      pressed === 'clicked' && Array.isArray(lit) && Array.isArray(where) && done === true && left.length === 0,
      JSON.stringify({ pressed, lit, where, done, left }),
    );

    // Inspect Workbook.
    await js(`document.querySelector('.rw-appmark')?.click(), 1`);
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Info')))`), 'the app menu', 4000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Info'))?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.pw-card[data-card="inspect"]'))`), 'the Inspect card', 5000).catch(() => {});
    await js(`document.querySelector('.pw-card[data-card="inspect"] .pw-card-tile')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.pw-menu [data-item="inspect"]'))`), 'Check for Issues', 3000).catch(() => {});
    await js(`document.querySelector('.pw-menu [data-item="inspect"]')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.di-inspect'))`), 'the inspector', 6000).catch(() => {});
    const kinds = await js(`[...document.querySelectorAll('.di-dialog .di-row h5')].map((n) => n.textContent)`);
    await js(`document.querySelector('.di-inspect')?.click(), 1`);
    const found = await until(() => js(`(() => { const rows = [...document.querySelectorAll('.di-dialog .di-row[data-found]')]; return rows.length ? Object.fromEntries(rows.map((r) => [r.dataset.kind, Number(r.dataset.found)])) : null; })()`), 'the results', 8000).catch(() => null);
    await fresh('sheet-inspect-results.png');
    check(
      'sheet-inspect: Inspect Workbook lists Excel\'s kinds and finds the note, the header, the hidden worksheet, the hidden row and column, the hidden name and the author',
      ['Comments', 'Hidden Worksheets', 'Hidden Rows and Columns', 'Hidden Names', 'Headers, Footers, and Watermarks'].every((k) => kinds.includes(k))
        && Boolean(found) && found.comments === 1 && found.headers === 1 && found.hiddenSheets === 1 && found.hiddenRowsColumns === 2 && found.hiddenNames === 1 && found.personal >= 1,
      JSON.stringify({ kinds, found }),
    );
    for (const k of ['comments', 'headers', 'hiddenSheets', 'hiddenRowsColumns', 'hiddenNames', 'personal']) {
      await js(`document.querySelector('.di-remove-${k}')?.click(), 1`);
      await until(() => js(`document.querySelector('.di-row[data-kind="${k}"]')?.dataset.found === '0'`), `${k} removed`, 10000).catch(() => {});
    }
    const after = await js(`Object.fromEntries([...document.querySelectorAll('.di-dialog .di-row[data-found]')].map((r) => [r.dataset.kind, Number(r.dataset.found)]))`);
    await js(`document.querySelector('.di-close')?.click(), 1`);
    const tabs = await until(() => js(`(() => { const t = [...document.querySelectorAll('.sh-tab:not(.sh-tab-add)')].map((n) => n.textContent.trim()); return t.length === 1 ? t : null; })()`), 'the tabs', 6000).catch(() => null);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Save')); b?.click(); return 1; })()`);
    await until(() => doc.meta({ id: session.id }).dirty === false, 'the save', 15000).catch(() => {});
    const saved = OoxmlPackage.read(fs.readFileSync(file));
    const sheet = saved.text(saved.partNames().find((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)));
    check(
      'sheet-inspect: Remove All takes each out, the note, the header, the hidden sheet, the hidden row and column (deleted, as Excel deletes them), the hidden name, the author, and the saved workbook holds none of them',
      ['comments', 'headers', 'hiddenSheets', 'hiddenRowsColumns', 'hiddenNames', 'personal'].every((k) => after[k] === 0) && Array.isArray(tabs) && tabs[0] === 'Staff'
        && !/hidden="1"/.test(sheet) && !/<headerFooter/.test(sheet) && !/Secret/.test(saved.text('xl/workbook.xml')) && !saved.partNames().some((n) => /comments/i.test(n)),
      JSON.stringify({ after, tabs }),
    );
  } catch (err) {
    check('sheet-inspect: the Inspect Workbook checks ran', false, err.stack || err.message);
  }
}
