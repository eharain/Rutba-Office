// Worksheets: what an error and a function are, said where they are met.
// A cell showing #DIV/0! wears Excel's green corner, and its tip says why
// it shows the error; Insert Function lists each function with what it
// takes and what it does, and its search finds a function by what it does.
// The template gallery's search covers every template whatever the rail
// has chosen, and says so. Run alone with RUTBA_VERIFY_ONLY=sheethelp.

import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifySheetHelp(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const file = path.join(dir, 'error-help.xlsx');
  try {
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Share', 'Total'], ['=10/B3', 40], ['', 0]] }] }));
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`/DIV/.test(document.querySelector('.sh-cell[data-ref="A2"]')?.textContent || '')`), 'the error in A2', 8000).catch(() => {});
    const cell = await js(`(() => {
      const c = document.querySelector('.sh-cell[data-ref="A2"]');
      if (!c) return null;
      const corner = getComputedStyle(c, '::before');
      return { text: c.textContent.trim(), tip: c.dataset.tip || '', err: c.classList.contains('err'), corner: corner.borderTopColor, plain: document.querySelector('.sh-cell[data-ref="B2"]')?.dataset.tip || null };
    })()`);
    check('sheets: a cell showing #DIV/0! wears the green corner, and its tip says it divides by zero',
      cell?.err && cell.tip === '#DIV/0!: Divides by zero' && /30, 142, 62/.test(cell.corner) && !cell.plain,
      JSON.stringify(cell));

    // Formulas → Insert Function: each with what it takes and does; found by what it does.
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Formulas')?.click(), 1`);
    await until(() => js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].some((n) => n.textContent.trim() === 'Insert Function')`), 'the Formulas tab', 4000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Insert Function')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.sh-fn-item[data-function="SUM"]'))`), 'the Insert Function dialog', 4000).catch(() => {});
    const sum = await js(`(() => { const i = document.querySelector('.sh-fn-item[data-function="SUM"]'); return i ? { who: i.querySelector('.who')?.textContent.trim(), what: i.querySelector('.what')?.textContent.trim() } : null; })()`);
    await js(`(() => { const i = document.querySelector('.rw-dialog input'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, 'mean'); i.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-fn-item[data-function="AVERAGE"]'))`), 'AVERAGE found by what it does', 3000).catch(() => {});
    const found = await js(`[...document.querySelectorAll('.sh-fn-item')].map((n) => n.dataset.function)`);
    if (process.env.RUTBA_VERIFY_CAPTURE) { win.webContents.invalidate(); await wait(700); fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'sheet-insert-function.png'), (await win.webContents.capturePage()).toPNG()); }
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((n) => n.textContent.trim() === 'Cancel')?.click(), 1`);
    await until(() => js(`!document.querySelector('.sh-fn-item')`), 'the dialog closed', 3000).catch(() => {});
    check('sheets: Insert Function lists each function with what it takes and does, and finds AVERAGE when "mean" is typed',
      sum?.who === 'SUM(number1, …)' && sum.what === 'Add numbers and ranges' && found.includes('AVERAGE'),
      JSON.stringify({ sum, found: found.slice(0, 8) }));

    // The template gallery, with Presentations chosen on its rail: a search for budget finds the workbook.
    const home = await open('home');
    const hjs = (code) => home.webContents.executeJavaScript(code);
    await until(() => hjs(`Boolean(document.querySelector('.home-more'))`), 'the More templates tile', 6000);
    await hjs(`document.querySelector('.home-more').click(), 1`);
    await until(() => hjs(`Boolean(document.querySelector('.tg-rail-item[data-view="slides"]')) && document.querySelectorAll('.tg-card[data-template]').length > 3`), 'the gallery', 8000);
    await hjs(`document.querySelector('.tg-rail-item[data-view="slides"]').click(), 1`);
    await until(() => hjs(`[...document.querySelectorAll('.tg-card[data-template]')].every((n) => n.dataset.template.startsWith('deck-'))`), 'the decks', 4000).catch(() => {});
    const setSearch = (value) => hjs(`(() => { const i = document.querySelector('.tg-head .rw-search input'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(value)}); i.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await setSearch('budget');
    await until(() => hjs(`Boolean(document.querySelector('.tg-card[data-template="sheet-budget"]'))`), 'the budget found', 4000).catch(() => {});
    const searched = await hjs(`({
      cards: [...document.querySelectorAll('.tg-card[data-template]')].map((n) => n.dataset.template),
      said: document.querySelector('[data-role="searched-all"]')?.textContent.trim() || '',
      on: [...document.querySelectorAll('.tg-rail-item.on')].map((n) => n.dataset.view),
    })`);
    // A rail item chosen puts the search away and shows what it holds.
    await hjs(`document.querySelector('.tg-rail-item[data-view="slides"]').click(), 1`);
    await until(() => hjs(`document.querySelector('.tg-head .rw-search input')?.value === ''`), 'the search cleared', 3000).catch(() => {});
    const after = await hjs(`({ query: document.querySelector('.tg-head .rw-search input')?.value, decks: [...document.querySelectorAll('.tg-card[data-template]')].every((n) => n.dataset.template.startsWith('deck-')), said: Boolean(document.querySelector('[data-role="searched-all"]')) })`);
    await hjs(`document.querySelector('.tg-close')?.click(), 1`);
    check('templates: a search with Presentations chosen still finds the budget workbook, says it looked among every template, and choosing a rail item puts the search away',
      searched.cards.includes('sheet-budget') && /every template/.test(searched.said) && searched.on.length === 0 && after.query === '' && after.decks && !after.said,
      JSON.stringify({ searched, after }));
    await wait(200);

    const complaints = [...await errorsIn(win), ...await errorsIn(home)];
    check('sheets: the error and function help checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the error and function help checks ran', false, err?.message || String(err));
  }
}
