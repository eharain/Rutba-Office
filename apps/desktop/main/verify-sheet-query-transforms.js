// Worksheets: the Power Query Editor's newer transforms — a month-by-column
// table unpivoted on its Product column into rows, the month names cut to
// their first three letters with Extract, pivoted back into columns under
// the short names with Pivot Column, and loaded. Run alone with
// RUTBA_VERIFY_ONLY=querytransforms.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

export async function verifySheetQueryTransforms({ open, check, until, wait, errorsIn, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'query-transforms.xlsx');
  try {
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Data', rows: [['Product', 'January', 'February', 'March'], ['Pens', 10, 12, null], ['Ink', 5, null, 7]] }] }));
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('sheet').id });
    const text = (ref) => String((model().cells || []).find((c) => c.ref === ref)?.text ?? '');
    const clickText = (selector, label) => js(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(selector)})].find((x) => x.textContent.trim() === ${JSON.stringify(label)}); if (!b) return false; b.click(); return true; })()`);
    const heads = () => js(`[...document.querySelectorAll('.pq-grid th[data-col]')].map((th) => th.textContent.trim())`);
    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 8000);
    await clickText('.rw-tab', 'Data');
    await until(() => js(`[...document.querySelectorAll('button')].some((b) => b.textContent.trim() === 'From Table/Range')`), 'the Data tab', 4000);
    await clickText('button', 'From Table/Range');
    await until(() => js(`document.querySelectorAll('.pq-grid th[data-col]').length === 4`), 'the Power Query Editor', 6000);

    await js(`document.querySelector('.pq-grid th[data-col="Product"]').click()`);
    await wait(100);
    await clickText('.pq-tool', 'Unpivot Other Columns');
    const unpivoted = await until(async () => JSON.stringify(await heads()) === '["Product","Attribute","Value"]', 'the columns unpivoted', 4000).then(() => true).catch(() => false);
    const rows = await js(`document.querySelectorAll('.pq-grid tbody tr').length`);
    check('sheets: Unpivot Other Columns turns the month columns into rows of Product, Attribute and Value, a blank month giving none', unpivoted && rows === 4, JSON.stringify({ heads: await heads(), rows }));

    await js(`document.querySelector('.pq-grid th[data-col="Attribute"]').click()`);
    await wait(100);
    await js(`(() => { const b = [...document.querySelectorAll('.pq-tool')].find((x) => x.textContent.trim() === 'Extract'); b?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b?.click(); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.trim() === 'First characters…')`), 'the Extract menu', 3000);
    await clickText('.rw-menu button', 'First characters…');
    await until(() => js(`Boolean(document.querySelector('.pq-ask-ok'))`), 'the Extract form', 3000);
    await js(`(() => { const v = document.querySelector('input.pq-ask-count'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(v, '3'); v.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(150);
    await js(`document.querySelector('.pq-ask-ok').click()`);
    const months = await until(async () => {
      const cells = await js(`[...document.querySelectorAll('.pq-grid tbody tr')].map((tr) => [...tr.querySelectorAll('td')].map((td) => td.textContent.trim()))`);
      return cells.length === 4 && cells.every((r) => r.some((c) => /^(Jan|Feb|Mar)$/.test(c)));
    }, 'the months cut short', 4000).then(() => true).catch(() => false);
    const steps = await js(`[...document.querySelectorAll('.pq-step-row .pq-step')].map((b) => b.textContent.trim())`);
    check('sheets: Extract → First characters keeps the first three letters of each month, listed as a step', months && steps.some((s) => /first 3 characters/.test(s)), JSON.stringify({ months, steps }));

    // Pivot Column: the short month names back into columns, the values added up.
    await js(`document.querySelector('.pq-grid th[data-col="Attribute"]').click()`);
    await wait(100);
    await clickText('.pq-tool', 'Pivot Column');
    await until(() => js(`Boolean(document.querySelector('.pq-ask-ok'))`), 'the Pivot form', 3000);
    await js(`document.querySelector('.pq-ask-ok').click()`);
    const pivoted = await until(async () => JSON.stringify(await heads()) === '["Product","Jan","Feb","Mar"]', 'the months as columns again', 4000).then(() => true).catch(() => false);
    check('sheets: Pivot Column turns the short month names back into columns, the values added up for each product', pivoted, JSON.stringify(await heads()));
    await js(`document.querySelector('.pq-load').click()`);
    await until(() => text('B1') === 'Jan', 'the query loaded', 6000).catch(() => {});
    const loaded = ['A1', 'B1', 'C1', 'D1', 'A2', 'B2', 'C2', 'D2', 'A3', 'B3', 'C3', 'D3'].map(text);
    check('sheets: the unpivoted, extracted and pivoted query loads as a table of its rows', loaded.join(',') === 'Product,Jan,Feb,Mar,Pens,10,12,,Ink,5,,7', loaded.join(','));
    const complaints = await errorsIn(win);
    check('sheets: the query transforms report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the query transform checks ran', false, err?.message || String(err));
  }
}
