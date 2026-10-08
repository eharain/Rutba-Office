// Worksheets: Power Query → Merge Queries and Append Queries, in the editor
// — the people list loaded as a query, then the orders opened in the editor
// and merged with it on their customer key (inner join), their names brought
// in; and the extra orders appended. Run alone with RUTBA_VERIFY_ONLY=querymerge.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

export async function verifySheetQueriesMerge({ open, check, until, wait, errorsIn, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'queries-merge.xlsx');
  try {
    fs.writeFileSync(file, buildXlsx({ sheets: [
      { name: 'Orders', rows: [['Order', 'Customer', 'Total'], [1, 'C1', 100], [2, 'C2', 250], [3, 'C9', 40]] },
      { name: 'People', rows: [['Id', 'Name'], ['C1', 'Kim'], ['C2', 'Lee']] },
      { name: 'Late', rows: [['Order', 'Customer', 'Total'], [4, 'C2', 80]] },
    ] }));
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('sheet').id });
    const text = (ref) => String((model().cells || []).find((c) => c.ref === ref)?.text ?? '');
    const clickText = (selector, label) => js(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(selector)})].find((x) => x.textContent.trim() === ${JSON.stringify(label)}); if (!b) return false; b.click(); return true; })()`);
    const heads = () => js(`[...document.querySelectorAll('.pq-grid th[data-col]')].map((th) => th.textContent.trim())`);
    const choose = (selector, value) => js(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return false; Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('change', { bubbles: true })); return true; })()`);
    const fromHere = async (sheet, name) => {
      await clickText('.sh-tab, .sh-tabs button, [data-sheet]', sheet);
      await doc.apply({ id: sessionFor('sheet').id, ops: [{ op: 'sheet', name: sheet }] });
      await wait(300);
      await clickText('button', 'From Table/Range');
      await until(() => js(`document.querySelectorAll('.pq-grid th[data-col]').length > 0`), `the editor on ${sheet}`, 6000).catch(() => {});
      if (name) {
        await js(`(() => { const n = document.querySelector('input.pq-name'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(name)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
        await wait(100);
      }
    };
    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 8000);
    await clickText('.rw-tab', 'Data');
    await wait(300);

    // The people, loaded as a query.
    await fromHere('People', 'People list');
    await js(`document.querySelector('.pq-load').click()`);
    await until(() => (model().queries || []).length === 1, 'the people query', 6000).catch(() => {});

    // The orders in the editor: merged with the people on Customer = Id, only matching rows.
    await fromHere('Orders', 'Orders named');
    await js(`document.querySelector('.pq-grid th[data-col="Customer"]').click()`);
    await wait(100);
    await clickText('.pq-tool', 'Merge Queries');
    await until(() => js(`Boolean(document.querySelector('select.pq-ask-with'))`), 'the merge form', 4000).catch(() => {});
    const offered = await js(`[...document.querySelectorAll('select.pq-ask-with option')].map((o) => o.textContent)`);
    await choose('select.pq-ask-how', 'inner');
    await wait(100);
    await js(`document.querySelector('.pq-ask-ok').click()`);
    await until(() => js(`Boolean(document.querySelector('select.pq-ask-withOn'))`), 'the key form', 4000).catch(() => {});
    await choose('select.pq-ask-withOn', 'Id');
    await wait(100);
    await js(`document.querySelector('.pq-ask-ok').click()`);
    await until(async () => (await heads()).includes('Name'), 'the merge', 5000).catch(() => {});
    const merged = { heads: await heads(), rows: await js(`document.querySelectorAll('.pq-grid tbody tr').length`) };
    check('sheets: Merge Queries joins the orders to the people query on their key, only the matching rows, the names brought in',
      offered.some((o) => /Query: People list/.test(o)) && JSON.stringify(merged.heads) === '["Order","Customer","Total","Name"]' && merged.rows === 2,
      JSON.stringify({ offered, merged }));

    // Append: there is no table of the late orders, so a query of them is made first, then appended.
    await js(`document.querySelector('.pq-load').click()`);
    await until(() => (model().queries || []).length === 2, 'the merged query', 6000).catch(() => {});
    check('sheets: the merged query loads its rows with the names', text('D2') === 'Kim' && text('D3') === 'Lee' && text('A4') === '', `${model().sheet}: ${['A1', 'D1', 'D2', 'D3', 'A4'].map(text).join(',')}`);
    await fromHere('Late', 'Late orders');
    await js(`document.querySelector('.pq-load').click()`);
    await until(() => (model().queries || []).length === 3, 'the late query', 6000).catch(() => {});
    await fromHere('Orders', 'All orders');
    await clickText('.pq-tool', 'Append Queries');
    await until(() => js(`Boolean(document.querySelector('select.pq-ask-with'))`), 'the append form', 4000).catch(() => {});
    const late = await js(`[...document.querySelectorAll('select.pq-ask-with option')].find((o) => /Late orders/.test(o.textContent))?.value ?? null`);
    await choose('select.pq-ask-with', late);
    await wait(100);
    await js(`document.querySelector('.pq-ask-ok').click()`);
    await until(() => js(`document.querySelectorAll('.pq-grid tbody tr').length === 4`), 'the append', 5000).catch(() => {});
    const appended = await js(`[...document.querySelectorAll('.pq-grid tbody tr')].map((r) => r.children[1]?.textContent)`);
    check('sheets: Append Queries puts another query\'s rows after these', JSON.stringify(appended) === '["1","2","3","4"]', JSON.stringify(appended));
    await js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => b.textContent.trim() === 'Cancel')?.click(), 1`);

    const complaints = await errorsIn(win);
    check('sheets: Merge and Append report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the Merge and Append checks ran', false, err?.message || JSON.stringify(err));
  }
}
