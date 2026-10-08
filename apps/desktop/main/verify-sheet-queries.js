// Worksheets: Data → Get & Transform — From Table/Range opens the Power
// Query Editor on the list round the cell; a column taken away and the rows
// filtered there, Close & Load writes the result as a table on a sheet of
// its own and lists the query in Queries & Connections; the source changed,
// Refresh All brings it in; Edit opens the query again, and a step taken
// away loads the rows it had kept out; Recent Sources reads a CSV file into
// the workbook. Run alone with RUTBA_VERIFY_ONLY=queries.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

export async function verifySheetQueries({ open, check, until, wait, errorsIn, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'queries.xlsx');
  const csv = path.join(dir, 'stock.csv');
  try {
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Data', rows: [['Region', 'Rep', 'Units', 'Sales'], ['East', 'Kim', 3, 300], ['East', 'Lee', 2, 250], ['North', 'Ann', 5, 400], ['West', 'Bo', 1, 90], ['West', 'Cy', 4, 210]] }] }));
    fs.writeFileSync(csv, 'Item,Qty\nInk,4\nPaper,6\nToner,1\n');
    const win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const id = () => sessionFor('sheet').id;
    const model = () => doc.model({ id: id() });
    const text = (ref) => String((model().cells || []).find((c) => c.ref === ref)?.text ?? '');
    const clickText = (selector, label) => js(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(selector)})].find((x) => x.textContent.trim() === ${JSON.stringify(label)}); if (!b) return false; b.click(); return true; })()`);
    const heads = () => js(`[...document.querySelectorAll('.pq-grid th[data-col]')].map((th) => th.textContent.trim())`);
    const rowsShown = () => js(`document.querySelectorAll('.pq-grid tbody tr').length`);
    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 8000);
    await clickText('.rw-tab', 'Data');
    await until(() => js(`[...document.querySelectorAll('button')].some((b) => b.textContent.trim() === 'From Table/Range')`), 'the Data tab', 4000);

    // From Table/Range, the active cell A1 in the list.
    await clickText('button', 'From Table/Range');
    await until(() => js(`document.querySelectorAll('.pq-grid th[data-col]').length === 4`), 'the Power Query Editor', 6000).catch(() => {});
    check('sheets: From Table/Range opens the Power Query Editor on the list round the cell, its headings as columns', JSON.stringify(await heads()) === '["Region","Rep","Units","Sales"]' && (await rowsShown()) === 5, JSON.stringify(await heads()));

    // Units picked and taken away; the rows with sales under 250 filtered out.
    await js(`document.querySelector('.pq-grid th[data-col="Units"]').click()`);
    await wait(100);
    await clickText('.pq-tool', 'Remove Column');
    await until(async () => (await heads()).length === 3, 'the column removed', 4000).catch(() => {});
    await js(`document.querySelector('.pq-grid th[data-col="Sales"]').click()`);
    await wait(100);
    await clickText('.pq-tool', 'Filter');
    await until(() => js(`Boolean(document.querySelector('.pq-ask-ok'))`), 'the filter form', 3000);
    await js(`(() => {
      const op = document.querySelector('select.pq-ask-op');
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(op, 'greaterOrEqual');
      op.dispatchEvent(new Event('change', { bubbles: true }));
      const v = document.querySelector('input.pq-ask-value');
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(v, '250');
      v.dispatchEvent(new Event('input', { bubbles: true }));
      return 1;
    })()`);
    await wait(150);
    await js(`document.querySelector('.pq-ask-ok').click()`);
    await until(async () => (await rowsShown()) === 3, 'the rows filtered', 4000).catch(() => {});
    const steps = await js(`[...document.querySelectorAll('.pq-step-row .pq-step')].map((b) => b.textContent.trim())`);
    check('sheets: the editor removes a column and filters the rows, each listed as an applied step', (await rowsShown()) === 3 && steps.length === 2 && /Units/.test(steps[0]) && /Sales/.test(steps[1]), `${await rowsShown()} rows; ${steps.join(' | ')}`);

    // Named and loaded.
    await js(`(() => {
      const n = document.querySelector('input.pq-name');
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, 'Big sales');
      n.dispatchEvent(new Event('input', { bubbles: true }));
      return 1;
    })()`);
    await wait(100);
    await js(`document.querySelector('.pq-load').click()`);
    await until(() => model().sheet === 'Big sales', 'the query loaded', 6000).catch(() => {});
    check('sheets: Close & Load writes the result as a table on a sheet named after the query',
      model().sheet === 'Big sales' && ['A1', 'B1', 'C1'].map(text).join(',') === 'Region,Rep,Sales' && ['C2', 'C3', 'C4'].map(text).join(',') === '300,250,400' && text('A5') === ''
        && (model().tables || []).some((t) => t.ref === 'A1:C4'),
      `${model().sheet}: ${['A1', 'B1', 'C1', 'C2', 'C3', 'C4', 'A5'].map(text).join(',')}`);
    await until(() => js(`Boolean(document.querySelector('.pq-item[data-query]'))`), 'Queries & Connections', 4000).catch(() => {});
    check('sheets: the query is listed in Queries & Connections with the rows it loaded', await js(`(() => { const it = document.querySelector('.pq-item'); return Boolean(it) && /Big sales/.test(it.textContent) && /3 rows loaded/.test(it.textContent); })()`), await js(`document.querySelector('.pq-pane')?.textContent || 'no pane'`));

    // The source changes behind it; Refresh All brings Bo's sale in.
    await doc.apply({ id: id(), ops: [{ op: 'sheet', name: 'Data' }, { op: 'setCell', row: 4, col: 3, value: 600 }, { op: 'sheet', name: 'Big sales' }] });
    await js(`document.querySelector('.pq-refresh-all').click()`);
    await until(() => text('C5') === '600', 'the refresh', 6000).catch(() => {});
    check('sheets: Refresh All runs the query again on its source as it is now, the table grown to fit', text('A5') === 'West' && text('C5') === '600' && (model().tables || []).some((t) => t.ref === 'A1:C5'), `${['A5', 'C5'].map(text).join(',')} ${(model().tables || []).map((t) => t.ref).join(' ')}`);

    // Edit: the filter step taken away loads every row.
    await js(`document.querySelector('.pq-edit').click()`);
    await until(() => js(`document.querySelectorAll('.pq-step-row').length === 2`), 'the editor again', 5000).catch(() => {});
    await js(`document.querySelectorAll('.pq-step-x')[1].click()`);
    await until(async () => (await rowsShown()) === 5, 'the filter gone', 4000).catch(() => {});
    await js(`document.querySelector('.pq-load').click()`);
    await until(() => text('C6') === '210', 'the query loaded again', 6000).catch(() => {});
    check('sheets: Edit opens the query in the editor, and a step taken away loads the rows it kept out', text('C6') === '210' && text('A7') === '' && model().queries?.length === 1 && model().queries[0].steps === 1, `${['C2', 'C6', 'A7'].map(text).join(',')} steps ${model().queries?.[0]?.steps}`);

    // Recent Sources: a CSV file read into the workbook.
    await js(`localStorage.setItem('sheets.recentSources', ${JSON.stringify(JSON.stringify([{ source: { kind: 'csv', path: csv }, sourceText: 'stock.csv' }]))}); 1`);
    await clickText('button', 'Recent Sources');
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.trim() === 'stock.csv')`), 'the Recent Sources menu', 3000);
    await clickText('.rw-menu button', 'stock.csv');
    await until(async () => JSON.stringify(await heads()) === '["Item","Qty"]', 'the CSV in the editor', 6000).catch(() => {});
    await js(`document.querySelector('.pq-load').click()`);
    await until(() => model().sheet === 'stock', 'the CSV loaded', 6000).catch(() => {});
    check('sheets: Recent Sources reads a CSV file into the workbook as a query, its first row the headings', model().sheet === 'stock' && ['A1', 'B1', 'A4', 'B4'].map(text).join(',') === 'Item,Qty,Toner,1' && model().queries?.length === 2, `${model().sheet}: ${['A1', 'B1', 'A4', 'B4'].map(text).join(',')}`);

    const complaints = await errorsIn(win);
    check('sheets: Get & Transform reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the Get & Transform checks ran', false, err.message);
  }
}
