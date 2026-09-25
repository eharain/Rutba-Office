// Worksheets: Data → Consolidate and Data → Forecast Sheet.
//
// Consolidate: from a Summary sheet, two regions' ranges added to the
// dialog's list and consolidated by their labels — the union of the rows
// and columns, summed, labels round it — then again with links to the
// source data, the detail rows folded under Excel's outline. Forecast Sheet:
// from three years of monthly revenue, the dialog draws its preview (the
// forecast and its confidence band) and Create makes a new sheet in front
// of the data with the table of FORECAST.ETS formulas and the chart, the
// band shaded. Run alone with RUTBA_VERIFY_ONLY=analysis.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const serial = (y, m) => Date.UTC(y, m, 1) / 86400000 + 25569;
const HISTORY = Array.from({ length: 36 }, (_, i) => [serial(2023, i), Math.round(4200 + 45 * i + 650 * Math.sin((2 * Math.PI * (i - 2)) / 12) + 90 * Math.sin(i * 1.7))]);

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} where the fixtures are written
 */
export async function verifyAnalysis(h, { dir }) {
  const { open, check, until, wait, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'analysis.xlsx');
  fs.writeFileSync(file, buildXlsx({ sheets: [
    { name: 'Summary', rows: [['']] },
    { name: 'East', rows: [['Product', 'Q1', 'Q2'], ['Pens', 120, 140], ['Ink', 60, 75], ['Paper', 300, 280]] },
    { name: 'West', rows: [['Product', 'Q2', 'Q3'], ['Ink', 40, 44], ['Folders', 90, 95], ['Pens', 80, 85]] },
    { name: 'Sales', rows: [['Month', 'Revenue'], ...HISTORY], styles: { 'A1:B1': { bold: true }, 'A2:A37': { numFmt: 'mmm yyyy' }, 'B2:B37': { numFmt: '#,##0' } } },
  ] }));
  try {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');
    const model = () => doc.model({ id: session.id });
    await until(() => js(`Boolean(document.querySelector('.sh-tab'))`), 'the grid', 8000);

    const clickIn = async (title) => {
      const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    const tab = (name) => js(`(() => { [...document.querySelectorAll('.rw-ribbon button, .rw-tabs button, button')].find((b) => b.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const sheetTab = async (name) => {
      await js(`(() => { [...document.querySelectorAll('.sh-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
      return until(() => model().activeSheet === name, 'the ' + name + ' sheet', 5000).catch(() => false);
    };
    const setField = (selector, v) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no field ' + ${JSON.stringify(selector)};
      const proto = el.tagName === 'SELECT' ? window.HTMLSelectElement.prototype : window.HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(v)});
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return 'set';
    })()`);
    const clickSel = (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'no ' + ${JSON.stringify(selector)}; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    const refresh = async () => { win.webContents.invalidate(); await wait(700); };
    const texts = (refs) => js(`${JSON.stringify(refs)}.map((r) => document.querySelector('.sh-cell[data-ref="' + r + '"]')?.textContent.trim() ?? '')`);

    // ── Data → Consolidate by labels, from the Summary sheet's A1.
    await sheetTab('Summary');
    await tab('Data');
    const opened = await clickIn('Consolidate');
    const dialog = await until(() => js(`Boolean(document.querySelector('.sh-cons-ref'))`), 'the Consolidate dialog', 5000).catch(() => false);
    const prefilled = await js(`document.querySelector('.sh-cons-ref')?.value || ''`);
    for (const r of ['East!$A$1:$C$4', 'West!$A$1:$C$4']) {
      await setField('.sh-cons-ref', r);
      await wait(120);
      await clickSel('.sh-cons-add');
      await until(() => js(`Boolean(document.querySelector('.sh-cons-list [data-ref="${r.replace(/\$/g, '\\$')}"]'))`), 'the reference listed', 3000).catch(() => {});
    }
    await js(`(() => { for (const c of ['.sh-cons-top', '.sh-cons-left']) { const b = document.querySelector(c); if (!b.checked) b.click(); } return 1; })()`);
    const listed = await js(`[...document.querySelectorAll('.sh-cons-list [data-ref]')].map((n) => n.dataset.ref)`);
    await wait(300);
    await refresh();
    await refresh();
    await capture(win, 'sheets-consolidate-dialog.png');
    await clickSel('.sh-cons-ok');
    const done = await until(async () => (await texts(['A4']))[0] === 'Paper', 'the consolidated table', 8000).catch(() => false);
    const grid = await texts(['B1', 'C1', 'D1', 'A2', 'B2', 'C2', 'D2', 'A3', 'C3', 'A5', 'C5', 'D5']);
    check('sheets: Data → Consolidate adds references to its list and sums them by their labels — rows and columns met by name, in the order first met, labels round the result',
      opened === 'clicked' && dialog === true && /^Summary!\$A\$1$/.test(prefilled) && listed.length === 2 && done === true
        && JSON.stringify(grid) === JSON.stringify(['Q1', 'Q2', 'Q3', 'Pens', '120', '220', '85', 'Ink', '115', 'Folders', '90', '95']),
      `${opened}; prefilled ${prefilled}; listed ${JSON.stringify(listed)}; grid ${JSON.stringify(grid)}`);
    await refresh();
    await capture(win, 'sheets-consolidate.png');

    // Again, below, with links to the source data: formulas, the detail folded under the outline.
    doc.apply({ id: session.id, ops: [{ op: 'select', row: 8, col: 0 }] });
    await js(`(() => { const g = document.querySelector('.sh-grid'); g.scrollTop += 1; g.scrollTop -= 1; g.dispatchEvent(new Event('scroll')); return 1; })()`);
    await wait(300);
    await clickIn('Consolidate');
    await until(() => js(`Boolean(document.querySelector('.sh-cons-links'))`), 'the dialog again', 5000).catch(() => {});
    const remembered = await js(`[...document.querySelectorAll('.sh-cons-list [data-ref]')].map((n) => n.dataset.ref)`);
    await js(`(() => { const b = document.querySelector('.sh-cons-links'); if (!b.checked) b.click(); return 1; })()`);
    await clickSel('.sh-cons-ok');
    const linked = await until(() => js(`Boolean(document.querySelector('.sh-gutter.rows')) && Boolean(document.querySelector('.sh-ol-box'))`), 'the outline', 8000).catch(() => false);
    const m = model();
    const hiddenDetail = !m.rows.some((r) => r.index === 9);
    const pens = m.cells.find((c) => c.ref === 'D12');
    check('sheets: Create links to source data writes formulas over detail rows folded under Excel\'s outline — the dialog remembered its references',
      remembered.length === 2 && linked === true && hiddenDetail && pens?.text === '220' && pens?.isFormula === true,
      `remembered ${JSON.stringify(remembered)}; outline ${linked}; detail hidden ${hiddenDetail}; D12 ${JSON.stringify(pens && { text: pens.text, formula: pens.formula })}`);

    // ── Data → Forecast Sheet from the monthly revenue.
    await sheetTab('Sales');
    doc.apply({ id: session.id, ops: [{ op: 'select', row: 3, col: 1 }] });
    await js(`(() => { const g = document.querySelector('.sh-grid'); g.scrollTop += 1; g.scrollTop -= 1; g.dispatchEvent(new Event('scroll')); return 1; })()`);
    await wait(300);
    await tab('Data');
    const fcOpened = await clickIn('Forecast Sheet');
    const previewed = await until(() => js(`Boolean(document.querySelector('.sh-fc-preview svg .band'))`), 'the preview with its band', 10000).catch(() => false);
    const fields = await js(`({ timeline: document.querySelector('.sh-fc-timeline')?.value, values: document.querySelector('.sh-fc-values')?.value, end: document.querySelector('.sh-fc-end')?.value })`);
    await refresh();
    await capture(win, 'sheets-forecast-dialog.png');
    check('sheets: Data → Forecast Sheet opens on the timeline and values round the cell, and previews the forecast with its confidence band',
      fcOpened === 'clicked' && previewed === true && fields.timeline === 'Sales!$A$2:$A$37' && fields.values === 'Sales!$B$2:$B$37' && /^\d{4}-\d{2}-\d{2}$/.test(fields.end || ''),
      `${fcOpened}; preview ${previewed}; ${JSON.stringify(fields)}`);

    // Column preview, back to Line, then Create.
    await clickSel('.sh-fc-kind[data-kind="column"]');
    const columns = await until(() => js(`(() => { const s = document.querySelector('.sh-fc-preview svg'); return Boolean(s) && !s.querySelector('.band') && s.querySelectorAll('path').length > 20; })()`), 'the column preview', 8000).catch(() => false);
    await clickSel('.sh-fc-kind[data-kind="line"]');
    await until(() => js(`Boolean(document.querySelector('.sh-fc-preview svg .band'))`), 'the line preview', 8000).catch(() => {});
    await clickSel('.sh-fc-create');
    const made = await until(() => model().activeSheet === 'Sheet5' && model().sheets.join(',') === 'Summary,East,West,Sheet5,Sales', 'the forecast sheet', 10000).catch(() => false);
    const chart = await until(() => js(`Boolean(document.querySelector('.sh-drawing svg .band'))`), 'the chart with its band', 10000).catch(() => false);
    const heads = await texts(['A1', 'B1', 'C1', 'D1', 'E1']);
    const cm = model();
    doc.save({ id: session.id });
    const saved = new SheetView(fs.readFileSync(file));
    const c38 = { formula: saved.calc.getInput('Sheet5', 37, 2), text: String(saved.calc.getValue('Sheet5', 37, 2)) };
    check('sheets: Create makes the forecast sheet in front of the data — Excel\'s table with FORECAST.ETS past the last month, and the chart with its shaded band',
      columns === true && made === true && chart === true
        && JSON.stringify(heads) === JSON.stringify(['Month', 'Revenue', 'Forecast(Revenue)', 'Lower Confidence Bound(Revenue)', 'Upper Confidence Bound(Revenue)'])
        && /^=FORECAST\.ETS\(A38,/.test(c38?.formula || '') && Number(String(c38?.text || '').replace(/,/g, '')) > 3000,
      `columns ${columns}; sheet ${made} (${cm.sheets.join(',')}); chart ${chart}; heads ${JSON.stringify(heads)}; C38 ${JSON.stringify(c38 && { text: c38.text, formula: c38.formula })}`);
    await refresh();
    await capture(win, 'sheets-forecast-sheet.png');
    // The table where history meets the forecast.
    await js(`(() => { document.querySelector('.sh-grid').scrollTop = 520; return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.sh-rowheads .sh-head')].some((n) => n.textContent.trim() === '45')`), 'the forecast rows', 5000).catch(() => {});
    await refresh();
    await capture(win, 'sheets-forecast-table.png');

    const complaints = await errorsIn(win);
    check('sheets: the consolidate and forecast checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the consolidate and forecast checks ran', false, err.message);
  }
}
