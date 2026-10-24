// Worksheets 2.18: a list worked on in the window. Conditional formatting's
// Apply adds its rule and the cells show it; the filter arrows in the
// heading row open the value list, and OK hides the rows left unticked; a
// chart redraws as soon as a figure it shows is edited; Chart Design sets
// the chart's title; and the Sort dialog names the columns from the list's
// heading row, which stays at the top. Run alone with
// RUTBA_VERIFY_ONLY=sheetlists.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc, sessionFor
 * @param {{ dir: string }} where the file is written
 */
export async function verifySheetLists(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'lists-2-18.xlsx');
  fs.writeFileSync(file, buildXlsx({ sheets: [{
    name: 'Data',
    rows: [
      ['Item', 'Region', 'Qty', '', 'Month', 'Takings'],
      ['Ink', 'North', 4, '', 'Jan', 120],
      ['Paper', 'South', 6, '', 'Feb', 80],
      ['Pens', 'North', 2, '', 'Mar', 95],
      ['Tape', 'South', 9],
    ],
    styles: { 'A1:C1': { bold: true } },
  }] }));
  try {
    const win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const capture = async (name) => { if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG()); };
    const cell = (ref) => js(`(document.querySelector('.sh-cell[data-ref="${ref}"]')?.innerText || '').trim()`);
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="C5"]'))`), 'the grid', 8000);
    const session = sessionFor('sheet');
    const nudge = () => js(`(() => { const g = document.querySelector('.sh-grid'); g.scrollTop += 1; g.scrollTop -= 1; g.dispatchEvent(new Event('scroll')); return 1; })()`);
    const apply = async (...ops) => { const r = await doc.apply({ id: session.id, ops }); await nudge(); await wait(350); return r; };
    const tab = (name) => js(`(() => { const b = [...document.querySelectorAll('.rw-tabs button, .rw-tabs [role="tab"]')].find((n) => n.textContent.trim() === ${JSON.stringify(name)}); if (!b) return 'no tab ' + ${JSON.stringify(name)}; b.click(); return 'clicked'; })()`);
    const press = async (label) => {
      const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.textContent || '').trim() === ${JSON.stringify(label)} && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${label} button`, 3000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(label)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    const setValue = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no ' + ${JSON.stringify(selector)};
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return 'set';
    })()`);
    const dialogButton = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-dialog button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no ' + ${JSON.stringify(label)}; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    const background = (ref) => js(`(() => { const c = document.querySelector('.sh-cell[data-ref="${ref}"]'); return c ? getComputedStyle(c).backgroundColor : null; })()`);

    // ── Conditional formatting: Qty greater than 5, shown in red.
    await apply({ op: 'select', row: 1, col: 2 }, { op: 'select', row: 4, col: 2, extend: true });
    await tab('Home');
    const cfOpened = await press('Conditional Formatting');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog select'))`), 'Conditional formatting', 4000).catch(() => {});
    const typed = await setValue('.rw-dialog input:not([disabled])', '5');
    await wait(150);
    const applied = await dialogButton('Apply');
    const shaded = await until(async () => (await background('C3')) !== (await background('C2')) && (await background('C5')) === (await background('C3')), 'the rule shown', 5000).then(() => true).catch(() => false);
    const cfErrors = await errorsIn(win);
    check('worksheets: conditional formatting\'s Apply adds the rule and the cells over 5 are shaded at once',
      cfOpened === 'clicked' && typed === 'set' && applied === 'clicked' && shaded && cfErrors.length === 0,
      JSON.stringify({ cfOpened, typed, applied, C2: await background('C2'), C3: await background('C3'), C5: await background('C5'), cfErrors }));

    // ── Filter arrows: Data → Filter puts one in each heading; South unticked hides its rows.
    await apply({ op: 'select', row: 1, col: 0 });
    await tab('Data');
    const filterOn = await press('Filter');
    const arrows = await until(() => js(`document.querySelectorAll('.sh-filter-arrow').length`), 'the filter arrows', 5000).then(() => js(`[...document.querySelectorAll('.sh-cell .sh-filter-arrow')].map((a) => a.closest('.sh-cell')?.dataset.ref)`)).catch(() => []);
    await js(`(() => { document.querySelector('.sh-cell[data-ref="B1"] .sh-filter-arrow')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); return 1; })()`);
    const listed = await until(() => js(`document.querySelectorAll('.sh-filter-menu .sh-filter-value[data-value]').length === 2`), 'the value list', 5000)
      .then(() => js(`[...document.querySelectorAll('.sh-filter-menu .sh-filter-value[data-value]')].map((v) => v.dataset.value)`)).catch(() => []);
    await js(`(() => { document.querySelector('.sh-filter-menu .sh-filter-value[data-value="South"] input')?.click(); return 1; })()`);
    await wait(150);
    await js(`(() => { [...document.querySelectorAll('.sh-filter-menu button')].find((b) => b.textContent.trim() === 'OK')?.click(); return 1; })()`);
    const hid = await until(async () => !(await js(`Boolean(document.querySelector('.sh-cell[data-ref="A3"]'))`)) && (await cell('A4')) === 'Pens', 'the South rows hidden', 5000).then(() => true).catch(() => false);
    const marked = await js(`Boolean(document.querySelector('.sh-cell[data-ref="B1"] .sh-filter-arrow.on'))`);
    await capture('sheet-filter-arrows.png');
    check('worksheets: the heading row has filter arrows; an arrow lists the column\'s values, and OK hides the rows left unticked',
      filterOn === 'clicked' && JSON.stringify(arrows.slice(0, 3)) === '["A1","B1","C1"]' && JSON.stringify(listed) === '["North","South"]' && hid && marked,
      JSON.stringify({ filterOn, arrows, listed, hid, marked, A3: await cell('A3'), heads: await js(`document.querySelectorAll('.sh-cell.filter-head').length`) }));
    await js(`(() => { document.querySelector('.sh-cell[data-ref="B1"] .sh-filter-arrow')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-filter-menu .sh-filter-clear'))`), 'the menu again', 4000).catch(() => {});
    const cleared = await js(`(() => { const b = document.querySelector('.sh-filter-menu .sh-filter-clear'); if (!b) return 'no Clear filter'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    const back = await until(async () => (await cell('A3')) === 'Paper', 'the rows back', 5000).then(() => true).catch(() => false);
    check('worksheets: Clear filter from the arrow\'s menu shows the rows again', cleared === 'clicked' && back, JSON.stringify({ cleared, A3: await cell('A3') }));

    // ── A chart follows its cells: Feb's takings edited, the chart drawn again at once.
    await apply({ op: 'select', row: 1, col: 4 }, { op: 'insertChart', kind: 'column' });
    const chartSvg = () => js(`document.querySelector('.sh-drawing[data-kind="chart"] svg')?.outerHTML || ''`);
    await until(async () => (await chartSvg()).length > 0, 'the chart', 6000).catch(() => {});
    const before = await chartSvg();
    await apply({ op: 'setCell', row: 2, col: 5, value: '400' });
    const followed = await until(async () => { const now = await chartSvg(); return now !== before && now.includes('>400<'); }, 'the chart redrawn', 6000).then(() => true).catch(() => false);
    check('worksheets: a chart is drawn again as soon as a figure it shows is edited',
      before.length > 0 && followed, JSON.stringify({ drawn: before.length > 0, followed }));

    // ── Chart Design: the picked chart's title.
    const box = await js(`(() => { const d = document.querySelector('.sh-drawing[data-kind="chart"]'); if (!d) return null; const r = d.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    if (box) {
      wc.sendInputEvent({ type: 'mouseMove', x: box.x, y: box.y });
      wc.sendInputEvent({ type: 'mouseDown', x: box.x, y: box.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: box.x, y: box.y, button: 'left', clickCount: 1 });
    }
    const designTab = await until(() => js(`Boolean([...document.querySelectorAll('.rw-tabs button, .rw-tabs [role="tab"]')].find((n) => n.textContent.trim() === 'Chart Design'))`), 'the Chart Design tab', 5000).then(() => true).catch(() => false);
    await tab('Chart Design');
    await until(() => js(`Boolean(document.querySelector('.sh-chart-title'))`), 'the title field', 4000).catch(() => {});
    const titled = await setValue('.sh-chart-title', 'Takings by month');
    await js(`(() => { document.querySelector('.sh-chart-title')?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })); return 1; })()`);
    const drawnTitle = await until(async () => (await chartSvg()).includes('Takings by month'), 'the title drawn', 5000).then(() => true).catch(() => false);
    await capture('sheet-chart-design.png');
    check('worksheets: a picked chart has a Chart Design tab, and its title field sets the chart\'s title',
      Boolean(box) && designTab && titled === 'set' && drawnTitle, JSON.stringify({ box, designTab, titled, drawnTitle }));
    await js(`(() => { document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); return 1; })()`);

    // ── Sort: the dialog names the columns from the heading row, which stays at the top.
    await apply({ op: 'select', row: 2, col: 1 });
    await tab('Data');
    const sortOpened = await press('Sort');
    await until(() => js(`Boolean(document.querySelector('.sh-sort-col-0'))`), 'the Sort dialog', 5000).catch(() => {});
    const offered = await js(`[...document.querySelectorAll('.sh-sort-col-0 option')].map((o) => o.textContent.trim())`);
    const said = await js(`Boolean(document.querySelector('.sh-sort-header'))`);
    await setValue('.sh-sort-col-0', '2');
    await setValue('.sh-sort-dir-0', 'desc');
    await wait(150);
    await js(`(() => { document.querySelector('.sh-sort-ok')?.click(); return 1; })()`);
    const sorted = await until(async () => (await cell('C2')) === '9', 'the sort', 5000).then(() => true).catch(() => false);
    const top = { A1: await cell('A1'), A2: await cell('A2'), C2: await cell('C2'), C5: await cell('C5') };
    check('worksheets: the Sort dialog names the columns from the list\'s heading row, and sorting leaves that row at the top',
      sortOpened === 'clicked' && JSON.stringify(offered) === '["Item","Region","Qty"]' && said && sorted && top.A1 === 'Item' && top.A2 === 'Tape' && top.C5 === '2',
      JSON.stringify({ sortOpened, offered, said, sorted, top }));

    const complaints = await errorsIn(win);
    check('worksheets: none of the list checks reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('worksheets: the list checks ran', false, err.message);
  }
}
