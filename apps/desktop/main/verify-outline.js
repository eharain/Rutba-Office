// Worksheets: Data → Group, the outline gutter, and Subtotal.
//
// Rows are grouped from the ribbon with the top row frozen, and the gutter
// Excel draws must appear beside the row headings without pushing a heading
// off its row: the − box at the summary row, the level buttons in the
// corner. The box folds the group, the level buttons fold and open it, and
// Shift+Alt+Left takes it away. Then Subtotal on the list: a total under
// each region, a Grand Total, the figures read off the grid, level 2 showing
// only the totals; Remove All puts the list back, and the saved file is the
// fixture as it was. Run alone with RUTBA_VERIFY_ONLY=outline.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const LIST = [
  ['Region', 'Rep', 'Units', 'Sales'],
  ['East', 'Kim', 3, 300],
  ['East', 'Lee', 2, 250],
  ['North', 'Ann', 5, 400],
  ['West', 'Bo', 1, 90],
  ['West', 'Cy', 4, 210],
  ['West', 'Di', 2, 120],
];

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture
 * @param {{ dir: string }} where the fixture is written
 */
export async function verifyOutline(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, capture } = h;
  const file = path.join(dir, 'outline.xlsx');
  fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sales', rows: LIST }] }));
  const before = new SheetView(fs.readFileSync(file));
  const original = LIST.map((row, r) => row.map((_, c) => before.displayValue(r, c).text));
  try {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="D7"]'))`), 'the grid', 8000);

    const clickIn = async (title) => {
      const find = `[...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    const tab = (name) => js(`(() => { [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const menuItem = async (pattern) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => ${pattern}.test(b.textContent)))`), 'the menu', 4000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => ${pattern}.test(b.textContent)); if (!b) return 'no item'; b.click(); return 'clicked'; })()`);
    };
    const heads = () => js(`[...document.querySelectorAll('.sh-rowheads .sh-head')].map((n) => n.textContent.trim())`);
    const selectRows = async (a, b) => {
      await js(`(() => {
        const h = (label) => [...document.querySelectorAll('.sh-rowheads .sh-head')].find((n) => n.textContent.trim() === label);
        h(${JSON.stringify(String(a))}).dispatchEvent(new MouseEvent('click', { bubbles: true }));
        return 1;
      })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.sh-rowheads .sh-head.active')].find((n) => n.textContent.trim() === ${JSON.stringify(String(a))}))`), 'the first row selected', 4000).catch(() => {});
      await js(`(() => {
        const h = [...document.querySelectorAll('.sh-rowheads .sh-head')].find((n) => n.textContent.trim() === ${JSON.stringify(String(b))});
        h.dispatchEvent(new MouseEvent('click', { bubbles: true, shiftKey: true }));
        return 1;
      })()`);
      return until(() => js(`[...document.querySelectorAll('.sh-rowheads .sh-head.active')].length === ${b - a + 1}`), 'the rows selected', 4000).catch(() => false);
    };
    const box = (start, level = 1) => `.sh-ol-box[data-axis="row"][data-level="${level}"][data-start="${start}"]`;
    const refresh = async () => {
      // An off-screen window hands back its last painted frame.
      win.webContents.invalidate();
      await wait(700);
    };

    // The top row frozen first: the gutter and the pinned panes must live together.
    await tab('View');
    await until(() => js(`Boolean([...document.querySelectorAll('button')].find((b) => /Freeze Panes/.test(b.textContent)))`), 'the View tab', 4000);
    await js(`(() => { [...document.querySelectorAll('button')].find((b) => /Freeze Panes/.test(b.textContent)).click(); return 1; })()`);
    await menuItem(/Freeze top row/);
    const frozen = await until(() => js(`Boolean(document.querySelector('.sh-pin-rows .sh-cell[data-ref="A1"]'))`), 'the frozen top row', 5000).catch(() => false);

    // Rows 2 to 4 by their headings, then Data → Group.
    const picked = await selectRows(2, 4);
    await tab('Data');
    const pressed = await clickIn('Group');
    const grouped = await until(() => js(`Boolean(document.querySelector('${box(1)}:not(.folded)')) && Boolean(document.querySelector('.sh-gutter.rows'))`), 'the gutter and its − box', 6000).catch(() => false);
    await wait(300);
    const seen = await js(`(() => {
      const r = (n) => { if (!n) return null; const b = n.getBoundingClientRect(); return { left: b.left, right: b.right, top: b.top, bottom: b.bottom }; };
      const head = (label) => [...document.querySelectorAll('.sh-rowheads .sh-head')].find((n) => n.textContent.trim() === label);
      return {
        gutter: r(document.querySelector('.sh-rowheads > .sh-gutter.rows')),
        corner: r(document.querySelector('.sh-corner')),
        head1: r(head('1')), head2: r(head('2')), head5: r(head('5')),
        colA: r([...document.querySelectorAll('.sh-colheads .sh-head')].find((n) => n.textContent.trim() === 'A')),
        a1: r(document.querySelector('.sh-pin-rows .sh-cell[data-ref="A1"]')),
        a2: r(document.querySelector('.sh-cell[data-ref="A2"]')),
        box: r(document.querySelector('${box(1)}')),
        levels: [...document.querySelectorAll('.sh-ol-level[data-axis="row"]')].map((b) => b.textContent.trim()),
        lines: document.querySelectorAll('.sh-rowheads > .sh-gutter.rows polyline').length,
        tip: document.querySelector('${box(1)}')?.dataset.tip || '',
      };
    })()`);
    const near = (a, b, d = 1) => a != null && b != null && Math.abs(a - b) <= d;
    const aligned = seen.gutter && seen.head2 && seen.a2 && seen.colA
      && near(seen.head2.top, seen.a2.top) && near(seen.head2.bottom, seen.a2.bottom)
      && near(seen.head2.left, seen.gutter.right) && near(seen.head2.right, seen.a2.left)
      && near(seen.colA.left, seen.a2.left) && near(seen.corner.right, seen.head2.right)
      && near(seen.head1?.left, seen.head2.left) && near(seen.a1?.left, seen.a2.left);
    const boxAt = seen.box && seen.head5 && near((seen.box.top + seen.box.bottom) / 2, (seen.head5.top + seen.head5.bottom) / 2, 1.5)
      && seen.box.left >= seen.gutter.left && seen.box.right <= seen.gutter.right;
    check('sheets: Data → Group on rows 2 to 4 draws Excel\'s gutter beside the row headings — a bracket, the − box at row 5, the level buttons 1 2 in the corner',
      frozen === true && picked === true && pressed === 'clicked' && grouped === true && boxAt && seen.levels.join(',') === '1,2' && seen.lines >= 1 && /^Hide Detail — rows 2 to 4$/.test(seen.tip),
      `frozen ${frozen}, picked ${picked}, ${pressed}, grouped ${grouped}, box ${JSON.stringify(seen.box)} vs row 5 ${JSON.stringify(seen.head5)}, levels ${seen.levels}, lines ${seen.lines}, tip ${JSON.stringify(seen.tip)}`);
    check('sheets: the gutter widens the heading track without moving the grid — each heading still meets its row, the column headings their cells, the frozen row its column',
      Boolean(aligned), JSON.stringify({ gutter: seen.gutter, head2: seen.head2, a2: seen.a2, colA: seen.colA, corner: seen.corner, head1: seen.head1, a1: seen.a1 }));
    await refresh();
    await capture(win, 'sheets-outline.png');

    // The − box folds the group: rows 2 to 4 leave the grid, the box turns +.
    await js(`(() => { document.querySelector('${box(1)}').click(); return 1; })()`);
    const folded = await until(async () => {
      const h = await heads();
      return !h.includes('2') && !h.includes('4') && h.includes('5') && (await js(`Boolean(document.querySelector('${box(1)}.folded'))`));
    }, 'the group to fold', 5000).catch(() => false);
    const foldedTip = await js(`document.querySelector('${box(1)}')?.dataset.tip || ''`);
    check('sheets: the − box folds the group — rows 2 to 4 leave the grid and the box turns to + (Show Detail)', folded === true && /^Show Detail/.test(foldedTip), `${folded}; headings ${(await heads()).slice(0, 8).join(' ')}; tip ${foldedTip}`);

    // The level buttons: 2 shows everything, 1 folds again, 2 once more.
    const level = (n) => js(`(() => { document.querySelector('.sh-ol-level[data-axis="row"][data-level="${n}"]').click(); return 1; })()`);
    await level(2);
    const shown = await until(async () => (await heads()).includes('3') && (await js(`Boolean(document.querySelector('${box(1)}:not(.folded)'))`)), 'level 2 to show the rows', 5000).catch(() => false);
    await level(1);
    const levelOne = await until(async () => !(await heads()).includes('3'), 'level 1 to fold them', 5000).catch(() => false);
    await level(2);
    const again = await until(async () => (await heads()).includes('3'), 'level 2 again', 5000).catch(() => false);
    check('sheets: the level buttons fold and open the outline — 2 shows every row, 1 folds the group away', shown === true && levelOne === true && again === true, `2 ${shown}, 1 ${levelOne}, 2 ${again}`);

    // Shift+Alt+Left on the same rows takes the group away, gutter and all.
    await selectRows(2, 4);
    await js(`(() => { document.querySelector('.sh')?.focus(); return 1; })()`);
    await press(win.webContents, 'Left', { modifiers: ['shift', 'alt'] });
    const ungrouped = await until(() => js(`!document.querySelector('.sh-gutter') && !document.querySelector('.sh-ol-level')`), 'the gutter to go', 5000).catch(() => false);
    const flush = await js(`(() => { const h = [...document.querySelectorAll('.sh-rowheads .sh-head')].find((n) => n.textContent.trim() === '2')?.getBoundingClientRect(); const c = document.querySelector('.sh-cell[data-ref="A2"]')?.getBoundingClientRect(); return h && c ? Math.abs(h.right - c.left) <= 1 && Math.abs(h.left - document.querySelector('.sh-rowheads').getBoundingClientRect().left) <= 1 : false; })()`);
    check('sheets: Shift+Alt+Left ungroups the rows, and the headings go back to the edge', ungrouped === true && flush === true, `gone ${ungrouped}, flush ${flush}`);

    // Columns B and C by their headings: the gutter above the column
    // headings, the − box over column D, the level buttons down the corner.
    await js(`(() => {
      const h = (label) => [...document.querySelectorAll('.sh-colheads .sh-head')].find((n) => n.textContent.trim() === label);
      h('B').dispatchEvent(new MouseEvent('click', { bubbles: true }));
      return 1;
    })()`);
    await until(() => js(`Boolean([...document.querySelectorAll('.sh-colheads .sh-head.active')].find((n) => n.textContent.trim() === 'B'))`), 'column B selected', 4000).catch(() => {});
    await js(`(() => { [...document.querySelectorAll('.sh-colheads .sh-head')].find((n) => n.textContent.trim() === 'C').dispatchEvent(new MouseEvent('click', { bubbles: true, shiftKey: true })); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.sh-colheads .sh-head.active')].length === 2`), 'columns B and C selected', 4000).catch(() => {});
    await js(`(() => { document.querySelector('.sh')?.focus(); return 1; })()`);
    await press(win.webContents, 'Right', { modifiers: ['shift', 'alt'] });
    const colBox = '.sh-ol-box[data-axis="col"][data-level="1"][data-start="1"]';
    const colGrouped = await until(() => js(`Boolean(document.querySelector('${colBox}'))`), 'the column gutter', 5000).catch(() => false);
    await wait(300);
    const colSeen = await js(`(() => {
      const r = (n) => { if (!n) return null; const b = n.getBoundingClientRect(); return { left: b.left, right: b.right, top: b.top, bottom: b.bottom }; };
      const head = (label) => [...document.querySelectorAll('.sh-colheads .sh-head')].find((n) => n.textContent.trim() === label);
      return { box: r(document.querySelector('${colBox}')), d: r(head('D')), b: r(head('B')), b2: r(document.querySelector('.sh-cell[data-ref="B2"]')),
        gutter: r(document.querySelector('.sh-colheads > .sh-gutter.cols')),
        levels: [...document.querySelectorAll('.sh-ol-level[data-axis="col"]')].map((n) => n.textContent.trim()).join(',') };
    })()`);
    const colOk = Boolean(colSeen.box && colSeen.d && colSeen.gutter && colSeen.b && colSeen.b2)
      && near((colSeen.box.left + colSeen.box.right) / 2, (colSeen.d.left + colSeen.d.right) / 2, 1.5)
      && colSeen.box.bottom <= colSeen.d.top + 0.5 && near(colSeen.gutter.bottom, colSeen.b.top)
      && near(colSeen.b.left, colSeen.b2.left);
    await refresh();
    await capture(win, 'sheets-outline-columns.png');
    await js(`(() => { document.querySelector('${colBox}').click(); return 1; })()`);
    const colFolded = await until(() => js(`![...document.querySelectorAll('.sh-colheads .sh-head')].some((n) => n.textContent.trim() === 'B') && Boolean(document.querySelector('${colBox}.folded'))`), 'columns B and C folded', 5000).catch(() => false);
    await js(`(() => { document.querySelector('${colBox}').click(); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.sh-colheads .sh-head')].some((n) => n.textContent.trim() === 'B')`), 'columns back', 5000).catch(() => {});
    await js(`(() => { document.querySelector('.sh')?.focus(); return 1; })()`);
    await press(win.webContents, 'Left', { modifiers: ['shift', 'alt'] });
    const colGone = await until(() => js(`!document.querySelector('.sh-gutter')`), 'the column gutter to go', 5000).catch(() => false);
    check('sheets: whole columns group as columns — the gutter above the headings, the − box over column D, levels down the corner; it folds B and C away and Shift+Alt+Left ungroups',
      colGrouped === true && colOk && colSeen.levels === '1,2' && colFolded === true && colGone === true,
      `grouped ${colGrouped}, box ${JSON.stringify(colSeen.box)} vs D ${JSON.stringify(colSeen.d)}, levels ${colSeen.levels}, folded ${colFolded}, gone ${colGone}`);

    // Subtotal on the list: at each change in Region, Sum of Units and Sales.
    await js(`(() => { document.querySelector('.sh-cell[data-ref="B3"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-cell.active[data-ref="B3"]'))`), 'B3 active', 4000).catch(() => {});
    await clickIn('Subtotal');
    const dialog = await until(() => js(`Boolean(document.querySelector('.sh-sub-by')) && Boolean(document.querySelector('.sh-sub-col-3')?.checked)`), 'the Subtotal dialog', 5000).catch(() => false);
    const offered = await js(`[...document.querySelectorAll('.sh-sub-by option')].map((o) => o.textContent).join(',') + ' | ' + document.querySelector('.sh-sub-fn')?.value`);
    await js(`(() => { document.querySelector('.sh-sub-col-2').click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-sub-col-2')?.checked)`), 'Units ticked', 3000).catch(() => {});
    win.webContents.invalidate();
    await wait(500);
    await capture(win, 'sheets-subtotal-dialog.png');
    await js(`(() => { document.querySelector('.sh-sub-ok').click(); return 1; })()`);
    const cellText = (ref) => `(document.querySelector('.sh-cell[data-ref="${ref}"]')?.textContent || '')`;
    const totalled = await until(() => js(`${cellText('A11')} === 'Grand Total' && ${cellText('D11')} === '1370'`), 'the Grand Total', 8000).catch(() => false);
    const figures = await js(`({
      labels: ['A4', 'A6', 'A10', 'A11'].map((r) => r + '=' + (document.querySelector('.sh-cell[data-ref="' + r + '"]')?.textContent || '')),
      sales: ['D4', 'D6', 'D10', 'D11'].map((r) => document.querySelector('.sh-cell[data-ref="' + r + '"]')?.textContent || ''),
      units: ['C4', 'C6', 'C10', 'C11'].map((r) => document.querySelector('.sh-cell[data-ref="' + r + '"]')?.textContent || ''),
      bold: getComputedStyle(document.querySelector('.sh-cell[data-ref="A4"]')).fontWeight,
      levels: [...document.querySelectorAll('.sh-ol-level[data-axis="row"]')].map((b) => b.textContent.trim()).join(','),
      boxes: document.querySelectorAll('.sh-rowheads > .sh-gutter.rows .sh-ol-box').length,
    })`);
    check('sheets: Data → Subtotal adds a bold total under each region and a Grand Total — the figures read 550, 400, 420 and 1370 (units 5, 5, 7, 17)',
      dialog === true && /^Region,Rep,Units,Sales \| sum$/.test(offered) && totalled === true
        && figures.labels.join(' ') === 'A4=East Total A6=North Total A10=West Total A11=Grand Total'
        && figures.sales.join(',') === '550,400,420,1370' && figures.units.join(',') === '5,5,7,17' && Number(figures.bold) >= 600,
      `${offered}; ${JSON.stringify(figures)}`);
    check('sheets: the subtotals come with a three-level outline — buttons 1 2 3 and a box by every total',
      figures.levels === '1,2,3' && figures.boxes === 4, `levels ${figures.levels}, boxes ${figures.boxes}`);
    await refresh();
    await capture(win, 'sheets-subtotal.png');

    await js(`(() => { document.querySelector('.sh-ol-level[data-axis="row"][data-level="2"]').click(); return 1; })()`);
    const totalsOnly = await until(async () => (await heads()).slice(0, 6).join(',') === '1,4,6,10,11,12', 'level 2 to leave only the totals', 5000).catch(() => false);
    await refresh();
    await capture(win, 'sheets-subtotal-level2.png');
    check('sheets: level 2 leaves the header, the three totals and the Grand Total on screen', totalsOnly === true, (await heads()).slice(0, 8).join(','));
    await js(`(() => { document.querySelector('.sh-ol-level[data-axis="row"][data-level="3"]').click(); return 1; })()`);
    await until(async () => (await heads()).includes('2'), 'level 3', 5000).catch(() => {});

    // Remove All puts the list back.
    await clickIn('Subtotal');
    await until(() => js(`Boolean(document.querySelector('.sh-sub-remove:not(:disabled)'))`), 'Remove All', 5000).catch(() => {});
    await js(`(() => { document.querySelector('.sh-sub-remove').click(); return 1; })()`);
    const removed = await until(() => js(`${cellText('A7')} === 'West' && ${cellText('D7')} === '120' && !document.querySelector('.sh-gutter') && ${cellText('A8')} === ''`), 'the list as it was', 8000).catch(() => false);
    check('sheets: Subtotal → Remove All takes the totals and the outline away', removed === true, await js(`${cellText('A4')} + ' ' + ${cellText('A7')} + ' ' + ${cellText('A8')}`));

    // The pane unfrozen and the file saved: it is the fixture as it was.
    await tab('View');
    await until(() => js(`Boolean([...document.querySelectorAll('button')].find((b) => /Freeze Panes/.test(b.textContent)))`), 'the View tab', 4000).catch(() => {});
    await js(`(() => { [...document.querySelectorAll('button')].find((b) => /Freeze Panes/.test(b.textContent)).click(); return 1; })()`);
    await menuItem(/Unfreeze panes/);
    await until(() => js(`!document.querySelector('.sh-pin-rows')`), 'the pane unfrozen', 5000).catch(() => {});
    await wait(300);
    await clickIn('Save');
    const saved = await until(() => {
      try {
        const v = new SheetView(fs.readFileSync(file));
        return v.frozenPane().rows === 0 && v.displayValue(6, 0).text === 'West';
      } catch { return false; }
    }, 'the save to land', 8000).catch(() => false);
    const after = new SheetView(fs.readFileSync(file));
    const xml = after.pkg.text(after.workbook.partNameFor('Sales'));
    const same = LIST.every((row, r) => row.every((_, c) => after.displayValue(r, c).text === original[r][c])) && after.displayValue(7, 0).text === '';
    check('sheets: saved after Remove All, the workbook is the list it was — no outline, no SUBTOTAL, no frozen pane',
      saved === true && same && !/outlineLevel|SUBTOTAL|collapsed=|hidden=/.test(xml), `saved ${saved}, same ${same}, marks ${(xml.match(/outlineLevel|SUBTOTAL|collapsed=|hidden=/g) || []).join(',') || 'none'}`);

    const complaints = await errorsIn(win);
    check('sheets: the outline and subtotal checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the outline checks ran', false, err.message);
  }
}
