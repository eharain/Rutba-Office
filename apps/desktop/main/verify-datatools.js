// Worksheets: Data → Advanced (in place and copied), Clear, and Flash Fill.
//
// A list with a criteria range beside it: Advanced filters it in place by
// two OR'ed rows and Clear shows it all again; Advanced copies the rows
// under five units to another place, unique records only. Then Flash Fill:
// "Ada" typed beside "Lovelace, Ada" and Ctrl+E fills the other first
// names; "AL" typed in the next column and the ribbon's Flash Fill fills
// the initials. Every step is undone by Ctrl+Z at the end, and the file on
// disk was never written. Run alone with RUTBA_VERIFY_ONLY=datatools.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

const ROWS = [
  ['Region', 'Rep', 'Units', 'Sales', null, 'Region', 'Units'],
  ['East', 'Kim', 3, 300, null, 'East', '<5'],
  ['East', 'Lee', 2, 250, null, 'North'],
  ['North', 'Ann', 5, 400],
  ['West', 'Bo', 1, 90],
  ['Western', 'Cy', 4, 210],
  ['West', 'Bo', 1, 90],
  [],
  [],
  ['Name', 'First', 'Initials'],
  ['Lovelace, Ada'],
  ['Hopper, Grace'],
  ['Turing, Alan'],
  ['Babbage, Charles'],
];

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} where the fixture is written
 */
export async function verifyDataTools(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'datatools.xlsx');
  fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sales', rows: ROWS }] }));
  const bytes = fs.readFileSync(file);
  try {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');
    const model = () => doc.model({ id: session.id });
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="A14"]'))`), 'the grid', 8000);

    const clickIn = async (title) => {
      const find = `[...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    const tab = (name) => js(`(() => { [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const heads = () => js(`[...document.querySelectorAll('.sh-rowheads .sh-head')].map((n) => n.textContent.trim())`);
    const cell = (ref) => `(document.querySelector('.sh-cell[data-ref="${ref}"]')?.textContent || '')`;
    const setField = (sel, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(sel)});
      if (!el) return false;
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    })()`);
    // An empty cell has no node until it is active: the filled cell at the
    // start of its row is pressed, and the arrow keys walk across.
    const selectCell = async (ref) => {
      const [, letters, row] = /^([A-Z]+)(\d+)$/.exec(ref);
      const there = await js(`Boolean(document.querySelector('.sh-cell[data-ref="${ref}"]'))`);
      await js(`(() => { document.querySelector('.sh-cell[data-ref="${there ? ref : 'A' + row}"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
      if (!there) {
        await until(() => js(`document.querySelector('.sh-cell.active')?.dataset.ref === 'A${row}'`), 'A' + row + ' active', 4000).catch(() => {});
        await js(`(() => { document.querySelector('.sh')?.focus(); return 1; })()`);
        for (let i = 0; i < letters.charCodeAt(0) - 65; i++) await press(win.webContents, 'Right');
      }
      return until(() => js(`document.querySelector('.sh-cell.active')?.dataset.ref === '${ref}'`), ref + ' active', 4000).catch(() => false);
    };
    const typeIn = async (text) => {
      await js(`(() => { document.querySelector('.sh')?.focus(); return 1; })()`);
      for (const ch of text) {
        // A capital is Shift and the letter, as a keyboard sends it.
        const modifiers = /[A-Z]/.test(ch) ? ['shift'] : [];
        win.webContents.sendInputEvent({ type: 'keyDown', keyCode: ch, modifiers });
        win.webContents.sendInputEvent({ type: 'char', keyCode: ch, modifiers });
        win.webContents.sendInputEvent({ type: 'keyUp', keyCode: ch, modifiers });
        await wait(70);
      }
      await wait(150);
      await press(win.webContents, 'Return');
    };

    // Advanced, in place: Region East OR North, from the ranges typed in.
    await tab('Data');
    await clickIn('Advanced');
    const dialog = await until(() => js(`Boolean(document.querySelector('.sh-adv-criteria'))`), 'the Advanced Filter dialog', 5000).catch(() => false);
    const listOffered = await js(`document.querySelector('.sh-adv-list')?.value || ''`);
    await setField('.sh-adv-criteria', 'F1:F3');
    await until(() => js(`!document.querySelector('.sh-adv-ok')?.disabled`), 'OK', 3000).catch(() => {});
    await js(`(() => { document.querySelector('.sh-adv-ok').click(); return 1; })()`);
    const filtered = await until(async () => {
      const h = await heads();
      return h.includes('2') && h.includes('4') && !h.includes('5') && !h.includes('7') && h.includes('8');
    }, 'rows 5 to 7 to hide', 6000).catch(() => false);
    const toast = await js(`[...document.querySelectorAll('.rw-toast')].map((t) => t.textContent).join(' | ')`);
    check('sheets: Data → Advanced filters the list in place by two criteria rows, OR\'ed — East and North stay, the West rows hide',
      dialog === true && listOffered === 'A1:D7' && filtered === true && /3 rows of the list pass/.test(toast),
      `dialog ${dialog}, list ${listOffered}, filtered ${filtered}, headings ${(await heads()).slice(0, 9).join(' ')}, toast ${toast}`);
    win.webContents.invalidate();
    await wait(700);
    await capture(win, 'sheets-advanced-filter.png');

    await clickIn('Clear');
    const cleared = await until(async () => { const h = await heads(); return ['5', '6', '7'].every((x) => h.includes(x)); }, 'the rows to come back', 5000).catch(() => false);
    check('sheets: Data → Clear shows every row the advanced filter hid', cleared === true, (await heads()).slice(0, 9).join(' '));

    // Advanced, copied: Units under five to I1, unique records only.
    await clickIn('Advanced');
    await until(() => js(`Boolean(document.querySelector('.sh-adv-copy'))`), 'the dialog again', 5000).catch(() => {});
    const offeredAgain = await js(`document.querySelector('.sh-adv-criteria')?.value || ''`);
    await js(`(() => { document.querySelector('.sh-adv-copy').click(); document.querySelector('.sh-adv-unique').click(); return 1; })()`);
    await setField('.sh-adv-criteria', 'G1:G2');
    await setField('.sh-adv-to', 'I1');
    await until(() => js(`!document.querySelector('.sh-adv-ok')?.disabled && document.querySelector('.sh-adv-unique')?.checked`), 'OK', 3000).catch(() => {});
    win.webContents.invalidate();
    await wait(500);
    await capture(win, 'sheets-advanced-dialog.png');
    await js(`(() => { document.querySelector('.sh-adv-ok').click(); return 1; })()`);
    const copied = await until(() => js(`${cell('I1')} === 'Region' && ${cell('J5')} === 'Cy' && ${cell('L2')} === '300'`), 'the rows copied to I1', 6000).catch(() => false);
    const extract = await js(`[1, 2, 3, 4, 5, 6].map((r) => ['I', 'J', 'K', 'L'].map((c) => (document.querySelector('.sh-cell[data-ref="' + c + r + '"]')?.textContent || '')).join('|'))`);
    check('sheets: Data → Advanced copies the rows under five units to I1, headings first, the repeated Bo row once',
      offeredAgain === 'F1:F3' && copied === true && extract.join(' / ') === 'Region|Rep|Units|Sales / East|Kim|3|300 / East|Lee|2|250 / West|Bo|1|90 / Western|Cy|4|210 / |||',
      `offered ${offeredAgain}; ${extract.join(' / ')}`);
    win.webContents.invalidate();
    await wait(700);
    await capture(win, 'sheets-advanced-copy.png');

    // Flash Fill: "Ada" typed beside "Lovelace, Ada", then Ctrl+E.
    await selectCell('B11');
    await typeIn('Ada');
    await until(() => js(`${cell('B11')} === 'Ada'`), 'the example typed', 5000).catch(() => {});
    await selectCell('B11');
    await wait(300);
    await js(`(() => { document.querySelector('.sh')?.focus(); return 1; })()`);
    await press(win.webContents, 'e', { modifiers: ['control'] });
    const firsts = await until(() => js(`${cell('B12')} === 'Grace' && ${cell('B13')} === 'Alan' && ${cell('B14')} === 'Charles'`), 'the first names', 6000).catch(() => false);
    check('sheets: Ctrl+E after one example fills the first names out of "Last, First"', firsts === true, await js(`[${cell('B11')}, ${cell('B12')}, ${cell('B13')}, ${cell('B14')}].join(',')`));

    // Initials: "AL" typed, then Data → Flash Fill on the ribbon.
    await selectCell('C11');
    await typeIn('AL');
    await until(() => js(`${cell('C11')} === 'AL'`), 'the second example', 5000).catch(() => {});
    await selectCell('C11');
    await wait(300);
    await clickIn('Flash Fill');
    const initials = await until(() => js(`${cell('C12')} === 'GH' && ${cell('C13')} === 'AT' && ${cell('C14')} === 'CB'`), 'the initials', 6000).catch(() => false);
    check('sheets: Data → Flash Fill fills the initials from one example, first letter then last', initials === true, await js(`[${cell('C11')}, ${cell('C12')}, ${cell('C13')}, ${cell('C14')}].join(',')`));
    win.webContents.invalidate();
    await wait(700);
    await capture(win, 'sheets-flash-fill.png');

    const complaints = await errorsIn(win);
    check('sheets: the advanced filter and Flash Fill checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');

    // Everything undone, one Ctrl+Z at a time: the workbook is the fixture.
    await js(`(() => { document.querySelector('.sh')?.focus(); return 1; })()`);
    for (let i = 0; i < 12 && model().canUndo; i++) {
      await press(win.webContents, 'z', { modifiers: ['control'] });
      await wait(250);
    }
    const back = await until(() => js(`${cell('I1')} === '' && ${cell('B12')} === '' && ${cell('B11')} === '' && ${cell('C11')} === ''`), 'the fixture as it was', 6000).catch(() => false);
    const m = model();
    const names = (m.names || []).map((n) => n.name).join(',');
    check('sheets: Ctrl+Z takes every step back — the list, the copy, the examples and the fills — and the file on disk was never written',
      back === true && m.canUndo === false && !/_xlnm\./.test(names) && Buffer.compare(fs.readFileSync(file), bytes) === 0,
      `back ${back}, canUndo ${m.canUndo}, names ${names || 'none'}`);
  } catch (err) {
    check('sheets: the advanced filter and Flash Fill checks ran', false, err.message);
  }
}
