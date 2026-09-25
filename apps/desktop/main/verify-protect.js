// Worksheets: Review → Protect Workbook and Allow Edit Ranges.
//
// A sheet is hidden and shown again, and moved, from its tab's menu. Protect
// Workbook with a password (typed twice) locks the structure: the tab menu
// greys Insert, Delete, Rename, Move, Hide and Unhide, the + is greyed, a
// padlock sits by the tabs, and an added sheet is refused in Excel's words.
// Unprotect asks for the password; a wrong one is refused in Excel's words.
// Allow Edit Ranges makes a range with a password from the selection; the
// sheet is protected; typing into the range asks for the range's password
// once, then the typed key goes on into the cell; outside it, a locked cell
// is refused as before. Saved, the file carries workbookProtection and the
// protected range with SHA-512 hashes. Run alone with RUTBA_VERIFY_ONLY=protect.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { hashPassword } from '@rutba/ooxml/protection';

const ROWS = [
  ['Item', 'Qty', 'Price'],
  ['Pens', 4, 1.5],
  ['Ink', 2, 6],
  ['Paper', 10, 4.25],
];

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} where the fixture is written
 */
export async function verifyProtect(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'protect.xlsx');
  fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Orders', rows: ROWS }, { name: 'Summary', rows: [['=SUM(Orders!B2:B4)']] }, { name: 'Notes', rows: [['n']] }] }));
  try {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');
    const model = () => doc.model({ id: session.id });
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="C4"]'))`), 'the grid', 8000);

    const clickIn = async (title) => {
      const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    const tab = (name) => js(`(() => { [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const setField = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no field ' + ${JSON.stringify(selector)};
      const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
      setter.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return 'set';
    })()`);
    const clickSel = (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'no ' + ${JSON.stringify(selector)}; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    const tabMenu = async (sheet) => {
      await js(`(() => { const t = [...document.querySelectorAll('.sh-tab')].find((n) => n.textContent.trim() === ${JSON.stringify(sheet)}); t.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: t.getBoundingClientRect().left + 10, clientY: t.getBoundingClientRect().top - 4 })); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the tab menu', 4000).catch(() => {});
      return js(`[...document.querySelectorAll('.rw-menu button')].map((b) => ({ label: b.textContent.trim(), disabled: b.disabled }))`);
    };
    const menuPick = (pattern) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => ${pattern}.test(b.textContent.trim())); if (!b) return 'no item'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    const closeMenu = () => js(`(() => { window.dispatchEvent(new MouseEvent('mousedown')); return 1; })()`);
    const tabNames = () => js(`[...document.querySelectorAll('.sh-tab:not(.sh-tab-add)')].map((t) => t.textContent.trim())`);
    const select = async (ref) => {
      await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="${ref}"]'))`), ref, 4000).catch(() => {});
      await js(`(() => { document.querySelector('.sh-cell[data-ref="${ref}"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
      return until(() => js(`document.querySelector('.sh-cell.active')?.dataset.ref === '${ref}'`), ref + ' active', 4000).catch(() => false);
    };
    const typeText = async (text) => {
      for (const ch of text) {
        win.webContents.sendInputEvent({ type: 'keyDown', keyCode: ch });
        win.webContents.sendInputEvent({ type: 'char', keyCode: ch });
        win.webContents.sendInputEvent({ type: 'keyUp', keyCode: ch });
        await wait(70);
      }
    };
    const refresh = async () => { win.webContents.invalidate(); await wait(700); };
    const cellText = (ref) => model().cells?.find((c) => c.ref === ref)?.text ?? '';

    // ── The tab's Hide, Unhide and Move right, before anything is locked.
    const menu1 = await tabMenu('Notes');
    const hid = await menuPick(/^Hide$/);
    const gone = await until(async () => !(await tabNames()).includes('Notes'), 'the Notes tab to go', 4000).catch(() => false);
    await tabMenu('Orders');
    const unhid = await menuPick(/^Unhide "Notes"$/);
    const back = await until(async () => (await tabNames()).includes('Notes'), 'the Notes tab to come back', 4000).catch(() => false);
    await tabMenu('Orders');
    const moved = await menuPick(/^Move right$/);
    const order = await until(async () => (await tabNames()).join(',') === 'Summary,Orders,Notes', 'the tabs reordered', 4000).catch(() => false);
    check('sheets: a tab\'s menu hides a sheet and Unhide brings it back; Move right moves it one place, in the file\'s own order',
      menu1.some((i) => i.label === 'Hide' && !i.disabled) && hid === 'clicked' && gone === true && unhid === 'clicked' && back === true && moved === 'clicked' && order === true && model().sheets.join(',') === 'Summary,Orders,Notes',
      `${hid} ${unhid} ${moved}; tabs ${(await tabNames()).join(',')}`);
    await js(`(() => { [...document.querySelectorAll('.sh-tab')].find((n) => n.textContent.trim() === 'Orders').click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="C4"]'))`), 'Orders again', 4000).catch(() => {});

    // ── Review → Protect Workbook with a password, typed twice.
    await tab('Review');
    const pw = await clickIn('Protect Workbook');
    const asked = await until(() => js(`Boolean(document.querySelector('.sh-protect-password'))`), 'the Protect Workbook dialog', 4000).catch(() => false);
    await setField('.sh-protect-password', 'Budget2026');
    await until(() => js(`Boolean(document.querySelector('.sh-protect-again'))`), 'the confirm box', 3000).catch(() => {});
    await setField('.sh-protect-again', 'Budget2026');
    await until(() => js(`document.querySelector('.sh-protect-ok')?.disabled === false`), 'OK to be pressable', 3000).catch(() => {});
    await refresh();
    await refresh();
    await capture(win, 'sheets-protect-workbook-dialog.png');
    await clickSel('.sh-protect-ok');
    const locked = await until(() => model().workbookProtection?.structure === true, 'the structure locked', 8000).catch(() => false);
    const label = await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Unprotect Workbook' && b.getAttribute('aria-pressed') === 'true'))`), 'the pressed Unprotect Workbook button', 4000).catch(() => false);
    const items = await tabMenu('Orders');
    const greyed = ['Insert sheet', 'Delete sheet…', 'Rename sheet…', 'Move left', 'Move right', 'Hide'].every((l) => items.find((i) => i.label === l)?.disabled === true);
    const plus = await js(`document.querySelector('.sh-tab-add')?.disabled === true && Boolean(document.querySelector('.sh-tabs-lock'))`);
    await refresh();
    await capture(win, 'sheets-protect-workbook-menu.png');
    await closeMenu();
    check('sheets: Review → Protect Workbook with a password locks the structure — the button pressed, the tab menu greys Insert, Delete, Rename, Move and Hide, the + greyed and a padlock by the tabs',
      pw === 'clicked' && asked === true && locked === true && model().workbookProtection.hasPassword === true && label === true && greyed && plus === true,
      `${pw}; locked ${locked}; label ${label}; menu ${JSON.stringify(items)}; plus ${plus}`);

    let refused = '';
    try { doc.apply({ id: session.id, ops: [{ op: 'addSheet' }] }); } catch (err) { refused = err.message; }
    check('sheets: while the workbook is protected a new sheet is refused in Excel\'s words, and nothing changes',
      /^Workbook is protected and cannot be changed\.$/.test(refused) && model().sheets.length === 3, refused);

    // ── Unprotect Workbook asks for the password; a wrong one is refused.
    await clickIn('Unprotect Workbook');
    await until(() => js(`Boolean(document.querySelector('.sh-password'))`), 'the password dialog', 4000).catch(() => {});
    await setField('.sh-password', 'budget2026');
    await clickSel('.sh-password-ok');
    await until(() => js(`Boolean(document.querySelector('.sh-password-error'))`), 'the refusal', 8000).catch(() => false);
    const wrong = await js(`document.querySelector('.sh-password-error')?.textContent || ''`);
    await refresh();
    await capture(win, 'sheets-protect-wrong-password.png');
    await setField('.sh-password', 'Budget2026');
    await clickSel('.sh-password-ok');
    const unlocked = await until(() => model().workbookProtection?.structure === false, 'the structure unlocked', 8000).catch(() => false);
    const closed = await until(() => js(`!document.querySelector('.sh-password')`), 'the dialog gone', 4000).catch(() => false);
    check('sheets: Unprotect Workbook asks for the password — a wrong one is refused in Excel\'s words, the right one unlocks the structure',
      /^The password you supplied is not correct\./.test(wrong) && unlocked === true && closed === true, `wrong ${JSON.stringify(wrong)}; unlocked ${unlocked}`);

    // Locked again (no password) for the saved file's sake further down.
    await clickIn('Protect Workbook');
    await until(() => js(`Boolean(document.querySelector('.sh-protect-ok'))`), 'the dialog', 4000).catch(() => {});
    await setField('.sh-protect-password', 'Budget2026');
    await until(() => js(`Boolean(document.querySelector('.sh-protect-again'))`), 'the confirm box', 3000).catch(() => {});
    await setField('.sh-protect-again', 'Budget2026');
    await clickSel('.sh-protect-ok');
    await until(() => model().workbookProtection?.structure === true, 'locked again', 8000).catch(() => {});

    // ── Allow Edit Ranges: a range with a password, from the selection B2:B4.
    await select('B2');
    await js(`(() => { document.querySelector('.sh-cell[data-ref="B4"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0, shiftKey: true })); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sh-cell.sel').length >= 3`), 'B2:B4 selected', 4000).catch(() => {});
    const opened = await clickIn('Allow Edit Ranges');
    await until(() => js(`Boolean(document.querySelector('.sh-ranges-new'))`), 'the Allow Edit Ranges dialog', 4000).catch(() => {});
    await clickSel('.sh-ranges-new');
    await until(() => js(`Boolean(document.querySelector('.sh-range-title'))`), 'the New Range form', 4000).catch(() => {});
    const prefilled = await js(`document.querySelector('.sh-range-ref')?.value || ''`);
    await setField('.sh-range-title', 'Quantities');
    await setField('.sh-range-password', 'ledger');
    await until(() => js(`Boolean(document.querySelector('.sh-range-again'))`), 'the confirm box', 3000).catch(() => {});
    await setField('.sh-range-again', 'ledger');
    await refresh();
    await capture(win, 'sheets-edit-range-new.png');
    await clickSel('.sh-range-ok');
    const listed = await until(() => js(`Boolean(document.querySelector('.sh-ranges-row[data-title="Quantities"]'))`), 'the range in the list', 8000).catch(() => false);
    const row = await js(`document.querySelector('.sh-ranges-row[data-title="Quantities"]')?.textContent || ''`);
    await refresh();
    await capture(win, 'sheets-edit-ranges.png');
    check('sheets: Review → Allow Edit Ranges → New makes a range from the selection with a password, listed with its cells',
      opened === 'clicked' && /^B2:B4$/.test(prefilled) && listed === true && /Quantities/.test(row) && /B2:B4/.test(row)
        && JSON.stringify(model().editRanges) === JSON.stringify([{ title: 'Quantities', ref: 'B2:B4', hasPassword: true }]),
      `${opened}; prefilled ${prefilled}; row ${JSON.stringify(row)}; ranges ${JSON.stringify(model().editRanges)}`);

    // Protect Sheet… from the dialog, no password.
    await clickSel('.sh-ranges-protect');
    await until(() => js(`Boolean(document.querySelector('.sh-protect-ok'))`), 'the Protect Sheet dialog', 4000).catch(() => {});
    await clickSel('.sh-protect-ok');
    const sheetLocked = await until(() => model().protection?.sheet === true, 'the sheet protected', 8000).catch(() => false);

    // Typing into B3: Unlock Range asks for the password, then the key goes on.
    await select('B3');
    await js(`document.querySelector('.sh')?.focus(), 'ok'`);
    await typeText('7');
    const unlockAsked = await until(() => js(`document.querySelector('.rw-dialog-head')?.textContent === 'Unlock Range'`), 'Unlock Range', 6000).catch(() => false);
    const message = await js(`document.querySelector('.rw-dialog .sh-protect-note')?.textContent || ''`);
    await refresh();
    await capture(win, 'sheets-unlock-range.png');
    await setField('.sh-password', 'ledger');
    await clickSel('.sh-password-ok');
    const editing = await until(() => js(`document.querySelector('.sh-editor')?.value === '7'`), 'the edit to start with the 7', 8000).catch(() => false);
    await typeText('5');
    await press(win.webContents, 'Return', { char: true });
    const took = await until(() => cellText('B3') === '75', 'B3 to take 75', 6000).catch(() => false);
    check('sheets: typing into a password range on the protected sheet asks for its password once (Unlock Range), then the key typed goes into the cell',
      sheetLocked === true && unlockAsked === true && /Quantities/.test(message) && editing === true && took === true,
      `protected ${sheetLocked}; asked ${unlockAsked}; editing ${editing}; B3 ${cellText('B3')}`);

    // A locked cell outside the range is refused as before; B4 in the range takes edits without asking again.
    await select('A3');
    await js(`document.querySelector('.sh')?.focus(), 'ok'`);
    await typeText('X');
    await press(win.webContents, 'Return', { char: true });
    await wait(600);
    await select('B4');
    await js(`document.querySelector('.sh')?.focus(), 'ok'`);
    await typeText('9');
    await press(win.webContents, 'Return', { char: true });
    const b4 = await until(() => cellText('B4') === '9', 'B4 to take 9', 6000).catch(() => false);
    const noAsk = await js(`!document.querySelector('.sh-password')`);
    check('sheets: outside the range a locked cell is refused as before; the rest of the range takes edits without asking again',
      cellText('A3') === 'Ink' && b4 === true && noAsk === true, `A3 ${cellText('A3')}; B4 ${cellText('B4')}; asked again ${!noAsk}`);

    // Saved: the workbook's protection and the range, hashed as Excel hashes them.
    await wait(300);
    doc.save({ id: session.id });
    const pkg = OoxmlPackage.read(fs.readFileSync(file));
    const wb = pkg.text('xl/workbook.xml');
    const wp = /<workbookProtection\b[^>]*\/>/.exec(wb)?.[0] || '';
    const salt = /workbookSaltValue="([^"]+)"/.exec(wp)?.[1] || '';
    const hash = /workbookHashValue="([^"]+)"/.exec(wp)?.[1] || '';
    const sheetPart = pkg.partNames().find((n) => /worksheets\/sheet1\.xml$/.test(n));
    const sx = pkg.text(sheetPart);
    const pr = /<protectedRange\b[^>]*\/>/.exec(sx)?.[0] || '';
    check('sheets: the saved workbook carries workbookProtection (SHA-512, salt, 100,000 rounds, the hash of the password) and the protected range with its own hash, after sheetProtection',
      /workbookAlgorithmName="SHA-512"/.test(wp) && /lockStructure="1"/.test(wp) && /workbookSpinCount="100000"/.test(wp) && hash === hashPassword('Budget2026', { salt, spinCount: 100000 })
        && /sqref="B2:B4"/.test(pr) && /name="Quantities"/.test(pr) && /algorithmName="SHA-512"/.test(pr) && sx.indexOf('<sheetProtection') < sx.indexOf('<protectedRanges'),
      `${wp.slice(0, 160)}…; ${pr.slice(0, 120)}…`);

    // The refusal of A3 above is the point of that check; anything else is not.
    const complaints = (await errorsIn(win)).filter((t) => !/A3 is locked/.test(t));
    check('sheets: the protection checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the protection checks ran', false, err.message);
  }
}
