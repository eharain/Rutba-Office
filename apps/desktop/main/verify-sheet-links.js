// Worksheets: a note on a cell, a link on a cell, and a link put there.
//
// A workbook that carried notes showed "[object Object]" as their tooltip
// and nothing else; a workbook that carried links showed plain text. This
// opens one with a note and a link, reads what the grid makes of them,
// follows the link to the other sheet, puts a new link on a cell through
// the dialog, saves, and reads the file back through the engine. The
// address link is never clicked: a check run opens nothing outside the
// suite. Run alone with RUTBA_VERIFY_ONLY=links.

import fs from 'node:fs';
import { SheetView } from '@rutba/sheet-view';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture
 */
export async function verifySheetLinks(h, { file }) {
  const { open, check, until, wait, press, errorsIn, capture } = h;
  if (!file) return check('sheets: the notes workbook was made', false, 'no fixture');
  const clickIn = async (win, title) => {
    const find = `[...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
    await until(() => win.webContents.executeJavaScript(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
    return win.webContents.executeJavaScript(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
  };
  try {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.sh-cell').length > 4`), 'the grid', 8000);

    const shown = await js(`(() => {
      const noted = [...document.querySelectorAll('.sh-cell.noted')];
      const linked = [...document.querySelectorAll('.sh-cell.link')];
      const mark = noted[0] ? getComputedStyle(noted[0], '::after') : null;
      return { noted: noted.map((c) => c.dataset.ref + ':' + (c.dataset.tip || '')), linked: linked.map((c) => c.dataset.ref + ':' + (c.dataset.tip || '')),
        mark: mark ? mark.borderTopColor : null, underline: linked[0] ? getComputedStyle(linked[0]).textDecorationLine : null, chips: document.querySelector('.rw-status')?.textContent || '' };
    })()`);
    check('sheets: a cell with a note wears a red corner and tells the note on hover; a cell with a link is underlined and tells where it goes',
      shown.noted.length === 1 && /^B2:Kim Lee: Check this figure/.test(shown.noted[0]) && shown.mark === 'rgb(208, 54, 47)' && shown.linked.length === 1 && /^A4:The ledger row — Ledger!A2 \(Ctrl\+click to open\)$/.test(shown.linked[0]) && /underline/.test(shown.underline || '') && /1 note/.test(shown.chips),
      JSON.stringify(shown).slice(0, 300));

    // Ctrl+click follows the link to the ledger, on the other sheet.
    await js(`(() => { document.querySelector('.sh-cell.link').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, ctrlKey: true })); return 1; })()`);
    const followed = await until(() => js(`document.querySelector('.sh-tab.active')?.textContent === 'Ledger' && document.querySelector('.sh-cell.active')?.dataset.ref === 'A2'`), 'the link to be followed', 5000).catch(() => false);
    check('sheets: Ctrl+click on a link goes to the place it names, on the other sheet', followed === true, await js(`(document.querySelector('.sh-tab.active')?.textContent || '?') + ' ' + (document.querySelector('.sh-cell.active')?.dataset.ref || '?')`));

    // Back on Sales, a link is put on C2 through the dialog: Ctrl+K, an
    // address, a tip, Add link. The cell is underlined at once and the
    // status bar says where it goes.
    await js(`(() => { [...document.querySelectorAll('.sh-tab')].find((t) => t.textContent === 'Sales')?.click(); return 1; })()`);
    await until(() => js(`document.querySelector('.sh-tab.active')?.textContent === 'Sales'`), 'the Sales sheet', 4000);
    await js(`(() => { document.querySelector('.sh-cell[data-ref="C2"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); return 1; })()`);
    await until(() => js(`document.querySelector('.sh-cell.active')?.dataset.ref === 'C2'`), 'C2 to be active', 4000);
    await press(win.webContents, 'k', { modifiers: ['control'] });
    const dialog = await until(() => js(`Boolean(document.querySelector('.rw-dialog .sh-link-to'))`), 'the Link dialog', 4000).catch(() => false);
    await js(`(() => {
      const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
      const to = document.querySelector('.rw-dialog .sh-link-to'); set.call(to, 'rutba.io/office'); to.dispatchEvent(new Event('input', { bubbles: true }));
      const tip = document.querySelector('.rw-dialog .sh-link-tip'); set.call(tip, 'The suite'); tip.dispatchEvent(new Event('input', { bubbles: true }));
      return 1;
    })()`);
    const says = await until(() => js(`/Opens https:\\/\\/rutba\\.io\\/office/.test(document.querySelector('.rw-dialog .sh-link-says')?.textContent || '')`), 'the dialog to read the address', 3000).catch(() => false);
    await js(`(() => { document.querySelector('.rw-dialog .sh-link-ok')?.click(); return 1; })()`);
    const linked = await until(() => js(`document.querySelector('.sh-cell[data-ref="C2"]')?.classList.contains('link') && /rutba\\.io\\/office/.test(document.querySelector('.rw-status')?.textContent || '')`), 'the link on C2', 5000).catch(() => false);
    check('sheets: Ctrl+K puts a link on the cell — a bare domain read as an address — and the status bar says where it goes', dialog === true && says === true && linked === true, `dialog ${dialog}, read ${says}, linked ${linked}: ${await js(`document.querySelector('.sh-cell[data-ref="C2"]')?.dataset.tip || 'no tip'`)}`);
    await wait(400);
    await capture(win, 'sheet-links.png');

    // Saved, the file carries it as Excel would: an external relationship.
    await clickIn(win, 'Save');
    const saved = await until(() => {
      try {
        const links = SheetView.open(fs.readFileSync(file)).workbook.hyperlinks('Sales');
        return links.some((l) => l.ref === 'C2' && l.href === 'https://rutba.io/office' && l.tooltip === 'The suite');
      } catch { return false; }
    }, 'the save to land', 8000).catch(() => false);
    const reread = SheetView.open(fs.readFileSync(file));
    const rels = reread.pkg.text(reread.workbook.partNameFor('Sales').replace(/worksheets\//, 'worksheets/_rels/') + '.rels');
    check('sheets: the saved file carries the link as an external relationship, beside the note and the link it had', saved === true && /rutba\.io\/office" TargetMode="External"/.test(rels) && reread.workbook.hyperlinks('Sales').length === 2 && reread.comments.get('Sales')?.size === 1, `${reread.workbook.hyperlinks('Sales').map((l) => l.ref + '→' + (l.href || l.location)).join(', ')}; external: ${/TargetMode="External"/.test(rels)}`);

    // Ctrl+K on a linked cell offers to change or remove it; Remove takes it off.
    await press(win.webContents, 'k', { modifiers: ['control'] });
    await until(() => js(`Boolean(document.querySelector('.rw-dialog .sh-link-remove'))`), 'the dialog with Remove', 4000).catch(() => {});
    const offered = await js(`(document.querySelector('.rw-dialog .sh-link-to')?.value || '') + ' | ' + (document.querySelector('.rw-dialog .rw-dialog-title, .rw-dialog h2, .rw-dialog header')?.textContent || '')`);
    await js(`(() => { document.querySelector('.rw-dialog .sh-link-remove')?.click(); return 1; })()`);
    const removed = await until(() => js(`!document.querySelector('.sh-cell[data-ref="C2"]')?.classList.contains('link') && !document.querySelector('.rw-dialog')`), 'the link to go', 5000).catch(() => false);
    check('sheets: Ctrl+K on a linked cell shows the link it has, and Remove takes it off', /rutba\.io\/office/.test(offered) && removed === true, `offered ${JSON.stringify(offered)}, removed ${removed}`);

    // A note put on a cell through the dialog wears the corner and tells
    // itself; the saved file carries it as Excel does — the comments part,
    // the VML box, and the sheet pointing at both; Delete takes it off.
    await js(`(() => { document.querySelector('.sh-cell[data-ref="C3"]')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="C3"].active'))`), 'C3 active', 4000).catch(() => {});
    await press(win.webContents, 'F2', { modifiers: ['shift'] });
    const noteDialog = await until(() => js(`Boolean(document.querySelector('.sh-note-text'))`), 'the note dialog', 4000).catch(() => false);
    await js(`(() => { const set = (sel, v) => { const el = document.querySelector(sel); const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype; Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, v); el.dispatchEvent(new Event('input', { bubbles: true })); }; set('.sh-note-text', 'Confirm with the supplier.'); set('.sh-note-author', 'Sam Roe'); return 1; })()`);
    await until(() => js(`!document.querySelector('.sh-note-ok')?.disabled`), 'the OK button', 3000).catch(() => {});
    await js(`(() => { document.querySelector('.sh-note-ok')?.click(); return 1; })()`);
    const noted = await until(() => js(`(() => { const c = document.querySelector('.sh-cell[data-ref="C3"]'); return Boolean(c) && c.classList.contains('noted') && /Sam Roe: Confirm with the supplier\\./.test(c.dataset.tip || '') ? c.dataset.tip : false; })()`), 'the noted cell', 5000).catch(() => false);
    const counted = await js(`/2 notes/.test(document.querySelector('.rw-status')?.textContent || '')`);
    check('sheets: Shift+F2 puts a note on the cell through the dialog — it wears the corner, tells itself, and the status bar counts it', noteDialog === true && noted !== false && counted === true, `dialog ${noteDialog}, tip ${JSON.stringify(noted)}, counted ${counted}`);

    await clickIn(win, 'Save');
    await until(() => {
      try {
        return SheetView.open(fs.readFileSync(file)).workbook.comments('Sales').some((n) => n.ref === 'C3');
      } catch { return false; }
    }, 'the save to land', 8000).catch(() => false);
    const withNote = SheetView.open(fs.readFileSync(file));
    const notes = withNote.workbook.comments('Sales');
    const sheetXml = withNote.pkg.text(withNote.workbook.partNameFor('Sales'));
    const vmlRel = withNote.pkg.rels(withNote.workbook.partNameFor('Sales')).find((r) => /vmlDrawing$/.test(r.Type));
    const vmlName = vmlRel ? 'xl/' + vmlRel.Target.replace(/^\.\.\//, '') : null;
    const vml = vmlName && withNote.pkg.has(vmlName) ? withNote.pkg.text(vmlName) : '';
    const legacyId = /<legacyDrawing r:id="([^"]+)"/.exec(sheetXml)?.[1];
    check('sheets: the saved file carries the note as Excel does — the comments part, the VML box, and the sheet pointing at both',
      notes.some((n) => n.ref === 'C3' && n.author === 'Sam Roe' && n.text === 'Confirm with the supplier.') && notes.some((n) => n.ref === 'B2' && n.author === 'Kim Lee')
        && Boolean(vmlRel) && legacyId === vmlRel?.Id && /ObjectType="Note"/.test(vml) && /<x:Row>2<\/x:Row>\s*<x:Column>2<\/x:Column>/.test(vml)
        && withNote.pkg.contentTypes().defaults.get('vml') === 'application/vnd.openxmlformats-officedocument.vmlDrawing',
      `notes ${JSON.stringify(notes)}; legacyDrawing ${legacyId} vs rel ${vmlRel?.Id}; box ${/ObjectType="Note"/.test(vml)}`);

    // The keys go to the grid, which the dialog took them from.
    await js(`(() => { document.querySelector('.sh')?.focus(); return 1; })()`);
    await press(win.webContents, 'F2', { modifiers: ['shift'] });
    await until(() => js(`Boolean(document.querySelector('.sh-note-remove'))`), 'the note dialog with Delete', 4000).catch(() => {});
    const offeredNote = await js(`document.querySelector('.sh-note-text')?.value || ''`);
    await js(`(() => { document.querySelector('.sh-note-remove')?.click(); return 1; })()`);
    const unnoted = await until(() => js(`(() => { const c = document.querySelector('.sh-cell[data-ref="C3"]'); return Boolean(c) && !c.classList.contains('noted'); })()`), 'the note to go', 4000).catch(() => false);
    check('sheets: Shift+F2 on a noted cell shows the note, and Delete takes it off', offeredNote === 'Confirm with the supplier.' && unnoted === true, `offered ${JSON.stringify(offeredNote)}, gone ${unnoted}`);

    const complaints = await errorsIn(win);
    check('sheets: the link and note checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the link and note checks ran', false, err.message);
  }
  return undefined;
}
