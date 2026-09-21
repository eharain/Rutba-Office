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
import { readPageSetup } from '@rutba/sheet-view/print';

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

    // Home → Format as Table: from A1, the block of data round it (A1:C4)
    // becomes a table — the header row and the banding painted at once, and
    // the saved file carrying the table as Excel keeps one.
    await js(`(() => { const c = document.querySelector('.sh-cell[data-ref="A1"]'); c?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
    await wait(200);
    const clicked = await clickIn(win, 'Format as Table');
    const menuShown = await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => /Medium, banded rows/.test(b.textContent)))`), 'the table styles', 3000).catch(() => false);
    await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => /Medium, banded rows/.test(b.textContent))?.click(); return 1; })()`);
    const paintRead = `(() => {
      const bg = (ref) => { const el = document.querySelector('.sh-cell[data-ref="' + ref + '"]'); return el ? getComputedStyle(el).backgroundColor : ''; };
      const clear = (v) => !v || v === 'rgba(0, 0, 0, 0)' || v === 'transparent' || v === 'rgb(255, 255, 255)';
      const head = bg('A1'); const band = bg('A3'); const plain = bg('A2');
      if (clear(head) || clear(band)) return false;
      const weight = getComputedStyle(document.querySelector('.sh-cell[data-ref="A1"]')).fontWeight;
      return { head, band, plainClear: clear(plain), weight };
    })()`;
    // `until` answers true, not what it polled: read the paint again once it is there.
    const paintOk = await until(() => js(paintRead), 'the table paint', 5000).catch(() => false);
    const painted = paintOk === true ? await js(paintRead) : false;
    check('sheets: Home → Format as Table makes the block round the cell a table — the header row and the banding painted at once',
      painted !== false && painted.plainClear === true && Number(painted.weight) >= 600, `${clicked}; menu ${menuShown}; ${JSON.stringify(painted)}`);
    await capture(win, 'sheets-table.png');

    await clickIn(win, 'Save');
    await until(() => {
      try { return SheetView.open(fs.readFileSync(file)).workbook.tables().length > 0; } catch { return false; }
    }, 'the table to land in the file', 8000).catch(() => false);
    const withTable = SheetView.open(fs.readFileSync(file));
    const tables = withTable.workbook.tables();
    const tableSheetXml = withTable.pkg.text(withTable.workbook.partNameFor('Sales'));
    check('sheets: the file carries the table as Excel keeps one — the part with its style and columns, the sheet\'s rels and tableParts',
      tables.length === 1 && tables[0].sheet === 'Sales' && tables[0].ref === 'A1:C4' && tables[0].styleName === 'TableStyleMedium2' && tables[0].showRowStripes === true
        && tables[0].columns.join(',') === 'Region,Q1,Q2' && withTable.pkg.has(tables[0].part) && /<tableParts count="1"><tablePart r:id="[^"]+"\/><\/tableParts>/.test(tableSheetXml),
      `${tables.length} table(s): ${JSON.stringify(tables.map((t) => [t.name, t.ref, t.styleName, t.columns]))}`);

    // Page Layout → Print Area: B2:C3, selected with Shift and the arrows,
    // becomes what prints — the sheet's _xlnm.Print_Area, as the page setup
    // reads it — and Clear takes it away.
    await js(`(() => { document.querySelector('.sh-cell[data-ref="B2"]')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
    await until(() => js(`document.querySelector('.sh-cell.active')?.dataset.ref === 'B2'`), 'B2 to be active', 4000).catch(() => {});
    await js(`document.querySelector('.sh')?.focus(), 'focused'`);
    await wait(150);
    // Each extension is waited for: a key sent before the grid holds focus is lost.
    const selectedCells = () => js(`document.querySelectorAll('.sh-cell.sel').length`);
    for (let tries = 0; tries < 3 && (await selectedCells()) < 2; tries++) {
      await press(win.webContents, 'Down', { modifiers: ['shift'] });
      await until(async () => (await selectedCells()) >= 2, 'the selection to reach B3', 1200).catch(() => {});
    }
    for (let tries = 0; tries < 3 && (await selectedCells()) < 4; tries++) {
      await press(win.webContents, 'Right', { modifiers: ['shift'] });
      await until(async () => (await selectedCells()) >= 4, 'the selection to reach C3', 1200).catch(() => {});
    }
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Page Layout')?.click(), 'tab'`);
    await wait(200);
    const areaRead = `(async () => { const all = await window.rutbaOffice.doc.sessions({}); const mine = all.filter((s) => s.kind === 'sheet').pop(); const s = await window.rutbaOffice.doc.pageSetup({ id: mine.id }); return s.area || ''; })()`;
    const pickArea = async (label) => {
      await clickIn(win, 'Print Area');
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.includes(${JSON.stringify(label)})))`), 'the print area menu', 3000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.includes(${JSON.stringify(label)})); if (!b) return 'no item'; b.click(); return 'picked'; })()`);
    };
    const picked = await pickArea('Set print area');
    const areaSet = await until(async () => (await js(areaRead)) === 'B2:C3', 'the print area', 5000).catch(() => false);
    const said = await js(`[...document.querySelectorAll('.rw-toast, [class*="toast"]')].map((t) => t.textContent.trim()).join(' | ')`);
    const selectedNow = await js(`(() => { const a = document.querySelector('.sh-cell.active'); return (a ? a.dataset.ref : '?') + ' selected ' + document.querySelectorAll('.sh-cell.sel').length; })()`);
    check('sheets: Page Layout → Print Area makes the selection what prints, kept as the sheet\'s print area', picked === 'picked' && areaSet === true, `${picked}; area ${JSON.stringify(await js(areaRead))}; toast ${JSON.stringify(said)}; ${selectedNow}`);
    const pickedClear = await pickArea('Clear print area');
    const cleared = await until(async () => (await js(areaRead)) === '', 'the print area to clear', 5000).catch(() => false);
    check('sheets: Clear print area takes it away', pickedClear === 'picked' && cleared === true, `${pickedClear}; area ${JSON.stringify(await js(areaRead))}`);

    // Home → Increase indent twice, then Text orientation → Rotate text up,
    // on A2: the words move in by two of Excel's indent units and stand on
    // end (the row growing to hold them), and the saved file keeps both in
    // the cell's alignment.
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Home')?.click(), 'tab'`);
    await js(`(() => { document.querySelector('.sh-cell[data-ref="A2"]')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
    await until(() => js(`document.querySelector('.sh-cell.active')?.dataset.ref === 'A2'`), 'A2 to be active', 4000).catch(() => {});
    const padRead = `(() => { const c = document.querySelector('.sh-cell[data-ref="A2"]'); return c ? parseFloat(getComputedStyle(c).paddingLeft) : -1; })()`;
    const padBefore = await js(padRead);
    const clickedIn = await clickIn(win, 'Increase indent');
    await until(async () => (await js(padRead)) > padBefore, 'the first indent', 4000).catch(() => {});
    const padOne = await js(padRead);
    await clickIn(win, 'Increase indent');
    await until(async () => (await js(padRead)) > padOne, 'the second indent', 4000).catch(() => {});
    const padTwo = await js(padRead);
    check('sheets: Home → Increase indent moves the words in the cell in by one Excel indent unit a press',
      clickedIn === 'clicked' && padOne > padBefore && padTwo > padOne && Math.round(padTwo - padBefore) === 18, `${clickedIn}; padding ${padBefore} → ${padOne} → ${padTwo}`);

    const heightBefore = await js(`document.querySelector('.sh-cell[data-ref="A2"]')?.getBoundingClientRect().height || 0`);
    await clickIn(win, 'Text orientation');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.includes('Rotate text up')))`), 'the orientation menu', 3000).catch(() => {});
    const pickedUp = await js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.includes('Rotate text up')); if (!b) return 'no item'; b.click(); return 'picked'; })()`);
    // rotate(-90deg) is the matrix (0, -1, 1, 0, …), give or take a rounding.
    const turnRead = `(() => { const s = document.querySelector('.sh-cell[data-ref="A2"] .sh-rot'); if (!s) return null; const m = /matrix\\(([^)]*)\\)/.exec(getComputedStyle(s).transform); return m ? m[1].split(',').slice(0, 4).map((n) => Math.round(Number(n) * 1000) / 1000).join(',') : getComputedStyle(s).transform; })()`;
    const turned = await until(async () => (await js(turnRead)) === '0,-1,1,0', 'the words to stand on end', 4000).catch(() => false);
    const grown = await until(async () => (await js(`document.querySelector('.sh-cell[data-ref="A2"]')?.getBoundingClientRect().height || 0`)) > heightBefore + 10, 'the row to grow', 4000).catch(() => false);
    const heightAfter = await js(`document.querySelector('.sh-cell[data-ref="A2"]')?.getBoundingClientRect().height || 0`);
    check('sheets: Text orientation → Rotate text up stands the words in the cell on end, and the row grows to hold them',
      pickedUp === 'picked' && turned === true && grown === true, `${pickedUp}; turn ${JSON.stringify(await js(turnRead))}; row ${heightBefore} → ${heightAfter}`);
    await capture(win, 'sheets-orientation.png');

    await clickIn(win, 'Save');
    const stateInFile = () => { const v = SheetView.open(fs.readFileSync(file)); v.select(1, 0); return v.formatState(); };
    await until(() => { try { return stateInFile().rotation === 90; } catch { return false; } }, 'the rotation to land in the file', 8000).catch(() => false);
    const kept = stateInFile();
    check('sheets: the file keeps the indent and the rotation in the alignment of the cell, as Excel reads them',
      kept.indent === 2 && kept.rotation === 90 && kept.align === 'left', JSON.stringify({ indent: kept.indent, rotation: kept.rotation, align: kept.align }));

    // Page Layout → Orientation, Print Titles and Width write the file's
    // own page setup — the one the print dialog and Excel start from — and
    // the saved file keeps all three.
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Page Layout')?.click(), 'tab'`);
    await wait(200);
    const setupRead = `(async () => { const all = await window.rutbaOffice.doc.sessions({}); const mine = all.filter((s) => s.kind === 'sheet').pop(); const s = await window.rutbaOffice.doc.pageSetup({ id: mine.id }); return [s.orientation, s.fit, s.repeatRows, s.paper].join(' '); })()`;
    const pickPage = async (button, label) => {
      await clickIn(win, button);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith(${JSON.stringify(label)})))`), `the ${button} menu`, 3000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith(${JSON.stringify(label)})); if (!b) return 'no item'; b.click(); return 'picked'; })()`);
    };
    const setupBefore = await js(setupRead);
    const pickedLandscape = await pickPage('Orientation', 'Landscape');
    const landscape = await until(async () => (await js(setupRead)).startsWith('landscape'), 'landscape', 5000).catch(() => false);
    const pickedTitles = await pickPage('Print Titles', 'Repeat row 1');
    const titled = await until(async () => (await js(setupRead)).split(' ')[2] === '1', 'the repeated row', 5000).catch(() => false);
    const pickedWidth = await pickPage('Width', '1 page');
    const fitted = await until(async () => (await js(setupRead)).split(' ')[1] === 'width', 'fit to width', 5000).catch(() => false);
    check('sheets: Page Layout → Orientation, Print Titles and Width write the page setup of the file, the one the print dialog and Excel start from',
      pickedLandscape === 'picked' && pickedTitles === 'picked' && pickedWidth === 'picked' && landscape === true && titled === true && fitted === true,
      `${setupBefore} → ${await js(setupRead)}; ${pickedLandscape} ${pickedTitles} ${pickedWidth}`);
    await clickIn(win, 'Save');
    await until(() => { try { return readPageSetup(SheetView.open(fs.readFileSync(file)), 'Sales').fit === 'width'; } catch { return false; } }, 'the setup to land in the file', 8000).catch(() => false);
    const inFile = readPageSetup(SheetView.open(fs.readFileSync(file)), 'Sales');
    check('sheets: the saved file keeps the orientation, the repeated row and the fit, as Excel reads them',
      inFile.orientation === 'landscape' && inFile.repeatRows === 1 && inFile.fit === 'width', JSON.stringify([inFile.orientation, inFile.repeatRows, inFile.fit]));

    const complaints = await errorsIn(win);
    check('sheets: the link and note checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the link and note checks ran', false, err.message);
  }
  return undefined;
}
