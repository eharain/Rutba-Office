// Worksheets: a few words in a cell formatted while it is typed in (owner,
// 2026-10-21). A word chosen in the cell editor made bold from the ribbon
// turns the editor rich, showing it bold; Enter puts the words and their
// looks in and moves down; saved and opened again they are there; a cell
// whose words have looks opens in the rich editor, where Ctrl+I makes a
// chosen word italic; and Esc leaves the cell as it was.

import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

export async function verifySheetRichEdit(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'rich-cell.xlsx');
  try {
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Price list today'], ['next']] }] }));
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'sheet' && s.path === file).pop();
    const runsAt = (id, r) => doc.model({ id }).cells?.find((c) => c.ref === r)?.runs || null;
    const looks = (runs) => (runs || []).map((r) => [r.text, Boolean(r.bold), Boolean(r.italic)]);
    const editA1 = async () => {
      await js(`(() => { const c = document.querySelector('.sh-cell[data-ref="A1"]'); c.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); c.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, button: 0 })); c.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return 1; })()`);
    };
    const bold = () => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => /^Bold/.test(n.dataset.tip || '')); if (!b) return 'no button'; b.click(); return 'pressed'; })()`);
    const enter = (sel) => js(`document.querySelector(${JSON.stringify(sel)})?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })), 1`);

    // "list" chosen in the plain editor, Bold from the ribbon: the editor turns rich, "list" bold in it.
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="A1"]'))`), 'the grid', 8000);
    await editA1();
    await until(() => js(`Boolean(document.querySelector('input.sh-editor'))`), 'the cell editor', 5000).catch(() => {});
    await js(`(() => { const i = document.querySelector('input.sh-editor'); i.focus(); i.setSelectionRange(6, 10); return 1; })()`);
    const pressed = await bold();
    await until(() => js(`Boolean(document.querySelector('.sh-rich'))`), 'the rich cell editor', 4000).catch(() => {});
    const shown = await js(`[...document.querySelectorAll('.sh-rich span')].map((s) => [s.textContent, /700/.test(s.style.fontWeight)])`);
    await enter('.sh-rich');
    const kept = await until(() => looks(runsAt(session.id, 'A1')).some(([t, b]) => t === 'list' && b) || null, 'the runs in the cell', 5000).then(() => true, () => false);
    await until(() => js(`document.querySelector('.sh-cell.active')?.dataset.ref === 'A2'`), 'the move down', 4000).catch(() => {});
    const moved = await js(`document.querySelector('.sh-cell.active')?.dataset.ref || ''`);
    check('sheets: a word chosen in the cell editor and made bold from the ribbon shows bold in a rich editor, and Enter puts the words and their looks in and moves down',
      pressed === 'pressed' && JSON.stringify(shown) === JSON.stringify([['Price ', false], ['list', true], [' today', false]]) && kept && moved === 'A2',
      JSON.stringify({ pressed, shown, runs: looks(runsAt(session.id, 'A1')), moved }));

    await doc.save({ id: session.id, path: file });
    const again = await doc.open({ path: file });
    const reopened = looks(runsAt(again.id, 'A1'));
    doc.close({ id: again.id });
    check('sheets: the words formatted in the cell are there when the file is opened again', JSON.stringify(reopened) === JSON.stringify([['Price ', false, false], ['list', true, false], [' today', false, false]]), JSON.stringify(reopened));

    // Opened again for editing: the rich editor at once; "Price" chosen, Ctrl+I.
    await editA1();
    const richAtOnce = await until(() => js(`Boolean(document.querySelector('.sh-rich'))`), 'the rich editor at once', 5000).then(() => true, () => false);
    await js(`(() => { const el = document.querySelector('.sh-rich'); el.focus(); const t = document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode(); const r = document.createRange(); r.setStart(t, 0); r.setEnd(t, 5); const s = getSelection(); s.removeAllRanges(); s.addRange(r); return 1; })()`);
    await wait(150);
    await js(`document.querySelector('.sh-rich').dispatchEvent(new KeyboardEvent('keydown', { key: 'i', ctrlKey: true, bubbles: true, cancelable: true })), 1`);
    await wait(150);
    await enter('.sh-rich');
    const italic = await until(() => looks(runsAt(session.id, 'A1')).some(([t, , i]) => t === 'Price' && i) || null, 'Price italic', 5000).then(() => true, () => false);
    check('sheets: a cell whose words have looks opens in the rich editor, where Ctrl+I makes a chosen word italic and the bold word stays bold',
      richAtOnce && italic && looks(runsAt(session.id, 'A1')).some(([t, b]) => t === 'list' && b), JSON.stringify(looks(runsAt(session.id, 'A1'))));

    // Esc: the cell as it was.
    const before = JSON.stringify(looks(runsAt(session.id, 'A1')));
    await editA1();
    await until(() => js(`Boolean(document.querySelector('.sh-rich'))`), 'the rich editor', 5000).catch(() => {});
    await js(`(() => { const el = document.querySelector('.sh-rich'); el.focus(); const s = getSelection(); s.selectAllChildren(el); return 1; })()`);
    await wait(100);
    await bold();
    await wait(100);
    await js(`document.querySelector('.sh-rich')?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })), 1`);
    await until(() => js(`!document.querySelector('.sh-rich')`), 'the editor gone', 4000).catch(() => {});
    check('sheets: Esc in the rich cell editor leaves the cell as it was', JSON.stringify(looks(runsAt(session.id, 'A1'))) === before, `${before} -> ${JSON.stringify(looks(runsAt(session.id, 'A1')))}`);
  } catch (err) {
    check('sheets: the rich cell editor checks ran', false, err.stack || err.message);
  }
}
