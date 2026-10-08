// Presentations: Table Design.
//
// A table put on a blank slide from Insert → Table is selected, and the
// Table Design tab is there: Banded Columns on writes the table's bandCol,
// Header Row off its firstRow, and Shading colours every cell (none was
// typed in), the stage drawing each. The tab goes when nothing is selected.
// Run alone with RUTBA_VERIFY_ONLY=decktabledesign.
import fs from 'node:fs';
import path from 'node:path';
import { buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the deck is written
 */
export async function verifyDeckTableDesign(h, { dir }) {
  const { open, check, until, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'table-design.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ slides: [{ layout: 'blank' }] }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('deck');
    const table = () => doc.model({ id: session.id }).slide.shapes.find((s) => s.kind === 'table');
    const tabs = () => js(`[...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim())`);
    const tab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    const button = (label) => `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)} || n.getAttribute('aria-label') === ${JSON.stringify(label)})`;
    const press = (label) => js(`(() => { const b = ${button(label)}; if (!b) return 'missing'; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); b.click(); return 'pressed'; })()`);
    const pick = async (label, item) => {
      await press(label);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(item)}))`), `"${item}" in ${label}`, 3000);
      return js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(item)}).click(), 'picked'`);
    };

    const before = await tabs();
    await tab('Insert');
    await until(() => js(`Boolean(${button('Table')})`), 'Insert → Table', 3000);
    await pick('Table', '3 × 3');
    await until(async () => Boolean(table()) && (await tabs()).includes('Table Design'), 'the table and Table Design', 5000).catch(() => {});
    check('slides: a table selected brings the Table Design tab, which was not there before',
      !before.includes('Table Design') && (await tabs()).includes('Table Design') && Boolean(table()), `${before.includes('Table Design')} → ${(await tabs()).includes('Table Design')}`);

    await tab('Table Design');
    await until(() => js(`Boolean(${button('Banded Columns')})`), 'the Table Design controls', 3000);
    await press('Banded Columns');
    await until(() => table()?.table?.flags?.bandCol === true, 'banded columns on', 4000).catch(() => {});
    await press('Header Row');
    await until(() => table()?.table?.flags?.firstRow === false, 'the header row off', 4000).catch(() => {});
    const flags = table()?.table?.flags || {};
    check('slides: Table Style Options turn Banded Columns on and Header Row off on the table', flags.bandCol === true && flags.firstRow === false && flags.bandRow === true, JSON.stringify(flags));

    await pick('Shading', 'Orange');
    await until(() => (table()?.table?.cells || []).flat().every((c) => String(c.fill || '').toLowerCase() === '#ffc000'), 'every cell orange', 4000).catch(() => {});
    const fills = (table()?.table?.cells || []).flat().map((c) => String(c.fill || '').toLowerCase());
    check('slides: Shading with no cell typed in colours every cell of the table', fills.length === 9 && fills.every((f) => f === '#ffc000'), fills.join(','));

    const complaints = await errorsIn(win);
    check('slides: Table Design reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the Table Design checks ran', false, err.message);
  }
}
