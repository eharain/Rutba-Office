// Word: tables — merged cells drawn as one, and Table Layout.
//
// A document whose table merges a column's cells down three rows (w:vMerge)
// and another cell across two columns and down two rows draws each merge as
// one cell reaching over the places it covers, as Word draws it, not as the
// first cell and blank cells under it, and so are cells merged from Table
// Layout. Its Cell Alignment, Repeat Header Rows and Distribute Columns
// each change the table on the page (the rest of the tab is checked by
// verify-word-table.js). Run alone with RUTBA_VERIFY_ONLY=tabletools.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { readZip, writeZip } from '@rutba/ooxml/zip';

const cell = (text, { w = 2000, span = 1, merge = null } = {}) =>
  `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}${merge === 'restart' ? '<w:vMerge w:val="restart"/>' : merge ? '<w:vMerge/>' : ''}</w:tcPr><w:p>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p></w:tc>`;
const row = (...cells) => `<w:tr>${cells.join('')}</w:tr>`;
const table = (cols, rows) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid>${cols.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;

/** A document of two tables: one merged down a column, one merged across and down. */
function mergedDocx() {
  const base = buildDocx({ styles: true, paragraphs: [{ text: 'Merged cells' }, { text: 'Between the tables' }, { text: 'After the tables' }] });
  const { entries } = readZip(Buffer.from(base));
  const part = entries.find((e) => e.name === 'word/document.xml');
  const down = table([3000, 2000, 2000], [
    row(cell('Region', { w: 3000 }), cell('Q1'), cell('Q2')),
    row(cell('North and South', { w: 3000, merge: 'restart' }), cell('120'), cell('135')),
    row(cell('', { w: 3000, merge: 'continue' }), cell('140'), cell('150')),
    row(cell('', { w: 3000, merge: 'continue' }), cell('160'), cell('170')),
  ]);
  const block = table([2000, 2000, 2000], [
    row(cell('Notes', { w: 4000, span: 2, merge: 'restart' }), cell('X')),
    row(cell('', { w: 4000, span: 2, merge: 'continue' }), cell('Y')),
    row(cell('A'), cell('B'), cell('C')),
  ]);
  const xml = part.data.toString('utf8');
  const at = (text) => xml.indexOf('</w:p>', xml.indexOf(text)) + '</w:p>'.length;
  const first = at('Merged cells');
  const second = at('Between the tables');
  part.data = Buffer.from(xml.slice(0, first) + down + xml.slice(first, second) + block + xml.slice(second), 'utf8');
  return writeZip(entries);
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordTableTools(h, { dir }) {
  const { open, check, until, errorsIn } = h;
  const file = path.join(dir, 'merged-cells.docx');
  try {
    fs.writeFileSync(file, mergedDocx());
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.wd-page table.wd-table').length >= 2`), 'the tables to be drawn', 8000).catch(() => {});
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-merged-cells.png'), (await win.webContents.capturePage()).toPNG());
    const tables = await js(`[...document.querySelectorAll('.wd-page table.wd-table')].map((t) => ({
      rows: [...t.tBodies[0].rows].map((r) => [...r.cells].map((c) => c.innerText.trim() + (c.rowSpan > 1 ? '^' + c.rowSpan : '') + (c.colSpan > 1 ? '<' + c.colSpan : '')).join('|')),
      tall: (() => { const c = [...t.querySelectorAll('td')].find((d) => d.rowSpan > 1); const rows = [...t.tBodies[0].rows]; return c ? Math.round(c.getBoundingClientRect().height) + '/' + Math.round(rows.slice(c.parentElement.rowIndex, c.parentElement.rowIndex + c.rowSpan).reduce((s, r) => s + r.getBoundingClientRect().height, 0)) : null; })(),
    }))`);
    const [down, block] = tables;
    check('word: a cell merged down a column is one cell over the three rows it covers',
      down?.rows.join(' / ') === 'Region|Q1|Q2 / North and South^3|120|135 / 140|150 / 160|170' && down.tall && down.tall.split('/')[0] === down.tall.split('/')[1],
      `${down?.rows.join(' / ')}; ${down?.tall} px tall over its rows`);
    check('word: a cell merged across two columns and down two rows is one cell over the four places',
      block?.rows.join(' / ') === 'Notes^2<2|X / Y / A|B|C',
      block?.rows.join(' / ') || 'no second table');
    const complaints = await errorsIn(win);
    check('word: drawing merged cells reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the merged-cell checks ran', false, err.message);
  }

  // Table Layout: there while the caret is in a table, and each of its
  // controls doing what it says to the table on the page and in the file.
  const grid = path.join(dir, 'table-layout.docx');
  try {
    fs.writeFileSync(grid, buildDocx({ styles: true, paragraphs: [
      { text: 'Before the table' },
      { table: { rows: [['Region', 'Q1', 'Q2'], ['North', '120', '135'], ['South', '140', '150'], ['East', '160', '170']] } },
      { text: 'After the table' },
    ] }));
    const win = await open('word', grid);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = () => h.doc.model({ id: h.sessionFor('doc').id });
    const index = (text) => model().blocks.findIndex((b) => b.text === text);
    await until(() => js(`document.querySelectorAll('.wd-page table.wd-table').length === 1`), 'the table to be drawn', 8000);
    // A DOM range from one paragraph to another, then the mouse released — how the editor reads a selection.
    const select = (from, to = from) => js(`(() => {
      const el = (i) => document.querySelector('.wd-page [data-block="' + i + '"]');
      const a = el(${index(from)}); const b = el(${index(to)});
      a.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
      const r = document.createRange();
      const first = (n) => { const w = document.createTreeWalker(n, NodeFilter.SHOW_TEXT); return w.nextNode() || n; };
      r.setStart(first(a), 0); r.setEnd(first(b), 0);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return true;
    })()`);
    const tabs = () => js(`[...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim())`);
    const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    const button = (label) => `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)})`;
    const press = (label) => js(`(() => { const b = ${button(label)}; if (!b) return 'missing'; if (b.disabled) return 'disabled'; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); b.click(); return 'pressed'; })()`);
    const pick = async (label, item) => {
      await press(label);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(item)}))`), `"${item}" in ${label}`, 3000);
      return js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(item)}).click(), 'picked'`);
    };
    const rows = () => js(`[...document.querySelector('.wd-page table.wd-table').tBodies[0].rows].map((r) => [...r.cells].map((c) => c.innerText.trim() + (c.rowSpan > 1 ? '^' + c.rowSpan : '')).join('|'))`);

    // The tab itself, its rows and columns and its Merge and Split are verify-word-table.js's.
    await select('North');
    await until(async () => (await tabs()).includes('Table Layout'), 'Table Layout to appear', 3000);
    await clickTab('Table Layout');
    await until(() => js(`Boolean(${button('Distribute Columns')})`), 'the Table Layout controls', 3000);

    await select('South', 'East');
    await until(() => js(`${button('Merge Cells')}?.disabled === false`), 'Merge Cells to wake', 3000).catch(() => {});
    await press('Merge Cells');
    await until(async () => (await rows()).some((r) => r.includes('^2')), 'the merge', 4000).catch(() => {});
    const merged = await rows();
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-table-layout.png'), (await win.webContents.capturePage()).toPNG());
    check('word: cells merged from Table Layout are drawn as one cell, South over East\'s row',
      merged[2]?.startsWith('South') && merged[2].includes('^2') && merged[3] === '160|170',
      merged.join(' / '));

    await select('120');
    await pick('Cell Alignment', 'Centre');
    await until(() => js(`(() => { const p = document.querySelector('.wd-page [data-block="${index('120')}"]'); return getComputedStyle(p.closest('td')).verticalAlign === 'middle'; })()`), 'the cell centred', 4000).catch(() => {});
    const centred = await js(`getComputedStyle(document.querySelector('.wd-page [data-block="${index('120')}"]').closest('td')).verticalAlign`);
    check('word: Cell Alignment → Centre sets the cell\'s words in its middle',
      centred === 'middle' && model().blocks[index('120')].cellVAlign === 'center', `${centred}; model ${model().blocks[index('120')].cellVAlign}`);

    await select('Region');
    await press('Repeat Header Rows');
    await until(() => js(`${button('Repeat Header Rows')}?.getAttribute('aria-pressed') === 'true' || ${button('Repeat Header Rows')}?.classList.contains('pressed')`), 'the header row pressed', 4000).catch(() => {});
    const header = model().blocks[index('Region')].rowHeader === true && !model().blocks[index('North')].rowHeader;
    check('word: Repeat Header Rows makes the top row the table\'s header', header, `Region ${model().blocks[index('Region')].rowHeader}; North ${model().blocks[index('North')].rowHeader}`);

    await press('Distribute Columns');
    await until(() => js(`(() => { const cols = [...document.querySelectorAll('.wd-page table.wd-table col')].map((c) => parseFloat(c.style.width)); return cols.length === 3 && cols.every((w) => Math.abs(w - cols[0]) < 0.5); })()`), 'the columns shared out', 4000).catch(() => {});
    const widths = await js(`[...document.querySelectorAll('.wd-page table.wd-table col')].map((c) => c.style.width)`);
    check('word: Distribute Columns makes every column the same width', widths.length === 3 && new Set(widths).size === 1, widths.join(', '));

    const complaints = await errorsIn(win);
    check('word: Table Layout reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the Table Layout checks ran', false, err.message);
  }
}
