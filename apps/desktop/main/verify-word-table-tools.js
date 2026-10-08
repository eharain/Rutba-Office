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

/** A document of two tables: one in a Table Grid style the styles part defines, a cell with its own red border and shading; one with no lines at all. */
function linesDocx() {
  const base = buildDocx({ styles: true, paragraphs: [{ text: 'Lines' }, { text: 'Between' }, { text: 'After' }] });
  const { entries } = readZip(Buffer.from(base));
  const part = entries.find((e) => e.name === 'word/document.xml');
  const styles = entries.find((e) => e.name === 'word/styles.xml');
  styles.data = Buffer.from(styles.data.toString('utf8').replace('</w:styles>', '<w:style w:type="table" w:styleId="GridStyle"><w:name w:val="Grid Style"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8" w:color="000000"/><w:left w:val="single" w:sz="8" w:color="000000"/><w:bottom w:val="single" w:sz="8" w:color="000000"/><w:right w:val="single" w:sz="8" w:color="000000"/><w:insideH w:val="single" w:sz="8" w:color="000000"/><w:insideV w:val="single" w:sz="8" w:color="000000"/></w:tblBorders></w:tblPr></w:style></w:styles>'), 'utf8');
  const tc = (text, pr = '') => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${pr}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const styled = `<w:tbl><w:tblPr><w:tblStyle w:val="GridStyle"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>`
    + `<w:tr>${tc('Styled')}${tc('Red', '<w:tcBorders><w:top w:val="single" w:sz="24" w:color="FF0000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>')}</w:tr><w:tr>${tc('Grid')}${tc('Lines')}</w:tr></w:tbl>`;
  const bare = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${tc('No')}${tc('lines')}</w:tr></w:tbl>`;
  const xml = part.data.toString('utf8');
  const at = (text) => xml.indexOf('</w:p>', xml.indexOf(text)) + '</w:p>'.length;
  const first = at('>Lines<');
  const second = at('>Between<');
  part.data = Buffer.from(xml.slice(0, first) + styled + xml.slice(first, second) + bare + xml.slice(second), 'utf8');
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

  // A table's lines as the file gives them: its style's, a cell's own over
  // them, a cell's shading; a table that gives none drawn with none.
  const lines = path.join(dir, 'table-lines.docx');
  try {
    fs.writeFileSync(lines, linesDocx());
    const win = await open('word', lines);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.wd-page table.wd-table').length >= 2`), 'the tables to be drawn', 8000).catch(() => {});
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-table-lines.png'), (await win.webContents.capturePage()).toPNG());
    const looks = await js(`(() => {
      const td = (text) => [...document.querySelectorAll('.wd-page table.wd-table td')].find((c) => c.innerText.trim() === text);
      const of = (text) => { const c = td(text); if (!c) return null; const s = getComputedStyle(c); return { top: s.borderTopWidth + ' ' + s.borderTopStyle + ' ' + s.borderTopColor, left: s.borderLeftWidth + ' ' + s.borderLeftStyle, bg: s.backgroundColor, outline: s.outlineStyle }; };
      return { styled: of('Grid'), red: of('Red'), bare: of('No') };
    })()`);
    check('word: a table in a style that rules it is drawn with the style\'s lines',
      looks.styled?.top === '1px solid rgb(0, 0, 0)' && looks.styled?.left === '1px solid', JSON.stringify(looks.styled));
    check('word: a cell\'s own border and shading are drawn over the table\'s',
      // w:sz="24" is 3pt, four pixels.
      /^4px solid rgb\(255, 0, 0\)$/.test(looks.red?.top || '') && looks.red?.bg === 'rgb(255, 255, 0)', JSON.stringify(looks.red));
    check('word: a table that gives no lines is drawn with none, only the faint gridlines',
      looks.bare?.top.startsWith('0px') && looks.bare?.outline === 'dashed', JSON.stringify(looks.bare));
    const complaints = await errorsIn(win);
    check('word: drawing a table\'s lines reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the table-lines checks ran', false, err.message);
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

  // A header row repeated: a table too long for its page, its top row made
  // the header, draws that row again at the head of the next page's piece.
  const long = path.join(dir, 'header-rows.docx');
  try {
    const body = Array.from({ length: 45 }, (_, i) => [`Row ${i + 1}`, String(100 + i), String(200 + i)]);
    fs.writeFileSync(long, buildDocx({ styles: true, paragraphs: [{ text: 'A long table' }, { table: { rows: [['Region', 'Q1', 'Q2'], ...body] } }, { text: 'After the long table' }] }));
    const win = await open('word', long);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = () => h.doc.model({ id: h.sessionFor('doc').id });
    const index = (text) => model().blocks.findIndex((b) => b.text === text);
    await until(() => js(`document.querySelectorAll('.wd-page table.wd-table[data-table]').length >= 2`), 'the table split across pages', 10000);
    await js(`(() => {
      const a = document.querySelector('.wd-page [data-block="${index('Region')}"]');
      a.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
      const r = document.createRange(); r.setStart(a.firstChild || a, 0); r.collapse(true);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return true;
    })()`);
    await until(() => js(`[...document.querySelectorAll('.rw-tab')].some((t) => t.textContent.trim() === 'Table Layout')`), 'Table Layout', 3000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Table Layout').click(), 'tab'`);
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Repeat Header Rows'))`), 'Repeat Header Rows', 3000);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Repeat Header Rows'); b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); b.click(); return 'pressed'; })()`);
    await until(() => js(`Boolean(document.querySelector('.wd-page table.wd-table[data-part="1"] tr[data-repeat]'))`), 'the header drawn again', 6000).catch(() => {});
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await js(`document.querySelector('.wd-page table.wd-table[data-part="1"]')?.scrollIntoView({ block: 'start' }), 'scrolled'`);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-header-rows.png'), (await win.webContents.capturePage()).toPNG());
    }
    const pieces = await js(`[...document.querySelectorAll('.wd-page table.wd-table[data-table]')].map((t) => {
      const rows = [...t.tBodies[0].rows];
      return {
        from: Number(t.dataset.rowFrom || 0),
        first: rows[0] ? [...rows[0].cells].map((c) => c.innerText.trim()).join('|') : '',
        repeat: rows.filter((r) => r.dataset.repeat).length,
        locked: rows[0]?.dataset.repeat ? rows[0].isContentEditable === false : null,
        second: rows.find((r) => !r.dataset.repeat) ? [...rows.find((r) => !r.dataset.repeat).cells].map((c) => c.innerText.trim())[0] : '',
      };
    })`);
    const [top, next] = pieces;
    check('word: a header row is drawn again at the head of the table\'s next page, not typed in there, and the rows carry on after it',
      model().blocks[index('Region')].rowHeader === true && top?.repeat === 0 && next?.repeat === 1 && next.first === 'Region|Q1|Q2' && next.locked === true && next.second === `Row ${next.from}`,
      JSON.stringify(pieces));
    const complaints = await errorsIn(win);
    check('word: repeating a header row reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the header-row checks ran', false, err.message);
  }
}
