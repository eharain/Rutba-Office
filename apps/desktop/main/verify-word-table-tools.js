// Word: tables — merged cells drawn as one.
//
// A document whose table merges a column's cells down three rows (w:vMerge)
// and another cell across two columns and down two rows draws each merge as
// one cell reaching over the places it covers, as Word draws it, not as the
// first cell and blank cells under it. Run alone with
// RUTBA_VERIFY_ONLY=tabletools.
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
}
