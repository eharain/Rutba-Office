// Word: an OpenDocument text, opened as it was made.
//
// An .odt written the way LibreOffice writes one — a heading, a run in its
// looks, a bulleted list and a numbered one with a level inside it, a table
// whose first column spans two rows, a picture — opens as its pages: the
// lists labelled as their list styles say, the table a table, the picture
// drawn. Run alone with RUTBA_VERIFY_ONLY=odt.
import fs from 'node:fs';
import path from 'node:path';
import { readZip, writeZip, ZipEntry } from '@rutba/ooxml/zip';
import { writeOdt } from '@rutba/office-formats/odf-write';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAABAAAAAICAIAAAB/FOjAAAAAE0lEQVR4nGOQizpBEmIY1UALDQAzrqABidudowAAAABJRU5ErkJggg==', 'base64');
const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3"';

const STYLES = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS}><office:styles>`
  + '<style:style style:name="Standard" style:family="paragraph"/>'
  + '<style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Standard" style:default-outline-level="1"><style:text-properties fo:font-size="18pt" fo:font-weight="bold"/></style:style>'
  + '</office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/></style:page-layout></office:automatic-styles>'
  + '<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>';

const CONTENT = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS}><office:automatic-styles>`
  + '<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>'
  + '<style:style style:name="Table1.A" style:family="table-column"><style:table-column-properties style:column-width="5cm"/></style:style>'
  + '<style:style style:name="Table1.B" style:family="table-column"><style:table-column-properties style:column-width="3cm"/></style:style>'
  + '<text:list-style style:name="L1"><text:list-level-style-bullet text:level="1" text:bullet-char="•"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment fo:text-indent="-0.635cm" fo:margin-left="1.27cm"/></style:list-level-properties></text:list-level-style-bullet></text:list-style>'
  + '<text:list-style style:name="L2"><text:list-level-style-number text:level="1" style:num-suffix="." style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment fo:text-indent="-0.635cm" fo:margin-left="1.27cm"/></style:list-level-properties></text:list-level-style-number>'
  + '<text:list-level-style-number text:level="2" style:num-suffix=")" style:num-format="a"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment fo:text-indent="-0.635cm" fo:margin-left="1.905cm"/></style:list-level-properties></text:list-level-style-number></text:list-style>'
  + '</office:automatic-styles><office:body><office:text>'
  + '<text:h text:style-name="Heading_20_1" text:outline-level="1">The OpenDocument showcase</text:h>'
  + '<text:p text:style-name="Standard">Plain and <text:span text:style-name="T1">bold</text:span>.</text:p>'
  + '<text:list text:style-name="L1"><text:list-item><text:p>First bullet</text:p></text:list-item></text:list>'
  + '<text:list text:style-name="L2"><text:list-item><text:p>Step one</text:p><text:list><text:list-item><text:p>Inside</text:p></text:list-item></text:list></text:list-item><text:list-item><text:p>Step two</text:p></text:list-item></text:list>'
  + '<table:table table:name="Table1"><table:table-column table:style-name="Table1.A"/><table:table-column table:style-name="Table1.B" table:number-columns-repeated="2"/>'
  + '<table:table-row><table:table-cell><text:p>Region</text:p></table:table-cell><table:table-cell><text:p>Q1</text:p></table:table-cell><table:table-cell><text:p>Q2</text:p></table:table-cell></table:table-row>'
  + '<table:table-row><table:table-cell table:number-rows-spanned="2"><text:p>North and South</text:p></table:table-cell><table:table-cell><text:p>120</text:p></table:table-cell><table:table-cell><text:p>135</text:p></table:table-cell></table:table-row>'
  + '<table:table-row><table:covered-table-cell/><table:table-cell><text:p>140</text:p></table:table-cell><table:table-cell><text:p>150</text:p></table:table-cell></table:table-row>'
  + '</table:table>'
  + '<text:p text:style-name="Standard"><draw:frame draw:name="Logo" text:anchor-type="as-char" svg:width="4cm" svg:height="2cm"><draw:image xlink:href="Pictures/logo.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></text:p>'
  + '<text:p text:style-name="Standard">After the picture.</text:p>'
  + '</office:text></office:body></office:document-content>';

/** The file: our own archive's mimetype and manifest, with these parts and the picture in it. */
function showcase() {
  const { entries } = readZip(Buffer.from(writeOdt({ blocks: [] })));
  for (const e of entries) {
    if (e.name === 'content.xml') e.data = Buffer.from(CONTENT, 'utf8');
    if (e.name === 'styles.xml') e.data = Buffer.from(STYLES, 'utf8');
  }
  const pic = new ZipEntry({ name: 'Pictures/logo.png', method: 0, crc: 0, compressedSize: 0, uncompressedSize: 0, compressed: Buffer.alloc(0), dosTime: 0, dosDate: 0x2821, flags: 0, externalAttrs: 0, comment: Buffer.alloc(0) });
  pic.data = PNG;
  return writeZip([...entries, pic]);
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the file is written
 */
export async function verifyWordOdt(h, { dir }) {
  const { open, check, until, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'showcase.odt');
  try {
    fs.writeFileSync(file, showcase());
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('doc');
    await until(() => js(`document.querySelectorAll('.wd-page table.wd-table').length > 0`), 'the table to be drawn', 8000).catch(() => {});
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-odt.png'), (await win.webContents.capturePage()).toPNG());

    const shown = await js(`(() => {
      const page = document.querySelector('.wd-page');
      const table = page?.querySelector('table.wd-table');
      return {
        text: page ? page.innerText : '',
        rows: table ? [...table.querySelectorAll('tr')].map((r) => [...r.children].map((c) => c.innerText.trim()).join('|')) : [],
        pictures: page ? [...page.querySelectorAll('img')].filter((i) => i.naturalWidth > 0).map((i) => Math.round(i.getBoundingClientRect().width / Math.max(1, i.getBoundingClientRect().height) * 10) / 10) : [],
      };
    })()`);
    const model = doc.model({ id: session.id });
    const text = (b) => (b.runs || []).map((r) => r.text || '').join('');
    const label = (t) => model.listLabels?.[model.blocks.findIndex((b) => text(b) === t)]?.label;
    check('word: an .odt opens as its pages, the heading and the words in their looks',
      session.converted?.from === 'odt' && /The OpenDocument showcase/.test(shown.text) && /After the picture\./.test(shown.text),
      `from ${session.converted?.from}; ${shown.text.replace(/\s+/g, ' ').slice(0, 160)}`);
    // The place under the span is the span's continuation: empty, the words not written twice.
    const spanned = shown.rows[1]?.startsWith('North and South|') && shown.rows[2] === '|140|150';
    const labels = ['First bullet', 'Step one', 'Inside', 'Step two'].map(label);
    const drawnLabels = /•\s*First bullet/.test(shown.text) && /1\.\s*Step one/.test(shown.text) && /a\)\s*Inside/.test(shown.text) && /2\.\s*Step two/.test(shown.text);
    check('word: the .odt\'s lists are labelled as its list styles say, a level inside counted on its own',
      labels.join(' ') === '• 1. a) 2.' && drawnLabels,
      `${labels.join(' ')}; drawn ${drawnLabels}`);
    check('word: the .odt\'s table is a table, the place under its span the span\'s continuation',
      shown.rows[0] === 'Region|Q1|Q2' && spanned,
      `${shown.rows.join(' / ') || 'no table'}; spanned ${spanned}`);
    check('word: the .odt\'s picture is drawn at its proportions',
      shown.pictures.length === 1 && Math.abs(shown.pictures[0] - 2) < 0.15,
      `${shown.pictures.length} picture(s), ${shown.pictures.join(', ')} wide to high`);

    const complaints = await errorsIn(win);
    check('word: opening an .odt reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');

    // Saved, it is written back as an .odt whole, and opens again as it was.
    const saved = path.join(dir, 'showcase-saved.odt');
    const wrote = doc.save({ id: session.id, path: saved });
    const again = await open('word', saved);
    const jsAgain = (code) => again.webContents.executeJavaScript(code);
    await until(() => jsAgain(`document.querySelectorAll('.wd-page table.wd-table').length > 0`), 'the saved file\'s table', 8000).catch(() => {});
    const back = await jsAgain(`(() => { const page = document.querySelector('.wd-page'); return { text: page ? page.innerText : '', pictures: page ? [...page.querySelectorAll('img')].filter((i) => i.naturalWidth > 0).length : 0, tables: page ? page.querySelectorAll('table.wd-table').length : 0 }; })()`);
    const labelsBack = /•\s*First bullet/.test(back.text) && /1\.\s*Step one/.test(back.text) && /a\)\s*Inside/.test(back.text) && /2\.\s*Step two/.test(back.text);
    check('word: an .odt saved from the suite opens again with its lists labelled, its table and its picture',
      wrote?.format === 'odt' && labelsBack && back.tables === 1 && back.pictures === 1 && /The OpenDocument showcase/.test(back.text),
      `${wrote?.format}; labels ${labelsBack}; ${back.tables} table(s), ${back.pictures} picture(s)`);
  } catch (err) {
    check('word: the .odt checks ran', false, err.message);
  }
}
