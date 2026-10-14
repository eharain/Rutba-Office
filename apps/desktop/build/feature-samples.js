// The samples the feature screenshots open (see make-screens.js and
// main/screen-features.js). Each is invented: the people, the companies and
// the figures are made up, the addresses are on example domains, and the
// pictures are the landscapes the screens build draws, or the printed pages
// the text recognition tests read.

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { buildPptx, Deck } from '@rutba/presentation';
import { decodePng } from '@rutba/ocr/png';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { secretLetter } from '../main/verify-review-fixtures.js';
import { PST_MESSAGES, OLM_MESSAGES, ACCOUNT_MESSAGES, writeOlm } from '../../../tests/fixtures/mail-archives.js';
import { writePst } from '../../../tests/fixtures/pst-writer.js';

const require = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, '../../..');

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The person the archive fixtures write to, in the pictures: invented. */
const OWNER = { name: 'Imogen Hartley', address: 'imogen@harbourlane.example' };
// The fixtures' own recipient, found as the first message's addressee
// rather than named here.
const THEIRS = PST_MESSAGES.Inbox[0].to[0].address;
const swap = (p) => (p && p.address === THEIRS ? OWNER : p);

/** Archive messages with their own recipient replaced by OWNER. */
export function readdressed(messages) {
  return (messages || []).map((m) => ({ ...m, from: swap(m.from), to: (m.to || []).map(swap), ...(m.cc ? { cc: m.cc.map(swap) } : {}) }));
}

/** The account's messages from the same matter, readdressed. */
export const ACCOUNT_MAIL = readdressed(ACCOUNT_MESSAGES);

/** Writes every feature sample into `dir`. */
export function writeFeatureSamples(dir) {
  const at = (name) => path.join(dir, name);
  const pdf = require(path.join(repo, 'packages/pdf/src/index.js'));
  const ocrFixture = (name) => path.join(repo, 'packages/ocr/test/fixtures', name);

  /* Documents ----------------------------------------------------------- */

  // A tenancy draft two people have been through: insertions, deletions and
  // two comments, each by name.
  {
    const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'Tenancy agreement: draft three', style: 'Title', bold: true, size: 30 }] }));
    const t = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const ins = (id, who, text) => `<w:ins w:id="${id}" w:author="${who}" w:date="2026-10-12T10:00:00Z"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:ins>`;
    const del = (id, who, text) => `<w:del w:id="${id}" w:author="${who}" w:date="2026-10-12T11:00:00Z"><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
    const head = (text) => `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="26"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
    const body = head('1. The premises')
      + `<w:p>${t('The landlord lets to the tenant the ground floor shop at 12 Harbour Lane, ')}${ins(11, 'Imogen Hartley', 'together with the yard behind it, ')}${t('for use as a café and for no other purpose.')}</w:p>`
      + head('2. Rent')
      + `<w:p><w:commentRangeStart w:id="0"/>${t('The rent is ')}${del(12, 'Tomas Lindqvist', '£1,450')}${ins(13, 'Tomas Lindqvist', '£1,380')}${t(' a month, paid on the first working day of each month')}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>${t('.')}</w:p>`
      + `<w:p>${t('The rent is reviewed ')}${del(14, 'Imogen Hartley', 'every year')}${ins(15, 'Imogen Hartley', 'on the third anniversary of the start date')}${t(', and can only go up by the change in the retail price index since the last review.')}</w:p>`
      + head('3. Repairs')
      + `<w:p><w:commentRangeStart w:id="1"/>${t('The tenant keeps the inside of the shop in good repair. The landlord keeps the roof, the outside walls and the drains in good repair.')}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>`
      + `<w:p>${ins(16, 'Tomas Lindqvist', 'The tenant may put up a sign above the shop window without asking first. ')}${t('Any other change to the outside of the building needs the written consent of the landlord, which will not be held back without a good reason.')}</w:p>`
      + head('4. Ending the tenancy')
      + `<w:p>${t('Either side can end this agreement by giving ')}${del(17, 'Imogen Hartley', 'three')}${ins(18, 'Imogen Hartley', 'six')}${t(' months of notice in writing.')}</w:p>`;
    pkg.write_('word/document.xml', pkg.text('word/document.xml').replace('<w:sectPr/>', `${body}<w:sectPr/>`));
    pkg.addPart(
      'word/comments.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W}>`
        + '<w:comment w:id="0" w:author="Imogen Hartley" w:initials="IH" w:date="2026-10-12T12:00:00Z"><w:p><w:r><w:t>Agreed on the phone on Friday. Can we say it includes the service charge?</w:t></w:r></w:p></w:comment>'
        + '<w:comment w:id="1" w:author="Tomas Lindqvist" w:initials="TL" w:date="2026-10-12T14:00:00Z"><w:p><w:r><w:t>Does the yard wall count as an outside wall? Worth saying.</w:t></w:r></w:p></w:comment>'
        + '</w:comments>',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
    );
    pkg.addRelationshipTo('word/document.xml', `${REL}/comments`, 'comments.xml');
    fs.writeFileSync(at('Tenancy agreement.docx'), pkg.write());
  }

  // A form made of content controls and filled in: a name, a choice from a
  // list, a date and a tick.
  {
    const pkg = OoxmlPackage.read(buildDocx({
      styles: true,
      paragraphs: [
        { text: 'Volunteer sign-up', style: 'Title', bold: true, size: 30 },
        { text: 'Harbour Lane Community Garden, spring planting days' },
        { text: 'Name: ' },
        { text: 'Team: ' },
        { text: 'First day: ' },
        { text: 'I can bring my own tools: ' },
      ],
    }));
    pkg.addPart('word/settings.xml', Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W}><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/></w:settings>`, 'utf8'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
    pkg.addRelationshipTo('word/document.xml', `${REL}/settings`, 'settings.xml');
    const v = openDocx(pkg.write());
    const end = (block) => (v.render({ pages: false }).blocks.find((b) => b.index === block)?.text || '').length;
    const put = (block, kind, props, value) => {
      v.setSelection({ block, offset: end(block) });
      const place = v.insertControl({ kind, props });
      if (value !== undefined) v.setControlValue({ ...place, value });
    };
    put(2, 'plainText', { title: 'Full name', tag: 'name' }, 'Imogen Hartley');
    put(3, 'dropDownList', { title: 'Team', tag: 'team', items: [{ text: 'Planting', value: 'Planting' }, { text: 'Watering', value: 'Watering' }, { text: 'Compost', value: 'Compost' }] }, 'Planting');
    put(4, 'date', { title: 'First day', tag: 'first' }, '2026-03-14');
    put(5, 'checkBox', { title: 'Own tools', tag: 'tools' }, true);
    fs.writeFileSync(at('Volunteer sign-up.docx'), v.save());
  }

  // A letter that names one person everywhere a document can keep words,
  // for Redact (the review checks' own letter; the person is invented).
  fs.writeFileSync(at('Offer letter.docx'), secretLetter());

  /* Worksheets ---------------------------------------------------------- */

  // A year of orders by region and product, and a PivotTable of it.
  {
    const regions = ['North', 'South', 'East', 'West'];
    const products = ['Coffee beans', 'Tea', 'Pastries', 'Gift cards'];
    const price = { 'Coffee beans': 14, Tea: 9, Pastries: 4, 'Gift cards': 25 };
    const rows = [['Region', 'Product', 'Quarter', 'Orders', 'Revenue']];
    let n = 7;
    for (const q of ['Q1', 'Q2', 'Q3', 'Q4']) {
      for (const r of regions) {
        for (const p of products) {
          n = (n * 37 + 11) % 97;
          const orders = 20 + n;
          rows.push([r, p, q, orders, orders * price[p]]);
        }
      }
    }
    const v = SheetView.open(
      buildXlsx({ sheets: [{ name: 'Orders', rows, styles: { 'A1:E1': { bold: true, fill: '#1F3A5F', colour: '#FFFFFF' }, [`E2:E${rows.length}`]: { numFmt: '#,##0' } } }] }),
      { viewportWidth: 1100, viewportHeight: 620 },
    );
    v.setColWidth(1, 110);
    v.createPivot({ name: 'Revenue by region', source: `Orders!A1:E${rows.length}`, target: { sheet: 'Orders', row: 0, col: 6 }, rowFields: ['Region'], colFields: ['Quarter'], dataFields: [{ field: 'Revenue', subtotal: 'sum' }] });
    fs.writeFileSync(at('Orders 2026.xlsx'), v.save());
  }

  // Results against a target, with data bars, a colour scale and a rule
  // that marks who is short.
  {
    const team = [['Amira', 92], ['Ben', 64], ['Chloe', 118], ['Dev', 81], ['Elif', 103], ['Femi', 47], ['Grace', 99], ['Hugo', 72], ['Ines', 110], ['Jonah', 58], ['Kasia', 87], ['Luca', 95]];
    const rows = [['Name', 'Target', 'Sold', 'Of target', 'Returns'], ...team.map(([name, sold], i) => [name, 90, sold, `=C${i + 2}/B${i + 2}`, ((i * 7) % 9) + 1])];
    const v = SheetView.open(
      buildXlsx({ sheets: [{ name: 'Team', rows, styles: { 'A1:E1': { bold: true, fill: '#1F3A5F', colour: '#FFFFFF' }, [`D2:D${rows.length}`]: { numFmt: '0%' } } }] }),
      { viewportWidth: 1100, viewportHeight: 620 },
    );
    v.setColWidth(0, 110);
    v.setColWidth(2, 150);
    v.setColWidth(3, 110);
    const rule = (col, spec) => {
      v.select(1, col);
      v.select(team.length, col, { extend: true });
      v.addConditionalRule(spec);
    };
    rule(2, { kind: 'dataBar' });
    rule(4, { kind: 'colorScale' });
    rule(3, { kind: 'cellIs', operator: 'lessThan', values: [0.8], look: 'red' });
    v.select(0, 0);
    fs.writeFileSync(at('Team results.xlsx'), v.save());
  }

  // An empty sheet and a picture of a printed table, for Data from Picture.
  fs.writeFileSync(at('Stock count.xlsx'), Buffer.from(buildXlsx({ sheets: [{ name: 'Stock', rows: [['Stock count, from the printed sheet']], styles: { 'A1:A1': { bold: true } } }] })));
  fs.copyFileSync(ocrFixture('table-ruled.png'), at('Printed stock sheet.png'));

  /* Pictures ------------------------------------------------------------ */

  // A printed letter as a picture, for Copy Text.
  fs.mkdirSync(at('Scans'), { recursive: true });
  fs.copyFileSync(ocrFixture('eng-serif.png'), path.join(at('Scans'), 'Printed letter.png'));

  /* Mail ---------------------------------------------------------------- */

  // Two old archives, an Outlook .pst and a Mac .olm, for searching across
  // them: the archive checks' own messages, but addressed to an invented
  // person, since the checks' recipient carries a real name.
  const dir2 = at('Mail archives');
  fs.mkdirSync(dir2, { recursive: true });
  writePst(
    path.join(dir2, 'Northwind 2019.pst'),
    [
      { name: 'Inbox', messages: readdressed(PST_MESSAGES.Inbox), children: [{ name: 'Projects', messages: readdressed(PST_MESSAGES['Inbox/Projects']) }] },
      { name: 'Sent Items', messages: readdressed(PST_MESSAGES['Sent Items']) },
    ],
    { name: 'Northwind 2019' },
  );
  writeOlm(path.join(dir2, 'Contoso archive.olm'), Object.fromEntries(Object.entries(OLM_MESSAGES).map(([k, v]) => [k, readdressed(v)])));

  /* Presentations ------------------------------------------------------- */

  // A launch plan, and the same plan as a colleague sent it back: a figure
  // changed, the title nudged, a note added and two slides swapped round.
  {
    const original = buildPptx({
      title: 'Launch plan',
      slides: [
        { layout: 'title', title: 'Launch plan for spring', body: 'Harbour Lane Coffee, draft two' },
        { layout: 'obj', title: 'Budget', body: ['Fit-out £18,000', 'Equipment £9,500', 'Opening week £2,400'] },
        { layout: 'obj', title: 'Timeline', body: ['Lease signed in January', 'Fit-out through February', 'Doors open on 14 March'] },
        { layout: 'obj', title: 'Risks', body: ['Supplier lead times', 'Hiring two baristas in time'] },
      ],
    });
    fs.writeFileSync(at('Launch plan.pptx'), Buffer.from(original));
    const d = Deck.open(original);
    const body = d.slide(1).shapes.find((sh) => sh.placeholder?.type !== 'title' && sh.text);
    d.setText(1, body.id, ['Fit-out £21,500', 'Equipment £9,500', 'Opening week £2,400'].map((text) => ({ runs: [{ text }] })));
    const title = d.slide(1).shapes.find((sh) => sh.placeholder?.type === 'title');
    d.setGeometry(1, title.id, { ...title.geometry, x: title.geometry.x + 60 });
    d.addTextBox(1, { x: 120, y: 430, w: 420, h: 60, paragraphs: [{ runs: [{ text: 'The quote from Brightwell came in higher' }] }], name: 'Reviewer note' });
    d.moveSlide(3, 2);
    fs.writeFileSync(at('Launch plan (reviewed).pptx'), Buffer.from(d.save()));
  }

  // One slide of a title and two points, for Design Ideas to offer looks for.
  fs.writeFileSync(
    at('Design ideas.pptx'),
    Buffer.from(buildPptx({ title: 'Our first year', slides: [{ layout: 'obj', title: 'Our first year', body: ['Twelve thousand cups served', 'Three new suppliers, all local'] }] })),
  );

  // A short talk with speaker notes, for the presenter view.
  {
    const d = Deck.open(
      buildPptx({
        title: 'Opening week',
        slides: [
          { layout: 'title', title: 'Opening week', body: 'What we learned in the first seven days' },
          { layout: 'obj', title: 'What sold', body: ['Flat whites, by a distance', 'Cinnamon buns by eleven', 'Very little tea'] },
          { layout: 'obj', title: 'What we change', body: ['A second grinder', 'Bake twice a day', 'Open at seven'] },
        ],
      }),
    );
    d.setNotes(0, 'Thank everyone for coming. Keep this short: two minutes, then the numbers.');
    d.setNotes(1, 'Flat whites were 41% of sales. Buns sold out every day before noon.');
    d.setNotes(2, 'Ask for a decision on the grinder today.');
    fs.writeFileSync(at('Opening week.pptx'), Buffer.from(d.save()));
  }

  /* PDF ----------------------------------------------------------------- */

  // An application form with fields to fill (the PDF tests' own form).
  const { formPdf } = require(path.join(repo, 'packages/pdf/test/form-fixture.js'));
  fs.writeFileSync(at('Application form.pdf'), formPdf());

  // A scanned letter: a picture of printed words and nothing to select,
  // until Recognise Text reads it.
  {
    const png = decodePng(fs.readFileSync(ocrFixture('eng-serif.png')));
    const rgb = Buffer.alloc(png.width * png.height * 3);
    for (let i = 0, o = 0; i < png.data.length; i += 4, o += 3) {
      rgb[o] = png.data[i];
      rgb[o + 1] = png.data[i + 1];
      rgb[o + 2] = png.data[i + 2];
    }
    const doc = new pdf.PdfDocument({ title: 'Scanned letter', created: '2026-10-13T09:00:00Z' });
    const ref = doc.addImage({ width: png.width, height: png.height, data: rgb, colorSpace: 'DeviceRGB' });
    doc.addPage().image(ref, 36, 48, (png.width / 150) * 72, (png.height / 150) * 72);
    fs.writeFileSync(at('Scanned letter.pdf'), doc.toBuffer());
  }
}
