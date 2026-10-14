// What the windows make of a document read (record.js): an expense row from
// a receipt or an invoice in each of the six languages, its figures as read
// only when the sums agree, the row as it goes into a sheet (a date a date,
// amounts numbers), Mail's note over an attached invoice, names for a form's
// tick boxes left to right and right to left, and the limits said for the
// languages chosen.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { documentFields } from '@rutba/ocr/fields';
import { documentRecord, recordCells, shownValue, dateSerial, invoiceNote, checkboxLabels, ocrLimits, RECORD_FIELDS } from '@rutba/ocr/record';

/** Words from rows of [text, x] pieces, as fields.test.js lays them. */
function page(rows, { height = 24, rtl = false, width = 1600, pageHeight = 2200, languages = ['eng'] } = {}) {
  const words = [];
  const lines = [];
  rows.forEach(([y, pieces]) => {
    const first = words.length;
    for (const [text, x] of pieces) {
      let at = x;
      for (const t of text.split(' ')) {
        const w = [...t].length * 12;
        const wx = rtl ? at - w : at;
        words.push({ text: t, box: { x: wx, y, width: w, height }, confidence: 95 });
        at = rtl ? wx - 12 : wx + w + 12;
      }
    }
    lines.push({ words: [first, words.length] });
  });
  return { width, height: pageHeight, words, lines, languages };
}

const ENGLISH = [
  [60, [['Northgate Supplies Ltd', 100]]],
  [140, [['INVOICE', 100]]],
  [220, [['Invoice No.', 1000], ['INV-0042', 1300]]],
  [260, [['Date', 1000], ['14/10/2026', 1300]]],
  [300, [['Due date', 1000], ['13/11/2026', 1300]]],
  [400, [['Description', 100], ['Qty', 800], ['Unit price', 1000], ['Amount', 1300]]],
  [440, [['Printer paper', 100], ['2', 800], ['12.50', 1000], ['£25.00', 1300]]],
  [480, [['Desk lamp', 100], ['1', 800], ['40.00', 1000], ['£40.00', 1300]]],
  [660, [['Subtotal', 1000], ['£65.00', 1300]]],
  [700, [['VAT at 20%', 1000], ['£13.00', 1300]]],
  [740, [['Total due', 1000], ['£78.00', 1300]]],
];

/** A receipt in each language, its figures in Western digits, as the window's language and English read it. */
const PAGES = {
  eng: page(ENGLISH),
  ara: page([
    [100, [['مؤسسة النخيل التجارية', 1450]]],
    [160, [['فاتورة', 1450]]],
    [200, [['رقم الفاتورة:', 1450], ['675078', 700]]],
    [240, [['التاريخ:', 1450], ['14/10/2026', 700]]],
    [400, [['الوصف', 1450], ['الكمية', 1000], ['المبلغ', 500]]],
    [440, [['ورق طباعة', 1450], ['2', 1000], ['50.00 ر.س', 500]]],
    [480, [['حبر أسود', 1450], ['1', 1000], ['50.00 ر.س', 500]]],
    [560, [['المجموع الفرعي:', 900], ['100.00 ر.س', 500]]],
    [600, [['ضريبة القيمة المضافة 15%:', 900], ['15.00 ر.س', 500]]],
    [640, [['الإجمالي:', 900], ['115.00 ر.س', 500]]],
  ], { rtl: true, languages: ['ara', 'eng'] }),
  urd: page([
    [100, [['الفلاح ٹریڈرز', 1450]]],
    [160, [['رسید', 1450]]],
    [200, [['رسید نمبر:', 1450], ['448502', 700]]],
    [240, [['تاریخ:', 1450], ['02/06/2026', 700]]],
    [560, [['ذیلی کل:', 900], ['Rs 1,000.00', 500]]],
    [600, [['سیلز ٹیکس 17%:', 900], ['Rs 170.00', 500]]],
    [640, [['کل رقم:', 900], ['Rs 1,170.00', 500]]],
  ], { rtl: true, languages: ['urd', 'eng'] }),
  hin: page([
    [60, [['शर्मा ट्रेडर्स', 100]]],
    [140, [['बीजक', 100]]],
    [220, [['बीजक संख्या', 1000], ['2026-0137', 1300]]],
    [260, [['दिनांक', 1000], ['14/10/2026', 1300]]],
    [660, [['उप-योग', 1000], ['₹2,000.00', 1300]]],
    [700, [['जीएसटी 18%', 1000], ['₹360.00', 1300]]],
    [740, [['कुल राशि', 1000], ['₹2,360.00', 1300]]],
  ], { languages: ['hin', 'eng'] }),
  chi_sim: page([
    [60, [['北京华星商贸有限公司', 100]]],
    [140, [['发票', 100]]],
    [220, [['发票号码', 1000], ['20261014', 1300]]],
    [260, [['开票日期', 1000], ['2026年10月14日', 1300]]],
    [660, [['小计', 1000], ['¥1,000.00', 1300]]],
    [700, [['税额', 1000], ['¥130.00', 1300]]],
    [740, [['合计', 1000], ['¥1,130.00', 1300]]],
  ], { languages: ['chi_sim', 'eng'] }),
  jpn: page([
    [60, [['さくら印刷株式会社', 100]]],
    [140, [['請求書', 100]]],
    [220, [['請求書番号', 1000], ['2026-0137', 1300]]],
    [260, [['発行日', 1000], ['2026年10月14日', 1300]]],
    [660, [['小計', 1000], ['¥14,400', 1300]]],
    [700, [['消費税', 1000], ['¥1,440', 1300]]],
    [740, [['合計', 1000], ['¥15,840', 1300]]],
  ], { languages: ['jpn', 'eng'] }),
};

const EXPECTED = {
  eng: { date: '2026-10-14', supplier: 'Northgate Supplies Ltd', number: 'INV-0042', subtotal: 65, tax: 13, total: 78, currency: 'GBP' },
  ara: { date: '2026-10-14', supplier: 'مؤسسة النخيل التجارية', number: '675078', subtotal: 100, tax: 15, total: 115, currency: 'SAR' },
  urd: { date: '2026-06-02', number: '448502', subtotal: 1000, tax: 170, total: 1170, currency: 'PKR' },
  hin: { date: '2026-10-14', supplier: 'शर्मा ट्रेडर्स', number: '2026-0137', subtotal: 2000, tax: 360, total: 2360, currency: 'INR' },
  chi_sim: { date: '2026-10-14', number: '20261014', subtotal: 1000, tax: 130, total: 1130, currency: 'CNY' },
  jpn: { date: '2026-10-14', number: '2026-0137', subtotal: 14400, tax: 1440, total: 15840, currency: 'JPY' },
};

for (const [lang, read] of Object.entries(PAGES)) {
  test(`documentRecord: a ${lang} bill gives one row, its figures as read where its sums agree`, () => {
    const row = documentRecord(documentFields(read));
    for (const [k, v] of Object.entries(EXPECTED[lang])) assert.equal(row.values[k], v, `${lang} ${k}`);
    assert.equal(row.consistent, true, `${lang}: the sums agree`);
    assert.equal(row.check.subtotal || row.check.tax || row.check.total, false, `${lang}: nothing of the figures to check`);
    assert.equal(row.found, true);
    assert.ok(row.boxes.total && row.boxes.total.width > 0, `${lang}: the total has its box`);
  });
}

test('documentRecord: figures on a page whose sums do not agree are marked to check, and a missing total too', () => {
  const wrong = page(ENGLISH.map(([y, p]) => [y, p.map(([t, x]) => [t === '£78.00' ? '£87.00' : t, x])]));
  const row = documentRecord(documentFields(wrong));
  assert.equal(row.consistent, false);
  assert.equal(row.values.total, 87);
  assert.deepEqual([row.check.subtotal, row.check.tax, row.check.total], [true, true, true]);
  assert.equal(row.check.number, false, 'a number read surely is not marked');
  const none = documentRecord(documentFields({ words: [], lines: [] }));
  assert.equal(none.found, false);
  assert.equal(none.check.total, true, 'a total not found is marked');
  assert.equal(none.check.subtotal, false, 'an empty subtotal is not');
  assert.deepEqual(Object.keys(none.values), RECORD_FIELDS);
  assert.equal(documentRecord(null).kind, 'other');
});

test('documentRecord: a date or a number the reading was unsure of is marked to check', () => {
  const read = page(ENGLISH);
  for (const w of read.words) if (w.text === 'INV-0042') w.confidence = 52;
  const row = documentRecord(documentFields(read));
  assert.equal(row.values.number, 'INV-0042');
  assert.equal(row.check.number, true);
  assert.equal(row.check.date, false);
});

test('recordCells: a date a date, amounts numbers with their decimals, the number kept as text', () => {
  assert.equal(dateSerial('2026-10-14'), 46309);
  assert.equal(dateSerial('1900-03-01'), 61);
  assert.equal(dateSerial('not a date'), null);
  const cells = recordCells({ date: '2026-10-14', supplier: 'Harbour Print Co.', number: '00342', subtotal: '181.08', tax: '9.05', total: '1,190.13', currency: 'USD' });
  assert.deepEqual(cells, [
    { value: 46309, format: 'yyyy-mm-dd' },
    { value: 'Harbour Print Co.', text: true },
    { value: '00342', text: true },
    { value: 181.08, format: '#,##0.00' },
    { value: 9.05, format: '#,##0.00' },
    { value: 1190.13, format: '#,##0.00' },
    { value: 'USD', text: true },
  ]);
  const yen = recordCells({ date: '14/10/2026', total: '¥15,840', currency: 'jpy' });
  assert.deepEqual(yen[0], { value: 46309, format: 'yyyy-mm-dd' }, 'a corrected date in another form is still a date');
  assert.deepEqual(yen[5], { value: 15840, format: '#,##0' });
  assert.deepEqual(yen[3], { value: null });
  const odd = recordCells({ date: 'last Tuesday', total: 'unclear' });
  assert.deepEqual([odd[0], odd[5]], [{ value: 'last Tuesday', text: true }, { value: 'unclear', text: true }]);
  assert.equal(shownValue('total', 78, 'GBP'), '78.00');
  assert.equal(shownValue('total', 15840, 'JPY'), '15840');
  assert.equal(shownValue('date', '2026-10-14'), '2026-10-14');
  assert.equal(shownValue('number', null), '');
});

test('invoiceNote: who it is from, the total, the due date and the number of an invoice; nothing for a letter', () => {
  const note = invoiceNote(documentFields(PAGES.eng));
  assert.equal(note.kind, 'invoice');
  assert.equal(note.from.value, 'Northgate Supplies Ltd');
  assert.equal(note.total.value, 78);
  assert.equal(note.total.sure, true);
  assert.equal(note.dueDate.value, '2026-11-13');
  assert.equal(note.number.value, 'INV-0042');
  assert.equal(note.currency, 'GBP');
  for (const k of ['from', 'total', 'dueDate', 'number']) assert.ok(note[k].box.width > 0, `${k} has its box`);
  const letter = page([
    [60, [['14 October 2026', 1200]]],
    [140, [['Dear Ms Patel,', 100]]],
    [200, [['Thank you for your letter of last month about the new office and the plans for the opening.', 100]]],
    [400, [['Yours sincerely,', 100]]],
  ]);
  assert.equal(invoiceNote(documentFields(letter)), null);
  assert.equal(invoiceNote(null), null);
  const urdu = invoiceNote(documentFields(PAGES.urd));
  assert.equal(urdu.kind, 'receipt');
  assert.equal(urdu.total.value, 1170);
});

test('checkboxLabels: the words beside each box, up to the next box, right to left too; Check Box n when none', () => {
  const box = (x, y) => ({ box: { x, y, width: 24, height: 24 }, checked: false });
  const words = [
    { text: 'New', box: { x: 140, y: 100, width: 40, height: 22 } },
    { text: 'Renewal', box: { x: 330, y: 100, width: 80, height: 22 } },
    { text: 'Change', box: { x: 530, y: 100, width: 70, height: 22 } },
    { text: 'of', box: { x: 610, y: 100, width: 20, height: 22 } },
    { text: 'address', box: { x: 640, y: 100, width: 70, height: 22 } },
    { text: 'Far', box: { x: 1400, y: 100, width: 40, height: 22 } },
    { text: 'New', box: { x: 140, y: 200, width: 40, height: 22 } },
  ];
  const labels = checkboxLabels([box(100, 100), box(290, 100), box(490, 100), box(100, 200), box(100, 300)], words);
  assert.deepEqual(labels, ['New', 'Renewal', 'Change of address', 'New 2', 'Check Box 5']);
  const rtl = checkboxLabels([box(1000, 100), box(700, 100)], [
    { text: 'جدید', box: { x: 930, y: 100, width: 50, height: 22 } },
    { text: 'تجدید', box: { x: 630, y: 100, width: 60, height: 22 } },
  ], { rtl: true });
  assert.deepEqual(rtl, ['جدید', 'تجدید']);
});

test('ocrLimits: Arabic-Indic and Urdu digits where Arabic or Urdu is read, Nastaliq where Urdu is', () => {
  assert.deepEqual(ocrLimits(['eng']), []);
  assert.deepEqual(ocrLimits(['hin', 'eng']), []);
  assert.deepEqual(ocrLimits(['ara', 'eng']), ['digits']);
  assert.deepEqual(ocrLimits(['urd', 'eng']), ['digits', 'nastaliq']);
});

test('on the pool: an invoice the suite drew gives its row with sums that agree, and a form its tick boxes by name', async () => {
  const { createOcrHost } = await import('@rutba/ocr');
  const fs = await import('node:fs');
  const host = createOcrHost({ threads: 2 });
  try {
    const fixture = (name) => fs.readFileSync(new URL(`./fixtures/${name}`, import.meta.url));
    const truth = JSON.parse(fixture('invoice-eng.json'));
    const [invoice, form] = await Promise.all([
      host.recognize(fixture('invoice-eng.png'), { languages: ['eng'], document: true }),
      host.recognize(fixture('form-eng.png'), { languages: ['eng'], document: true }),
    ]);
    const row = documentRecord(invoice.document);
    assert.equal(row.kind, 'invoice');
    assert.equal(row.consistent, true, 'the sums agree');
    for (const k of ['number', 'date', 'subtotal', 'tax', 'total', 'currency']) assert.equal(row.values[k], truth.fields[k], k);
    // The engine boxes the name's first two words far taller than "Ltd" (a line of the address taken in), so the
    // rows keep it apart: the supplier comes back short of it, as a person sees in the preview and puts right.
    assert.equal(row.values.supplier, 'Northgate Supplies');
    assert.deepEqual(RECORD_FIELDS.filter((f) => row.check[f]), [], 'nothing to check');
    const note = invoiceNote(invoice.document);
    assert.equal(note.dueDate.value, truth.fields.dueDate);
    assert.equal(form.document.kind, 'form');
    const boxes = form.document.checkboxes;
    assert.equal(boxes.length, 6);
    assert.deepEqual(checkboxLabels(boxes, form.document.words), ['New', 'Renewal', 'Change', 'Email', 'Phone', 'Post']);
    const want = JSON.parse(fixture('form-eng.json')).checkboxes;
    assert.deepEqual(boxes.map((b) => b.checked), want.map((c) => c.checked));
  } finally {
    await host.stop();
  }
});
