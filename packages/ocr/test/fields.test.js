// The fields of bills, invoices and receipts (fields.js), proved on words
// laid out here with known answers: amounts and dates in each language's
// way of writing them, labels beside and above their values, left to right
// and right to left, a page turned a little, the sums choosing between two
// readings of a figure, and the guess at a page's kind.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { documentFields, parseAmount, parseDate, parseNumber, parsePercent, westernDigits } from '@rutba/ocr/fields';

/**
 * Words from rows of text: each row a list of [text, x] pieces at a height,
 * a piece's words 12 px a letter and a space apart. Right to left, a piece
 * is laid from its x leftwards, its first word rightmost.
 */
function page(rows, { height = 24, rtl = false, slope = 0, width = 1600, pageHeight = 2200 } = {}) {
  const words = [];
  const lines = [];
  rows.forEach(([y, pieces]) => {
    const first = words.length;
    for (const [text, x] of pieces) {
      let at = x;
      for (const t of text.split(' ')) {
        const w = [...t].length * 12;
        const wx = rtl ? at - w : at;
        words.push({ text: t, box: { x: wx, y: Math.round(y + slope * (wx + w / 2)), width: w, height }, confidence: 95 });
        at = rtl ? wx - 12 : wx + w + 12;
      }
    }
    lines.push({ words: [first, words.length] });
  });
  return { width, height: pageHeight, words, lines, languages: ['eng'] };
}

const value = (r, k) => r.fields[k]?.value ?? null;

test('westernDigits: Arabic-Indic, Urdu, Devanagari and full-width digits, and the Arabic separators', () => {
  assert.equal(westernDigits('١٬٢٣٤٫٥٦'), '1,234.56');
  assert.equal(westernDigits('۱۲۳۴'), '1234');
  assert.equal(westernDigits('१२,३४५.००'), '12,345.00');
  assert.equal(westernDigits('１２３'), '123');
});

test('parseAmount: separators by locale, currency, and what is not an amount', () => {
  const v = (t) => parseAmount(t)?.value ?? null;
  assert.equal(v('£1,234.56'), 1234.56);
  assert.equal(v('1.234,56 €'), 1234.56);
  assert.equal(v('1 234,56'), 1234.56);
  assert.equal(v('₹1,23,456.00'), 123456);
  assert.equal(v('١٬٢٣٤٫٥٦ ر.س'), 1234.56);
  assert.equal(v('Rs ۱,۲۳۴.۵۰'), 1234.5);
  assert.equal(v('¥12,800'), 12800);
  assert.equal(v('12,800円'), 12800);
  assert.equal(v('¥1, 582. 70'), 1582.7, 'the engine spaced the separators');
  assert.equal(v('4 x ¥450   ¥1,800'), 1800, 'the last figure');
  assert.equal(v('20%'), null, 'a rate');
  assert.equal(v('VAT at 20% £750.81'), 750.81);
  assert.equal(v('no figures'), null);
  assert.equal(parseAmount('£1,234.56').currency, 'GBP');
  assert.equal(parseAmount('١٬٢٣٤٫٥٦ ر.س').currency, 'SAR');
  assert.equal(parseAmount('₹500').currency, 'INR');
  assert.equal(parseAmount('1,234.56').decimals, true);
  assert.equal(parseAmount('12345').decimals, false);
  assert.equal(parsePercent('ضريبة ١٥٪'), null, 'the Arabic percent sign is not read as a rate here');
  assert.equal(parsePercent('Sales tax 17.5 %'), 17.5);
});

test('parseDate: year first, day first, month names in four languages, 年月日', () => {
  assert.equal(parseDate('2026-10-14'), '2026-10-14');
  assert.equal(parseDate('2026/10/14'), '2026-10-14');
  assert.equal(parseDate('2026年10月14日'), '2026-10-14');
  assert.equal(parseDate('2026 年 5 月 9'), '2026-05-09', 'the engine lost the 日');
  assert.equal(parseDate('14/10/2026'), '2026-10-14');
  assert.equal(parseDate('04/10/2026'), '2026-10-04', 'day first');
  assert.equal(parseDate('04/10/2026', { monthFirst: true }), '2026-04-10');
  assert.equal(parseDate('10/14/2026'), '2026-10-14', 'month first when the first cannot be a month');
  assert.equal(parseDate('14.10.26'), '2026-10-14');
  assert.equal(parseDate('14 Oct 2026'), '2026-10-14');
  assert.equal(parseDate('October 14, 2026'), '2026-10-14');
  assert.equal(parseDate('١٤ أكتوبر ٢٠٢٦'), '2026-10-14');
  assert.equal(parseDate('14 اکتوبر 2026'), '2026-10-14');
  assert.equal(parseDate('14 अक्टूबर 2026'), '2026-10-14');
  assert.equal(parseDate('१४/१०/२०२६'), '2026-10-14');
  assert.equal(parseDate('31/02/2026'), null, 'no such day');
  assert.equal(parseDate('Thank you'), null);
});

test('parseNumber: the first run holding a digit, three characters or more', () => {
  assert.equal(parseNumber(': INV-0042'), 'INV-0042');
  assert.equal(parseNumber('# 2026-0137'), '2026-0137');
  assert.equal(parseNumber('No. ٤٤٨٥٠٢'), '448502');
  assert.equal(parseNumber('A1'), null);
});

/** The Word gallery's invoice, as words: number and dates under their labels, sums beside theirs. */
const invoiceRows = [
  [100, [['Northgate Supplies Ltd', 160]]],
  [104, [['INVOICE', 1200]]],
  [140, [['12 High Street, Leeds', 160]]],
  [300, [['BILL TO', 160], ['INVOICE NUMBER', 900], ['DATE', 1300]]],
  [340, [['Riverside Clinic', 160], ['INV-0042', 900], ['14 Oct 2026', 1300]]],
  [380, [['PAYMENT DUE', 900]]],
  [420, [['13 Nov 2026', 900]]],
  [520, [['Description', 160], ['Quantity', 800], ['Unit price', 1000], ['Amount', 1300]]],
  [560, [['Printer paper', 160], ['2', 800], ['£12.50', 1000], ['£25.00', 1300]]],
  [600, [['Desk lamp', 160], ['1', 800], ['£40.00', 1000], ['£40.00', 1300]]],
  [660, [['Subtotal', 1000], ['£65.00', 1300]]],
  [700, [['VAT at 20%', 1000], ['£13.00', 1300]]],
  [740, [['Total due', 1000], ['£78.00', 1300]]],
  [840, [['Thank you for your business.', 160]]],
];

test('documentFields: an invoice laid out as the Word template lays it', () => {
  const r = documentFields(page(invoiceRows));
  assert.equal(r.kind, 'invoice');
  assert.equal(value(r, 'supplier'), 'Northgate Supplies Ltd');
  assert.equal(value(r, 'number'), 'INV-0042');
  assert.equal(value(r, 'date'), '2026-10-14');
  assert.equal(value(r, 'dueDate'), '2026-11-13');
  assert.equal(value(r, 'subtotal'), 65);
  assert.equal(value(r, 'tax'), 13);
  assert.equal(value(r, 'total'), 78);
  assert.equal(value(r, 'taxRate'), 20);
  assert.equal(r.currency, 'GBP');
  assert.deepEqual(r.items.map((i) => [i.description, i.quantity, i.unitPrice, i.amount]), [['Printer paper', 2, 12.5, 25], ['Desk lamp', 1, 40, 40]]);
  assert.deepEqual(r.checks, { itemsAddUp: true, totalAddsUp: true, taxRateHolds: true, consistent: true });
  assert.equal(r.fields.total.checked, true);
  assert.ok(r.fields.total.box.x >= 1300 && r.fields.total.box.y === 740, 'the total is boxed where it was read');
});

test('documentFields: the same page turned a little still pairs labels with values along their rows', () => {
  const r = documentFields(page(invoiceRows, { slope: 0.03 }));
  assert.ok(Math.abs(r.slope - 0.03) < 0.002, `levelled by ${r.slope}`);
  assert.equal(value(r, 'total'), 78);
  assert.equal(value(r, 'number'), 'INV-0042');
  assert.equal(r.checks.consistent, true);
});

test('documentFields: sums that do not agree are said so, and a lost decimal point is put back when the sums want it', () => {
  const wrong = invoiceRows.map(([y, p]) => [y, p.map(([t, x]) => [t === '£78.00' ? '£87.00' : t, x])]);
  const r = documentFields(page(wrong));
  assert.equal(r.checks.totalAddsUp, false);
  assert.equal(r.checks.consistent, false);
  assert.equal(r.fields.total.checked, false);
  const lost = invoiceRows.map(([y, p]) => [y, p.map(([t, x]) => [t === '£78.00' ? '£7800' : t, x])]);
  const back = documentFields(page(lost));
  assert.equal(value(back, 'total'), 78);
  assert.equal(back.fields.total.repaired, true);
  assert.equal(back.checks.consistent, true);
});

test('documentFields: the sums choose between a figure as first read and as read again', () => {
  const misread = invoiceRows.map(([y, p]) => [y, p.map(([t, x]) => [t === '£13.00' ? '£18.00' : t, x])]);
  const read = page(misread);
  const collect = [];
  documentFields(read, { collect });
  assert.ok(collect.length >= 5, 'the figures beside the labels and the items\' amounts are asked for');
  // The second reading of every box: the true figure where the first was wrong.
  const tax = read.words.find((w) => w.text === '£18.00').box;
  const again = collect.map((box) => ({ box, readings: [{ text: box.x <= tax.x && box.x + box.width >= tax.x + tax.width && box.y <= tax.y && box.y + box.height >= tax.y + tax.height ? '£13.00' : '', confidence: 90 }] }));
  const r = documentFields(read, { again });
  assert.equal(value(r, 'tax'), 13);
  assert.equal(r.checks.consistent, true);
});

test('documentFields: a value the reading passed over is found by reading its row again', () => {
  const dropped = invoiceRows.filter(([y]) => y !== 740).concat([[740, [['Total due', 1000]]]]).sort((a, b) => a[0] - b[0]);
  const read = page(dropped);
  const collect = [];
  documentFields(read, { collect });
  const row = collect.find((b) => b.y <= 740 && b.y + b.height >= 764 && b.x >= 1000);
  assert.ok(row, 'the row beyond the label is asked for');
  const r = documentFields(read, { again: [{ box: row, readings: [{ text: '£78.00', confidence: 80 }] }] });
  assert.equal(value(r, 'total'), 78);
  assert.equal(r.checks.totalAddsUp, true);
});

test('documentFields: an Arabic invoice right to left, Arabic-Indic digits, labels with colons', () => {
  const read = page([
    [100, [['مؤسسة النخيل التجارية', 1450]]],
    [200, [['رقم الفاتورة:', 1450], ['٦٧٥٠٧٨', 700]]],
    [240, [['التاريخ:', 1450], ['١٤/١٠/٢٠٢٦', 700]]],
    [280, [['تاريخ الاستحقاق:', 1450], ['١٣/١١/٢٠٢٦', 700]]],
    [400, [['الوصف', 1450], ['الكمية', 1000], ['المبلغ', 500]]],
    [440, [['ورق طباعة', 1450], ['٢', 1000], ['٥٠٫٠٠ ر.س', 500]]],
    [480, [['حبر أسود', 1450], ['١', 1000], ['٥٠٫٠٠ ر.س', 500]]],
    [560, [['المجموع الفرعي:', 900], ['١٠٠٫٠٠ ر.س', 500]]],
    [600, [['ضريبة القيمة المضافة ١٥%:', 900], ['١٥٫٠٠ ر.س', 500]]],
    [640, [['الإجمالي:', 900], ['١١٥٫٠٠ ر.س', 500]]],
  ], { rtl: true });
  read.languages = ['ara', 'eng'];
  const r = documentFields(read);
  assert.equal(r.kind, 'invoice');
  assert.equal(value(r, 'supplier'), 'مؤسسة النخيل التجارية');
  assert.equal(value(r, 'number'), '675078');
  assert.equal(value(r, 'date'), '2026-10-14');
  assert.equal(value(r, 'dueDate'), '2026-11-13');
  assert.equal(value(r, 'subtotal'), 100);
  assert.equal(value(r, 'tax'), 15);
  assert.equal(value(r, 'total'), 115);
  assert.equal(value(r, 'taxRate'), 15);
  assert.equal(r.currency, 'SAR');
  assert.deepEqual(r.items.map((i) => i.amount), [50, 50]);
  assert.equal(r.checks.consistent, true);
});

test('documentFields: a till roll with no heading, its items above the subtotal, read as a receipt', () => {
  const read = page([
    [40, [['Meadow Lane Bakery', 120]]],
    [80, [['RECEIPT', 220]]],
    [120, [['Receipt No: 448502', 20], ['', 0]]],
    [160, [['Date: 02/06/2026 14:32', 20]]],
    [200, [['Sourdough loaf', 20], ['$4.50', 480]]],
    [240, [['2 x $1.75', 20]]],
    [242, [['$3.50', 480]]],
    [280, [['Coffee', 20], ['$2.80', 480]]],
    [340, [['Subtotal', 20], ['$10.80', 480]]],
    [380, [['Tax', 20], ['$0.54', 480]]],
    [420, [['TOTAL', 20], ['$11.34', 480]]],
    [480, [['Cash', 20], ['$20.00', 480]]],
    [520, [['Change', 20], ['$8.66', 480]]],
  ].map(([y, p]) => [y, p.filter(([t]) => t)]), { width: 630, pageHeight: 1100 });
  const r = documentFields(read);
  assert.equal(r.kind, 'receipt');
  assert.equal(value(r, 'supplier'), 'Meadow Lane Bakery');
  assert.equal(value(r, 'number'), '448502');
  assert.equal(value(r, 'date'), '2026-06-02');
  assert.equal(value(r, 'total'), 11.34, 'not the cash handed over');
  assert.deepEqual(r.items.map((i) => i.amount), [4.5, 3.5, 2.8]);
  assert.equal(r.checks.itemsAddUp, true);
  assert.equal(r.checks.totalAddsUp, true);
});

test('documentFields: Japanese labels, a yen total with no decimals, 年月日', () => {
  const read = page([
    [60, [['さくら印刷株式会社', 100]]],
    [140, [['請求書番号', 100], ['2026-0137', 700]]],
    [180, [['発行日', 100], ['2026年10月14日', 700]]],
    [300, [['小計', 600], ['¥14,400', 900]]],
    [340, [['消費税', 600], ['¥1,440', 900]]],
    [380, [['合計', 600], ['¥15,840', 900]]],
  ]);
  read.languages = ['jpn', 'eng'];
  const r = documentFields(read);
  assert.equal(value(r, 'number'), '2026-0137');
  assert.equal(value(r, 'date'), '2026-10-14');
  assert.equal(value(r, 'total'), 15840);
  assert.equal(r.currency, 'JPY');
  assert.equal(r.checks.totalAddsUp, true);
});

test('documentFields: a letter and a form are told from a bill, and a page of nothing much is other', () => {
  const letter = page([
    [60, [['Kestrel Stationers', 900]]],
    [200, [['14 October 2026', 100]]],
    [300, [['Dear Ms Patel,', 100]]],
    [360, [['Thank you for your letter of last month about the delivery of the paper and the ink you ordered.', 100]]],
    [400, [['We have looked into it and the parcel left our store on the morning of the ninth as planned.', 100]]],
    [440, [['A second parcel with the missing folders is on its way to you and should arrive this week.', 100]]],
    [520, [['Yours sincerely,', 100]]],
  ]);
  assert.equal(documentFields(letter).kind, 'letter');
  const form = page([
    [60, [['Application form', 500]]],
    [160, [['Name:', 100], ['Jane Patel', 400]]],
    [200, [['Address:', 100]]],
    [240, [['Date of birth:', 100]]],
    [280, [['Signature:', 100]]],
  ]);
  assert.equal(documentFields(form).kind, 'form');
  assert.equal(documentFields(page([[60, [['Holiday photos 2026', 100]]]])).kind, 'other');
  assert.equal(documentFields({ words: [], lines: [] }).kind, 'other');
});

test('recognize with document: a receipt the suite drew comes back as its fields, its items and sums that agree, on the pool', async () => {
  const { createOcrHost } = await import('@rutba/ocr');
  const fs = await import('node:fs');
  const host = createOcrHost({ threads: 1 });
  try {
    const picture = fs.readFileSync(new URL('./fixtures/receipt-eng.png', import.meta.url));
    const r = await host.recognize(picture, { languages: ['eng'], document: true });
    const d = r.document;
    assert.equal(d.kind, 'receipt');
    assert.equal(value(d, 'supplier'), 'Harbour Print Co.');
    assert.equal(value(d, 'number'), '342457');
    assert.equal(value(d, 'date'), '2025-09-04');
    assert.equal(value(d, 'subtotal'), 181.08);
    assert.equal(value(d, 'tax'), 9.05);
    assert.equal(value(d, 'total'), 190.13);
    assert.equal(d.currency, 'USD');
    assert.deepEqual(d.items.map((i) => i.amount), [80.31, 32.27, 38.87, 5.02, 13.18, 11.43]);
    assert.equal(d.checks.consistent, true);
    assert.ok(d.words.length >= r.words.length, 'the words of both readings');
    assert.ok(d.again.length >= 3 && d.again.every((a) => a.readings.length), 'the figures were read again');
    const box = d.fields.total.box;
    assert.ok(box.x > 315 && box.y > 850 && box.y < 960, `the total is boxed where it is printed: ${JSON.stringify(box)}`);
  } finally {
    await host.stop();
  }
});
