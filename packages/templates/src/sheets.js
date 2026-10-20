// The gallery's workbooks, written through the suite's own sheet engine.
//
// Each starts as a workbook File > New would make from a specification
// (buildXlsx: the cells, their formulas and their looks), and is then opened
// in the sheet engine the Worksheets window edits with and given what a
// person would give it there: a theme from Page Layout, column widths, frozen
// headings, a table, rules that colour a cell, lists to choose from. A chart
// is anchored where it belongs, its figures the ones the engine worked out
// from the formulas, so it reads the same before and after a recalculation.

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { themeById } from '@rutba/office-formats/themes';

const MONEY = '£#,##0.00';
const MONEY0 = '£#,##0';
const MONEY_SIGNED = '£#,##0;[Red]-£#,##0';
const PERCENT = '0.0%';

/** Column letters to index, "A" is 0. */
const col = (letters) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

/** A sheet spec's rows, filled from a map of "A1" to value. */
function grid(cells) {
  const rows = [];
  for (const [ref, value] of Object.entries(cells)) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    const r = Number(m[2]) - 1;
    const c = col(m[1]);
    while (rows.length <= r) rows.push([]);
    const row = rows[r];
    while (row.length <= c) row.push(null);
    row[c] = value;
  }
  return rows;
}

/**
 * A workbook: the sheets built, opened in the engine, the figures every
 * chart plots read from its calculation, the workbook built again with
 * those charts anchored, then `finish` runs the window's own operations on
 * it (widths, panes, tables, rules) before it is saved.
 */
function compose({ theme, sheets, charts = () => [], finish = null }) {
  const first = SheetView.open(Buffer.from(buildXlsx({ sheets })), { viewportWidth: 1100, viewportHeight: 620 });
  const value = (sheet, ref) => {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    const v = first.calc.getValue(sheet, Number(m[2]) - 1, col(m[1]));
    return typeof v === 'number' && Number.isFinite(v) ? Math.round(v * 100) / 100 : v;
  };
  const values = (sheet, from, to) => {
    const a = /^([A-Z]+)(\d+)$/.exec(from);
    const b = /^([A-Z]+)(\d+)$/.exec(to);
    const out = [];
    for (let r = Number(a[2]); r <= Number(b[2]); r++) {
      for (let c = col(a[1]); c <= col(b[1]); c++) out.push(value(sheet, String.fromCharCode(65 + c) + r));
    }
    return out;
  };
  const withCharts = sheets.map((s) => ({ ...s, drawings: [...(s.drawings || []), ...charts({ value, values }).filter((c) => c.sheet === s.name).map(({ sheet, ...c }) => c)] }));
  const view = SheetView.open(Buffer.from(buildXlsx({ sheets: withCharts })), { viewportWidth: 1100, viewportHeight: 620 });
  // Page Layout > Themes: the workbook's theme, its colours and fonts.
  view.setWorkbookTheme({ theme });
  // Each sheet's title, in A1, gets a row tall enough for its size.
  for (const s of sheets) {
    view.selectSheet(s.name);
    view.setRowHeight(0, 40);
  }
  view.selectSheet(sheets[0].name);
  if (finish) finish(view);
  view.selectSheet(sheets[0].name);
  view.select(0, 0);
  return Buffer.from(view.save());
}

/** The window's own operations, on a view: a range selected, then something done to it. */
const ops = (view) => ({
  sheet: (name) => view.selectSheet(name),
  widths: (map) => { for (const [letters, px] of Object.entries(map)) view.setColWidth(col(letters), px); },
  height: (row, px) => view.setRowHeight(row - 1, px),
  range(ref) {
    const [a, b] = ref.split(':');
    const pa = /^([A-Z]+)(\d+)$/.exec(a);
    const pb = /^([A-Z]+)(\d+)$/.exec(b || a);
    view.select(Number(pa[2]) - 1, col(pa[1]));
    view.select(Number(pb[2]) - 1, col(pb[1]), { extend: true });
    return this;
  },
  format(ref, delta) { this.range(ref); view.setFormat(delta); },
  merge(ref) { this.range(ref); view.mergeSelection(); },
  freeze: (rows, cols = 0) => view.freezePanes(rows, cols),
  table(ref, style) { this.range(ref); view.formatAsTable({ style, stripes: true }); },
  rule(ref, spec) { this.range(ref); view.addConditionalRule(spec); },
  list(ref, items) { this.range(ref); view.addValidationRule({ kind: 'list', items }); },
});

/** The header row's look: white words on the theme's first accent. */
const header = (p) => ({ bold: true, fill: '#' + p.accent1, colour: '#FFFFFF', align: 'left' });
const title = (p) => ({ bold: true, size: 20, colour: '#' + p.dk2 });
const note = (p) => ({ colour: '#' + p.dk2 });
const total = (p, numFmt) => ({ bold: true, fill: '#' + p.lt2, numFmt, border: true });

/* ---- the templates ------------------------------------------------------- */

function budget() {
  const theme = 'meadow';
  const p = themeById(theme).palette;
  const spending = [
    ['Housing', 1100, 1100], ['Utilities', 180, 164], ['Groceries', 420, 455], ['Transport', 160, 138],
    ['Insurance', 90, 90], ['Phone and internet', 60, 60], ['Health', 50, 35], ['Going out', 120, 152],
    ['Savings', 400, 400], ['Other', 100, 82],
  ];
  const cells = {
    A1: 'Monthly budget', A2: 'Month', B2: 'January',
    A4: 'Income', B4: 'Planned', C4: 'Actual', D4: 'Difference',
    A5: 'Salary', B5: 3200, C5: 3200, D5: '=C5-B5',
    A6: 'Other income', B6: 300, C6: 250, D6: '=C6-B6',
    A7: 'Total income', B7: '=SUM(B5:B6)', C7: '=SUM(C5:C6)', D7: '=C7-B7',
    A9: 'Spending', B9: 'Planned', C9: 'Actual', D9: 'Difference',
    A20: 'Total spending', B20: '=SUM(B10:B19)', C20: '=SUM(C10:C19)', D20: '=B20-C20',
    A22: 'Left over', B22: '=B7-B20', C22: '=C7-C20', D22: '=C22-B22',
    A23: 'Saved this month', C23: '=IFERROR((C22+C18)/C7,0)',
    A25: 'Type the planned and actual figures in the green columns; the totals, the differences and the charts follow.',
  };
  spending.forEach(([name, plan, actual], i) => {
    const r = 10 + i;
    Object.assign(cells, { [`A${r}`]: name, [`B${r}`]: plan, [`C${r}`]: actual, [`D${r}`]: `=B${r}-C${r}` });
  });
  const sheets = [{
    name: 'Budget',
    rows: grid(cells),
    styles: {
      A1: title(p), A2: { bold: true, colour: '#' + p.dk2 }, B2: { bold: true },
      'A4:D4': header(p), 'A9:D9': header(p),
      'B5:D6': { numFmt: MONEY0 }, 'B10:D19': { numFmt: MONEY0 },
      'A7:D7': total(p, MONEY0), 'A20:D20': total(p, MONEY0),
      'D5:D7': { numFmt: MONEY_SIGNED }, 'D10:D19': { numFmt: MONEY_SIGNED },
      D7: total(p, MONEY_SIGNED), D20: total(p, MONEY_SIGNED),
      'A22:D22': { bold: true, numFmt: MONEY_SIGNED, border: true }, A22: { bold: true, border: true },
      A23: { colour: '#' + p.dk2 }, C23: { numFmt: PERCENT, bold: true },
      A25: { colour: '#' + p.dk2 },
    },
  }];
  const charts = ({ values }) => [
    {
      sheet: 'Budget', kind: 'chart', chartKind: 'column', title: 'Planned and actual spending',
      categories: { ref: 'Budget!$A$10:$A$19', values: spending.map((s) => s[0]) },
      series: [
        { name: 'Planned', nameRef: 'Budget!$B$9', ref: 'Budget!$B$10:$B$19', values: values('Budget', 'B10', 'B19') },
        { name: 'Actual', nameRef: 'Budget!$C$9', ref: 'Budget!$C$10:$C$19', values: values('Budget', 'C10', 'C19') },
      ],
      from: { col: 5, row: 1 }, to: { col: 13, row: 17 },
    },
    {
      sheet: 'Budget', kind: 'chart', chartKind: 'doughnut', title: 'Where the money went',
      categories: { ref: 'Budget!$A$10:$A$19', values: spending.map((s) => s[0]) },
      series: [{ name: 'Actual', nameRef: 'Budget!$C$9', ref: 'Budget!$C$10:$C$19', values: values('Budget', 'C10', 'C19') }],
      from: { col: 5, row: 18 }, to: { col: 13, row: 34 },
    },
  ];
  return compose({
    theme, sheets, charts,
    finish(view) {
      const o = ops(view);
      o.widths({ A: 190, B: 100, C: 100, D: 104, E: 24 });
      o.format('B5:C6', { fill: '#' + p.lt2 });
      o.format('B10:C19', { fill: '#' + p.lt2 });
      // Over budget is red, under it green, as the difference column reads.
      o.rule('D10:D19', { kind: 'cellIs', operator: 'lessThan', values: [0], look: 'red' });
      o.rule('D22', { kind: 'cellIs', operator: 'lessThan', values: [0], look: 'red' });
    },
  });
}

function invoiceSheet() {
  const theme = 'basalt';
  const p = themeById(theme).palette;
  const items = [
    ['Design consultation (hours)', 6, 65], ['Website pages', 5, 180], ['Photography, half day', 1, 350], ['Hosting, twelve months', 1, 120],
  ];
  const cells = {
    A1: 'INVOICE',
    A2: 'Your Company Ltd', A3: 'Street address, Town, Postcode', A4: 'hello@example.com  |  07000 000000',
    C2: 'Invoice number', D2: 'INV-0001', C3: 'Date', D3: '=TODAY()', C4: 'Payment due', D4: '=D3+30',
    A6: 'Bill to', A7: '[Customer name]', A8: '[Company]', A9: '[Street address, Town, Postcode]',
    A11: 'Description', B11: 'Quantity', C11: 'Unit price', D11: 'Amount',
    C21: 'Subtotal', D21: '=SUM(D12:D19)',
    C22: 'VAT rate', D22: 0.2,
    C23: 'VAT', D23: '=ROUND(D21*D22,2)',
    C24: 'Total due', D24: '=D21+D23',
    A26: 'Pay by bank transfer to Your Company Ltd, sort code 00-00-00, account 00000000, quoting the invoice number.',
    A27: 'Thank you for your business.',
  };
  for (let i = 0; i < 8; i++) {
    const r = 12 + i;
    const item = items[i];
    if (item) Object.assign(cells, { [`A${r}`]: item[0], [`B${r}`]: item[1], [`C${r}`]: item[2] });
    cells[`D${r}`] = `=IF(B${r}*C${r}=0,"",B${r}*C${r})`;
  }
  const sheets = [{
    name: 'Invoice',
    rows: grid(cells),
    styles: {
      A1: { bold: true, size: 26, colour: '#' + p.accent1 },
      A2: { bold: true, size: 12 }, 'A3:A4': note(p),
      'C2:C4': { bold: true, colour: '#' + p.accent1, align: 'right' }, D2: { align: 'right' }, 'D3:D4': { numFmt: 'd mmm yyyy', align: 'right' },
      A6: { bold: true, colour: '#' + p.accent1 },
      'A11:D11': header(p), 'B11:D11': { ...header(p), align: 'right' },
      'A12:D19': { border: true }, 'B12:B19': { border: true, align: 'right' }, 'C12:D19': { border: true, numFmt: MONEY },
      'C21:C24': { bold: true, align: 'right' }, D21: { numFmt: MONEY }, D22: { numFmt: '0%' }, D23: { numFmt: MONEY },
      'C24:D24': { bold: true, fill: '#' + p.lt2, numFmt: MONEY, align: 'right', border: true },
      A26: note(p), A27: { bold: true, colour: '#' + p.accent1 },
    },
  }];
  return compose({
    theme, sheets,
    finish(view) {
      const o = ops(view);
      o.widths({ A: 330, B: 90, C: 120, D: 120 });
    },
  });
}

function timesheet() {
  const theme = 'harbour';
  const p = themeById(theme).palette;
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const cells = {
    A1: 'Timesheet',
    A3: 'Employee', B3: '[Name]', A4: 'Week starting', B4: '=TODAY()-WEEKDAY(TODAY(),3)',
    D3: 'Hourly rate', E3: 15.5, D4: 'Standard hours', E4: 37.5, D5: 'Overtime rate', E5: 1.5,
    A7: 'Day', B7: 'Date', C7: 'Start', D7: 'Finish', E7: 'Break (minutes)', F7: 'Hours',
    A15: 'Total', F15: '=SUM(F8:F14)',
    A17: 'Standard hours', B17: '=MIN(F15,E4)', A18: 'Overtime hours', B18: '=MAX(0,F15-E4)',
    A19: 'Pay', B19: '=ROUND(B17*E3+B18*E3*E5,2)',
    A21: 'Type each day\'s start and finish as times (9:00, 17:30) and the break in minutes; the hours and the pay are worked out.',
  };
  days.forEach((day, i) => {
    const r = 8 + i;
    Object.assign(cells, {
      [`A${r}`]: day, [`B${r}`]: i ? `=B${r - 1}+1` : '=B4',
      [`F${r}`]: `=IF(OR(C${r}="",D${r}=""),0,ROUND((D${r}-C${r})*24-E${r}/60,2))`,
    });
    if (i < 5) Object.assign(cells, { [`C${r}`]: 0.375, [`D${r}`]: i === 4 ? 0.6875 : 0.729166666666667, [`E${r}`]: 30 });
  });
  const sheets = [{
    name: 'Timesheet',
    rows: grid(cells),
    styles: {
      A1: title(p), 'A3:A4': { bold: true, colour: '#' + p.dk2 }, 'D3:D5': { bold: true, colour: '#' + p.dk2 },
      B4: { numFmt: 'd mmm yyyy', align: 'left' }, E3: { numFmt: MONEY }, E4: { numFmt: '0.0' }, E5: { numFmt: '0.0"x"' },
      'A7:F7': header(p),
      'A8:A14': { bold: true, border: true }, 'B8:B14': { numFmt: 'ddd d mmm', border: true }, 'C8:D14': { numFmt: 'h:mm', border: true },
      'E8:E14': { border: true }, 'F8:F14': { numFmt: '0.00', border: true },
      'A15:F15': total(p, '0.00'),
      'A17:A19': { bold: true }, 'B17:B18': { numFmt: '0.00' }, B19: { numFmt: MONEY, bold: true },
      A21: note(p),
    },
  }];
  return compose({
    theme, sheets,
    finish(view) {
      const o = ops(view);
      o.widths({ A: 130, B: 120, C: 80, D: 120, E: 120, F: 80 });
      o.format('C8:E14', { fill: '#' + p.lt2 });
      o.freeze(7, 0);
    },
  });
}

function loan() {
  const theme = 'rutba';
  const p = themeById(theme).palette;
  const cells = {
    A1: 'Loan calculator',
    A3: 'Loan amount', B3: 200000, A4: 'Annual interest rate', B4: 0.045, A5: 'Term in years', B5: 25, A6: 'Payments a year', B6: 12,
    A8: 'Payment', B8: '=PMT(B4/B6,B5*B6,-B3)', A9: 'Number of payments', B9: '=B5*B6',
    A10: 'Total paid', B10: '=B8*B9', A11: 'Total interest', B11: '=B10-B3',
    A12: 'Change the figures in blue.',
    D3: 'Year', E3: 'Opening balance', F3: 'Paid', G3: 'Interest', H3: 'Principal', I3: 'Closing balance',
  };
  for (let y = 1; y <= 30; y++) {
    const r = 3 + y;
    const off = `D${r}>$B$5`;
    Object.assign(cells, {
      [`D${r}`]: y,
      [`E${r}`]: `=IF(${off},"",${y === 1 ? '$B$3' : `I${r - 1}`})`,
      [`F${r}`]: `=IF(${off},"",$B$8*$B$6)`,
      [`I${r}`]: `=IF(${off},"",MAX(0,ROUND(-FV($B$4/$B$6,D${r}*$B$6,-$B$8,$B$3),2)))`,
      [`H${r}`]: `=IF(${off},"",E${r}-I${r})`,
      [`G${r}`]: `=IF(${off},"",F${r}-H${r})`,
    });
  }
  const sheets = [{
    name: 'Loan',
    rows: grid(cells),
    styles: {
      A1: title(p), 'A3:A6': { bold: true, colour: '#' + p.dk2 },
      B3: { numFmt: MONEY0, colour: '#' + p.accent1, bold: true }, B4: { numFmt: '0.00%', colour: '#' + p.accent1, bold: true },
      'B5:B6': { colour: '#' + p.accent1, bold: true },
      'A8:B8': { bold: true, size: 14, fill: '#' + p.lt2 }, B8: { bold: true, size: 14, fill: '#' + p.lt2, numFmt: MONEY },
      'A9:A11': { bold: true }, 'B10:B11': { numFmt: MONEY },
      A12: note(p),
      'D3:I3': { ...header(p), wrap: true }, 'D4:D33': { align: 'center', border: true }, 'E4:I33': { numFmt: MONEY0, border: true },
    },
  }];
  const charts = ({ values }) => {
    const balances = values('Loan', 'I4', 'I28');
    return [{
      sheet: 'Loan', kind: 'chart', chartKind: 'column', title: 'Balance left at the end of each year',
      categories: { ref: 'Loan!$D$4:$D$28', values: Array.from({ length: 25 }, (_, i) => String(i + 1)) },
      series: [{ name: 'Closing balance', nameRef: 'Loan!$I$3', ref: 'Loan!$I$4:$I$28', values: balances }],
      from: { col: 0, row: 14 }, to: { col: 3, row: 30 },
    }];
  };
  return compose({
    theme, sheets, charts,
    finish(view) {
      const o = ops(view);
      o.widths({ A: 170, B: 120, C: 24, D: 56, E: 120, F: 100, G: 100, H: 100, I: 120 });
      o.height(3, 36);
      o.freeze(3, 0);
    },
  });
}

function todo() {
  const theme = 'studio';
  const p = themeById(theme).palette;
  const tasks = [
    ['Draft the plan', 'Sam', 2, 'High', 'Done', 'Shared on Monday'],
    ['Book the venue', 'Alex', 5, 'High', 'In progress', 'Two quotes in'],
    ['Order materials', 'Priya', 9, 'Medium', 'Not started', ''],
    ['Send invitations', 'Sam', 12, 'Medium', 'Not started', 'After the venue is booked'],
    ['Arrange catering', 'Jordan', 14, 'Low', 'Not started', ''],
    ['Confirm speakers', 'Alex', 16, 'High', 'In progress', 'Three of four agreed'],
  ];
  const cells = {
    A1: 'To-do list', A2: 'Done', B2: '=IFERROR(COUNTIF(E5:E16,"Done")/COUNTA(A5:A16),0)',
    A4: 'Task', B4: 'Owner', C4: 'Due', D4: 'Priority', E4: 'Status', F4: 'Notes',
    A18: 'Choose a priority and a status from the lists in their cells; finished tasks turn green.',
  };
  tasks.forEach((t, i) => {
    const r = 5 + i;
    Object.assign(cells, { [`A${r}`]: t[0], [`B${r}`]: t[1], [`C${r}`]: `=TODAY()+${t[2]}`, [`D${r}`]: t[3], [`E${r}`]: t[4], [`F${r}`]: t[5] });
  });
  const sheets = [{
    name: 'Tasks',
    rows: grid(cells),
    styles: { A1: title(p), A2: { bold: true, colour: '#' + p.dk2 }, B2: { numFmt: '0%', bold: true, colour: '#' + p.accent1 }, 'C5:C16': { numFmt: 'd mmm yyyy', align: 'left' }, A18: note(p) },
  }];
  return compose({
    theme, sheets,
    finish(view) {
      const o = ops(view);
      o.widths({ A: 220, B: 100, C: 110, D: 90, E: 110, F: 240 });
      o.table('A4:F16', 'TableStyleMedium2');
      o.freeze(4, 0);
      o.list('D5:D16', ['High', 'Medium', 'Low']);
      o.list('E5:E16', ['Not started', 'In progress', 'Done']);
      o.rule('E5:E16', { kind: 'containsText', text: 'Done', look: 'green' });
      o.rule('D5:D16', { kind: 'containsText', text: 'High', look: 'red' });
    },
  });
}

function ledger() {
  const theme = 'paperkite';
  const p = themeById(theme).palette;
  const entries = [
    ['Sales', 'Card takings, week 1', 1840, 0], ['Rent', 'Shop rent', 0, 950], ['Supplies', 'Stock order', 0, 612.4],
    ['Sales', 'Card takings, week 2', 2015.5, 0], ['Utilities', 'Electricity', 0, 138.2], ['Wages', 'Part-time help', 0, 720],
    ['Sales', 'Card takings, week 3', 1762, 0], ['Supplies', 'Packaging', 0, 96.75], ['Sales', 'Online orders', 488.3, 0], ['Other', 'Bank charges', 0, 12.5],
  ];
  const categories = ['Sales', 'Rent', 'Supplies', 'Utilities', 'Wages', 'Other'];
  const cells = {
    A1: 'Accounts ledger',
    A4: 'Date', B4: 'Category', C4: 'Description', D4: 'Money in', E4: 'Money out', F4: 'Balance',
    A5: '=TODAY()-30', B5: 'Opening', C5: 'Opening balance', F5: 2500,
    H4: 'Summary', I4: 'Amount',
    H5: 'Money in', I5: '=SUM(D6:D29)', H6: 'Money out', I6: '=SUM(E6:E29)', H7: 'Closing balance', I7: '=F5+I5-I6',
    H9: 'Category', I9: 'In', J9: 'Out',
    A31: 'Add a row for each payment in or out; the balance runs down column F and the summary adds up each category.',
  };
  categories.forEach((c, i) => {
    const r = 10 + i;
    Object.assign(cells, { [`H${r}`]: c, [`I${r}`]: `=SUMIF($B$6:$B$29,H${r},$D$6:$D$29)`, [`J${r}`]: `=SUMIF($B$6:$B$29,H${r},$E$6:$E$29)` });
  });
  for (let i = 0; i < 24; i++) {
    const r = 6 + i;
    const e = entries[i];
    if (e) Object.assign(cells, { [`A${r}`]: `=A5+${i * 3 + 1}`, [`B${r}`]: e[0], [`C${r}`]: e[1], [`D${r}`]: e[2] || null, [`E${r}`]: e[3] || null });
    cells[`F${r}`] = `=IF(AND(D${r}="",E${r}=""),"",$F$5+SUM($D$6:D${r})-SUM($E$6:E${r}))`;
  }
  const sheets = [{
    name: 'Ledger',
    rows: grid(cells),
    styles: {
      A1: title(p), 'A5:A29': { numFmt: 'd mmm yyyy', align: 'left' }, 'D5:F29': { numFmt: MONEY },
      'H4:I4': header(p), 'H9:J9': header(p), 'I5:I7': { numFmt: MONEY, border: true }, 'H5:H7': { bold: true, border: true },
      'H7:I7': total(p, MONEY), H7: { bold: true, fill: '#' + p.lt2, border: true },
      'H10:H15': { border: true }, 'I10:J15': { numFmt: MONEY, border: true }, A31: note(p),
    },
  }];
  return compose({
    theme, sheets,
    finish(view) {
      const o = ops(view);
      o.widths({ A: 100, B: 96, C: 200, D: 100, E: 100, F: 110, G: 24, H: 130, I: 100, J: 100 });
      o.table('A4:F29', 'TableStyleLight9');
      o.freeze(4, 0);
      o.list('B6:B29', categories);
      o.rule('F6:F29', { kind: 'cellIs', operator: 'lessThan', values: [0], look: 'red' });
    },
  });
}

/** The workbooks, by template id. */
function expenses() {
  const theme = 'harbour';
  const p = themeById(theme).palette;
  const claims = [
    [-20, 'Travel', 'Train to the client, return', 64.4, 0],
    [-20, 'Meals', 'Lunch with the client', 38.5, 6.42],
    [-14, 'Mileage', 'Site visit, 46 miles', 20.7, 0],
    [-9, 'Accommodation', 'Hotel, one night', 112, 18.67],
    [-8, 'Supplies', 'Printer paper and ink', 27.99, 4.67],
  ];
  const cells = {
    A1: 'Expense report',
    A2: 'Name', B2: '[Your name]', D2: 'Period', E2: '[Month and year]',
    A3: 'Department', B3: '[Department]', D3: 'Manager', E3: '[Name]',
    A5: 'Date', B5: 'Category', C5: 'Description', D5: 'Amount', E5: 'VAT included', F5: 'Receipt',
    C21: 'Total claimed', D21: '=SUM(D6:D19)', E21: '=SUM(E6:E19)',
    C22: 'Less advances', D22: 0,
    C23: 'Owed to you', D23: '=D21-D22',
    A25: 'Choose a category from the list, keep each receipt and tick it off. Mileage is the miles times the rate your employer pays.',
    A27: 'Signed', B27: '', D27: 'Approved', E27: '',
  };
  claims.forEach(([days, cat, what, amount, vat], i) => {
    const r = 6 + i;
    Object.assign(cells, { [`A${r}`]: `=TODAY()${days}`, [`B${r}`]: cat, [`C${r}`]: what, [`D${r}`]: amount, [`E${r}`]: vat, [`F${r}`]: 'Yes' });
  });
  const sheets = [{
    name: 'Expenses',
    rows: grid(cells),
    styles: {
      A1: title(p), 'A2:A3': { bold: true, colour: '#' + p.dk2 }, 'D2:D3': { bold: true, colour: '#' + p.dk2 },
      'A6:A19': { numFmt: 'd mmm yyyy', align: 'left' }, 'D6:E19': { numFmt: MONEY },
      'C21:E21': total(p, MONEY), C22: { bold: true }, D22: { numFmt: MONEY }, 'C23:D23': { bold: true, numFmt: MONEY, border: true },
      A25: note(p), A27: { bold: true }, D27: { bold: true },
    },
  }];
  const charts = ({ values }) => [{
    sheet: 'Expenses', kind: 'chart', chartKind: 'pie', title: 'Claimed by category',
    categories: { ref: 'Expenses!$B$6:$B$10', values: claims.map((c) => c[1]) },
    series: [{ name: 'Amount', nameRef: 'Expenses!$D$5', ref: 'Expenses!$D$6:$D$10', values: values('Expenses', 'D6', 'D10') }],
    from: { col: 7, row: 1 }, to: { col: 13, row: 16 },
  }];
  return compose({
    theme, sheets, charts,
    finish(view) {
      const o = ops(view);
      o.widths({ A: 100, B: 120, C: 240, D: 96, E: 100, F: 80 });
      o.table('A5:F19', 'TableStyleMedium9');
      o.freeze(5, 0);
      o.list('B6:B19', ['Travel', 'Mileage', 'Accommodation', 'Meals', 'Supplies', 'Phone', 'Other']);
      o.list('F6:F19', ['Yes', 'No']);
      o.rule('F6:F19', { kind: 'containsText', text: 'No', look: 'red' });
    },
  });
}

function inventory() {
  const theme = 'saffron';
  const p = themeById(theme).palette;
  const stock = [
    ['SKU-001', 'Notebook, A5 lined', 'Stationery', 140, 50, 2.1],
    ['SKU-002', 'Ballpoint pens, box of 10', 'Stationery', 32, 40, 3.75],
    ['SKU-003', 'Printer paper, 500 sheets', 'Paper', 18, 25, 4.6],
    ['SKU-004', 'Envelopes, C5, pack of 50', 'Paper', 60, 20, 5.2],
    ['SKU-005', 'Stapler', 'Desk', 12, 10, 6.99],
    ['SKU-006', 'Desk lamp', 'Desk', 4, 6, 18.5],
    ['SKU-007', 'Sticky notes, 12 pads', 'Stationery', 75, 30, 4.25],
  ];
  const cells = {
    A1: 'Stock list', A2: 'Items to reorder', B2: '=COUNTIF(H5:H24,"Reorder")', D2: 'Stock value', E2: '=SUM(G5:G24)',
    A4: 'Code', B4: 'Item', C4: 'Category', D4: 'In stock', E4: 'Reorder at', F4: 'Unit cost', G4: 'Value', H4: 'Status',
    A26: 'Count what is on the shelf into In stock; an item at or below its reorder level is marked to reorder.',
  };
  for (let i = 0; i < 20; i++) {
    const r = 5 + i;
    const s = stock[i];
    if (s) Object.assign(cells, { [`A${r}`]: s[0], [`B${r}`]: s[1], [`C${r}`]: s[2], [`D${r}`]: s[3], [`E${r}`]: s[4], [`F${r}`]: s[5] });
    cells[`G${r}`] = `=IF(D${r}="","",D${r}*F${r})`;
    cells[`H${r}`] = `=IF(D${r}="","",IF(D${r}<=E${r},"Reorder","OK"))`;
  }
  const sheets = [{
    name: 'Stock',
    rows: grid(cells),
    styles: {
      A1: title(p), A2: { bold: true, colour: '#' + p.dk2 }, B2: { bold: true, colour: '#' + p.accent1 },
      D2: { bold: true, colour: '#' + p.dk2 }, E2: { bold: true, numFmt: MONEY0, colour: '#' + p.accent1 },
      'F5:G24': { numFmt: MONEY }, A26: note(p),
    },
  }];
  return compose({
    theme, sheets,
    finish(view) {
      const o = ops(view);
      o.widths({ A: 90, B: 220, C: 110, D: 80, E: 90, F: 90, G: 96, H: 90 });
      o.table('A4:H24', 'TableStyleMedium3');
      o.freeze(4, 0);
      o.list('C5:C24', ['Stationery', 'Paper', 'Desk', 'Other']);
      o.rule('H5:H24', { kind: 'containsText', text: 'Reorder', look: 'red' });
      o.rule('H5:H24', { kind: 'containsText', text: 'OK', look: 'green' });
    },
  });
}

function planner() {
  const theme = 'orchard';
  const p = themeById(theme).palette;
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const cols = ['B', 'C', 'D', 'E', 'F', 'G', 'H'];
  const cells = {
    A1: 'Weekly planner', A2: 'Week starting', B2: '=TODAY()-WEEKDAY(TODAY(),3)',
    A4: '',
    A20: 'Priorities this week', E20: 'Notes',
    A21: '1.', A22: '2.', A23: '3.',
  };
  days.forEach((d, i) => {
    cells[`${cols[i]}4`] = d;
    cells[`${cols[i]}5`] = `=$B$2+${i}`;
  });
  const hours = ['8:00', '9:00', '10:00', '11:00', '12:00', '13:00', '14:00', '15:00', '16:00', '17:00', '18:00', 'Evening'];
  hours.forEach((h, i) => { cells[`A${6 + i}`] = h; });
  cells.B7 = 'Team meeting';
  cells.D9 = 'Lunch with Sam';
  cells.F12 = 'Dentist';
  cells.G6 = 'Market';
  const sheets = [{
    name: 'Week',
    rows: grid(cells),
    styles: {
      A1: title(p), A2: { bold: true, colour: '#' + p.dk2 }, B2: { numFmt: 'd mmmm yyyy', bold: true, align: 'left' },
      'B4:H4': header(p), 'B5:H5': { numFmt: 'd mmm', colour: '#' + p.dk2, align: 'center', fill: '#' + p.lt2 },
      'A6:A17': { bold: true, colour: '#' + p.dk2 },
      A20: { bold: true, colour: '#' + p.accent1 }, E20: { bold: true, colour: '#' + p.accent1 }, 'A21:A23': { colour: '#' + p.dk2 },
    },
  }];
  return compose({
    theme, sheets,
    finish(view) {
      const o = ops(view);
      o.widths({ A: 112, B: 130, C: 130, D: 130, E: 130, F: 130, G: 130, H: 130 });
      for (let r = 6; r <= 17; r++) o.height(r, 30);
      // Home > Borders > All borders, in the theme's light grey, so each hour is a box to write in.
      const line = { style: 'thin', colour: '#C8CCD2' };
      o.format('B6:H17', { border: { top: line, bottom: line, left: line, right: line }, wrap: true, valign: 'top' });
      o.format('G6:H17', { fill: '#' + p.lt2 });
      o.freeze(5, 1);
    },
  });
}

function grades() {
  const theme = 'meadow';
  const p = themeById(theme).palette;
  const pupils = [
    ['Amira Khan', 78, 85, 92, 88], ['Ben Thompson', 64, 70, 58, 72], ['Chloe Martin', 91, 94, 89, 95],
    ['Daniel Osei', 55, 61, 67, 60], ['Ella Wright', 82, 79, 85, 81], ['Farhan Ali', 70, 76, 74, 79],
    ['Grace Lee', 88, 90, 93, 91], ['Harry Evans', 47, 52, 58, 49],
  ];
  const cells = {
    A1: 'Grade book', A2: 'Class', B2: '[Class and subject]', D2: 'Pass mark', E2: 50,
    A4: 'Pupil', B4: 'Test 1', C4: 'Test 2', D4: 'Project', E4: 'Exam', F4: 'Average', G4: 'Grade', H4: 'Result',
    A31: 'Class average', B31: '=IFERROR(AVERAGE(B5:B29),"")', C31: '=IFERROR(AVERAGE(C5:C29),"")', D31: '=IFERROR(AVERAGE(D5:D29),"")', E31: '=IFERROR(AVERAGE(E5:E29),"")', F31: '=IFERROR(AVERAGE(F5:F29),"")',
    A33: 'Grades: A from 85, B from 70, C from 60, D from 50, otherwise E. Change the pass mark in E2.',
  };
  for (let i = 0; i < 25; i++) {
    const r = 5 + i;
    const s = pupils[i];
    if (s) Object.assign(cells, { [`A${r}`]: s[0], [`B${r}`]: s[1], [`C${r}`]: s[2], [`D${r}`]: s[3], [`E${r}`]: s[4] });
    cells[`F${r}`] = `=IF(COUNT(B${r}:E${r})=0,"",ROUND(AVERAGE(B${r}:E${r}),1))`;
    cells[`G${r}`] = `=IF(F${r}="","",IF(F${r}>=85,"A",IF(F${r}>=70,"B",IF(F${r}>=60,"C",IF(F${r}>=50,"D","E")))))`;
    cells[`H${r}`] = `=IF(F${r}="","",IF(F${r}>=$E$2,"Pass","Needs support"))`;
  }
  const sheets = [{
    name: 'Grades',
    rows: grid(cells),
    styles: {
      A1: title(p), A2: { bold: true, colour: '#' + p.dk2 }, D2: { bold: true, colour: '#' + p.dk2 }, E2: { bold: true },
      'F5:F29': { numFmt: '0.0', bold: true }, 'G5:G29': { align: 'center', bold: true },
      'A31:F31': total(p, '0.0'), A33: note(p),
    },
  }];
  const charts = ({ values }) => [{
    sheet: 'Grades', kind: 'chart', chartKind: 'bar', title: 'Average by pupil',
    categories: { ref: 'Grades!$A$5:$A$12', values: pupils.map((x) => x[0]) },
    series: [{ name: 'Average', nameRef: 'Grades!$F$4', ref: 'Grades!$F$5:$F$12', values: values('Grades', 'F5', 'F12') }],
    from: { col: 9, row: 1 }, to: { col: 16, row: 18 },
  }];
  return compose({
    theme, sheets, charts,
    finish(view) {
      const o = ops(view);
      o.widths({ A: 170, B: 70, C: 70, D: 70, E: 70, F: 80, G: 64, H: 120 });
      o.table('A4:H29', 'TableStyleMedium7');
      o.freeze(4, 1);
      o.rule('H5:H29', { kind: 'containsText', text: 'Needs support', look: 'red' });
      o.rule('G5:G29', { kind: 'containsText', text: 'A', look: 'green' });
    },
  });
}

export const WORKBOOKS = {
  'sheet-budget': budget,
  'sheet-invoice': invoiceSheet,
  'sheet-timesheet': timesheet,
  'sheet-loan': loan,
  'sheet-todo': todo,
  'sheet-ledger': ledger,
  'sheet-expenses': expenses,
  'sheet-inventory': inventory,
  'sheet-planner': planner,
  'sheet-grades': grades,
};
