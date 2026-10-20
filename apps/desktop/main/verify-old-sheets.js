// Worksheets: the old spreadsheets (the 2.8 plan): a Lotus 1-2-3 .wk1 with
// its formula worked out again, a .wk3 with its sheets in their tabs, a
// dBASE table with its memo file beside it, SYLK and DIF, each opened as a
// workbook saved on as .xlsx, and what a reader could not keep said once the
// file is open. The files are made here, record by record. Run alone with
// RUTBA_VERIFY_ONLY=old-sheets.
import fs from 'node:fs';
import path from 'node:path';

const rec = (type, body) => { const b = Buffer.alloc(4 + body.length); b.writeUInt16LE(type, 0); b.writeUInt16LE(body.length, 2); body.copy(b, 4); return b; };
const u16 = (n) => { const b = Buffer.alloc(2); b.writeUInt16LE(n & 0xffff); return b; };
const f64 = (n) => { const b = Buffer.alloc(8); b.writeDoubleLE(n); return b; };
const cell = (fmt, col, row) => Buffer.concat([Buffer.from([fmt]), u16(col), u16(row)]);
const rel = (off) => ((off & 0x3fff) | 0x8000 | (off < 0 ? 0x4000 : 0)) & 0xffff;
const str = (s) => Buffer.from(`${s}\0`, 'latin1');
const ext = (n) => { const b = Buffer.alloc(10); const e = Math.floor(Math.log2(Math.abs(n))); b.writeBigUInt64LE(BigInt(Math.round((Math.abs(n) / 2 ** e) * 2 ** 63))); b.writeUInt16LE((n < 0 ? 0x8000 : 0) | (e + 16383), 8); return b; };

function wk1() {
  const sum = Buffer.concat([Buffer.from([0x02]), u16(rel(0)), u16(rel(-2)), u16(rel(0)), u16(rel(-1)), Buffer.from([0x50, 1, 0x03])]);
  const cellFn = Buffer.concat([Buffer.from([0x06]), str('row'), Buffer.from([0x01]), u16(0), u16(0), Buffer.from([0x6c, 0x03])]);
  return Buffer.concat([
    rec(0x00, u16(0x0406)),
    rec(0x08, Buffer.concat([u16(0), Buffer.from([16])])),
    rec(0x0f, Buffer.concat([cell(0xff, 0, 0), str("'Item")])),
    rec(0x0f, Buffer.concat([cell(0xff, 1, 0), str('"Price')])),
    rec(0x0f, Buffer.concat([cell(0xff, 0, 1), str("'Tea")])),
    rec(0x0e, Buffer.concat([cell(0x22, 1, 1), f64(3.5)])),
    rec(0x0f, Buffer.concat([cell(0xff, 0, 2), str("'Cake")])),
    rec(0x0d, Buffer.concat([cell(0x22, 1, 2), u16(12)])),
    rec(0x0f, Buffer.concat([cell(0xff, 0, 3), str("'Total")])),
    rec(0x10, Buffer.concat([cell(0x22, 1, 3), f64(15.5), u16(sum.length), sum])),
    rec(0x10, Buffer.concat([cell(0xff, 2, 3), f64(7), u16(cellFn.length), cellFn])),
    rec(0x01, Buffer.alloc(0)),
  ]);
}

function wk3() {
  const at = (row, sheet, col) => Buffer.concat([u16(row), Buffer.from([sheet, col])]);
  return Buffer.concat([
    rec(0x00, Buffer.concat([u16(0x1000), Buffer.alloc(24)])),
    rec(0x16, Buffer.concat([at(0, 0, 0), str("'North")])),
    rec(0x17, Buffer.concat([at(1, 0, 0), ext(1234.5)])),
    rec(0x16, Buffer.concat([at(0, 1, 0), str("'South")])),
    rec(0x18, Buffer.concat([at(1, 1, 0), u16(99 << 1)])),
  ]);
}

function dbf() {
  const fields = [['NAME', 'C', 10], ['PRICE', 'N', 7, 2], ['NOTES', 'M', 10]];
  const records = [['Tea', '   3.50', '         1'], ['*', 'Gone', '   1.00', '          '], ['Cake', '  12.00', '         2']];
  const headerSize = 32 + fields.length * 32 + 1;
  const recordSize = 1 + fields.reduce((n, f) => n + f[2], 0);
  const out = Buffer.alloc(headerSize + recordSize * records.length + 1, 0x20);
  out.fill(0, 0, headerSize);
  out[0] = 0x83; out[1] = 124; out[2] = 10; out[3] = 21;
  out.writeUInt32LE(records.length, 4);
  out.writeUInt16LE(headerSize, 8);
  out.writeUInt16LE(recordSize, 10);
  out[29] = 0x03;
  fields.forEach(([name, type, length, decimals = 0], i) => { const a = 32 + i * 32; out.write(name, a, 'latin1'); out.write(type, a + 11, 'latin1'); out[a + 16] = length; out[a + 17] = decimals; });
  out[headerSize - 1] = 0x0d;
  records.forEach((r, k) => {
    const a = headerSize + k * recordSize;
    const gone = r[0] === '*';
    out[a] = gone ? 0x2a : 0x20;
    let off = 1;
    (gone ? r.slice(1) : r).forEach((v, c) => { out.write(String(v).padEnd(fields[c][2]).slice(0, fields[c][2]), a + off, 'latin1'); off += fields[c][2]; });
  });
  out[out.length - 1] = 0x1a;
  const memo = Buffer.alloc(512 * 3, 0);
  memo.write('Green, from Assam\x1a\x1a', 512, 'latin1');
  memo.write('Lemon sponge\x1a\x1a', 1024, 'latin1');
  return { table: out, memo };
}

const SLK = ['ID;PWXL;N;E', 'C;Y1;X1;K"Month"', 'C;X2;K"Sales"', 'C;Y2;X1;K"Jan"', 'C;X2;K120', 'C;Y3;X1;K"Feb"', 'C;X2;K80', 'C;Y4;X1;K"Total"', 'C;X2;K200;ESUM(R[-2]C:R[-1]C)', 'E'].join('\r\n');
const DIF = ['TABLE', '0,1', '"Stock"', 'VECTORS', '0,2', '""', 'TUPLES', '0,2', '""', 'DATA', '0,0', '""', '-1,0', 'BOT', '1,0', '"Pens"', '0,40', 'V', '-1,0', 'BOT', '1,0', '"Paper"', '0,25', 'V', '-1,0', 'EOD'].join('\r\n');

export async function verifyOldSheets(h, { dir }) {
  const { open, check, until, wait, errorsIn, sessionFor } = h;
  const folder = path.join(dir, 'old-sheets');
  fs.mkdirSync(folder, { recursive: true });
  const files = {
    wk1: path.join(folder, 'prices.wk1'),
    wk3: path.join(folder, 'regions.wk3'),
    dbf: path.join(folder, 'teas.dbf'),
    slk: path.join(folder, 'sales.slk'),
    dif: path.join(folder, 'stock.dif'),
  };
  fs.writeFileSync(files.wk1, wk1());
  fs.writeFileSync(files.wk3, wk3());
  const { table, memo } = dbf();
  fs.writeFileSync(files.dbf, table);
  fs.writeFileSync(path.join(folder, 'teas.dbt'), memo);
  fs.writeFileSync(files.slk, Buffer.from(SLK, 'latin1'));
  fs.writeFileSync(files.dif, Buffer.from(DIF, 'latin1'));

  /** A file opened in Worksheets: what the cells show, the tabs, and what it said. */
  const openAndRead = async (file, refs, settle) => {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="${settle}"]')) && (document.querySelector('.sh-cell[data-ref="${settle}"]').innerText.trim() !== '')`), `the cells of ${path.basename(file)}`, 10000).catch(() => {});
    await wait(600);
    return {
      win,
      shown: await js(`(() => {
        const cell = (ref) => document.querySelector('.sh-cell[data-ref="' + ref + '"]')?.innerText.trim() ?? null;
        return { cells: Object.fromEntries(${JSON.stringify(refs)}.map((r) => [r, cell(r)])), tabs: [...document.querySelectorAll('.sh-tab')].map((t) => t.innerText.trim()), said: [...document.querySelectorAll('.rw-toast')].map((t) => t.innerText.trim()).join(' | ') };
      })()`),
      session: sessionFor('sheet'),
    };
  };

  try {
    const a = await openAndRead(files.wk1, ['A1', 'B1', 'B2', 'B4', 'C4'], 'B4');
    check('old sheets: a Lotus 1-2-3 .wk1 opens with its labels and numbers, and its @SUM worked out again as SUM',
      a.session.converted?.from === 'wk1' && a.shown.cells.A1 === 'Item' && a.shown.cells.B4 === '$15.50' && a.shown.cells.C4 === '7',
      JSON.stringify(a.shown.cells));
    check('old sheets: what the .wk1 kept as a value is said: a formula using @CELL',
      /formula was kept as its value: it uses @CELL/.test(a.shown.said) && /Opened from WK1/.test(a.shown.said), a.shown.said);
    const b = await openAndRead(files.wk3, ['A1', 'A2'], 'A2');
    check('old sheets: a .wk3 opens with its sheets in their tabs, its 80-bit numbers read',
      b.shown.tabs.length === 2 && b.shown.tabs[0].startsWith('A') && b.shown.tabs[1].startsWith('B') && b.shown.cells.A1 === 'North' && /1,?234\.5/.test(b.shown.cells.A2 || ''),
      JSON.stringify(b.shown));
    const c = await openAndRead(files.dbf, ['A1', 'C1', 'A2', 'C2', 'A3', 'C3', 'A4'], 'C3');
    check('old sheets: a dBASE table opens as its fields and records, its memos read from the .dbt beside it, the deleted record left out and said',
      c.shown.cells.A1 === 'NAME' && c.shown.cells.C2 === 'Green, from Assam' && c.shown.cells.A3 === 'Cake' && c.shown.cells.C3 === 'Lemon sponge' && !c.shown.cells.A4 && /1 deleted record was left out/.test(c.shown.said),
      JSON.stringify(c.shown));
    const d = await openAndRead(files.slk, ['A1', 'B4'], 'B4');
    check('old sheets: a SYLK file opens with its formula from R1C1 worked out',
      d.session.converted?.from === 'slk' && d.shown.cells.A1 === 'Month' && d.shown.cells.B4 === '200', JSON.stringify(d.shown.cells));
    const e = await openAndRead(files.dif, ['A1', 'B2'], 'B2');
    check('old sheets: a DIF file opens with its values, the sheet named by its title',
      e.shown.cells.A1 === 'Pens' && e.shown.cells.B2 === '25' && e.shown.tabs.some((t) => t.startsWith('Stock')), JSON.stringify(e.shown));
    const complaints = [];
    for (const x of [a, b, c, d, e]) complaints.push(...(await errorsIn(x.win)));
    check('old sheets: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('old sheets: the check ran to the end', false, err?.message || String(err));
  }
}
