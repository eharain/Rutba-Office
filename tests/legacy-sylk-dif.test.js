// SYLK and DIF (the 2.8 plan's old spreadsheets): cells, formulas from
// R1C1 into A1, number formats, column widths and names from SYLK; the
// values and the title from DIF.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readSylk, isSylk, r1c1ToA1 } from '../packages/office-formats/src/sylk.js';
import { readDif, isDif } from '../packages/office-formats/src/dif.js';

const bytes = (s) => new Uint8Array(Buffer.from(s, 'latin1'));

test('R1C1 references become A1 ones, relative from the cell and absolute with $', () => {
  assert.equal(r1c1ToA1('SUM(R[-3]C:R[-1]C)', 3, 1), 'SUM(B1:B3)');
  assert.equal(r1c1ToA1('RC[-1]*2', 4, 2), 'B5*2');
  assert.equal(r1c1ToA1('R1C1+R2C3', 9, 9), '$A$1+$C$2');
  assert.equal(r1c1ToA1('ROUND(R[-1]C,2)&"RC"', 1, 0), 'ROUND(A1,2)&"RC"');
  assert.equal(r1c1ToA1('R[-5]C', 1, 0), '#REF!');
});

test('a SYLK file Excel writes is read: words, numbers, a formula, formats, widths and a name', () => {
  const slk = [
    'ID;PWXL;N;E',
    'P;PGeneral',
    'P;P0.00',
    'F;P0;DG0G8;M255',
    'B;Y4;X2;D0 0 3 1',
    'F;W1 1 20',
    'F;W2 2 12',
    'C;Y1;X1;K"Item"',
    'C;X2;K"Price"',
    'C;Y2;X1;K"Tea;; green"',
    'F;P1;FF2R;Y2;X2',
    'C;X2;K3.5',
    'C;Y3;X1;K"Cake"',
    'C;X2;K12',
    'C;Y4;X1;K"Total"',
    'F;SD;Y4;X2',
    'C;X2;K15.5;ESUM(R[-2]C:R[-1]C)',
    'NN;NPrices;ER2C2:R3C2',
    'E',
  ].join('\r\n');
  assert.ok(isSylk(bytes(slk)));
  const { sheets, names } = readSylk(bytes(slk));
  const [s] = sheets;
  assert.deepEqual(s.rows[0], ['Item', 'Price']);
  assert.deepEqual(s.rows[1], ['Tea; green', 3.5]);
  assert.equal(s.rows[3][1], '=SUM(B2:B3)');
  assert.equal(s.styles.B2.numFmt, '0.00');
  assert.equal(s.styles.B2.align, 'right');
  assert.equal(s.styles.B4.bold, true);
  assert.deepEqual(s.widths, { A: 20, B: 12 });
  assert.deepEqual(names, [{ name: 'Prices', ref: '$B$2:$B$3' }]);
});

test('a DIF file is read: its title for the sheet, numbers, words and the true and false values', () => {
  const dif = [
    'TABLE', '0,1', '"Prices"',
    'VECTORS', '0,2', '""',
    'TUPLES', '0,3', '""',
    'DATA', '0,0', '""',
    '-1,0', 'BOT', '1,0', '"Item"', '1,0', '"Price"',
    '-1,0', 'BOT', '1,0', '"Tea"', '0,3.5', 'V',
    '-1,0', 'BOT', '1,0', '"Cake ""large"""', '0,1', 'TRUE',
    '-1,0', 'EOD',
  ].join('\r\n');
  assert.ok(isDif(bytes(dif)));
  const { sheets } = readDif(bytes(dif));
  assert.equal(sheets[0].name, 'Prices');
  assert.deepEqual(sheets[0].rows, [['Item', 'Price'], ['Tea', 3.5], ['Cake "large"', true]]);
  assert.equal(isDif(bytes('ID;P')), false);
});
