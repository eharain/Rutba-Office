// Contacts' CSV import: today's Google headings read (companies, job titles,
// each address's and number's kind), Outlook's ANSI file read as ANSI, a
// preview that says what each column was read as and lets one be changed,
// what was left out said, a card with no address not added again on the
// next import, and vCard's CELL kept as the "mobile" the editor names.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { readContactsCsvReport, readContactsCsv, decodeText, readVCards, writeVCard } from '../packages/contacts/src/index.js';
import { createContactsService } from '../apps/desktop/main/contacts.js';

// The headings Google Contacts writes today (Export, Google CSV).
const GOOGLE = [
  'First Name,Middle Name,Last Name,Phonetic First Name,Phonetic Middle Name,Phonetic Last Name,Name Prefix,Name Suffix,Nickname,File As,Organization Name,Organization Title,Organization Department,Birthday,Notes,Photo,Labels,E-mail 1 - Label,E-mail 1 - Value,E-mail 2 - Label,E-mail 2 - Value,Phone 1 - Label,Phone 1 - Value,Phone 2 - Label,Phone 2 - Value,Address 1 - Label,Address 1 - Formatted,Address 1 - Street,Address 1 - City,Address 1 - PO Box,Address 1 - Region,Address 1 - Postal Code,Address 1 - Country,Address 1 - Extended Address,Website 1 - Label,Website 1 - Value',
  'Kim,,Lee,,,,Dr,,,,Tech Style Ltd,Head of Product,Design,1985-03-15,Prefers email,https://lh3.example/photo.jpg,* myContacts ::: Suppliers,* Work,kim@example.com,Home,kim@home.example,Mobile,+44 7700 900123,Google Voice,+44 20 7946 0000,Work,"1 Long Lane, London",1 Long Lane,London,,,EC1A 9PN,United Kingdom,Floor 2,Profile,https://kim.example',
  'Bo,,Lund,,,,,,,,,,,,,,* myContacts,,,,,Home,0113 496 0000,,,,,,,,,,,,,',
  ',,,,,,,,,,,,,,,,* myContacts,,,,,,,,,,,,,,,,,,,',
].join('\r\n');

test("today's Google headings: company, job title, department, kinds and labels, and no myContacts tag", () => {
  const [kim, bo] = readContactsCsv(GOOGLE);
  assert.equal(kim.name.full, 'Dr Kim Lee');
  assert.equal(kim.org, 'Tech Style Ltd');
  assert.equal(kim.title, 'Head of Product');
  assert.equal(kim.department, 'Design');
  assert.deepEqual(kim.emails.map((e) => [e.value, e.type]), [['kim@example.com', 'work'], ['kim@home.example', 'home']]);
  assert.deepEqual(kim.phones.map((p) => [p.value, p.type, p.label]), [['+44 7700 900123', 'mobile', null], ['+44 20 7946 0000', 'other', 'Google Voice']]);
  assert.equal(kim.addresses[0].type, 'work');
  assert.equal(kim.addresses[0].street, 'Floor 2\n1 Long Lane');
  assert.equal(kim.addresses[0].postcode, 'EC1A 9PN');
  assert.equal(kim.urls[0].value, 'https://kim.example');
  assert.deepEqual(kim.categories, ['Suppliers']);
  assert.deepEqual(bo.categories, []);
  assert.equal(bo.phones[0].type, 'home');
});

test('the report: each column and what it was read as, the rows left out and the columns not read', () => {
  const r = readContactsCsvReport(GOOGLE);
  assert.equal(r.contacts.length, 2);
  assert.deepEqual(r.skipped, [4], 'the empty row, by its line in the file');
  assert.deepEqual(r.unread, ['Photo'], 'a column with values that nothing reads is named');
  const of = (h) => r.columns.find((c) => c.header === h);
  assert.equal(of('Organization Name').field, 'org');
  assert.equal(of('Organization Title').field, 'jobTitle');
  assert.equal(of('E-mail 1 - Label').field, 'type');
  assert.equal(of('Phonetic First Name').field, '');
  assert.equal(of('Organization Name').sample, 'Tech Style Ltd');

  // A column changed: Notes left out, Nickname read as the company.
  const changed = readContactsCsvReport('Name,Nickname,Notes,Email\nAnn Ek,Ek Ltd,Met at the fair,ann@example.com\n', { columns: { 1: 'org', 2: '' } });
  assert.equal(changed.contacts[0].org, 'Ek Ltd');
  assert.equal(changed.contacts[0].nickname, null);
  assert.equal(changed.contacts[0].note, null);
  assert.deepEqual(changed.unread, ['Notes']);
  assert.equal(changed.columns[1].auto, 'nickname', 'what it was matched to is kept beside the change');

  // A company and a phone, with no name or address, is a card, not left out.
  assert.equal(readContactsCsv('Company,Business Phone\nAcme,0113 496 0001\n').length, 1);
});

test("Outlook's ANSI file is read as ANSI, a UTF-8 one as UTF-8, and a semicolon file by its separator", () => {
  const ansi = Buffer.from([...Buffer.from('First Name,Last Name,E-mail Address\r\n'), ...Buffer.from('Zo'), 0xeb, ...Buffer.from(',M'), 0xfc, ...Buffer.from('ller,zoe@example.com\r\n')]);
  const [zoe] = readContactsCsv(ansi);
  assert.equal(zoe.name.full, 'Zoë Müller');
  assert.equal(decodeText(Buffer.from('Zoë', 'utf8')), 'Zoë');
  assert.equal(decodeText(Buffer.from([0xef, 0xbb, 0xbf, ...Buffer.from('Zoë', 'utf8')])).trim(), 'Zoë');
  assert.equal(decodeText(Buffer.from([0xff, 0xfe, ...Buffer.from('Zoë', 'utf16le')])), 'Zoë');
  const semi = readContactsCsv('Name;Email\n"Lee; Kim";kim@example.com\n');
  assert.equal(semi[0].name.full, 'Lee; Kim');
  assert.equal(semi[0].emails[0].value, 'kim@example.com');
});

test("vCard's CELL is the card's mobile, and a mobile is written back as CELL", () => {
  const [c] = readVCards('BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Ann Ek\r\nTEL;TYPE=CELL:+46 70 000 0000\r\nEND:VCARD\r\n');
  assert.equal(c.phones[0].type, 'mobile');
  assert.match(writeVCard(c, { version: '3.0' }), /TEL;TYPE=cell(,pref)?:\+46 70 000 0000/);
});

test('the service: a preview first, the import says what it left out, and a card with no address is not added twice', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-csv-'));
  try {
    const file = path.join(dir, 'google.csv');
    fs.writeFileSync(file, GOOGLE);
    const contacts = createContactsService({ stores: { dir }, broadcast: null });
    const seen = contacts.peekCsv({ path: file });
    assert.equal(seen.count, 2);
    assert.equal(seen.sample[0].display, 'Dr Kim Lee');
    assert.equal(contacts.count(), 0, 'a preview keeps nothing');
    const first = contacts.importFile({ path: file });
    assert.deepEqual([first.added, first.skipped, first.unread], [2, 1, ['Photo']]);
    const again = contacts.importFile({ path: file });
    assert.deepEqual([again.added, again.same], [0, 2], 'Bo has no address, and is the same Bo the second time');
    assert.equal(contacts.count(), 2);

    // A card kept with vCard's CELL before it was read as mobile is shown as the mobile it is.
    const stored = path.join(dir, 'old.vcf');
    fs.writeFileSync(stored, 'BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Cy Old\r\nEMAIL:cy@example.com\r\nTEL;TYPE=CELL:+1 555 0100\r\nEND:VCARD\r\n');
    contacts.importFile({ path: stored });
    contacts._dav.state().contacts.find((c) => c.name.full === 'Cy Old').phones[0].type = 'cell';
    assert.equal(contacts.list({ query: 'Cy' })[0].phones[0].type, 'mobile');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
