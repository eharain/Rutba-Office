// vCard, read and written — Apple's groups and labels, Google's export,
// Outlook's 2.1 with quoted-printable, 4.0's data URIs and short birthdays,
// and the CSV address books.

import test from 'node:test';
import assert from 'node:assert/strict';
import { readVCards, writeVCards, displayName, initials, splitName, readContactsCsv, parseCsv } from '../packages/contacts/src/index.js';

const CRLF = (s) => s.replace(/\n/g, '\r\n');

test('an Apple card: groups with labels, a photo, a structured name and address', () => {
  const png = Buffer.from('89504e470d0a1a0a0000000d49484452', 'hex');
  const cards = readVCards(CRLF(`BEGIN:VCARD
VERSION:3.0
PRODID:-//Apple Inc.//macOS 14.0//EN
N:Lee;Kim;Alex;Dr.;
FN:Dr. Kim Alex Lee
ORG:Tech Style Ltd;Engineering
TITLE:Head of Product
item1.EMAIL;type=INTERNET;type=pref:kim@example.com
item1.X-ABLabel:_$!<Work>!$_
item2.EMAIL;type=INTERNET:kim.lee@home.example
item2.X-ABLabel:Personal
TEL;type=CELL;type=VOICE;type=pref:+44 7700 900123
item3.ADR;type=WORK:;;1 Long Lane\\, Floor 2;London;;EC1A 9PN;United Kingdom
item3.X-ABLabel:Office
BDAY:1985-03-15
NOTE:Prefers email\\; calls after 3pm.
URL;type=pref:https://example.com/kim
PHOTO;ENCODING=b;TYPE=PNG:${png.toString('base64')}
CATEGORIES:Suppliers,Friends
UID:8f3c1e2a-1
END:VCARD
`));
  assert.equal(cards.length, 1);
  const c = cards[0];
  assert.equal(c.name.full, 'Dr. Kim Alex Lee');
  assert.deepEqual([c.name.given, c.name.family, c.name.middle, c.name.prefix], ['Kim', 'Lee', 'Alex', 'Dr.']);
  assert.equal(c.org, 'Tech Style Ltd');
  assert.equal(c.department, 'Engineering');
  assert.equal(c.title, 'Head of Product');
  assert.equal(c.emails.length, 2);
  assert.equal(c.emails[0].value, 'kim@example.com');
  assert.equal(c.emails[0].pref, true);
  assert.equal(c.emails[0].label, 'Work', 'Apple\'s label, unwrapped from its markers');
  assert.equal(c.emails[1].label, 'Personal');
  assert.equal(c.phones[0].value, '+44 7700 900123');
  assert.equal(c.phones[0].type, 'cell');
  assert.equal(c.addresses[0].street, '1 Long Lane, Floor 2');
  assert.equal(c.addresses[0].city, 'London');
  assert.equal(c.addresses[0].postcode, 'EC1A 9PN');
  assert.equal(c.addresses[0].country, 'United Kingdom');
  assert.equal(c.addresses[0].label, 'Office');
  assert.equal(c.birthday, '1985-03-15');
  assert.equal(c.note, 'Prefers email; calls after 3pm.');
  assert.equal(c.urls[0].value, 'https://example.com/kim');
  assert.ok(c.photo.data.equals(png));
  assert.equal(c.photo.mediaType, 'image/png');
  assert.deepEqual(c.categories, ['Suppliers', 'Friends']);
  assert.equal(c.uid, '8f3c1e2a-1');
  assert.equal(displayName(c), 'Dr. Kim Alex Lee');
  assert.equal(initials(c), 'KL');
});

test('an Outlook 2.1 card: bare parameters and quoted-printable with soft breaks', () => {
  const cards = readVCards(CRLF(`BEGIN:VCARD
VERSION:2.1
N:M=C3=BCller;J=C3=BCrgen
FN;ENCODING=QUOTED-PRINTABLE;CHARSET=UTF-8:J=C3=BCrgen M=C3=BCller
TEL;WORK;VOICE:+49 30 1234567
TEL;CELL:+49 170 1234567
EMAIL;PREF;INTERNET:juergen@example.de
ADR;WORK;PREF:;;Unter den Linden 1;Berlin;;10117;Germany
LABEL;WORK;ENCODING=QUOTED-PRINTABLE;CHARSET=UTF-8:Unter den Linden 1=0D=0A=
Berlin 10117
NOTE;ENCODING=QUOTED-PRINTABLE;CHARSET=UTF-8:Erste Zeile=0D=0A=
Zweite Zeile
END:VCARD
`));
  const c = cards[0];
  assert.equal(c.name.full, 'Jürgen Müller');
  assert.equal(c.phones[0].type, 'work');
  assert.equal(c.phones[1].type, 'cell');
  assert.equal(c.emails[0].value, 'juergen@example.de');
  assert.equal(c.emails[0].pref, true);
  assert.equal(c.addresses[0].city, 'Berlin');
  assert.equal(c.note, 'Erste Zeile\r\nZweite Zeile', 'a soft break joins the lines, a hard one is kept');
});

test('a 4.0 card: a data URI photo, a tel: URI, a birthday without a year', () => {
  const cards = readVCards(CRLF(`BEGIN:VCARD
VERSION:4.0
FN:Sam Patel
N:Patel;Sam;;;
EMAIL;PREF=1:sam@example.org
TEL;VALUE=uri;TYPE=work:tel:+1-555-0100
PHOTO:data:image/jpeg;base64,/9j/4AAQSkZJRg==
BDAY:--0315
KIND:individual
END:VCARD
BEGIN:VCARD
VERSION:4.0
FN:Ops desk
KIND:group
EMAIL:ops@example.org
END:VCARD
`));
  assert.equal(cards.length, 2);
  const [sam, ops] = cards;
  assert.equal(sam.emails[0].pref, true);
  assert.equal(sam.phones[0].value, '+1-555-0100');
  assert.equal(sam.phones[0].type, 'work');
  assert.equal(sam.photo.mediaType, 'image/jpeg');
  assert.equal(sam.photo.data.length, 10);
  assert.equal(sam.birthday, '--03-15');
  assert.equal(ops.kind, 'group');
  assert.equal(displayName(ops), 'Ops desk');
});

test('written and read back, in 3.0 and 4.0', () => {
  const contact = {
    uid: 'u-1',
    name: { full: 'Kim Lee', given: 'Kim', family: 'Lee', middle: '', prefix: '', suffix: '' },
    nickname: 'Kimmy',
    org: 'Tech Style Ltd',
    department: null,
    title: 'Head of Product',
    role: null,
    emails: [{ value: 'kim@example.com', type: 'work', pref: true, label: null }, { value: 'kim@home.example', type: 'home', pref: false, label: 'Weekend' }],
    phones: [{ value: '+44 7700 900123', type: 'mobile', pref: true, label: null }],
    addresses: [{ street: '1 Long Lane, Floor 2', city: 'London', region: '', postcode: 'EC1A 9PN', country: 'United Kingdom', type: 'work', pref: true, label: null }],
    birthday: '1985-03-15',
    anniversary: null,
    note: 'Line one;\nline two, with a comma.',
    urls: [{ value: 'https://example.com/kim', type: 'other', label: null }],
    photo: { data: Buffer.from('89504e470d0a1a0a', 'hex'), mediaType: 'image/png', uri: null },
    categories: ['Suppliers'],
    rev: null,
    kind: 'individual',
    extra: [{ name: 'X-TWITTER', params: {}, value: '@kim', group: null }],
  };
  for (const version of ['3.0', '4.0']) {
    const text = writeVCards([contact], { version });
    assert.ok(text.split('\r\n').every((l) => Buffer.byteLength(l, 'utf8') <= 75), `${version}: folded within 75 octets`);
    const back = readVCards(text)[0];
    assert.equal(back.version, version);
    assert.equal(back.name.full, 'Kim Lee');
    assert.equal(back.name.given, 'Kim');
    assert.equal(back.org, 'Tech Style Ltd');
    assert.equal(back.emails.length, 2);
    assert.equal(back.emails[0].value, 'kim@example.com');
    assert.equal(back.emails[0].pref, true, `${version}: the preferred address is marked`);
    assert.equal(back.emails[1].label, 'Weekend', `${version}: a label survives as a group`);
    assert.equal(back.phones[0].value, '+44 7700 900123');
    assert.equal(back.addresses[0].street, '1 Long Lane, Floor 2');
    assert.equal(back.addresses[0].postcode, 'EC1A 9PN');
    assert.equal(back.birthday, '1985-03-15');
    assert.equal(back.note, 'Line one;\nline two, with a comma.');
    assert.ok(back.photo.data.equals(contact.photo.data), `${version}: the photo bytes`);
    assert.equal(back.photo.mediaType, 'image/png');
    assert.deepEqual(back.categories, ['Suppliers']);
    assert.equal(back.extra[0].name, 'X-TWITTER');
    assert.equal(back.extra[0].value, '@kim');
  }
});

test('a Google Contacts CSV and an Outlook CSV read into the same shape', () => {
  const google = `Name,Given Name,Family Name,Birthday,Notes,Group Membership,E-mail 1 - Type,E-mail 1 - Value,E-mail 2 - Type,E-mail 2 - Value,Phone 1 - Type,Phone 1 - Value,Address 1 - Type,Address 1 - Street,Address 1 - City,Address 1 - Postal Code,Address 1 - Country,Organization 1 - Name,Organization 1 - Title
Kim Lee,Kim,Lee,1985-03-15,"Prefers email, calls after 3pm.",* myContacts ::: Suppliers,* Work,kim@example.com,Home,kim@home.example,Mobile,+44 7700 900123,Work,"1 Long Lane, Floor 2",London,EC1A 9PN,United Kingdom,Tech Style Ltd,Head of Product
`;
  const g = readContactsCsv(google);
  assert.equal(g.length, 1);
  assert.equal(g[0].name.given, 'Kim');
  assert.equal(g[0].emails[0].value, 'kim@example.com');
  assert.equal(g[0].emails[0].type, 'work');
  assert.equal(g[0].emails[1].value, 'kim@home.example');
  assert.equal(g[0].phones[0].value, '+44 7700 900123');
  assert.equal(g[0].phones[0].type, 'mobile');
  assert.equal(g[0].addresses[0].street, '1 Long Lane, Floor 2');
  assert.equal(g[0].addresses[0].postcode, 'EC1A 9PN');
  assert.equal(g[0].org, 'Tech Style Ltd');
  assert.equal(g[0].title, 'Head of Product');
  assert.equal(g[0].birthday, '1985-03-15');
  assert.deepEqual(g[0].categories, ['myContacts', 'Suppliers']);

  const outlook = `First Name,Middle Name,Last Name,Company,Job Title,Business Street,Business City,Business Postal Code,Business Country/Region,Home Phone,Mobile Phone,E-mail Address,E-mail 2 Address,Notes
Sam,,Patel,Example Org,Director,"2 High St",Leeds,LS1 1AA,United Kingdom,0113 000 0000,07700 900456,sam@example.org,sam.p@home.example,"Met at the fair"
`;
  const o = readContactsCsv(outlook);
  assert.equal(o[0].name.full, 'Sam Patel');
  assert.equal(o[0].emails.map((e) => e.value).join(','), 'sam@example.org,sam.p@home.example');
  assert.equal(o[0].phones.find((p) => p.type === 'mobile').value, '07700 900456');
  assert.equal(o[0].addresses[0].type, 'work');
  assert.equal(o[0].addresses[0].city, 'Leeds');
  assert.equal(o[0].note, 'Met at the fair');

  const plain = readContactsCsv('Name,Email\n"Lee, Kim",kim@example.com\nNobody,\n');
  assert.equal(plain.length, 2);
  assert.equal(plain[0].name.full, 'Lee, Kim');
  assert.equal(plain[0].emails[0].value, 'kim@example.com');
  assert.deepEqual(parseCsv('a,"b ""c"" d",e\n1,2,3\n'), [['a', 'b "c" d', 'e'], ['1', '2', '3']]);
});

test('a typed name splits the way a card keeps it', () => {
  assert.deepEqual(splitName('Dr. Kim Alex Lee'), { full: 'Dr. Kim Alex Lee', given: 'Kim', family: 'Lee', middle: 'Alex', prefix: 'Dr.', suffix: '' });
  assert.deepEqual(splitName('Madonna'), { full: 'Madonna', given: 'Madonna', family: '', middle: '', prefix: '', suffix: '' });
  assert.equal(initials({ name: { full: 'Ops desk' } }), 'OD');
});
