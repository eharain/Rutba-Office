// Print for Mail, Calendar, Contacts, Image and Pictures (2.27.0): each app writes what it prints
// as one page of HTML, and Chromium cuts it into sheets in the hidden print window. These are the
// pages as the text they are: the paper and the frame they share, pictures laid on sheets, and
// what each app writes. The route through the main process is in print-page-route.test.js.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { t, tn } from '@rutba/office-ui/messages';
import {
  PAPER_MM, PX_PER_MM, MOST_TILES, esc, pageSetup, printableArea, sizeWords, printFileStem, pageHtml, gridFor, pictureSheets, tilesFor, posterSheets,
} from '../apps/desktop/renderer/print-page.js';
import { sentWords, messagePrintBody, messagePrintHtml } from '../apps/desktop/renderer/apps/mail/print.js';
import { slotDays, eventsByDay, calendarPrintHtml } from '../apps/desktop/renderer/apps/calendar-print.js';
import { letterOf, cardRows, contactPrintHtml, contactListPrintHtml } from '../apps/desktop/renderer/apps/contacts-print.js';
import { imageSheets, imagePrintHtml } from '../apps/desktop/renderer/apps/image/print.js';
import { MOST_PICTURES, filePictures, picturesPrintHtml } from '../apps/desktop/renderer/apps/pictures/print.js';

const count = (text, piece) => text.split(piece).length - 1;
const near = (a, b) => Math.abs(a - b) < 1e-6;
const sheetsIn = (html) => count(html, '<section class="sheet');

/* ── the paper and the frame ────────────────────────────────────────────── */

test('the dialog\'s choices become a page setup: a known paper, a direction, four margins', () => {
  assert.deepEqual(pageSetup(), { paper: 'A4', orientation: 'portrait', margins: { top: 12.7, right: 12.7, bottom: 12.7, left: 12.7 } });
  assert.equal(pageSetup({ paper: 'Foolscap' }).paper, 'A4', 'a paper this file does not know is A4');
  assert.equal(pageSetup({ paper: 'Letter', orientation: 'landscape' }).orientation, 'landscape');
  assert.equal(pageSetup({ orientation: 'sideways' }).orientation, 'portrait');
  const m = pageSetup({ margins: { top: 0, right: -3, bottom: 'x', left: 25.4 } }).margins;
  assert.deepEqual(m, { top: 0, right: 12.7, bottom: 12.7, left: 25.4 }, 'no margin is a margin; a wrong one is the normal one');
});

test('the room inside the margins follows the way the paper is turned', () => {
  const portrait = printableArea(pageSetup({ paper: 'A4' }));
  assert.ok(near(portrait.width, 184.6) && near(portrait.height, 271.6));
  const landscape = printableArea(pageSetup({ paper: 'A4', orientation: 'landscape' }));
  assert.ok(near(landscape.width, 271.6) && near(landscape.height, 184.6));
  const narrow = printableArea(pageSetup({ paper: 'A5', margins: { top: 5, right: 10, bottom: 15, left: 20 } }));
  assert.ok(near(narrow.width, 118) && near(narrow.height, 190));
});

test('the frame writes the paper and margins as @page, and a policy that fetches nothing', () => {
  const html = pageHtml({ title: 'A <page>', options: { paper: 'A4' }, body: '<p>Body</p>', css: '.x { color: red; }' });
  assert.match(html, /@page \{ size: 210mm 297mm; margin: 12\.7mm 12\.7mm 12\.7mm 12\.7mm; \}/);
  assert.match(html, /<title>A &lt;page&gt;<\/title>/);
  assert.ok(html.includes("default-src 'none'; img-src data: rutba:; style-src 'unsafe-inline'; font-src data:"));
  assert.ok(!html.includes('https:'), 'nothing comes from the web unless the app says the reader let it');
  assert.ok(html.includes('<p>Body</p>') && html.includes('.x { color: red; }'));
  assert.match(html, /<html lang="en" dir="ltr">/);
  assert.ok(!/<script/i.test(html));
});

test('the frame turns the paper, takes the reader\'s language and direction, and lets remote pictures in only when told', () => {
  const html = pageHtml({ options: { paper: 'Letter', orientation: 'landscape', margins: { top: 10, right: 20, bottom: 30, left: 40 } }, remote: true, lang: 'ur', dir: 'rtl' });
  assert.match(html, /@page \{ size: 279\.4mm 215\.9mm; margin: 10mm 20mm 30mm 40mm; \}/);
  assert.ok(html.includes('img-src data: rutba: https:;'));
  assert.match(html, /<html lang="ur" dir="rtl">/);
  assert.match(pageHtml({ lang: '"><x', dir: 'up' }), /<html lang="&quot;&gt;&lt;x" dir="ltr">/);
  for (const paper of Object.keys(PAPER_MM)) assert.ok(pageHtml({ options: { paper } }).includes(`size: ${PAPER_MM[paper].width}mm ${PAPER_MM[paper].height}mm`), paper);
});

test('text is made safe, sizes are said as a list of attachments says them, and a name is made a file\'s', () => {
  assert.equal(esc(`<a href="x">Tom & 'Jerry'</a>`), '&lt;a href=&quot;x&quot;&gt;Tom &amp; &#39;Jerry&#39;&lt;/a&gt;');
  assert.equal(esc(null), '');
  assert.equal(sizeWords(0), '0 B');
  assert.equal(sizeWords(-4), '0 B');
  assert.equal(sizeWords(512), '512 B');
  assert.equal(sizeWords(48230), '47.1 KB');
  assert.equal(sizeWords(3.1 * 1024 * 1024), '3.1 MB');
  assert.equal(sizeWords(250 * 1024 * 1024), '250 MB');
  assert.equal(printFileStem('Q3: figures/plan?  <final>'), 'Q3 figures plan final');
  assert.equal(printFileStem(''), 'print');
  assert.equal(printFileStem('???'), 'print');
  assert.equal(printFileStem('x'.repeat(200)).length, 80);
});

/* ── pictures on sheets ─────────────────────────────────────────────────── */

test('a grid of pictures turns with the paper, and an odd count to a page is one', () => {
  assert.deepEqual(gridFor(6, 'portrait'), { cols: 2, rows: 3, perPage: 6 });
  assert.deepEqual(gridFor(6, 'landscape'), { cols: 3, rows: 2, perPage: 6 });
  assert.deepEqual(gridFor(2, 'portrait'), { cols: 1, rows: 2, perPage: 2 });
  assert.deepEqual(gridFor(9), { cols: 3, rows: 3, perPage: 9 });
  assert.deepEqual(gridFor(5), { cols: 1, rows: 1, perPage: 1 });
});

test('pictures are laid so many to a sheet, each sheet the size of the room inside the margins', () => {
  const pictures = Array.from({ length: 14 }, (_, i) => ({ url: `rutba://file/p${i}`, name: `Holiday <${i + 1}>.png` }));
  const one = pictureSheets({ pictures, perPage: 1 });
  assert.equal(one.sheets, 14);
  const six = pictureSheets({ pictures, perPage: 6, captions: true });
  assert.equal(six.sheets, 3);
  assert.equal(count(six.body, '<img '), 14);
  assert.equal(count(six.body, 'class="cap"'), 14);
  assert.ok(six.body.includes('Holiday &lt;1&gt;.png'), 'a file name is text, never markup');
  assert.ok(six.body.includes('width:184.6mm;height:271mm'), 'a hair shorter than the page, so no sheet spills onto a blank one');
  assert.ok(six.body.includes('grid-template-columns:repeat(2,minmax(0,1fr));grid-template-rows:repeat(3,minmax(0,1fr))'));
  const nine = pictureSheets({ pictures, perPage: 9, options: { paper: 'Letter', orientation: 'landscape' } });
  assert.equal(nine.sheets, 2);
  assert.equal(count(nine.body, 'class="cap"'), 0, 'no names unless asked');
  assert.equal(pictureSheets({ pictures: [] }).sheets, 0);
});

test('a picture at its own size is counted in sheets at 96 pixels to the inch', () => {
  assert.deepEqual(tilesFor({ width: 600, height: 400 }).sheets, 1);
  const big = tilesFor({ width: 3000, height: 2000, options: { paper: 'A4' } });
  assert.deepEqual([big.across, big.down, big.sheets], [5, 2, 10]);
  const turned = tilesFor({ width: 3000, height: 2000, options: { paper: 'A4', orientation: 'landscape' } });
  assert.deepEqual([turned.across, turned.down, turned.sheets], [3, 3, 9]);
  // A picture a breath wider than the sheet is not worth a second sheet.
  const fits = Math.floor(184.6 * PX_PER_MM);
  assert.equal(tilesFor({ width: fits + 1, height: 100 }).across, 1);
  assert.equal(tilesFor({ width: fits + 12, height: 100 }).across, 2);
  assert.equal(tilesFor({ width: 0, height: NaN }).sheets, 1);
});

test('a picture at its own size is cut across its sheets, left to right and then down', () => {
  const poster = posterSheets({ url: 'rutba://blob/7', width: 3000, height: 2000, options: { paper: 'A4' } });
  assert.equal(poster.sheets, 10);
  assert.deepEqual([poster.across, poster.down], [5, 2]);
  const places = [...poster.body.matchAll(/left:(-?[\d.]+)mm;top:(-?[\d.]+)mm;width:([\d.]+)mm;height:([\d.]+)mm/g)].map((m) => m.slice(1).map(Number));
  assert.equal(places.length, 10);
  assert.deepEqual(places[0].slice(0, 2), [0, 0]);
  assert.deepEqual(places[1].slice(0, 2), [-184.6, 0]);
  assert.deepEqual(places[5].slice(0, 2), [0, -271]);
  assert.deepEqual(places[9].slice(0, 2), [-738.4, -271]);
  assert.ok(places.every(([, , w, h]) => near(w, 793.75) && near(h, 529.17)), 'every sheet draws the whole picture at one size, moved to show its part');
  const small = posterSheets({ url: 'rutba://blob/8', width: 300, height: 200 });
  assert.equal(small.sheets, 1);
});

test('a picture too large for one print is refused before anything is laid out', () => {
  const poster = posterSheets({ url: 'rutba://blob/9', width: 100000, height: 100000 });
  assert.equal(poster.tooMany, true);
  assert.equal(poster.body, '');
  assert.ok(poster.sheets > MOST_TILES);
});

/* ── Mail ───────────────────────────────────────────────────────────────── */

const message = {
  subject: 'Figures <draft> & the meeting',
  from: [{ name: 'Amina Yusuf', address: 'amina@example.com' }],
  to: [{ name: 'Bilal Khan', address: 'bilal@example.com' }, { address: 'team@example.com' }],
  date: '2026-11-20T09:15:00Z',
  attachments: [{ filename: 'figures.xlsx', size: 48230 }, { filename: 'logo.png', size: 900, inline: true }, { size: 12 }],
  html: '<p>Hello <b>there</b>.</p>',
};

test('a message prints under its subject, who it is from and to, when, and what is attached', () => {
  const html = messagePrintHtml(message, { tag: 'en' });
  assert.ok(html.includes('<h1>Figures &lt;draft&gt; &amp; the meeting</h1>'));
  assert.ok(html.includes(`<th scope="row">${t('common.from')}</th><td>Amina Yusuf &lt;amina@example.com&gt;</td>`));
  assert.ok(html.includes('Bilal Khan &lt;bilal@example.com&gt;, team@example.com'));
  assert.ok(html.includes(`figures.xlsx (47.1 KB), ${t('common.file')} (12 B)`), 'an attachment with no name is still said');
  assert.ok(!html.includes('logo.png'), 'a picture drawn in the body is not listed as attached');
  assert.equal(count(html, '<th scope="row">'), 4, 'no line for Cc or Bcc when there is none');
  assert.ok(html.includes('2026'));
  assert.ok(html.includes('<div class="msg"><template shadowrootmode="open">'), 'the body is drawn in a tree of its own');
  assert.ok(html.includes('<p>Hello <b>there</b>.</p>'));
  assert.ok(html.includes(`<th scope="row">${t('common.cc')}</th><td>cc@example.com</td>`) === false);
  const copied = messagePrintHtml({ ...message, cc: [{ address: 'cc@example.com' }], bcc: [{ address: 'bcc@example.com' }] }, { tag: 'en' });
  assert.equal(count(copied, '<th scope="row">'), 6);
});

test('a message with no subject, and one with nothing in it, still print', () => {
  const html = messagePrintHtml({ from: [{ address: 'a@example.com' }] });
  assert.ok(html.includes(`<h1>${esc(t('mail.no-subject'))}</h1>`));
  assert.ok(html.includes('<pre class="plain"></pre>'));
  assert.doesNotThrow(() => messagePrintHtml(null));
});

test('remote pictures print only when the reader has let them be shown', () => {
  const m = { html: '<p><img alt="x" src="https://tracker.example/p.gif"><img src="cid:logo"></p>', inlineUrls: { logo: 'rutba://blob/3' } };
  const blocked = messagePrintBody(m);
  assert.ok(blocked.includes('src="" data-blocked-remote="https://tracker.example/p.gif"'), 'the address is kept as a note, and nothing is fetched');
  assert.ok(blocked.includes('src="rutba://blob/3"'), 'a picture the message carries is drawn from where it is held');
  assert.ok(!messagePrintHtml(m).includes('rutba: https:'));
  const shown = messagePrintBody(m, { remote: true });
  assert.ok(shown.includes('src="https://tracker.example/p.gif"'));
  assert.ok(messagePrintHtml(m, { remote: true }).includes('img-src data: rutba: https:;'));
  assert.ok(!messagePrintHtml(m, { remote: true, plain: true }).includes('rutba: https:'), 'plain text fetches nothing whatever was let');
});

test('with Safety on a link prints as its words; with it off the address stays', () => {
  const m = { html: '<p>See <a class="l" href="https://example.com/report" title="t">the report</a> and <a href=one href=\'two\'>this</a>.</p><map><area shape="rect" href="https://example.com/x"></map>' };
  const safe = messagePrintBody(m);
  assert.ok(!/href/i.test(safe), 'no address is left to click in a PDF of the print');
  assert.ok(safe.includes('<a class="l" title="t">the report</a>'));
  const open = messagePrintBody(m, { links: true });
  assert.ok(open.includes('href="https://example.com/report"'));
});

test('a message cannot close the tree it is drawn in', () => {
  const body = messagePrintBody({ html: '<p>a</p></template><h1>From: someone else</h1></TEMPLATE ><p>b</p>' });
  assert.ok(!/<\/template/i.test(body));
  assert.ok(body.includes('&lt;/template>') && body.includes('&lt;/TEMPLATE >'));
  const html = messagePrintHtml({ subject: 'S', html: '</template><style>.head { display: none }</style>' });
  assert.equal(count(html.toLowerCase(), '</template>'), 1, 'only the frame\'s own closing tag is one');
});

test('a message shown as plain text prints as plain text', () => {
  const both = { html: '<p>Rich <b>words</b></p>', text: 'Plain <words> & more' };
  assert.equal(messagePrintBody(both, { plain: true }), '<pre class="plain">Plain &lt;words&gt; &amp; more</pre>');
  const onlyHtml = messagePrintBody({ html: '<p>Rich <b>words</b></p>' }, { plain: true });
  assert.ok(onlyHtml.startsWith('<pre class="plain">') && onlyHtml.includes('Rich') && !onlyHtml.includes('<b>'));
  assert.equal(messagePrintBody({ text: 'Only text' }), '<pre class="plain">Only text</pre>');
});

test('when a message was sent is said in full, and not at all when it is not known', () => {
  assert.equal(sentWords(null), '');
  assert.equal(sentWords('not a date'), '');
  assert.match(sentWords('2026-11-20T09:15:00Z', 'en'), /2026/);
  assert.match(sentWords('2026-11-20T09:15:00Z', 'not a tag!'), /2026/, 'a tag the computer does not know falls back to its own');
});

/* ── Calendar ───────────────────────────────────────────────────────────── */

const at = (day, hour, minute = 0) => new Date(2026, 10, day, hour, minute).getTime();
const allDay = (from, days) => ({ allDay: true, start: Date.UTC(2026, 10, from), end: Date.UTC(2026, 10, from + days) });
const NOV1 = new Date(2026, 10, 1);
const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

test('an event is on the days it covers: an all-day one by its dates, a timed one by this computer\'s clock', () => {
  const timed = slotDays({ start: at(10, 9), end: at(10, 10) });
  assert.deepEqual([iso(timed.first), iso(timed.last)], ['2026-11-10', '2026-11-10']);
  const toMidnight = slotDays({ start: at(10, 22), end: at(11, 0) });
  assert.equal(iso(toMidnight.last), '2026-11-10', 'an event that ends at midnight is not on the next day');
  const overnight = slotDays({ start: at(10, 22), end: at(11, 2) });
  assert.equal(iso(overnight.last), '2026-11-11');
  const trip = slotDays(allDay(10, 3));
  assert.deepEqual([iso(trip.first), iso(trip.last)], ['2026-11-10', '2026-11-12']);
  const none = slotDays({ allDay: true, start: Date.UTC(2026, 10, 10), end: Date.UTC(2026, 10, 10) });
  assert.equal(iso(none.last), '2026-11-10');
});

test('each day lists its events, all-day first and then by time, and an event over several days says which day it is', () => {
  const slots = [
    { summary: 'Late', start: at(10, 15), end: at(10, 16) },
    { summary: 'Early', start: at(10, 8), end: at(10, 9) },
    { summary: 'Trip', ...allDay(10, 3) },
    { summary: 'Before', start: at(2, 9), end: at(2, 10) },
  ];
  const days = eventsByDay(slots, new Date(2026, 10, 8), new Date(2026, 10, 12));
  assert.deepEqual([...days.keys()].sort(), ['2026-11-10', '2026-11-11'], 'only the days shown, and the range\'s end is not one');
  assert.deepEqual(days.get('2026-11-10').map((s) => s.summary), ['Trip', 'Early', 'Late']);
  assert.deepEqual(days.get('2026-11-10')[0].dayOf, { n: 1, count: 3 });
  assert.deepEqual(days.get('2026-11-11')[0].dayOf, { n: 2, count: 3 });
  assert.equal(days.get('2026-11-10')[1].dayOf, undefined);
  assert.equal(eventsByDay(null, NOV1, NOV1).size, 0);
});

test('a month prints as six weeks of seven days, every event of a day in its square', () => {
  const slots = [
    ...Array.from({ length: 9 }, (_, i) => ({ summary: `Event ${i + 1}`, start: at(4, 8 + i), end: at(4, 8 + i, 45), colour: '#c0392b' })),
    { summary: 'Trip <abroad>', ...allDay(10, 3), colour: 'red;background:url(x)' },
    { start: at(5, 9, 5), end: at(5, 10) },
  ];
  const html = calendarPrintHtml({ view: 'month', from: NOV1, to: new Date(2026, 11, 13), cursor: new Date(2026, 10, 15), slots, title: 'November 2026', options: { orientation: 'landscape' } });
  assert.ok(html.includes('<h1>November 2026</h1>'));
  assert.ok(html.includes('size: 297mm 210mm'));
  assert.equal(count(html, '<td class="'), 42);
  assert.equal(count(html, '<td class="other">'), 12, 'the days of the next month are marked');
  assert.equal(count(html, '<th scope="col">'), 7);
  assert.ok(html.includes('<th scope="col">Sun</th>'));
  assert.ok(html.includes('<div class="n">1 Nov</div>') && html.includes('<div class="n">1 Dec</div>'));
  for (let i = 1; i <= 9; i++) assert.ok(html.includes(`<span class="s">Event ${i}</span>`), `nothing is folded away as "more": event ${i}`);
  assert.ok(html.includes('<span class="t">8:00</span>') && html.includes('<span class="t">9:05</span>'));
  assert.equal(count(html, 'Trip &lt;abroad&gt;'), 3, 'on each of its days');
  assert.ok(html.includes('border-inline-start-color:#c0392b'));
  assert.ok(!html.includes('url(x)'), 'a colour that is not one is not written into the page');
  assert.ok(html.includes(`<span class="s">${esc(t('common.untitled'))}</span>`));
  assert.ok(html.includes('.month tr { break-inside: avoid; }'), 'a busy month goes over between two weeks, never through one');
});

test('a week prints its seven days side by side, each event with its times and place', () => {
  const from = new Date(2026, 10, 8);
  const slots = [{ summary: 'Stand-up', location: 'Room <4>', start: at(9, 9), end: at(9, 9, 45) }, { summary: 'Trip', ...allDay(10, 3) }, { summary: 'Next week', start: at(16, 9), end: at(16, 10) }];
  const html = calendarPrintHtml({ view: 'week', from, to: new Date(2026, 10, 15), slots, title: 'Week of 8 Nov 2026' });
  assert.equal(count(html, '<th scope="col">'), 7);
  assert.ok(html.includes('<th scope="col">Sun 8</th>') && html.includes('<th scope="col">Sat 14</th>'));
  assert.ok(html.includes('<div class="t">9:00 – 9:45</div><div class="s">Stand-up</div><div class="l">Room &lt;4&gt;</div>'));
  assert.ok(html.includes(`<div class="t">${esc(t('calendar.day-n-count', { n: 2, count: 3 }))}</div>`));
  assert.ok(!html.includes('Next week'));
});

test('a day and the agenda print as a list under each date, and say so when there is nothing', () => {
  const slots = [
    { summary: 'Review & plan', location: 'HQ', start: at(10, 14), end: at(10, 15) },
    { summary: 'Holiday', ...allDay(10, 1) },
    { summary: 'Later', start: at(20, 9), end: at(20, 10) },
  ];
  const day = calendarPrintHtml({ view: 'day', from: new Date(2026, 10, 10), to: new Date(2026, 10, 11), slots, title: 'Tuesday' });
  assert.equal(count(day, '<section class="day">'), 1);
  assert.ok(day.indexOf('Holiday') < day.indexOf('Review &amp; plan'));
  assert.ok(day.includes(`<td class="t">${esc(t('calendar.all-day-2'))}</td>`));
  assert.ok(day.includes('<td class="t">14:00 – 15:00</td>'));
  assert.ok(day.includes('Review &amp; plan<span class="l"> · HQ</span>'));
  assert.ok(day.includes('November 10, 2026'));
  assert.ok(!day.includes('Later'));

  const agenda = calendarPrintHtml({ view: 'agenda', from: NOV1, to: new Date(2026, 11, 1), slots, title: 'Next 30 days' });
  assert.equal(count(agenda, '<section class="day">'), 2, 'only the days that hold something');
  assert.ok(agenda.indexOf('November 10') < agenda.indexOf('November 20'));

  for (const view of ['day', 'agenda']) {
    const empty = calendarPrintHtml({ view, from: new Date(2026, 10, 2), to: new Date(2026, 10, 3), slots, title: 'Quiet <day>' });
    assert.ok(empty.includes(`<p class="none">${esc(t('calendar.print.no-events'))}</p>`), view);
    assert.ok(empty.includes('<h1>Quiet &lt;day&gt;</h1>'));
  }
});

/* ── Contacts ───────────────────────────────────────────────────────────── */

const amina = {
  display: 'Dr. Amina Yusuf',
  title: 'Head of Purchasing',
  org: 'Indus <Textiles>',
  categories: ['Suppliers', 'Karachi'],
  emails: [{ type: 'work', value: 'amina@indus.example' }, { label: 'private', value: 'amina@home.example' }, { type: 'other', value: '' }],
  phones: [{ type: 'mobile', value: '+92 300 1234567' }, { type: 'other', value: '021 111' }],
  addresses: [{ type: 'work', street: '12 Mill Road', city: 'Karachi', postcode: '74000', country: 'Pakistan' }],
  birthday: '1984-03-12',
  urls: [{ value: 'https://indus.example' }],
  note: 'Prefers calls before noon.\nSecond line.',
};

test('a card sorts under its letter as the window\'s list has it', () => {
  assert.equal(letterOf({ display: 'amina' }), 'A');
  assert.equal(letterOf({ display: '  Zed' }), 'Z');
  assert.equal(letterOf({ display: '9 Lives' }), '#');
  assert.equal(letterOf({ display: '' }), '#');
  assert.equal(letterOf(null), '#');
});

test('a card\'s rows are named as the window names them, in its order, and an empty one is left out', () => {
  assert.deepEqual(cardRows(amina), [
    [t('contacts.work'), 'amina@indus.example'],
    ['private', 'amina@home.example'],
    [t('contacts.mobile'), '+92 300 1234567'],
    [t('contacts.phone'), '021 111'],
    [t('contacts.work'), '12 Mill Road\nKarachi\n74000\nPakistan'],
    [t('contacts.birthday-2'), '1984-03-12'],
    [t('contacts.web'), 'https://indus.example'],
    [t('contacts.notes'), 'Prefers calls before noon.\nSecond line.'],
  ]);
  assert.deepEqual(cardRows(null), []);
  assert.deepEqual(cardRows({ addresses: [{}] }), []);
});

test('the chosen contact prints as its card in full', () => {
  const html = contactPrintHtml(amina, { options: { paper: 'A5' } });
  assert.ok(html.includes('size: 148mm 210mm'));
  assert.ok(html.includes('<h1>Dr. Amina Yusuf</h1>'));
  assert.ok(html.includes('<p class="sub">Head of Purchasing · Indus &lt;Textiles&gt;</p>'));
  assert.ok(html.includes('<p class="tags">Suppliers, Karachi</p>'));
  assert.equal(count(html, '<th scope="row">'), 8);
  assert.ok(html.includes('<td>Prefers calls before noon.\nSecond line.</td>'), 'a note keeps its lines');
  const bare = contactPrintHtml({});
  assert.ok(bare.includes(`<h1>${esc(t('contacts.no-name'))}</h1>`));
  assert.ok(!bare.includes('<table class="fields">'));
  assert.doesNotThrow(() => contactPrintHtml(null));
});

test('a room prints what it is, how many it seats and what it has', () => {
  const html = contactPrintHtml({ display: 'Board room', kind: 'room', room: { capacity: 12, where: 'Floor 2', equipment: ['Projector', 'Phone'] } });
  assert.ok(html.includes(`<p class="sub">${esc([t('common.room'), tn(12, 'common.count-seat'), 'Floor 2'].join(', '))}</p>`));
  assert.ok(html.includes('<p class="tags">Projector, Phone</p>'));
});

test('the shown list prints a line a card under their letters, with how many there are', () => {
  const list = [
    amina,
    { display: 'Asad Ali', emails: [{ value: 'asad@example.com' }, { value: 'asad2@example.com' }], phones: [{ label: 'desk', value: '111' }, { type: 'fax', value: '222' }, { value: '' }] },
    { display: 'Board room', kind: 'room', room: { capacity: 12 } },
    { display: '', org: 'Nameless & Co' },
  ];
  const html = contactListPrintHtml(list);
  assert.ok(html.includes(`<h1>${esc(t('common.contacts'))}</h1>`));
  assert.ok(html.includes(`<p class="count">${esc(tn(4, 'contacts.count-contact'))}</p>`));
  assert.deepEqual([...html.matchAll(/<tr class="letter"><td colspan="4">([^<]+)<\/td>/g)].map((m) => m[1]), ['D', 'A', 'B', '#'], 'the list is printed in the order the window shows it');
  assert.equal(count(html, '<td class="n">'), 4);
  assert.ok(html.includes('<td class="o">Head of Purchasing · Indus &lt;Textiles&gt;</td>'));
  assert.ok(html.includes('<td class="e v">asad@example.com\nasad2@example.com</td>'));
  assert.ok(html.includes(`<td class="p v">desk 111\n${t('contacts.fax')} 222</td>`));
  assert.ok(html.includes(`<td class="o">${esc([t('common.room'), tn(12, 'common.count-seat')].join(', '))}</td>`));
  assert.ok(html.includes(`<td class="n">${esc(t('contacts.no-name'))}</td><td class="o">Nameless &amp; Co</td>`));
  const named = contactListPrintHtml([], { title: 'Suppliers <2026>.vcf' });
  assert.ok(named.includes('<h1>Suppliers &lt;2026&gt;.vcf</h1>'));
  assert.ok(named.includes(`<p class="count">${esc(tn(0, 'contacts.count-contact'))}</p>`));
  assert.doesNotThrow(() => contactListPrintHtml(null));
});

/* ── Image ──────────────────────────────────────────────────────────────── */

test('the picture in Image prints fitted to one sheet, or at its own size over as many as it takes', () => {
  assert.equal(imageSheets({ width: 3000, height: 2000 }), 1);
  assert.equal(imageSheets({ width: 3000, height: 2000, size: 'actual' }), 10);
  assert.equal(imageSheets({ width: 3000, height: 2000, size: 'actual', options: { paper: 'A3', orientation: 'landscape' } }), 6);
  const fit = imagePrintHtml({ url: 'rutba://blob/print-1', width: 3000, height: 2000, name: 'Poster <1>.png' });
  assert.equal(sheetsIn(fit), 1);
  assert.equal(count(fit, 'src="rutba://blob/print-1"'), 1);
  assert.ok(fit.includes('<title>Poster &lt;1&gt;.png</title>'));
  assert.ok(!fit.includes('class="cap"'));
  const actual = imagePrintHtml({ url: 'rutba://blob/print-1', width: 3000, height: 2000, size: 'actual' });
  assert.equal(sheetsIn(actual), 10);
  assert.equal(count(actual, '<section class="sheet tile"'), 10);
  assert.equal(sheetsIn(imagePrintHtml({ url: 'rutba://blob/print-1', width: 3000, height: 2000, size: 'actual', options: { orientation: 'landscape' } })), 9);
});

test('a picture whose own size takes more sheets than one print can be is refused in words', () => {
  const sheets = tilesFor({ width: 40000, height: 30000 }).sheets;
  assert.equal(sheets, 1740);
  assert.throws(
    () => imagePrintHtml({ url: 'rutba://blob/print-2', width: 40000, height: 30000, size: 'actual' }),
    (err) => err.message === tn(sheets, 'image.print.actual-size-picture', { most: MOST_TILES }) && /1,740 sheets/.test(err.message) && err.message.includes('200')
  );
  assert.doesNotThrow(() => imagePrintHtml({ url: 'rutba://blob/print-2', width: 40000, height: 30000 }), 'fitted to the page it is one sheet');
});

/* ── Pictures ───────────────────────────────────────────────────────────── */

test('a file prints as itself, or as each of its pages when it has several', () => {
  assert.deepEqual(filePictures({ url: 'rutba://file/a', name: 'a.png' }), [{ url: 'rutba://file/a', name: 'a.png' }]);
  assert.deepEqual(filePictures({ url: 'rutba://file/a', name: 'a.png', pages: 0 }), [{ url: 'rutba://file/a', name: 'a.png' }]);
  assert.deepEqual(filePictures({ url: 'rutba://file/s', name: 'scan.tif', pages: 3, pageUrl: (u, i) => `${u}?page=${i}` }), [
    { url: 'rutba://file/s?page=0', name: 'scan.tif (1)' },
    { url: 'rutba://file/s?page=1', name: 'scan.tif (2)' },
    { url: 'rutba://file/s?page=2', name: 'scan.tif (3)' },
  ]);
});

test('the shown pictures print as a contact sheet, with their names beneath when asked', () => {
  const pictures = Array.from({ length: 14 }, (_, i) => ({ url: `rutba://file/p${i}`, name: `Holiday ${i + 1}.jpg` }));
  assert.equal(sheetsIn(picturesPrintHtml({ pictures })), 14);
  const sheet = picturesPrintHtml({ pictures, perPage: 6, captions: true, title: 'Holiday <2026>' });
  assert.equal(sheetsIn(sheet), 3);
  assert.equal(count(sheet, 'class="cap"'), 14);
  assert.ok(sheet.includes('<title>Holiday &lt;2026&gt;</title>'));
  assert.ok(sheet.includes('<div class="cap">Holiday 14.jpg</div>'));
  assert.equal(sheetsIn(picturesPrintHtml({ pictures, perPage: 9, options: { paper: 'Letter', orientation: 'landscape' } })), 2);
});

test('more pictures than one print can be are refused in words, and the most is allowed', () => {
  const many = (n) => Array.from({ length: n }, (_, i) => ({ url: `rutba://file/p${i}`, name: `${i}.jpg` }));
  assert.equal(sheetsIn(picturesPrintHtml({ pictures: many(MOST_PICTURES), perPage: 9 })), 23);
  assert.throws(
    () => picturesPrintHtml({ pictures: many(MOST_PICTURES + 1) }),
    (err) => err.message === tn(MOST_PICTURES + 1, 'pictures.print.count-picture-more', { most: MOST_PICTURES }) && err.message.includes('201')
  );
});
