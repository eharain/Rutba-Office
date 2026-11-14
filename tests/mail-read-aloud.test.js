// Read Aloud in Mail (D-2026-11-13-34): the message open, every message in
// the list, or its unread ones, each read as who sent it, its subject, then
// its words with what it quotes left out. The list as drawn decides, so a
// filter or a search decides what Play all reads.
import test from 'node:test';
import assert from 'node:assert/strict';
import { wordsToRead, textWithoutQuotes, pageWithoutQuotes, piecesOf, queueOf, createMailReading, decodeEntities, languageOf, placeOfList, openingDoes, readButtons, readerStalled, marksReadAtEnd, marksReadAfterReading, createQuietOpen, readingStopsFor, refocusAfterReading } from '../apps/desktop/renderer/apps/mail/read-aloud.js';
import { arrange, buildThreads } from '../apps/desktop/renderer/apps/mail/parts.js';

test('a reply is read without the message it answers, its "> " lines or its signature', () => {
  const text = [
    'Thursday suits me.',
    '',
    'Shall we say ten?',
    '--',
    'Sam Okafor',
    'Head of Finance',
    '',
    'On Tue, 10 Nov 2026 at 09:12, Amina Yusuf <amina@work.example> wrote:',
    '> Can we meet this week?',
  ].join('\n');
  assert.deepEqual(wordsToRead({ text }), ['Thursday suits me.', 'Shall we say ten?']);
  // The "wrote:" a client wraps onto a line of its own.
  assert.equal(textWithoutQuotes('Yes.\nOn Tue, 10 Nov 2026 at 09:12, Amina Yusuf\nwrote:\n> Can we?'), 'Yes.');
  // Outlook: a rule, then From and Sent of the message answered.
  assert.equal(textWithoutQuotes('Done, thanks.\n________________________________\nFrom: Amina\nSent: 10 November 2026 09:12\nTo: Sam\nSubject: Report'), 'Done, thanks.');
  assert.equal(textWithoutQuotes('Done.\n\nFrom: Amina Yusuf\nSent: Tuesday\nSubject: Report\n\nThe report.'), 'Done.');
  assert.equal(textWithoutQuotes('Seen it.\n-----Original Message-----\nFrom: Amina'), 'Seen it.');
  // "From:" in the words themselves is not a header.
  assert.equal(textWithoutQuotes('From: the desk of the editor.\nAll well.'), 'From: the desk of the editor.\nAll well.');
});

test('a page is read without its quote, whichever program wrote the reply', () => {
  const gmail = '<div dir="ltr">Thursday suits me.</div><br><div class="gmail_quote"><div class="gmail_attr">On Tue, Amina wrote:</div><blockquote>Can we meet?</blockquote></div>';
  assert.deepEqual(wordsToRead({ html: gmail }), ['Thursday suits me.']);
  const outlook = '<p>Done, thanks.</p><div id="appendonsend"></div><hr><div id="divRplyFwdMsg"><b>From:</b> Amina</div><p>The report.</p>';
  assert.deepEqual(wordsToRead({ html: outlook }), ['Done, thanks.']);
  assert.equal(pageWithoutQuotes('<p>Yes.</p><blockquote><p>a</p><blockquote>b</blockquote></blockquote><p>Bye.</p>'), '<p>Yes.</p><p>Bye.</p>');
  // The text wins where a message carries both.
  assert.deepEqual(wordsToRead({ text: 'Plain words.', html: '<p>Page words.</p>' }), ['Plain words.']);
});

test('a forward with nothing added is read whole, and a web address as its site alone', () => {
  assert.deepEqual(wordsToRead({ text: '> The minutes are attached.\n> Amina' }), ['The minutes are attached. Amina']);
  assert.deepEqual(wordsToRead({ text: 'The agenda is at https://www.work.example/meetings/2026/11/agenda?id=42 for anyone.' }), ['The agenda is at work.example for anyone.']);
  assert.deepEqual(wordsToRead({ text: 'See <https://docs.example.org/a/b>.' }), ['See docs.example.org.']);
  // Lines of one paragraph are joined; a line of only marks is not read.
  assert.deepEqual(wordsToRead({ text: 'One line\nand the next.\n\n* * *\n\nNew paragraph.' }), ['One line and the next.', 'New paragraph.']);
  assert.deepEqual(wordsToRead({}), []);
});

test('a message is read as who sent it, its subject, then its words, each piece saying where it is', () => {
  const pieces = piecesOf({ from: [{ name: 'Amina Yusuf', address: 'amina@work.example' }], subject: 'Planning', text: 'Thursday?\n\nTen.' }, 'k1');
  assert.deepEqual(pieces.map((p) => [p.text, p.source.part]), [['From Amina Yusuf', 'from'], ['Subject: Planning', 'subject'], ['Thursday?', 'body'], ['Ten.', 'body']]);
  assert.ok(pieces.every((p) => p.source.key === 'k1' && p.language === null));
  const bare = piecesOf({ from: [{ address: 'lee@work.example' }], subject: '', text: '' }, 'k2');
  assert.deepEqual(bare.map((p) => p.text), ['From lee@work.example', '(no subject)']);
});

const row = (id, extra = {}) => ({ accountId: 'a', folder: 'Inbox', id, subject: `Message ${id}`, date: `2026-11-1${id}T09:00:00Z`, from: { address: 'x@work.example' }, ...extra });

test('Play all reads the list as drawn, in its order: a filter or a search decides which, and Play unread the unread of it', () => {
  const rows = [row('1'), row('2', { unread: true }), row('3', { flagged: true }), row('4', { unread: true, flagged: true })];
  const ids = (list) => list.map((r) => r.id);
  const drawn = (filter) => arrange(rows, { filter }).map((r) => ({ ...r, messages: [r] }));
  assert.deepEqual(ids(queueOf({ threads: drawn('all'), mode: 'all' })), ['4', '3', '2', '1'], 'newest first, as the list shows');
  assert.deepEqual(ids(queueOf({ threads: drawn('unread'), mode: 'all' })), ['4', '2'], 'the Unread filter on: only the unread');
  assert.deepEqual(ids(queueOf({ threads: drawn('starred'), mode: 'all' })), ['4', '3'], 'Starred');
  assert.deepEqual(ids(queueOf({ threads: drawn('all'), mode: 'unread' })), ['4', '2'], 'Play unread with no filter');
  assert.deepEqual(ids(queueOf({ threads: drawn('starred'), mode: 'unread' })), ['4'], 'Play unread within a filter');
  // Threads: the threads in the list's order, each thread's messages one after another as allRowsIn gives them.
  const threads = [{ messages: [row('7'), row('5', { unread: true })] }, { messages: [row('9', { unread: true })] }, { messages: [row('6')] }];
  assert.deepEqual(ids(queueOf({ threads, rowsIn: (th) => th.messages, mode: 'all' })), ['7', '5', '9', '6']);
  assert.deepEqual(ids(queueOf({ threads, rowsIn: (th) => th.messages, mode: 'unread' })), ['5', '9']);
  // Not threaded, allRowsIn gives the row alone.
  assert.deepEqual(ids(queueOf({ threads: [row('2'), row('1')], rowsIn: (r) => [r], mode: 'all' })), ['2', '1']);
  // buildThreads' own threads go through as they are.
  const built = buildThreads([row('5', { subject: 'Budget', messageId: 'b1' }), row('6', { subject: 'Re: Budget', inReplyTo: 'b1', references: ['b1'] })]);
  assert.deepEqual(ids(queueOf({ threads: built, rowsIn: (th) => th.messages, mode: 'all' })), built.flatMap((th) => th.messages.map((m) => m.id)));
});

/** A reader that reads at once: each read ends on the next turn, unless stopped. */
function fakeReader() {
  const reads = [];
  let gen = 0;
  return {
    reads,
    async read(pieces, { onEnd }) {
      const mine = ++gen;
      reads.push(pieces.map((p) => p.text));
      setImmediate(() => { if (mine === gen) onEnd(); });
      return true;
    },
    stop() { gen += 1; },
  };
}
const until = async (ok) => { for (let i = 0; i < 200 && !ok(); i++) await new Promise((r) => setImmediate(r)); };

test('the messages are read one after another, each asked for only when its turn comes, a gone one passed over', async () => {
  const reader = fakeReader();
  const asked = [];
  const shown = [];
  const states = [];
  let ended = 0;
  const messages = { 1: { from: [{ name: 'Amina' }], subject: 'One', text: 'First.' }, 3: { from: [{ name: 'Lee' }], subject: 'Three', text: 'Third.' } };
  const playing = createMailReading({
    reader,
    keyOf: (r) => r.id,
    load: async (r) => { asked.push(r.id); if (!messages[r.id]) throw new Error('gone'); return messages[r.id]; },
    onMessage: (r, n, count) => shown.push([r.id, n, count]),
    onState: (s) => states.push(s.playing),
    onEnd: () => { ended += 1; },
  });
  assert.equal(playing.play([]), false, 'nothing to play');
  assert.equal(playing.play([row('1'), row('2'), row('3')]), true);
  assert.deepEqual(asked, ['1'], 'only the first is asked for at the start');
  await until(() => ended === 1);
  assert.deepEqual(asked, ['1', '2', '3']);
  assert.deepEqual(reader.reads, [['From Amina', 'Subject: One', 'First.'], ['From Lee', 'Subject: Three', 'Third.']]);
  assert.deepEqual(shown, [['1', 1, 3], ['3', 3, 3]]);
  assert.equal(states.at(-1), false, 'not playing at the end');
  assert.equal(playing.state().playing, false);
});

test('next, the one before, and stop', async () => {
  const reads = [];
  const ends = [];
  const reader = {
    async read(pieces, { onEnd }) { reads.push(pieces[1].text); ends.push(onEnd); return true; }, // ends only when told
    stop() {},
  };
  let ended = 0;
  const msg = (id) => ({ from: [{ name: 'A' }], subject: id, text: 'Words.' });
  const playing = createMailReading({ reader, keyOf: (r) => r.id, load: async (r) => msg(r.id), onEnd: () => { ended += 1; } });
  playing.play([row('1'), row('2'), row('3')]);
  await until(() => reads.length === 1);
  playing.next();
  await until(() => reads.length === 2);
  assert.equal(playing.state().n, 2);
  playing.previous();
  await until(() => reads.length === 3);
  assert.deepEqual(reads, ['Subject: 1', 'Subject: 2', 'Subject: 1']);
  // A read that was left (next pressed) ending late does not move on.
  ends[0]();
  await until(() => false);
  assert.equal(reads.length, 3);
  playing.stop();
  ends[2]();
  await until(() => false);
  assert.equal(reads.length, 3, 'stopped: nothing more is read');
  assert.equal(ended, 0, 'stopping is not the end of the list');
  assert.equal(playing.state().playing, false);
});


test('what a page hides is not read: its title, a hidden preheader, Outlook\'s own copy, comments; and its characters as characters', () => {
  const html = [
    '<html><head><title>Newsletter 42</title><style>p{}</style></head><body>',
    '<div style="display:none;max-height:0">Preheader you never see</div>',
    '<span style="mso-hide:all">Hidden from Outlook</span>',
    '<!--[if mso]><p>Outlook copy of the text.</p><![endif]-->',
    '<!-- <p>a comment with a tag</p> -->',
    '<p>It&rsquo;s here &mdash; caf&eacute; at ten &#8217;til &#x2019;noon&#x2019;, &copy; us &amp; them, &bogus; kept.</p>',
    '</body></html>',
  ].join('');
  assert.deepEqual(wordsToRead({ html }), ['It’s here \u2014 café at ten ’til ’noon’, © us & them, &bogus; kept.']);
  assert.equal(decodeEntities('&Uuml;ber &ntilde; &ccedil; &#0; &#xD800; &#1114112;'), 'Über ñ ç  &#xD800; &#1114112;');
});

test('a forward is read as what it forwards, without its own line and the header under it', () => {
  const text = ['FYI, see below.', '', '---------- Forwarded message ---------', 'From: Amina Yusuf <amina@work.example>', 'Date: Tue, 10 Nov 2026 at 09:12', 'Subject: Minutes', 'To: Sam <sam@work.example>', '', 'The minutes are attached.', '', 'On Mon, 9 Nov 2026, Lee wrote:', '> older'].join('\n');
  assert.deepEqual(wordsToRead({ text }), ['FYI, see below.', 'The minutes are attached.']);
  assert.deepEqual(wordsToRead({ text: 'Begin forwarded message:\n\nFrom: Lee\nSubject: Hi\nDate: today\n\nHello there.' }), ['Hello there.']);
  const gmail = '<div>FYI</div><div class="gmail_quote"><div class="gmail_attr">---------- Forwarded message ---------<br>From: <strong>Amina</strong><br>Date: Tue<br>Subject: Minutes<br></div><br><div>The minutes are attached.</div></div>';
  assert.deepEqual(wordsToRead({ html: gmail }), ['FYI', 'The minutes are attached.']);
});

test('a reply\'s quote is found in German, French, Spanish, Italian and Dutch too', () => {
  assert.equal(textWithoutQuotes('Passt.\nAm Di., 10. Nov. 2026 um 09:12 Uhr schrieb Amina Yusuf <amina@work.example>:\n> Treffen?'), 'Passt.');
  assert.equal(textWithoutQuotes('Erledigt.\n\nVon: Amina Yusuf\nGesendet: Dienstag, 10. November 2026 09:12\nBetreff: Bericht'), 'Erledigt.');
  assert.equal(textWithoutQuotes('D\'accord.\nLe mar. 10 nov. 2026 à 09:12, Amina a écrit :\n> Réunion ?'), 'D\'accord.');
  assert.equal(textWithoutQuotes('Vale.\nEl mar, 10 nov 2026 a las 9:12, Amina escribió:\n> ¿Reunión?'), 'Vale.');
  assert.equal(textWithoutQuotes('Ok.\nIl giorno mar 10 nov 2026 alle ore 09:12 Amina ha scritto:\n> Riunione?'), 'Ok.');
  assert.equal(textWithoutQuotes('Prima.\nOp di 10 nov. 2026 om 09:12 schreef Amina <amina@work.example>:\n> Vergadering?'), 'Prima.');
  assert.equal(textWithoutQuotes('Gut.\n-----Ursprüngliche Nachricht-----\nVon: Amina'), 'Gut.');
  // "Am" opening an ordinary sentence is not a quote.
  assert.equal(textWithoutQuotes('Am I late?\nNo.'), 'Am I late?\nNo.');
  // Only German and Dutch put the writer after the verb: an English line that goes on past "wrote" is the message.
  assert.equal(textWithoutQuotes('On Friday, here is what I wrote about the plan:\nShip it.'), 'On Friday, here is what I wrote about the plan:\nShip it.');
});

test('each paragraph is read in the voice of its own script: Urdu, Arabic, Hindi, or the default', () => {
  assert.equal(languageOf('کل دس بجے ملتے ہیں۔'), 'ur');
  assert.equal(languageOf('نلتقي غدا في العاشرة.'), 'ar');
  assert.equal(languageOf('कल दस बजे मिलते हैं।'), 'hi');
  assert.equal(languageOf('See you at ten.'), null);
  assert.equal(languageOf('Meeting tomorrow: کل'), null, 'mostly Latin: the default');
  assert.equal(languageOf('2026'), null);
  const pieces = piecesOf({ from: [{ name: 'Amina' }], subject: 'x', text: 'کل دس بجے ملتے ہیں۔\n\nSee you.' }, 'k');
  assert.deepEqual(pieces.map((p) => p.language), [null, null, 'ur', null]);
});

test('reading stops when the list being read is not the one shown: account, folder, a folder of All accounts, view or scope', () => {
  const base = { accountId: 'a', folder: 'Inbox', role: 'inbox', view: 'mail', scope: 'folder', unified: false };
  const at = placeOfList(base);
  for (const change of [{ accountId: 'b' }, { folder: 'Sent' }, { view: 'files' }, { view: 'people' }, { scope: 'all' }]) assert.notEqual(placeOfList({ ...base, ...change }), at, JSON.stringify(change));
  assert.equal(placeOfList({ ...base, role: 'sent' }), at, 'one account: its folder decides, not the role');
  const all = { ...base, accountId: '*', unified: true };
  assert.notEqual(placeOfList({ ...all, role: 'sent' }), placeOfList(all), 'All accounts: another folder is another role');
  assert.equal(placeOfList({ ...all, folder: 'Other' }), placeOfList(all), 'All accounts: the folder of one account does not matter');
});

test('a message Read Aloud shows opens no draft, leaves no form and is not marked read; one clicked does as before', () => {
  assert.deepEqual(openingDoes({ quiet: true, unread: true, inDrafts: true, moved: true, formUntouched: true }), { editDraft: false, leaveForm: false, markRead: false });
  assert.deepEqual(openingDoes({ unread: true, moved: true, formUntouched: true }), { editDraft: false, leaveForm: true, markRead: true });
  assert.deepEqual(openingDoes({ inDrafts: true, moved: true, formUntouched: true }), { editDraft: true, leaveForm: false, markRead: false });
  assert.deepEqual(openingDoes({ inDrafts: true, readOnly: true, unread: true, moved: true, formUntouched: true }), { editDraft: false, leaveForm: true, markRead: false });
  assert.deepEqual(openingDoes({ unread: true, moved: true, formUntouched: false }), { editDraft: false, leaveForm: false, markRead: true }, 'a form opened since: kept');
});

test('a play\'s button stays pressed and live while its run goes on, and a message\'s own window has no Play all or Play unread', () => {
  const rows = [row('1', { unread: true }), row('2')];
  assert.deepEqual(readButtons({ rows, hasMessage: true, openKey: 'k' }), { open: { pressed: false, disabled: false }, all: { pressed: false, disabled: false }, unread: { pressed: false, disabled: false } });
  // Play unread under way, and the last unread one now read: still the pressed, live stop button.
  assert.deepEqual(readButtons({ reading: { mode: 'unread' }, rows: [row('1'), row('2')], hasMessage: true }).unread, { pressed: true, disabled: false });
  // Play all with the Unread filter, the list emptied as each was read.
  assert.deepEqual(readButtons({ reading: { mode: 'all' }, rows: [] }).all, { pressed: true, disabled: false });
  assert.deepEqual(readButtons({ rows: [] }).all, { pressed: false, disabled: true });
  assert.equal(readButtons({ rows: [row('2')] }).unread.disabled, true);
  const alone = readButtons({ rows, alone: true, hasMessage: true, reading: { mode: 'open', key: 'k' }, openKey: 'k' });
  assert.equal(alone.all, null);
  assert.equal(alone.unread, null);
  assert.deepEqual(alone.open, { pressed: true, disabled: false });
  assert.equal(readButtons({ reading: { mode: 'open', key: 'k' }, openKey: 'other', hasMessage: true }).open.pressed, false, 'another message open: Read aloud reads that one');
});

test('a message is marked read only once it has been read to its end; one skipped or stopped stays as it was', async () => {
  const ends = [];
  const reader = { async read(_p, { onEnd }) { ends.push(onEnd); return true; }, stop() {} };
  const read = [];
  const msg = { from: [{ name: 'A' }], subject: 's', text: 'Words.', unread: true };
  const playing = createMailReading({ reader, keyOf: (r) => r.id, load: async () => msg, onRead: (r) => read.push(r.id) });
  playing.play([row('1'), row('2'), row('3'), row('4')]);
  await until(() => ends.length === 1);
  playing.next(); // 1 skipped
  await until(() => ends.length === 2);
  ends[1](); // 2 read to its end
  await until(() => ends.length === 3);
  assert.deepEqual(read, ['2']);
  playing.stop(); // 3 stopped
  ends[2]();
  await until(() => false);
  assert.deepEqual(read, ['2'], 'stopped: not marked');
});

test('a voice that fails mid-message ends the run there, with nothing marked; the player\'s own stops are not taken for that', async () => {
  const ends = [];
  let stops = 0;
  let playing = null;
  const reader = { async read(_p, { onEnd }) { ends.push(onEnd); return true; }, stop() { stops += 1; playing?.stalled(); } };
  const read = [];
  let stalls = 0;
  let ended = 0;
  const msg = { from: [{ name: 'A' }], subject: 's', text: 'Words.' };
  playing = createMailReading({ reader, keyOf: (r) => r.id, load: async () => msg, onRead: (r) => read.push(r.id), onStall: () => { stalls += 1; }, onEnd: () => { ended += 1; } });
  playing.play([row('1'), row('2'), row('3')]);
  await until(() => ends.length === 1);
  playing.next(); // the reader is stopped, and says so: not a stall
  await until(() => ends.length === 2);
  playing.play([row('5'), row('6')]); // a new play stops the reader too: not a stall
  await until(() => ends.length === 3);
  assert.equal(stalls, 0);
  assert.ok(stops >= 2);
  assert.equal(playing.stalled(), true, 'the voice failed');
  assert.equal(stalls, 1);
  assert.equal(playing.state().playing, false);
  ends[2]();
  await until(() => false);
  assert.deepEqual(read, [], 'nothing marked read');
  assert.equal(ended, 0);
  assert.equal(playing.stalled(), false, 'once only');
  // The shared reader's state says when it stopped by itself.
  assert.equal(readerStalled({ playing: false, paused: false, index: 2, count: 5 }), true, 'a voice error, or no speech');
  assert.equal(readerStalled({ playing: false, paused: false, index: 5, count: 5 }), false, 'past its last sentence: the end, which it says');
  assert.equal(readerStalled({ playing: false, paused: true, index: 2, count: 5 }), false, 'paused');
  assert.equal(readerStalled({ playing: true, paused: false, index: 2, count: 5 }), false);
  assert.equal(readerStalled({ playing: false, paused: false, index: 0, count: 0 }), false, 'nothing to read');
});

test('only Play all and Play unread mark a message read at its end; Read aloud on the open message changes nothing', () => {
  assert.equal(marksReadAtEnd('all'), true);
  assert.equal(marksReadAtEnd('unread'), true);
  assert.equal(marksReadAtEnd('open'), false);
  assert.equal(marksReadAtEnd(null), false);
});

test('a hostile message is read, or not, quickly: a megabyte of what made the old patterns scan to the end from every place', () => {
  // Timed by this thread's own CPU time, as the other hostile-input tests are.
  // Without it, the process's CPU time counts the test runner's other threads too: a wider bound.
  const ownThread = typeof process.threadCpuUsage === 'function';
  const cpu = ownThread ? () => process.threadCpuUsage() : () => process.cpuUsage();
  const bound = ownThread ? 1500 : 4000;
  const MB = 1 << 20;
  const fill = (unit) => unit.repeat(Math.ceil(MB / unit.length));
  const hostile = [
    { html: fill('<div') },
    { html: fill('<div class="gmail_quote" ') },
    { html: '<!--' + fill('x') },
    { html: fill('<head') },
    { html: fill('<head>') },
    { html: fill('<title>') },
    { html: fill('<blockquote') },
    { html: fill('<blockquote>') },
    { html: fill('<div style="display:none">') + 'x' },
    { html: fill('<div class="gmail_quote">Forwarded message ') },
    { html: '<p style="' + fill(' ') },
    { html: fill('<p>a</p>\n') },
    { text: fill('\n') },
    { text: '> ' + fill('\n') },
    { text: 'a' + fill(' ') + 'b' },
    { text: fill('On x\n') },
    { text: fill('From: a\n') },
    { text: fill('-') },
    { text: fill('\n \n') },
    { text: fill('http://') },
    // One tag of a megabyte: an attribute with no value yet, or a long one (review 3).
    { html: '<div class=' + fill(' ') + '>x' },
    { html: '<div class="' + fill('a ') + '">x' },
    { html: '<div id=' + fill(' ') },
    { html: '<blockquote type=' + fill(' ') + '>x' },
    { html: '<span style=' + fill(' ') + '>x' },
    { html: '<div ' + fill('class=') + '>x' },
    { html: '<td style="font-size:0">' + fill('<p style="display:none">a') },
  ];
  for (const message of hostile) {
    const started = Date.now();
    const was = cpu();
    piecesOf({ from: [{ name: 'A' }], subject: 's', ...message }, 'k');
    const now = cpu();
    const ms = (now.user - was.user + now.system - was.system) / 1000;
    const what = JSON.stringify(message).slice(0, 40);
    assert.ok(ms < bound, `${what}: ${Math.round(ms)} ms of CPU time, ${Date.now() - started} ms by the clock`);
  }
});

test('a hidden element is left out to its own end, nested elements and all, and the ways a preheader hides count', () => {
  assert.deepEqual(wordsToRead({ html: '<div style="display:none">Hidden <div>inner</div> still hidden</div><p>Seen.</p>' }), ['Seen.']);
  for (const style of ['max-height:0;overflow:hidden', 'opacity:0', 'font-size:0px', 'visibility:hidden', 'color:#fff; font-size: 0', 'display: none !important']) {
    assert.deepEqual(wordsToRead({ html: `<span style="${style}">Preheader</span><p>Seen.</p>` }), ['Seen.'], style);
  }
  for (const style of ['opacity:0.5', 'font-size:14px', 'line-height:0', 'max-height:100px', 'min-height:0', 'max-height:0', 'height:0', 'font-size:0', 'opacity:0']) {
    assert.deepEqual(wordsToRead({ html: `<div style="${style}">Shown.</div>` }), ['Shown.'], `a block: ${style}`);
  }
  // A unquoted style hides as a quoted one does.
  assert.deepEqual(wordsToRead({ html: '<div style=display:none>Hidden</div><p>Seen.</p>' }), ['Seen.']);
  // MJML and hybrid layouts: font-size:0 on the cells that hold the whole message, sizes put back inside (review 3).
  const mjml = [
    '<html><head><title>Issue 42</title></head><body>',
    '<div style="display:none;font-size:1px;color:#ffffff;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;">Preview text you never see</div>',
    '<!--[if mso]><p>Outlook copy</p><![endif]-->',
    '<div style="font-size:0px;"><table><tr><td style="font-size:0px;padding:0;word-break:break-word;">',
    '<div style="font-family:Arial;font-size:16px;line-height:1.5;">Hello from the newsletter.</div>',
    '</td></tr></table></div></body></html>',
  ].join('');
  assert.deepEqual(wordsToRead({ html: mjml }), ['Hello from the newsletter.']);
  const hybridReply = '<table><tr><td style="font-size:0"><div style="font-size:14px">Thanks, that works.</div><div class="gmail_quote">On Tue, Lee wrote:<blockquote>The quoted words.</blockquote></div></td></tr></table>';
  assert.deepEqual(wordsToRead({ html: hybridReply }), ['Thanks, that works.']);
  // All quote, the page is read whole: still without its title, Outlook copy or hidden preheader.
  const allQuote = '<html><head><title>Issue 42</title></head><body><!--[if mso]>Outlook copy<![endif]--><div style="display:none">Preview</div><blockquote>Only the quote.</blockquote></body></html>';
  assert.deepEqual(wordsToRead({ html: allQuote }), ['Only the quote.']);
  // A p, li or td written without its end tag ends where HTML ends it.
  assert.deepEqual(wordsToRead({ html: '<p style="display:none">Hidden<p>Seen.</p>' }), ['Seen.']);
  assert.deepEqual(wordsToRead({ html: '<ul><li style="display:none">Hidden<li>Seen.</ul>' }), ['Seen.']);
  assert.deepEqual(wordsToRead({ html: '<table><tr><td style="display:none">Hidden<td>Seen.</table>' }), ['Seen.']);
  assert.deepEqual(wordsToRead({ html: '<div><p style="display:none">Hidden</div><p>Seen.</p>' }), ['Seen.']);
  // Not inside what it holds: a hidden cell's own table is all hidden.
  assert.deepEqual(wordsToRead({ html: '<table><tr><td style="display:none"><table><tr><td>Inner</td></tr></table></td><td>Seen.</td></tr></table>' }), ['Seen.']);
  // A head with no end ends at the body; a comment with no end hides the rest, as a browser has it.
  assert.deepEqual(wordsToRead({ html: '<html><head><meta charset="utf-8"><title>T</title><body><p>Words.</p></body>' }), ['Words.']);
  assert.deepEqual(wordsToRead({ html: '<p>Seen.</p><!-- never closed <p>gone' }), ['Seen.']);
  // Apple Mail forwards in a cited blockquote: what it forwards is read.
  const apple = '<div>FYI</div><blockquote type="cite"><div>Begin forwarded message:</div><div>From: Lee</div><div>Subject: Hi</div><br><div>Hello there.</div></blockquote>';
  assert.deepEqual(wordsToRead({ html: apple }), ['FYI', 'Hello there.']);
});

test('the mark of a message Read Aloud shows is taken by that selection only, and never reaches a later click', () => {
  let q = createQuietOpen();
  q.mark('a');
  assert.equal(q.take('a', 1), true, 'shown by Read Aloud');
  assert.equal(q.take('a', 2), true, 'a re-run because a certificate was trusted is quiet too');
  assert.equal(q.take('a', 2), false, 'a re-run for any other reason is not');
  q = createQuietOpen();
  q.mark('a');
  q.take('a', 1);
  assert.equal(q.take('b', 1), false, 'another message clicked');
  assert.equal(q.take('a', 1), false, 'and that one clicked again later: as any click');
  // The reading ends, stops or stalls on a message shown quietly: a click on it after is a click (review 3).
  q = createQuietOpen();
  q.mark('a');
  q.take('a', 1);
  q.clear();
  assert.equal(q.take('a', 1), false, 'clicked after the end');
  q = createQuietOpen();
  q.mark('a');
  q.take('a', 1);
  q.clear();
  assert.equal(q.take('a', 2), false, 'and a certificate re-run after the end is not quiet either');
  q = createQuietOpen();
  q.mark('a');
  q.take(null); // a folder changed: nothing selected
  assert.equal(q.take('a'), false, 'the mark went with the folder');
  q = createQuietOpen();
  q.mark('a');
  q.clear(); // reading stopped
  assert.equal(q.take('a'), false, 'the mark went with the stop');
  q = createQuietOpen();
  q.mark('a');
  q.unmark('a'); // the row was open already
  assert.equal(q.take('a'), false);
});

test('read to its end, a message is marked read in a run that marks, unless a person marked it by hand meanwhile', () => {
  const base = { mode: 'all', unread: true, readOnly: false, local: false, touched: false };
  assert.equal(marksReadAfterReading(base), true);
  assert.equal(marksReadAfterReading({ ...base, mode: 'unread' }), true);
  assert.equal(marksReadAfterReading({ ...base, mode: 'open' }), false, 'Read aloud on the open message changes nothing');
  assert.equal(marksReadAfterReading({ ...base, touched: true }), false, 'marked unread (or read) by hand during the reading');
  assert.equal(marksReadAfterReading({ ...base, unread: false }), false);
  assert.equal(marksReadAfterReading({ ...base, readOnly: true }), false);
  assert.equal(marksReadAfterReading({ ...base, local: true }), false, 'an archive or a file opened');
});

// Review of 95418693: a form opened or left while the reading keeps it out of sight.
test('a form opened while Read Aloud reads stops the reading, so it is seen', () => {
  assert.equal(readingStopsFor({ reading: true, opening: true }), true, 'Reply, Forward, New message, a draft, a message taken back');
  assert.equal(readingStopsFor({ reading: true, opening: true, hidden: false }), true, 'in the whole window too');
  assert.equal(readingStopsFor({ reading: false, opening: true }), false, 'nothing to stop');
});

test('a form left while the reading keeps it out of sight stops the reading first; one in sight does not', () => {
  assert.equal(readingStopsFor({ reading: true, leaving: true, hidden: true }), true);
  assert.equal(readingStopsFor({ reading: true, leaving: true, hidden: false }), false, 'a form the person can see');
  assert.equal(readingStopsFor({ reading: false, leaving: true, hidden: true }), false);
  assert.equal(readingStopsFor({ reading: true }), false, 'the reading moving on by itself');
  assert.equal(readingStopsFor(), false);
});

// Second review: the form took the keys from Search, a dialog or the list as a reading ended by itself.
test('the docked form takes the keys again on a stop the person asked for, or where the keys were not elsewhere', () => {
  assert.equal(refocusAfterReading({ asked: true }), true, 'Play pressed again, or the bar closed');
  assert.equal(refocusAfterReading({ idle: true }), true, 'the keys nowhere in particular');
  assert.equal(refocusAfterReading({ inPane: true }), true, 'the keys in the reading pane');
  assert.equal(refocusAfterReading({}), false, 'typing in Search or a dialog, or a folder chosen, as the reading ended');
  assert.equal(refocusAfterReading(), false);
});
