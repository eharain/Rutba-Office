// Read Aloud in Mail (D-2026-11-13-34): the message open, every message in
// the list, or its unread ones, each read as who sent it, its subject, then
// its words with what it quotes left out. The list as drawn decides, so a
// filter or a search decides what Play all reads.
import test from 'node:test';
import assert from 'node:assert/strict';
import { wordsToRead, textWithoutQuotes, pageWithoutQuotes, piecesOf, queueOf, createMailReading } from '../apps/desktop/renderer/apps/mail/read-aloud.js';
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
  // Threads: each thread's messages one after another, as allRowsIn gives them.
  const threaded = buildThreads([row('5', { subject: 'Budget', messageId: 'b1' }), row('6', { subject: 'Re: Budget', inReplyTo: 'b1', references: ['b1'] })]);
  const all = queueOf({ threads: threaded, rowsIn: (th) => th.messages, mode: 'all' });
  assert.deepEqual(all.map((r) => r.id).sort(), ['5', '6']);
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

