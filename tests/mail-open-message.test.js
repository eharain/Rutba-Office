// Delete, Move to, Star and the rest act on the message the person has open
// (18.0170, 13 Nov 2026). A message just clicked is on its way to the reading
// pane for a moment, and on a busy computer for some seconds, while the pane
// still shows the one before it; Delete pressed then took the message clicked,
// one the person could not see. Now, with nothing checked, the ribbon and its
// keys take the one the reading pane shows and none while it shows another,
// and the pane's own buttons (Forward, attachments, Unsubscribe, the source)
// act on the message the pane shows.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { openMessage, menuRows, afterOpenStep, rowKey, rowAsListed } from '../apps/desktop/renderer/apps/mail/parts.js';

const row = (id, folder = 'INBOX', accountId = 'a') => ({ accountId, folder, id });

test('the message selected is the one acted on once the reading pane shows it', () => {
  const del = row('9');
  assert.deepEqual(openMessage(del, rowKey(del)), del);
});

test('a message clicked while the pane still shows the one before it is not acted on', () => {
  const before = row('4');
  const clicked = row('9');
  assert.equal(openMessage(clicked, rowKey(before)), null, 'not the one clicked, unseen');
  assert.equal(openMessage(clicked, null), null, 'nor before anything is shown');
});

test('the same id in another folder or account is another message', () => {
  assert.equal(openMessage(row('9', 'Trash'), rowKey(row('9'))), null);
  assert.equal(openMessage(row('9', 'INBOX', 'b'), rowKey(row('9'))), null);
});

test('nothing selected, nothing to act on', () => {
  assert.equal(openMessage(null, rowKey(row('9'))), null);
});

const src = fs.readFileSync(new URL('../apps/desktop/renderer/apps/mail.js', import.meta.url), 'utf8').replace(/\r\n/g, '\n');

test('what the ribbon and its keys act on is the message open, and its buttons are greyed while there is none', () => {
  const targets = src.slice(src.indexOf('const targets = useCallback('), src.indexOf('const foldersFor = useCallback('));
  assert.match(targets, /const open = prefs\.layout === 'off' \? selected : openMessage\(selected, shownFor\);\n\s*if \(!open\) return \[\];/);
  assert.doesNotMatch(targets.slice(targets.indexOf('const open')), /\[selected\]/, 'the selection is not taken in place of the message open');
  assert.match(targets, /\[alone, checkedRows, selected, shownFor,/);
  assert.match(targets, /const noMessage = !checked\.size && !\(alone \|\| prefs\.layout === 'off' \? selected : openMessage\(selected, shownFor\)\);/);
  assert.doesNotMatch(src, /!selected && !checked\.size/, 'no ribbon button is enabled by the selection alone');
});

test('the pane says which message it shows, and one that will not open is not left showing the one before', () => {
  const opening = src.slice(src.indexOf('// Opening a message: the body'), src.indexOf('// A PDF or picture attached'));
  assert.match(opening, /setMessage\(m\);\n\s*setShownAt\(\{ accountId: a, folder: f, id \}\);/);
  const failed = opening.slice(opening.indexOf('}, (err) => {'));
  assert.match(failed, /if \(!current\) return;\n\s*setMessage\(null\);/);
  assert.match(failed, /setShownAt\(\{ accountId: a, folder: f, id \}\);/);
  assert.match(opening, /if \(!selected\) \{\n\s*setMessage\(null\);\n\s*setShownAt\(null\);/);
});

test('the pane\'s own buttons act on the message it shows, not the one selected', () => {
  for (const [what, re] of [
    ['Forward\'s attachments', /fromMessage: \{ \.\.\.shownAt, index:/],
    ['saving an attachment', /shell\.mail\.attachment\(\{ \.\.\.shownAt, index \}\)/],
    ['opening an attachment', /async \(index, meta, where = shownAt\) =>/],
    ['Unsubscribe', /shell\.mail\.unsubscribe\(\{ accountId: shownAt\.accountId, folder: shownAt\.folder, id: shownAt\.id \}\)/],
    ['the source', /shell\.mail\.source\(\{ accountId: shownAt\.accountId, folder: shownAt\.folder, id: shownAt\.id \}\)/],
    ['the invitation card', /<InviteCard shell=\{shell\} message=\{message\} where=\{shownAt\}/],
    ['the attachment notes', /<AttachmentNotes shell=\{shell\} where=\{shownAt\}/],
  ]) assert.match(src, re, what);
  assert.equal((src.match(/sendingFrom\(shownAt\?\.accountId \|\| accountId\)/g) || []).length, 2, 'reply and forward send from the shown message\'s account');
});

// A row's right-click menu acts on the row pointed at (the review of 18.0170):
// it took the message open, so Delete on a row that was not open deleted
// another. Pointed at the open row, a checked row, or in a message's own
// window, it is the ribbon's menu as before.
const conversation = { ...row('7'), messages: [row('5'), row('6'), row('7')] };

test('the menu on a row that is not open acts on that row', () => {
  const pointed = menuRows(row('9'), { open: row('4') });
  assert.deepEqual(pointed.rows, [row('9')]);
  assert.deepEqual(pointed.all, [row('9')]);
  assert.deepEqual(pointed.open, row('9'));
});

test('the menu on the open row, a checked row, or in its own window is the ribbon\'s', () => {
  assert.equal(menuRows(row('9'), { open: row('9') }), null, 'the open row');
  assert.equal(menuRows(row('9'), { checked: new Set([rowKey(row('9'))]) }), null, 'a checked row');
  assert.equal(menuRows(row('9'), { alone: true }), null, 'a message window');
  assert.equal(menuRows(null), null);
});

test('a conversation pointed at is the whole of it, and one holding the open message is the ribbon\'s', () => {
  const pointed = menuRows(conversation, { open: row('2'), threaded: true });
  assert.deepEqual(pointed.rows.map((r) => r.id), ['5', '6', '7']);
  assert.equal(menuRows(conversation, { open: row('6'), threaded: true }), null);
  assert.equal(menuRows(conversation, { checked: new Set([rowKey(row('5'))]), threaded: true }), null);
});

test('in a read-only archive the row can be copied but not changed', () => {
  const pointed = menuRows(row('9'), { readOnly: () => true });
  assert.deepEqual(pointed.rows, []);
  assert.deepEqual(pointed.all, [row('9')]);
});

test('the menu changes the rows pointed at, and opens the row before Reply, Forward and the rest run', () => {
  const menu = src.slice(src.indexOf('const rowMenu = (thread, e) =>'), src.indexOf('useWheelZoom('));
  for (const what of ['archive', 'flag', 'pin', 'unread', 'junk', 'delete']) assert.match(menu, new RegExp(`act\\('${what}', null, rows\\)`), what);
  assert.match(menu, /copyMessages\('copy', all\)/);
  assert.match(menu, /copyMessages\('cut', rows\)/);
  for (const id of ['mail.reply', 'mail.replyAll', 'mail.forward']) assert.ok(menu.includes(`opening('${id}')`), id);
  assert.match(menu, /const opening = \(id\) => on\(id, \(\) => openThen\(open, id\), draft\);/, 'opened first, and not offered on a draft');
  assert.match(menu, /on\('mail\.saveEml', \(\) => saveAsEml\(all\)/);
  assert.match(src, /onMenu=\{\(e\) => rowMenu\(thread, e\)\}/);
});

// What a Reply waiting for its row does: run, wait or drop (the review of the row menu).
const waiting = { key: rowKey(row('9')), form: 3 };
const now = (o) => ({ selectedKey: rowKey(row('9')), shownFor: rowKey(row('9')), loaded: true, form: 3, ...o });

test('a Reply from the row menu runs once the pane shows that row', () => {
  assert.equal(afterOpenStep(waiting, now()), 'run');
  assert.equal(afterOpenStep(null, now()), null, 'nothing waiting');
});

test('it waits while the pane still shows the message before', () => {
  assert.equal(afterOpenStep(waiting, now({ shownFor: rowKey(row('4')) })), 'wait');
  assert.equal(afterOpenStep(waiting, now({ shownFor: null, loaded: false })), 'wait');
});

test('another message chosen, something else started in the form, or a row that will not open: it is dropped', () => {
  assert.equal(afterOpenStep(waiting, now({ selectedKey: rowKey(row('5')) })), 'drop', 'another message');
  assert.equal(afterOpenStep(waiting, now({ selectedKey: null })), 'drop', 'nothing selected');
  assert.equal(afterOpenStep(waiting, now({ form: 4 })), 'drop', 'r, Ctrl+N or another form opened meanwhile');
  assert.equal(afterOpenStep(waiting, now({ loaded: false })), 'drop', 'shown, and it would not open');
});

test('the form being written is left before the row opens, and the pending command follows the form', () => {
  const then = src.slice(src.indexOf('const openThen = async (open, id) =>'), src.indexOf('const rowMenu'));
  assert.ok(then.indexOf('await leaveForm()') >= 0 && then.indexOf('await leaveForm()') < then.indexOf('setSelected('), 'left first, then opened');
  assert.ok(then.indexOf('form: composeKey.current') > then.indexOf('await leaveForm()'), 'the form recorded after it was left');
  assert.match(src, /afterOpenStep\(afterOpen\.current, \{ selectedKey: selected \? rowKey\(selected\) : null, shownFor, loaded: Boolean\(message\), form: composeKey\.current \}\)/);
  assert.match(src, /\}, \[selected, shownFor, message, compose\]\);/, 'looked at again when the form changes');
});

test('Safety turned off is held against the message the pane shows, not the selection', () => {
  assert.match(src, /const safety = !\(safetyOffFor && safetyOffFor === shownFor\);/);
  // Stable, so the message's own right-click menu (set up once) turns it off for the message shown then.
  assert.match(src, /const setSafety = useCallback\(\(on\) => setSafetyOffFor\(on \? null : shownForRef\.current\), \[\]\);/);
  assert.match(src, /onSafetyOff: \(\) => setSafety\(false\)/);
  assert.match(src, /useEffect\(\(\) => \{ setSafetyOffFor\(\(off\) => \(off === shownFor \? off : null\)\); \}, \[shownFor\]\);/);
});

test('with the reading pane off, Reply and Forward wait for the message selected to load', () => {
  assert.match(src, /const answerable = Boolean\(message\) && \(alone \|\| prefs\.layout !== 'off' \|\| Boolean\(openMessage\(selected, shownFor\)\)\);/);
  assert.equal((src.match(/if \(!answerable\) return;/g) || []).length, 2, 'reply and forward');
  for (const b of ["label={t('Reply')} disabled={!answerable}", "label={t('Reply all')} disabled={!answerable}", "label={t('Forward')} disabled={!answerable}"]) assert.ok(src.includes(b), b);
});

test('a message that will not open says so where it would have been', () => {
  assert.match(src, /const notOpened = openError && selected && shownFor === rowKey\(selected\) && !message \?/);
  assert.match(src, /title=\{t\('This message could not be opened'\)\}/);
  assert.match(src, /readerView \|\| notOpened \|\| fallback/, 'the message, else that it would not open');
  assert.equal((src.match(/\{paneBody\(/g) || []).length, 2, 'both places the pane is drawn');
});

// Star and Pin on the open message turned it on every time (14 Nov 2026): the
// selection holds only where the message is, never its star or pin.
test('the open message is acted on as the list has it, star and pin with it', () => {
  const listedRow = { ...row('9'), flagged: true, pinned: false, subject: 'x' };
  assert.deepEqual(rowAsListed(row('9'), { rows: [row('4'), listedRow] }), listedRow);
});

test('with no list, the open message gives its own star and pin', () => {
  const message = { flagged: true, pinned: true };
  assert.deepEqual(rowAsListed(row('9'), { message, shownFor: rowKey(row('9')) }), { ...row('9'), flagged: true, pinned: true });
  assert.deepEqual(rowAsListed(row('9'), { message, shownFor: rowKey(row('4')) }), row('9'), 'not another message\'s');
  assert.deepEqual(rowAsListed(row('9')), row('9'));
});

test('Star and Pin on the open message say so on it at once, and turn off as well as on', () => {
  assert.match(src, /setMessage\(\(m\) => \(m \? \{ \.\.\.m, flagged: on \} : m\)\)/);
  assert.match(src, /setMessage\(\(m\) => \(m \? \{ \.\.\.m, pinned: on \} : m\)\)/);
  assert.match(src, /if \(alone\) return selected \? writable\(\[listed\(selected\)\]\) : \[\];/);
  assert.match(src, /: \[listed\(open\)\]\);/);
});

// Read Aloud loaded each message twice, once to read it and again to show it.
test('a message Read Aloud loads is shown as loaded, not loaded again', () => {
  assert.match(src, /const loaded = quiet && moved \? readLoaded\.current\.get\(keyOf\(selected\)\) : undefined;/);
  assert.match(src, /loaded \? Promise\.resolve\(loaded\) : shell\.mail\.message\(/);
  assert.match(src, /if \(m\) readLoaded\.current\.set\(keyOf\(row\), m\);/);
});

// A docked form closed itself while Read Aloud showed what it read.
test('a docked form is kept, out of sight, while Read Aloud reads', () => {
  assert.match(src, /<div className="ml-docked-form" style=\{\{ display: reading \? 'none' : 'contents' \}\}>\{composeView\}<\/div>/);
});

// Review of 95418693: a form opened or left while Read Aloud reads went unseen.
const bodyOf = (name) => {
  const at = src.indexOf(`const ${name} = useCallback(`);
  assert.ok(at >= 0, name);
  return src.slice(at, src.indexOf('\n  }', at));
};

test('a form opened while Read Aloud reads stops the reading, and every form opens through that one door', () => {
  assert.match(bodyOf('setCompose'), /if \(readingStopsFor\(\{ reading: readingNow\.current, opening: Boolean\(next\) \}\)\) stopReadingRef\.current\?\.\(\);/);
  assert.match(src, /stopReadingRef\.current = stopReading;\n\s*readingNow\.current = Boolean\(reading\);/);
  // Nothing opens a form beside setCompose: updateCompose only changes the form already there.
  assert.deepEqual(src.match(/updateCompose\(/g), ['updateCompose(', 'updateCompose(']);
  assert.match(src, /onAccount=\{\(id\) => updateCompose\(\(d\) => \(\{ \.\.\.d, accountId: id \}\)\)\}/);
});

test('Reply, Reply all, Forward, New message, a message taken back and a draft opened all stop the reading', () => {
  for (const name of ['reply', 'forward', 'unsend', 'editDraft']) assert.match(bodyOf(name), /setCompose\(\{/, name);
  assert.match(src, /key: 'Mod\+N',\n\s*run: \(\) => \{\n\s*const fromId = newMailFrom;\n\s*setCompose\(\{/, 'New message');
});

test('a form left while Read Aloud keeps it out of sight is shown again first', () => {
  const leave = bodyOf('leaveForm');
  assert.match(leave, /if \(readingStopsFor\(\{ reading: readingNow\.current, leaving: true, hidden: formHidden\.current \}\)\) stopReadingRef\.current\?\.\(\);\n\s*const asked = /);
  assert.ok(leave.indexOf('readingStopsFor') < leave.indexOf('keepWriting'), 'before it is saved, closed or asked about');
  assert.match(src, /const hiddenForReading = Boolean\(reading\) && hasDocked;\n\s*formHidden\.current = hiddenForReading;/);
});

test('the reading over, the docked form takes the keys again', () => {
  assert.match(src, /if \(wasHidden\.current && !hiddenForReading && hasDocked\) document\.querySelector\('\.ml-docked-form \.ml-compose-body'\)\?\.focus\(\);/);
});

test('a re-run for the certificates loads the message again, not the copy Read Aloud loaded before them', () => {
  assert.match(src, /const loaded = quiet && moved \? readLoaded\.current\.get\(keyOf\(selected\)\) : undefined;/);
});
