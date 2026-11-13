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
import { openMessage, rowKey } from '../apps/desktop/renderer/apps/mail/parts.js';

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
