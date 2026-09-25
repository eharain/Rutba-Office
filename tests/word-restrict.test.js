/**
 * Review → Restrict Editing: w:documentProtection and the regions anyone
 * may edit, the engine's side.
 *
 * Word hashes a Restrict Editing password with Excel's salted, spun SHA-512
 * — but of the Word 97 key made from the password, not the password itself
 * ([MS-OFFCRYPTO] 2.3.7.4 and [MS-OI29500]). The key's tables are checked
 * against the spec's own numbers, the hash against a step-by-step rewrite of
 * it, and the attributes against what Word writes. Then the document: the
 * element in settings.xml where the schema puts it, Stop Protection leaving
 * it unenforced, exception regions as w:permStart/w:permEnd that survive
 * the paragraph being edited and split, and the view refusing what the
 * protection refuses.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { wordLegacyKey, wordLegacyHex, wordProtectionAttrs, checkWordPassword, WORD_MATRIX, WORD_INITIAL_CODE, legacyPasswordHash, hashPassword } from '@rutba/ooxml/protection';

/** Give a built document a settings part of its own. */
function withSettings(pkg, xml) {
  if (pkg.has('word/settings.xml')) return pkg.write_('word/settings.xml', xml);
  pkg.addPart('word/settings.xml', Buffer.from(xml, 'utf8'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
  pkg.addRelationshipTo('word/document.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml');
}

const settingsOf = (bytes) => OoxmlPackage.read(Buffer.from(bytes)).text('word/settings.xml');

test('the Word 97 key uses the spec\'s tables, and its low word is the 16-bit verifier', () => {
  // [MS-OFFCRYPTO] 2.3.7.4: the first and last rows of the encryption
  // matrix and the initial code table, as the spec prints them.
  assert.deepEqual(WORD_MATRIX[0], [0xaefc, 0x4dd9, 0x9bb2, 0x2745, 0x4e8a, 0x9d14, 0x2a09]);
  assert.deepEqual(WORD_MATRIX[7], [0x47d3, 0x8fa6, 0x0f6d, 0x1eda, 0x3db4, 0x7b68, 0xf6d0]);
  assert.deepEqual(WORD_MATRIX[14], [0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48c4]);
  assert.equal(WORD_INITIAL_CODE.length, 15);
  assert.equal(WORD_INITIAL_CODE[0], 0xe1f0);
  assert.equal(WORD_INITIAL_CODE[14], 0x4ec3);
  // Excel's legacy sheet hash of "test" is CBEB; Word's verifier is the same
  // function of the same bytes.
  assert.equal(legacyPasswordHash('test'), 'CBEB');
  assert.equal(wordLegacyKey('test') & 0xffff, 0xcbeb);
  // The key written low byte first: EB CB, then the high word's bytes.
  const key = wordLegacyKey('test');
  assert.equal(wordLegacyHex('test'), [key, key >>> 8, key >>> 16, key >>> 24].map((b) => (b & 0xff).toString(16).toUpperCase().padStart(2, '0')).join(''));
  assert.ok(wordLegacyHex('test').startsWith('EBCB'));
  // Passwords past fifteen characters are cut to fifteen, as Word cuts them.
  assert.equal(wordLegacyKey('abcdefghijklmnopq'), wordLegacyKey('abcdefghijklmno'));
});

test('Restrict Editing\'s hash is SHA-512 of the Word key, salted and spun, under Word\'s attribute names', () => {
  const salt = Buffer.from('00112233445566778899aabbccddeeff', 'hex').toString('base64');
  const a = wordProtectionAttrs('Secret!', { salt, spinCount: 100000 });
  assert.deepEqual(Object.keys(a), ['w:cryptProviderType', 'w:cryptAlgorithmClass', 'w:cryptAlgorithmType', 'w:cryptAlgorithmSid', 'w:cryptSpinCount', 'w:hash', 'w:salt']);
  assert.equal(a['w:cryptProviderType'], 'rsaAES');
  assert.equal(a['w:cryptAlgorithmSid'], '14', 'SHA-512');
  assert.equal(a['w:cryptSpinCount'], '100000');
  // Written out step by step: H0 = SHA-512(salt + key as UTF-16LE hex), then
  // Hn = SHA-512(Hn-1 + iterator), 100,000 times.
  let h = createHash('sha512').update(Buffer.from(salt, 'base64')).update(Buffer.from(wordLegacyHex('Secret!'), 'utf16le')).digest();
  for (let i = 0; i < 100000; i++) {
    const it = Buffer.alloc(4);
    it.writeUInt32LE(i, 0);
    h = createHash('sha512').update(h).update(it).digest();
  }
  assert.equal(a['w:hash'], h.toString('base64'));

  const read = Object.fromEntries(Object.entries(a).map(([k, v]) => [k.slice(2), v]));
  assert.equal(checkWordPassword('Secret!', read), true);
  assert.equal(checkWordPassword('secret!', read), false);
  assert.equal(checkWordPassword('', read), false);
  // The ISO names (algorithmName, hashValue, saltValue, spinCount) are read
  // too, whichever form of the password another writer hashed.
  const iso = { algorithmName: 'SHA-512', hashValue: hashPassword('Plain', { salt, spinCount: 1000 }), saltValue: salt, spinCount: '1000' };
  assert.equal(checkWordPassword('Plain', iso), true);
  assert.equal(checkWordPassword('plain', iso), false);
  assert.equal(checkWordPassword('anything', { edit: 'readOnly' }), true, 'no password: nothing to check');
});

test('w:documentProtection is written where the schema puts it, read back, and left unenforced by Stop Protection', () => {
  const pkg = OoxmlPackage.read(buildDocx({ paragraphs: ['One.', 'Two.'] }));
  withSettings(pkg, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:trackRevisions/><w:defaultTabStop w:val="720"/><w:compat/></w:settings>');
  const view = openDocx(pkg.write());
  assert.equal(view.protection(), null);
  view.setProtection({ edit: 'readOnly', formatting: true, password: 'Secret!' });
  const settings = settingsOf(view.save());
  assert.match(settings, /<w:trackRevisions\/><w:documentProtection w:edit="readOnly" w:formatting="1" w:enforcement="1" w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" w:hash="[^"]+" w:salt="[^"]+"\/><w:defaultTabStop/);
  const p = view.protection();
  assert.deepEqual({ edit: p.edit, enforced: p.enforced, formatting: p.formatting, hasPassword: p.hasPassword }, { edit: 'readOnly', enforced: true, formatting: true, hasPassword: true });

  assert.throws(() => view.stopProtection('secret!'), /That password is not right/);
  assert.equal(view.protection().enforced, true, 'a wrong password leaves it on');
  view.stopProtection('Secret!');
  assert.match(settingsOf(view.save()), /<w:documentProtection w:edit="readOnly" w:formatting="1" w:enforcement="0"\/>/);
  assert.equal(view.protection().enforced, false);
  assert.equal(view.protection().hasPassword, false);
});

test('Exceptions → Everyone marks paragraphs with w:permStart/w:permEnd that survive editing them', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Locked heading.', 'Fill in here.', 'And here.', 'Locked footer.'] }));
  view.setSelection({ block: 1, offset: 0 }, { block: 2, offset: 3 });
  view.setPermission(true);
  const xml = OoxmlPackage.read(view.save()).text('word/document.xml');
  assert.match(xml, /<w:permStart w:id="1" w:edGrp="everyone"\/><w:r><w:t[^>]*>Fill in here\./);
  assert.match(xml, /And here\.<\/w:t><\/w:r><w:permEnd w:id="1"\/><\/w:p>/);
  assert.deepEqual(view.protection().regions, [{ id: '1', from: 1, to: 2 }]);

  // Enforced read-only: the region's blocks can be edited, the rest cannot.
  view.setProtection({ edit: 'readOnly' });
  assert.deepEqual([0, 1, 2, 3].map((i) => view.canEditBlock(i)), [false, true, true, false]);

  // Typing at the region's start keeps the region's start before the words,
  // and Enter at the end of its last paragraph makes a new last paragraph
  // still inside it, as in Word.
  view.collapseTo({ block: 1, offset: 0 });
  view.insertText('Name: ');
  view.collapseTo({ block: 2, offset: 'And here.'.length });
  view.splitParagraph();
  view.insertText('Third line.');
  const after = OoxmlPackage.read(view.save()).text('word/document.xml');
  assert.match(after, /<w:permStart w:id="1" w:edGrp="everyone"\/><w:r><w:t[^>]*>Name: Fill in here\./);
  assert.deepEqual(view.protection().regions, [{ id: '1', from: 1, to: 3 }], 'the new paragraph is inside the region');
  assert.deepEqual([0, 1, 2, 3, 4].map((i) => view.canEditBlock(i)), [false, true, true, true, false]);

  // Exceptions are set before protection is enforced.
  view.setSelection({ block: 0, offset: 0 });
  assert.throws(() => view.setPermission(true), /before protection is enforced/);
  view.stopProtection('');
  view.setSelection({ block: 2, offset: 0 });
  view.setPermission(false);
  assert.equal(view.protection().regions.length, 0);
  assert.doesNotMatch(OoxmlPackage.read(view.save()).text('word/document.xml'), /w:perm(Start|End)/);
});

test('protected for tracked changes: recording is on and cannot be turned off; comments and forms edit nothing else', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Draft text.'] }));
  view.setProtection({ edit: 'trackedChanges' });
  assert.equal(view.recording, true);
  assert.match(settingsOf(view.save()), /<w:trackRevisions\/>/);
  assert.throws(() => view.setTrackChanges(false), /Track Changes stays on/);
  assert.equal(view.canEditBlock(0), true, 'edits are allowed, and recorded');
  view.collapseTo({ block: 0, offset: 0 });
  view.insertText('New ');
  assert.match(OoxmlPackage.read(view.save()).text('word/document.xml'), /<w:ins\b[^>]*>(?:(?!<\/w:ins>).)*New /s);

  view.stopProtection();
  view.setProtection({ edit: 'comments' });
  assert.equal(view.canEditBlock(0), false);
  view.setProtection({ edit: 'forms' });
  assert.equal(view.canEditBlock(0), false);
});

test('a file Word protected opens protected', () => {
  // Word's own element, as Word 2016 writes it, on a document with one region.
  const pkg = OoxmlPackage.read(buildDocx({ paragraphs: ['Terms.', 'Signature: ______'] }));
  const a = wordProtectionAttrs('Contract');
  withSettings(pkg, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/>' +
    `<w:documentProtection w:edit="readOnly" w:enforcement="1" ${Object.entries(a).map(([k, v]) => `${k}="${v}"`).join(' ')}/>` +
    '<w:defaultTabStop w:val="720"/></w:settings>');
  pkg.write_('word/document.xml', pkg.text('word/document.xml').replace(/(<w:p>)(<w:r><w:t[^>]*>Signature)/, '$1<w:permStart w:id="1482113094" w:edGrp="everyone"/>$2').replace(/(______<\/w:t><\/w:r>)(<\/w:p>)/, '$1<w:permEnd w:id="1482113094"/>$2'));
  const view = openDocx(pkg.write());
  const p = view.protection();
  assert.equal(p.enforced, true);
  assert.equal(p.edit, 'readOnly');
  assert.equal(p.hasPassword, true);
  assert.deepEqual(p.regions, [{ id: '1482113094', from: 1, to: 1 }]);
  assert.deepEqual([view.canEditBlock(0), view.canEditBlock(1)], [false, true]);
  assert.throws(() => view.stopProtection('contract'), /not right/);
  view.stopProtection('Contract');
  assert.equal(view.protection().enforced, false);
});

test('the document service refuses what the protection refuses, whichever way the edit comes', async () => {
  const { createDocumentService } = await import('../apps/desktop/main/documents.js');
  const fs = await import('node:fs');
  const os = await import('node:os');
  const path = await import('node:path');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-restrict-'));
  const file = path.join(dir, 'form.docx');
  fs.writeFileSync(file, buildDocx({ paragraphs: ['Fixed terms.', 'Name: ', 'More fixed terms.'] }));
  const service = createDocumentService({ holdBlob: () => ({}) });
  const s = service.open({ path: file, kind: 'doc' });
  const at = (block, offset) => ({ op: 'setSelection', anchor: { block, offset }, focus: { block, offset } });
  const refused = (ops) => {
    try {
      service.apply({ id: s.id, ops });
      return null;
    } catch (err) {
      return err.message;
    }
  };
  service.apply({ id: s.id, ops: [at(1, 0), { op: 'setPermission', on: true }, { op: 'setProtection', edit: 'readOnly' }] });

  assert.match(refused([at(0, 2), { op: 'insertText', text: 'X' }]), /not allowed because the selection is locked/);
  assert.match(refused([at(0, 2), { op: 'toggleFormat', tag: 'bold' }]), /not allowed/);
  assert.equal(refused([at(1, 6), { op: 'insertText', text: 'Ada' }]), null, 'inside the region');
  assert.match(refused([at(1, 0), { op: 'deleteBackward' }]), /not allowed/, 'Backspace would join the locked paragraph before');
  assert.match(refused([at(1, 9), { op: 'deleteForward' }]), /not allowed/, 'Delete would join the locked paragraph after');
  assert.match(refused([{ op: 'toggleTrackChanges', on: true }]), /not allowed/);

  service.apply({ id: s.id, ops: [{ op: 'stopProtection', password: '' }, { op: 'setProtection', edit: 'trackedChanges' }] });
  assert.match(refused([{ op: 'toggleTrackChanges', on: false }]), /Track Changes stays on/);
  assert.equal(refused([at(0, 0), { op: 'insertText', text: 'New ' }]), null);

  service.apply({ id: s.id, ops: [{ op: 'stopProtection', password: '' }, { op: 'setProtection', edit: 'comments' }] });
  assert.match(refused([at(0, 0), { op: 'insertText', text: 'Q' }]), /only comments/);
  assert.equal(refused([at(0, 0), { op: 'addComment', text: 'Why fixed?' }]), null, 'a comment is allowed');

  service.apply({ id: s.id, ops: [{ op: 'stopProtection', password: '' }, { op: 'setProtection', edit: 'none', formatting: true }] });
  assert.match(refused([at(0, 2), { op: 'toggleFormat', tag: 'bold' }]), /formatting is limited to styles/);
  assert.equal(refused([at(0, 2), { op: 'insertText', text: 'ok ' }]), null, 'formatting-only protection leaves the words editable');
  service.close({ id: s.id });
});
