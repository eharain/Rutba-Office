// Office Mail as the Windows mail app (Mail 2.25): a mailto: link Windows
// opens Office with is taken to Mail, and the installer offers Office for
// mailto: links and .eml files in Settings, Default apps, without taking
// either from the program the person has chosen.
import test from 'node:test';
import assert from 'node:assert/strict';
import { parseLink } from '../packages/office-shell/src/links.js';
import { installerScript } from '../apps/desktop/build/make-app-packs.js';
import { describePack } from '../apps/desktop/main/app-packs.js';
import { PACKS } from '@rutba/office-shell/contract';

test('a mailto: link opens Mail with the link to write to', () => {
  assert.deepEqual(parseLink('mailto:amina@work.example?subject=Budget%20review&cc=sam@checks.example'), {
    app: 'mail', query: { mailto: 'mailto:amina@work.example?subject=Budget%20review&cc=sam@checks.example' },
  });
  assert.deepEqual(parseLink('MAILTO:amina@work.example'), { app: 'mail', query: { mailto: 'MAILTO:amina@work.example' } });
  assert.deepEqual(parseLink('  mailto:amina@work.example  '), { app: 'mail', query: { mailto: 'mailto:amina@work.example' } });
});

test('nothing else in a link opens Mail: no other scheme, nothing written over lines, nothing very long', () => {
  assert.equal(parseLink('mailto:'), null);
  assert.equal(parseLink('mailto:a@b.example\r\nBcc: everyone@x.example'), null);
  assert.equal(parseLink(`mailto:a@b.example?body=${'x'.repeat(9000)}`), null);
  assert.equal(parseLink('mailtox:a@b.example'), null);
  assert.equal(parseLink('https://work.example/mailto:a@b.example'), null);
  assert.equal(parseLink('C:\\Users\\sam\\mailto.eml'), null);
  // The Talk invitation still opens Talk.
  assert.deepEqual(parseLink('rutba-office://talk/invite/abcdefgh1234'), { app: 'talk', query: { invite: 'abcdefgh1234' } });
});

test('the installer offers Office for mailto: and .eml in Default apps, and takes neither', () => {
  const script = installerScript({ dir: 'x', packs: PACKS.map(({ id }) => ({ ...describePack(id), id })) });
  const from = script.indexOf('!macro rutbaWritePack_mail');
  const mail = script.slice(from, script.indexOf('!macroend', from));
  assert.match(mail, /!insertmacro rutbaMailLinks/, 'with Mail installed');
  const links = script.slice(script.indexOf('!macro rutbaMailLinks'), script.indexOf('!macroend', script.indexOf('!macro rutbaMailLinks')));
  assert.match(links, /Software\\Classes\\RutbaOffice\.mailto" "URL Protocol" ""/);
  assert.match(links, /Software\\Classes\\RutbaOffice\.mailto\\shell\\open\\command" "" "\$\\"\$appExe\$\\" \$\\"%1\$\\""/);
  assert.match(links, /Capabilities\\URLAssociations" "mailto" "RutbaOffice\.mailto"/);
  assert.doesNotMatch(script, /Software\\Classes\\mailto"/, 'the mailto class itself is the person\'s choice, never written');
  // Rutba Office listed in Default apps, written again from the packs chosen at each install.
  const offer = script.slice(script.indexOf('!macro rutbaOfferApps'), script.indexOf('!macroend', script.indexOf('!macro rutbaOfferApps')));
  assert.match(offer, /Software\\RegisteredApplications" "Rutba Office" "Software\\Rutba Office\\Capabilities"/);
  assert.match(offer, /Capabilities" "ApplicationName" "Rutba Office"/);
  const writes = script.slice(script.indexOf('!macro rutbaWritePacks'), script.indexOf('!macroend', script.indexOf('!macro rutbaWritePacks')));
  assert.ok(writes.indexOf('DeleteRegKey SHELL_CONTEXT "Software\\Rutba Office\\Capabilities"') >= 0 && writes.indexOf('DeleteRegKey') < writes.indexOf('!insertmacro rutbaWritePack_'), 'cleared before the packs write theirs');
  assert.ok(writes.indexOf('!insertmacro rutbaOfferApps') > writes.lastIndexOf('!insertmacro rutbaWritePack_'), 'and offered after');
  // Each kind of file is offered too, so Default apps lists them under Rutba Office.
  assert.match(script, /Capabilities\\FileAssociations" "\.\$\{EXT\}" "\$\{CLASS\}"/);
  const forget = script.slice(script.indexOf('!macro rutbaForgetTypes'), script.indexOf('!macroend', script.indexOf('!macro rutbaForgetTypes')));
  assert.match(forget, /!insertmacro rutbaForgetMailLinks/, 'and let go of by the uninstaller');
  const gone = script.slice(script.indexOf('!macro rutbaForgetMailLinks'), script.indexOf('!macroend', script.indexOf('!macro rutbaForgetMailLinks')));
  assert.match(gone, /DeleteRegKey SHELL_CONTEXT "Software\\Classes\\RutbaOffice\.mailto"/);
  assert.match(gone, /DeleteRegValue SHELL_CONTEXT "Software\\RegisteredApplications" "Rutba Office"/);
  assert.match(gone, /DeleteRegKey SHELL_CONTEXT "Software\\Rutba Office\\Capabilities"/);
});

test('without Mail, no mailto: offer is written', () => {
  const packs = PACKS.filter(({ id }) => id !== 'mail').map(({ id }) => ({ ...describePack(id), id }));
  const script = installerScript({ dir: 'x', packs });
  assert.ok(!/!insertmacro rutbaMailLinks/.test(script.slice(script.indexOf('!macro rutbaWritePacks'))), 'not called');
});
