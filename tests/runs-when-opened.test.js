// What runs when it is opened: the extension as Windows takes it, so a name
// that ends in a space, a dot or a stream's name is still a program.
import test from 'node:test';
import assert from 'node:assert/strict';
import { runsWhenOpened, openedExtension } from '@rutba/office-shell/runs';

test('programs are known by the extension Windows opens them under', () => {
  for (const name of ['setup.exe', 'C:\\Temp\\run.EXE', 'run.exe ', 'run.exe.', 'run.exe::$DATA', 'report.pdf.js', '/tmp/x.sh', 'disk.iso', 'link.lnk']) {
    assert.equal(runsWhenOpened(name), true, name);
  }
  for (const name of ['letter.docx', 'photo.jpg', 'README', 'C:\\setup.exe\\notes.txt', 'archive.zip', '']) {
    assert.equal(runsWhenOpened(name), false, name);
  }
  assert.equal(openedExtension('run.exe::$DATA'), 'exe');
  assert.equal(openedExtension('noext'), '');
});
