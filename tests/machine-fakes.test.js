// The stand-ins for the computer's clipboard and global shortcuts (the gate
// in lanes): each holds what the real one would, of its own, so check runs
// can run side by side.
import test from 'node:test';
import assert from 'node:assert/strict';
import { fakeClipboard, fakeShortcuts } from '../packages/office-shell/src/electron/machine-fakes.js';

test('the fake clipboard keeps words, and words and a table together, as the system one does', async () => {
  const c = fakeClipboard();
  assert.equal(await c.readText(), '');
  await c.writeText('Pasted');
  assert.equal(await c.readText(), 'Pasted');
  await c.write([{ types: ['text/plain', 'text/html'], getType: async (t) => new Blob([t === 'text/plain' ? 'a\tb' : '<table><tr><td>a</td><td>b</td></tr></table>'], { type: t }) }]);
  const [item] = await c.read();
  assert.deepEqual(item.types, ['text/plain', 'text/html']);
  assert.equal(await (await item.getType('text/html')).text(), '<table><tr><td>a</td><td>b</td></tr></table>');
  assert.equal(await c.readText(), 'a\tb');
  await c.write([{ types: ['image/png'], getType: async () => new Blob([new Uint8Array([137, 80, 78, 71])], { type: 'image/png' }) }]);
  assert.equal(await c.readText(), '');
  assert.deepEqual([...new Uint8Array(await (await (await c.read())[0].getType('image/png')).arrayBuffer())], [137, 80, 78, 71]);
  await c.clear();
  assert.deepEqual(await c.read(), []);
});

test('two fake clipboards are two clipboards', async () => {
  const a = fakeClipboard();
  const b = fakeClipboard();
  await a.writeText('one');
  await b.writeText('two');
  assert.equal(await a.readText(), 'one');
  assert.equal(await b.readText(), 'two');
});

test('fake shortcuts are held once, pressed by a check, and let go', () => {
  const s = fakeShortcuts();
  let pressed = 0;
  assert.equal(s.register('CommandOrControl+Shift+F8', () => pressed++), true);
  assert.equal(s.register('commandorcontrol+shift+f8', () => {}), false);
  assert.equal(s.isRegistered('CommandOrControl+Shift+F8'), true);
  assert.equal(s.press('CommandOrControl+Shift+F8'), true);
  assert.equal(pressed, 1);
  s.unregister('CommandOrControl+Shift+F8');
  assert.equal(s.press('CommandOrControl+Shift+F8'), false);
  assert.deepEqual(s.registered(), []);
});
