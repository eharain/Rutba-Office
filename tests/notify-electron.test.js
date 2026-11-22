// Electron's side of the notifier and the icon by the clock (office-shell
// notify-electron.js, D-2026-11-22-23), run in a real Electron: the icon and
// the icon with its mark are made from the application's own picture, the
// menu's rows are rows Electron takes, and a check run shows nothing.
//
// No icon is put by the clock, no notification is raised and no window opens:
// the probe makes pictures and a menu in memory and ends. It starts behind
// the guard (guarded-electron.js), so no error box reaches the screen.

import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { tempDir } from './temp-dirs.js';
import { writeGuarded, runGuarded } from './guarded-electron.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const url = (...parts) => JSON.stringify(pathToFileURL(path.join(root, ...parts)).href);
const icon = path.join(root, 'apps', 'desktop', 'resources', process.platform === 'win32' ? 'icon.ico' : 'icon.png');

test('in Electron: the icon and its mark at every size, a menu Electron takes, and nothing shown in a check run', () => {
  const dir = tempDir('notify-electron-');
  const profile = tempDir('notify-electron-profile-');
  writeGuarded(dir, `import { app, Menu } from 'electron';
const { trayIcons, notifyParts, TRAY_FLAG } = await import(${url('packages', 'office-shell', 'src', 'electron', 'notify-electron.js')});
const { menuRows } = await import(${url('packages', 'office-shell', 'src', 'electron', 'tray.js')});
await app.whenReady();
const out = { flag: TRAY_FLAG };

const icons = trayIcons(${JSON.stringify(icon)});
out.images = typeof icons.plain !== 'string' && typeof icons.waiting !== 'string';
out.sizes = [];
if (out.images) {
  for (const scaleFactor of [1, 1.25, 1.5, 2, 3]) {
    const px = Math.round(16 * scaleFactor);
    const plain = icons.plain.toBitmap({ scaleFactor });
    const waiting = icons.waiting.toBitmap({ scaleFactor });
    const at = (buf, x, y) => [...buf.subarray((y * px + x) * 4, (y * px + x) * 4 + 4)];
    const c = Math.floor(px - px * 0.27 - 1);
    out.sizes.push({ px, whole: plain.length === px * px * 4 && waiting.length === px * px * 4, differs: !plain.equals(waiting), dot: at(waiting, c, c), cornerSame: at(plain, 1, 1).join() === at(waiting, 1, 1).join() });
  }
}
const none = trayIcons(${JSON.stringify(path.join(root, 'no-such-picture.ico'))});
out.missing = [typeof none.plain, typeof none.waiting];

const parts = notifyParts({ iconPath: ${JSON.stringify(icon)}, windows: { all: () => [], infoFor: () => ({}) }, check: true, installed: true });
const shown = parts.show({ id: 'chat:u1', kind: 'chat', title: 'Amina', body: 'hello', silent: false, sticky: false, actions: [{ id: 'answer', label: 'Answer' }] });
parts.flash('talk');
out.check = { supported: parts.supported(), closes: typeof shown.close, recorded: globalThis.__rutbaNotifications.map((n) => [n.id, n.kind, n.title, n.body, n.actions.join()]), canStart: parts.startup.can, startOn: parts.startup.get(), startSet: parts.startup.set(true) };
const source = notifyParts({ iconPath: ${JSON.stringify(icon)}, windows: { all: () => [], infoFor: () => ({}) }, installed: false });
out.source = { canStart: source.startup.can, startOn: source.startup.get(), startSet: source.startup.set(true), supported: typeof source.supported() };

const rows = menuRows({
  state: { waiting: { missed: 1, chat: 2, mail: 3, total: 6 }, events: [{ id: 'e1', summary: 'Planning', start: Date.now() + 3600000, end: Date.now() + 7200000 }], paused: false, pausedUntil: 0 },
  name: 'Rutba Office',
  apps: ['talk', 'mail', 'calendar'],
  does: { update: () => ({ version: '9.9.9', run: () => {} }) },
});
const menu = parts.buildMenu(rows);
out.menu = { rows: rows.length, items: menu.items.length, quit: menu.getMenuItemById('quit')?.label };

console.log('RESULT=' + JSON.stringify(out));
app.exit(0);
`);
  const said = runGuarded(dir, [`--user-data-dir=${profile}`], { timeout: 120000 });
  const out = JSON.parse(/RESULT=(.*)/.exec(said)?.[1] || 'null');
  assert.ok(out, said);
  assert.equal(out.flag, '--tray');

  assert.equal(out.images, true, 'the icon was read and drawn');
  assert.deepEqual(out.sizes.map((s) => s.px), [16, 20, 24, 32, 48]);
  for (const s of out.sizes) {
    assert.equal(s.whole, true, `a whole picture at ${s.px} pixels`);
    assert.equal(s.differs, true, `the mark at ${s.px} pixels`);
    assert.deepEqual(s.dot, [77, 72, 229, 255], `the dot at ${s.px} pixels`);
    assert.equal(s.cornerSame, true, `the rest of the icon at ${s.px} pixels`);
  }
  assert.deepEqual(out.missing, ['string', 'string'], 'a picture that cannot be read leaves its path for the tray to draw');

  assert.deepEqual(out.check, { supported: true, closes: 'function', recorded: [['chat:u1', 'chat', 'Amina', 'hello', 'answer']], canStart: false, startOn: false, startSet: false });
  assert.deepEqual(out.source, { canStart: false, startOn: false, startSet: false, supported: 'boolean' }, 'a copy run from the source never starts with Windows');

  assert.equal(out.menu.items, out.menu.rows);
  assert.equal(out.menu.quit, 'Quit Rutba Office');
});
