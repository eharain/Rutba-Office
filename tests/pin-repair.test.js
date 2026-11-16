// A pin to a copy of Office that is gone starts this copy instead: the Mail
// pin left on "C:\Program Files\Rutba Office" after a move to this person's
// own folder (owner, 16 Nov), its ImplicitAppShortcuts copy with it.

import test from 'node:test';
import { win32 as path } from 'node:path';
import assert from 'node:assert/strict';
import { repairPins as repairPinsIn, repairOf, shortcutPlaces, isDeadCopy, installerFolders, pinsDue, pinStamp } from '../apps/desktop/main/pin-repair.js';

// The installer's folders on this made-up computer: for everyone, and for this person.
const FOLDERS = installerFolders({ ProgramW6432: 'C:\\Program Files', ProgramFiles: 'C:\\Program Files', LOCALAPPDATA: 'C:\\Users\\a\\AppData\\Local' });
const repairPins = (io) => repairPinsIn({ folders: FOLDERS, ...io });
const APPDATA = 'C:\\Users\\a\\AppData\\Roaming';
const DESKTOP = 'C:\\Users\\a\\Desktop';
const OLD = 'C:\\Program Files\\Rutba Office\\Rutba Office.exe';
const NEW = 'C:\\Users\\a\\AppData\\Local\\Programs\\Rutba Office\\Rutba Office.exe';
const PINNED = path.join(APPDATA, 'Microsoft', 'Internet Explorer', 'Quick Launch', 'User Pinned');

// A small disk: the files that exist, and the shortcuts with what each holds.
function disk({ files, links }) {
  const all = new Set([...files, ...Object.keys(links)].map((f) => f.toLowerCase()));
  const writes = [];
  return {
    writes,
    links,
    exists: (f) => all.has(f.toLowerCase()),
    list: (dir) => {
      const prefix = dir.toLowerCase().replace(/\\?$/, '\\');
      const names = new Map();
      for (const f of Object.keys(links)) {
        if (!f.toLowerCase().startsWith(prefix)) continue;
        const [first, ...rest] = f.slice(prefix.length).split('\\');
        names.set(first, rest.length > 0);
      }
      if (names.size === 0 && ![...all].some((f) => f.startsWith(prefix))) throw new Error('ENOENT');
      return [...names].map(([name, isDirectory]) => ({ name, isDirectory }));
    },
    read: (f) => ({ ...links[f] }),
    write: (f, change) => { writes.push([f, change]); Object.assign(links[f], change); return true; },
  };
}

const HUSK = { files: [OLD, NEW, path.join(path.dirname(NEW), 'resources', 'app.asar')] };

test('the Mail pin and its ImplicitAppShortcuts copy, left on a copy that is only its program, start this copy', () => {
  const taskbar = path.join(PINNED, 'TaskBar', 'Rutba Mail.lnk');
  const implicit = path.join(PINNED, 'ImplicitAppShortcuts', '69639df789022856', 'Rutba Mail.lnk');
  const pin = { target: OLD, args: '--app=mail "--profile=default"', icon: OLD, iconIndex: 3, cwd: path.dirname(OLD), appUserModelId: 'co.techstyle.rutba.office.mail' };
  const d = disk({ ...HUSK, links: { [taskbar]: { ...pin }, [implicit]: { ...pin } } });
  const { changed, failed } = repairPins({ exe: NEW, places: shortcutPlaces({ appData: APPDATA, desktop: DESKTOP }), ...d });
  assert.equal(failed, 0);
  assert.deepEqual(changed.map((c) => c.file).sort(), [implicit, taskbar].sort());
  // Electron's writer sets an icon only when it is given its index too.
  for (const [, change] of d.writes) assert.deepEqual(change, { target: NEW, icon: NEW, iconIndex: 3, cwd: path.dirname(NEW) });
  for (const f of [taskbar, implicit]) {
    assert.equal(d.links[f].target, NEW);
    assert.equal(d.links[f].icon, NEW);
    assert.equal(d.links[f].cwd, path.dirname(NEW));
    // What the pin asks for stays: the app, the profile, the icon in the program, the app id.
    assert.equal(d.links[f].args, '--app=mail "--profile=default"');
    assert.equal(d.links[f].iconIndex, 3);
    assert.equal(d.links[f].appUserModelId, 'co.techstyle.rutba.office.mail');
  }
});

test('a pin to a folder with nothing left in it is repaired too', () => {
  const desk = path.join(DESKTOP, 'Rutba Office.lnk');
  const d = disk({ files: [NEW, path.join(path.dirname(NEW), 'resources', 'app.asar')], links: { [desk]: { target: OLD, icon: 'C:\\Users\\a\\icons\\office.ico' } } });
  repairPins({ exe: NEW, places: shortcutPlaces({ appData: APPDATA, desktop: DESKTOP }), ...d });
  assert.equal(d.links[desk].target, NEW);
  assert.equal(d.links[desk].icon, 'C:\\Users\\a\\icons\\office.ico', 'an icon from elsewhere stays');
});

test('a pin to a working copy, to this copy, or to another program is left alone', () => {
  const other = 'D:\\Office test\\Rutba Office.exe';
  const links = {
    [path.join(PINNED, 'TaskBar', 'Working.lnk')]: { target: other },
    [path.join(PINNED, 'TaskBar', 'Here.lnk')]: { target: NEW.toUpperCase() },
    [path.join(PINNED, 'TaskBar', 'Notepad.lnk')]: { target: 'C:\\Windows\\notepad.exe' },
    [path.join(PINNED, 'TaskBar', 'Nothing.lnk')]: {},
  };
  const d = disk({ files: [NEW, other, 'D:\\Office test\\resources\\app.asar'], links });
  assert.deepEqual(repairPins({ exe: NEW, places: shortcutPlaces({ appData: APPDATA, desktop: DESKTOP }), ...d }), { changed: [], failed: 0 });
  assert.equal(d.writes.length, 0);
});

test('only as deep as each place, a folder that is not there, and a shortcut it cannot read or write', () => {
  const deep = path.join(PINNED, 'ImplicitAppShortcuts', 'a', 'b', 'Deep.lnk');
  const unreadable = path.join(PINNED, 'TaskBar', 'Broken.lnk');
  const unwritable = path.join(PINNED, 'TaskBar', 'Locked.lnk');
  const fine = path.join(PINNED, 'TaskBar', 'Mail.lnk');
  const d = disk({ ...HUSK, links: { [deep]: { target: OLD }, [unreadable]: { target: OLD }, [unwritable]: { target: OLD }, [fine]: { target: OLD } } });
  const read = d.read;
  d.read = (f) => { if (f === unreadable) throw new Error('not a shortcut'); return read(f); };
  const write = d.write;
  d.write = (f, c) => { if (f === unwritable) throw new Error('access denied'); return write(f, c); };
  const places = [...shortcutPlaces({ appData: APPDATA, desktop: 'C:\\Users\\a\\Not here' })];
  const { changed, failed } = repairPins({ exe: NEW, places, ...d });
  assert.deepEqual(changed, [{ file: fine, from: OLD }]);
  assert.equal(failed, 1, 'the one it could not write is counted, so the next start tries it again');
  assert.equal(d.links[deep].target, OLD, 'three folders down is not a place pins are kept');
});

test('a copy is dead when its program is gone or has no app beside it', () => {
  const exists = (f) => [OLD, NEW, 'C:\\Users\\a\\AppData\\Local\\Programs\\Rutba Office\\resources\\app.asar'].map((x) => x.toLowerCase()).includes(f.toLowerCase());
  assert.equal(isDeadCopy(OLD, { exists }), true);
  assert.equal(isDeadCopy(NEW, { exists }), false);
  assert.equal(isDeadCopy('E:\\Gone\\Rutba Office.exe', { exists }), true);
  assert.equal(repairOf({ target: 'C:\\Program Files\\Rutba Office\\rutba office.EXE' }, NEW, { exists, folders: FOLDERS })?.target, NEW, 'the name in any case');
});

test('only the installer\'s folders: a copy in a folder of the person\'s choosing, a build, or a share that is away is left alone', () => {
  const exists = (f) => [NEW, path.join(path.dirname(NEW), 'resources', 'app.asar')].map((x) => x.toLowerCase()).includes(f.toLowerCase());
  for (const target of ['E:\\Rutba Office\\Rutba Office.exe', 'D:\\office\\apps\\desktop\\release\\win-unpacked\\Rutba Office.exe', '\\\\server\\apps\\Rutba Office\\Rutba Office.exe']) {
    assert.equal(repairOf({ target }, NEW, { exists, folders: FOLDERS }), null, target);
  }
  assert.deepEqual(FOLDERS, ['C:\\Program Files\\Rutba Office', 'C:\\Users\\a\\AppData\\Local\\Programs\\Rutba Office']);
});

test('the start folder follows when it names the dead folder in any spelling, or is gone', () => {
  const exists = (f) => [OLD, NEW, path.join(path.dirname(NEW), 'resources', 'app.asar'), 'C:\\Users\\a'].map((x) => x.toLowerCase()).includes(f.toLowerCase());
  const io = { exists, folders: FOLDERS };
  assert.equal(repairOf({ target: OLD, cwd: 'C:\\Program Files\\Rutba Office\\' }, NEW, io).cwd, path.dirname(NEW), 'a closing separator');
  assert.equal(repairOf({ target: OLD, cwd: 'C:\\Somewhere gone' }, NEW, io).cwd, path.dirname(NEW), 'a folder that is not there');
  assert.equal(repairOf({ target: 'C:\\Program Files\\Rutba Office\\\\Rutba Office.exe' }, NEW, io)?.target, NEW, 'a doubled separator');
  assert.equal(repairOf({ target: OLD, cwd: 'C:\\Users\\a' }, NEW, io).cwd, undefined, 'a start folder of its own stays');
});

test('the pins are looked at once for each version and folder, or again when the other folder holds only a program', () => {
  const asar = path.join(path.dirname(NEW), 'resources', 'app.asar');
  const io = (files) => ({ version: '2.26.0', exe: NEW, folders: FOLDERS, exists: (f) => files.map((x) => x.toLowerCase()).includes(f.toLowerCase()) });
  const clean = [NEW, asar];
  const stamp = (version, exe, files = clean) => pinStamp({ ...io(files), version, exe });
  assert.equal(pinsDue({ ...io(clean), last: '' }), true, 'never looked');
  assert.equal(pinsDue({ ...io(clean), last: stamp('2.25.0', NEW) }), true, 'a new version');
  assert.equal(pinsDue({ ...io(clean), last: stamp('2.26.0', OLD) }), true, 'the same version, from the other folder');
  assert.equal(pinsDue({ ...io(clean), last: stamp('2.26.0', NEW.toUpperCase()) }), false, 'looked already');
  assert.equal(pinsDue({ ...io([...clean, OLD]), last: stamp('2.26.0', NEW) }), true, 'the other folder has come to hold only its program');
  // Looked at with that husk there, which a copy for this person cannot clear: not again at every start.
  assert.equal(pinsDue({ ...io([...clean, OLD]), last: stamp('2.26.0', NEW, [...clean, OLD]) }), false, 'the same husk, looked at already');
  assert.equal(pinsDue({ ...io([...clean, OLD, 'C:\\Program Files\\Rutba Office\\resources\\app.asar']), last: stamp('2.26.0', NEW) }), false, 'a working copy in the other folder');
  assert.equal(pinsDue({ ...io(clean), last: stamp('2.26.0', NEW, [...clean, OLD]) }), true, 'the husk cleared: looked at once more, then recorded clean');
});
