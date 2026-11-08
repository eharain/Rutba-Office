// A setup: its list, its file with its media, finding sources again, the
// first start of 2.23 and undo (packages/studio/src/setup.js).

import test from 'node:test';
import assert from 'node:assert/strict';
import {
  normaliseSetup, newSetup, duplicateSetup, dropStream, normaliseIndex, withSetup, withoutSetup, markUsed, menuOrder, openingSetup, uniqueName,
  setupId, emptyHistory, historyPush, historyUndo, historyRedo, programOf, rematch, sceneMissing, migrate, writeSetupFile, readSetupFile, placeMedia,
  mediaPaths, SETUP_MIME, setupSettings, globalSettings, joinSettings,
} from '../packages/studio/src/setup.js';
import {
  normaliseStack, addItem, applyScene, saveScene, renameScene, duplicateScene, moveScene, removeScene, updateSceneFromFrame, updateItem, quickScenes,
  newSetupStack, materialiseScenes, backSoonScene, soundOnlyScene, BACK_SOON, SOUND_ONLY, recordKind, emptyStack,
} from '../packages/studio/src/stack.js';
import { unzip } from '../packages/studio/src/rvid.js';

const screen1 = { id: 'screen:0:0', name: 'Screen 1', kind: 'screen', display: { width: 1920, height: 1080 } };
const screen2 = { id: 'screen:1:0', name: 'Screen 2', kind: 'screen', display: { width: 2560, height: 1440 } };
const excel = { id: 'window:100:0', name: 'Q3 figures.xlsx - Excel', kind: 'window' };
const word = { id: 'window:200:0', name: 'Letter.docx - Word', kind: 'window' };
const cam = { id: 'cam-a', name: 'HD webcam' };
const mic = { id: 'mic-a', name: 'Headset' };

const setupWith = (patch = {}) => newSetup({ id: 'a', name: 'A', screen: screen1, now: 1, ...patch });

test('a new setup has the four quick scenes, the sources they need, and Screen and camera up', () => {
  const s = setupWith();
  assert.deepEqual(s.stack.scenes.map((sc) => sc.key), ['screen', 'window', 'screen-camera', 'camera']);
  assert.deepEqual(s.stack.items.filter((it) => it.on).map((it) => it.kind), ['camera', 'screen']);
  assert.deepEqual(s.stack.sounds.map((x) => [x.kind, x.live]), [['mic', true], ['system', false]]);
  assert.equal(s.stack.items.find((it) => it.kind === 'window').ref, '', 'a window is chosen, not assumed');
  assert.equal(s.stack.items.find((it) => it.kind === 'screen').ref, screen1.id);
  const noCam = newSetup({ id: 'b', name: 'B', screen: screen1, hasCamera: false });
  assert.deepEqual(noCam.stack.scenes.map((sc) => sc.key), ['screen', 'window']);
});

test('the quick scenes are ordinary scenes: renamed, copied, moved, updated from the frame and removed like any', () => {
  let st = setupWith().stack;
  st = renameScene(st, 'screen', 'Whole screen');
  assert.equal(st.scenes[0].name, 'Whole screen');
  st = duplicateScene(st, 'camera', 'Camera again');
  assert.deepEqual(st.scenes.map((sc) => sc.name).slice(-2), ['Camera', 'Camera again']);
  const added = st.scenes[st.scenes.length - 1].key;
  st = moveScene(st, added, 0);
  assert.equal(st.scenes[0].key, added);
  st = moveScene(st, added, 2);
  assert.equal(st.scenes[1].key, added, 'the gap past itself: one on');
  st = applyScene(st, st.scenes.find((sc) => sc.key === 'camera'));
  st = updateSceneFromFrame(st, 'screen');
  assert.deepEqual(st.scenes.find((sc) => sc.key === 'screen').layers.map((l) => l.id), ['camera-1']);
  st = removeScene(st, 'window');
  assert.ok(!st.scenes.some((sc) => sc.key === 'window'));
});

test('scenes are no longer held to nine, keep their keys, and old places in follow and rotate find them', () => {
  let st = emptyStack();
  ({ stack: st } = addItem(st, { kind: 'screen', ref: 'x', name: 'x' }));
  for (let i = 0; i < 12; i++) st = saveScene(st, `Scene ${i + 1}`);
  const again = normaliseStack(st);
  assert.equal(again.scenes.length, 12);
  assert.equal(new Set(again.scenes.map((sc) => sc.key)).size, 12);
  const old = normaliseStack({ items: [], scenes: [{ name: 'A', layers: [] }, { name: 'B', layers: [] }], follow: { 'own-1': 'mic-1' }, rotate: { every: 5, keys: ['own-0', 'wide'] } });
  assert.equal(old.follow[old.scenes[1].key], 'mic-1');
  assert.deepEqual(old.rotate.keys, [old.scenes[0].key, 'wide']);
});

test('Back soon is the placeholder and its music; Sound only is no picture; any other scene brings both back', () => {
  const st = setupWith().stack;
  const back = applyScene(st, backSoonScene());
  assert.equal(back.video, 'placeholder');
  assert.equal(back.sound, 'music');
  assert.equal(backSoonScene({ backSoon: 'Zurück' }).name, 'Zurück');
  assert.equal(backSoonScene().key, BACK_SOON);
  const on = applyScene(back, st.scenes[0]);
  assert.equal(on.video, 'on');
  assert.equal(on.sound, 'on');
  const sound = applyScene(st, soundOnlyScene());
  assert.equal(soundOnlyScene().key, SOUND_ONLY);
  assert.equal(recordKind(sound), 'sound');
  assert.deepEqual(normaliseStack(back).video, 'placeholder');
});

test('Auto is one mode, kept with the setup; anything else is off', () => {
  assert.equal(normaliseStack({ auto: 'rotate' }).auto, 'rotate');
  assert.equal(normaliseStack({ auto: 'follow' }).auto, 'follow');
  assert.equal(normaliseStack({ auto: 'sideways' }).auto, 'off');
  assert.equal(normaliseStack({}).auto, 'off');
  assert.equal(normaliseStack({}).previewFirst, false, 'straight on air is the default');
  assert.equal(normaliseStack({ previewFirst: true }).previewFirst, true);
});

test('settings part into those of the setup and those of every setup, and come back whole', () => {
  const studio = { size: 720, fps: 24, quality: 'high', folder: 'D:/x', countdown: 5, shot: 'selected', shotTimer: 3, minimise: true, liveCopy: false, isolated: true };
  const own = setupSettings(studio);
  const all = globalSettings(studio);
  assert.deepEqual(own, { size: 720, fps: 24, quality: 'high', folder: 'D:/x', liveCopy: false, isolated: true });
  assert.deepEqual(all, { countdown: 5, minimise: true, shot: 'selected', shotTimer: 3 });
  const joined = joinSettings(own, all);
  assert.equal(joined.size, 720);
  assert.equal(joined.countdown, 5);
  assert.equal(joined.isolated, true);
});

test('the list: newest used first, the tools\' own hidden, each app on the one it last used', () => {
  let index = normaliseIndex({});
  const a = setupWith({ id: 'a', name: 'A', now: 1 });
  const b = setupWith({ id: 'b', name: 'B', now: 5, app: 'broadcast' });
  const tools = normaliseSetup({ id: 't', name: 'For the AI tools', tools: true, usedAt: 9 });
  for (const s of [a, b, tools]) index = withSetup(index, s);
  assert.deepEqual(menuOrder(index).map((x) => x.id), ['b', 'a']);
  index = markUsed(index, 'a', 'recorder', 10);
  index = markUsed(index, 'b', 'broadcast', 11);
  assert.deepEqual(menuOrder(index).map((x) => x.id), ['b', 'a']);
  assert.equal(openingSetup(index, 'recorder'), 'a');
  assert.equal(openingSetup(index, 'broadcast'), 'b');
  index = markUsed(index, 'a', 'broadcast', 20);
  assert.equal(index.setups.find((x) => x.id === 'a').app, 'broadcast', 'a setup remembers the app it was last used in');
  assert.equal(openingSetup(index, 'broadcast'), 'a');
  const left = withoutSetup(index, 'a');
  assert.equal(openingSetup(left, 'recorder'), 'b', 'an app whose setup was deleted goes to the newest left');
  assert.equal(openingSetup(left, 'broadcast'), 'b');
  assert.equal(uniqueName('Show', ['show', 'Show 2']), 'Show 3');
  assert.equal(uniqueName('Show', []), 'Show');
  assert.ok(!['setup-1'].includes(setupId(['setup-1'], 1)));
  assert.deepEqual(normaliseIndex({ setups: [{ id: 'x' }, { id: 'x' }, { nope: 1 }] }).setups.map((x) => x.id), ['x']);
});

test('a copy is a setup of its own, and a stream removed drops out of a setup', () => {
  const a = { ...setupWith(), liveOn: ['yt', 'fb'] };
  const b = duplicateSetup(a, { id: 'b', name: 'B', now: 9 });
  assert.equal(b.id, 'b');
  assert.deepEqual(b.liveOn, ['yt', 'fb']);
  assert.deepEqual(dropStream(a, 'yt').liveOn, ['fb']);
});

test('undo steps back and redo forward; a new change drops what was undone', () => {
  let h = emptyHistory();
  assert.equal(historyUndo(h, 'now'), null);
  h = historyPush(h, 's1');
  h = historyPush(h, 's2');
  let r = historyUndo(h, 's3');
  assert.equal(r.stack, 's2');
  r = historyUndo(r.history, 's2');
  assert.equal(r.stack, 's1');
  r = historyRedo(r.history, 's1');
  assert.equal(r.stack, 's2');
  h = historyPush(r.history, 's2b');
  assert.equal(historyRedo(h, 'x'), null);
  h = emptyHistory();
  for (let i = 0; i < 300; i++) h = historyPush(h, i);
  assert.equal(h.back.length, 200);
  assert.equal(h.back[0], 100);
});

test('a window: the same title after its program restarts, else the same program\'s closest title, else not found', () => {
  assert.equal(programOf('Q3 figures.xlsx - Excel'), 'Excel');
  assert.equal(programOf('Report \u2014 Pages'), 'Pages');
  assert.equal(programOf('Notepad'), '');
  const base = (name, ref) => ({ ...setupWith(), stack: normaliseStack({ items: [{ id: 'w', kind: 'window', ref, name, on: true }] }) });
  // Its program restarted: a new capture id, the same title.
  let r = rematch(base(excel.name, 'window:999:0'), [excel, word, screen1], { cameras: [], mics: [] });
  assert.equal(r.setup.stack.items[0].ref, excel.id);
  assert.equal(r.found[0].how, 'title');
  assert.equal(r.missing.length, 0);
  // Same program, a sibling title.
  r = rematch(base('Q4 figures.xlsx - Excel', 'window:1:0'), [word, { id: 'window:3:0', name: 'Budget.xlsx - Excel', kind: 'window' }, excel], { cameras: [], mics: [] });
  assert.equal(r.setup.stack.items[0].ref, excel.id, 'the closest title of that program');
  assert.equal(r.found[0].how, 'program');
  assert.equal(r.setup.stack.items[0].name, excel.name);
  // None of that program: not found, and kept as it was.
  r = rematch(base('Pic.png - Paint', 'window:1:0'), [excel, word], { cameras: [], mics: [] });
  assert.equal(r.missing.length, 1);
  assert.equal(r.missing[0].name, 'Pic.png - Paint');
  assert.equal(r.setup.stack.items.length, 1, 'never dropped');
  assert.equal(r.setup.stack.items[0].ref, 'window:1:0');
  // The very same window, retitled.
  r = rematch(base('Old title - Word', word.id), [word], { cameras: [], mics: [] });
  assert.equal(r.found[0].how, 'same');
  assert.equal(r.setup.stack.items[0].name, word.name);
  // A window yet to be chosen is not lost.
  r = rematch(base('', ''), [word], { cameras: [], mics: [] });
  assert.equal(r.missing.length + r.found.length, 0);
});

test('a screen: same id and size, else same size, else the main screen saying which it was', () => {
  const base = (ref, size, name = 'Screen 2') => ({ ...setupWith(), stack: normaliseStack({ items: [{ id: 's', kind: 'screen', ref, name, size, on: true }] }) });
  let r = rematch(base(screen2.id, { w: 2560, h: 1440 }), [screen1, screen2], { cameras: [], mics: [] });
  assert.equal(r.found[0].how, 'same');
  r = rematch(base('screen:7:0', { w: 2560, h: 1440 }), [screen1, screen2], { cameras: [], mics: [] });
  assert.equal(r.found[0].how, 'size');
  assert.equal(r.setup.stack.items[0].ref, screen2.id);
  r = rematch(base('screen:7:0', { w: 3840, h: 2160 }), [screen1, screen2], { cameras: [], mics: [] });
  assert.equal(r.found[0].how, 'main');
  const it = r.setup.stack.items[0];
  assert.equal(it.ref, screen1.id);
  assert.equal(it.was, 'Screen 2 (3840 × 2160)');
  assert.deepEqual(it.size, { w: 3840, h: 2160 }, 'it goes on looking for the screen it was');
  // Found once the screen is back, and no longer says it was another.
  r = rematch(r.setup, [screen1, { id: 'screen:9:0', name: 'Big', kind: 'screen', display: { width: 3840, height: 2160 } }], { cameras: [], mics: [] });
  assert.equal(r.found[0].how, 'size');
  assert.equal(r.setup.stack.items[0].was, '');
});

test('a camera or microphone: the same id, else the same name, else not found; two alike stay two', () => {
  const base = (items, sounds = []) => ({ ...setupWith(), stack: normaliseStack({ items, sounds }) });
  const camItem = (id, ref, name) => ({ id, kind: 'camera', ref, name, on: true });
  let r = rematch(base([camItem('c', cam.id, 'old')]), [], { cameras: [cam], mics: [] });
  assert.equal(r.found[0].how, 'same');
  assert.equal(r.setup.stack.items[0].name, cam.name);
  r = rematch(base([camItem('c', 'elsewhere', cam.name)]), [], { cameras: [cam], mics: [] });
  assert.equal(r.found[0].how, 'name');
  assert.equal(r.setup.stack.items[0].ref, cam.id);
  r = rematch(base([camItem('c', 'gone', 'Studio camera')]), [], { cameras: [cam], mics: [] });
  assert.equal(r.missing[0].kind, 'camera');
  assert.equal(r.setup.stack.items[0].ref, 'gone');
  const twins = [{ id: 'x', name: 'Twin' }, { id: 'y', name: 'Twin' }];
  r = rematch(base([camItem('c1', 'p', 'Twin'), camItem('c2', 'q', 'Twin')]), [], { cameras: twins, mics: [] });
  assert.deepEqual(r.setup.stack.items.map((it) => it.ref).sort(), ['x', 'y']);
  // The default device: there is one, or there is not.
  r = rematch(base([camItem('c', '', '')]), [], { cameras: [cam], mics: [] });
  assert.equal(r.found[0].how, 'default');
  r = rematch(base([camItem('c', '', '')]), [], { cameras: [], mics: [] });
  assert.equal(r.missing.length, 1);
  // A microphone by its name.
  r = rematch(base([], [{ id: 'm', kind: 'mic', ref: 'old', name: mic.name, live: true }, { id: 'sys', kind: 'system', live: true }]), [], { cameras: [], mics: [mic] });
  assert.equal(r.found[0].how, 'name');
  assert.equal(r.setup.stack.sounds[0].ref, mic.id);
  r = rematch(base([], [{ id: 'm', kind: 'mic', ref: 'old', name: 'Gone', live: true }]), [], { cameras: [], mics: [mic] });
  assert.equal(r.missing[0].scope, 'sound');
});

test('a scene knows which of its sources are not found', () => {
  const st = setupWith().stack;
  const sc = st.scenes.find((x) => x.key === 'screen-camera');
  assert.deepEqual(sceneMissing(sc, [{ id: 'camera-1', scope: 'item', name: 'HD webcam' }]).map((m) => m.name), ['HD webcam']);
  assert.deepEqual(sceneMissing(st.scenes.find((x) => x.key === 'screen'), [{ id: 'camera-1', scope: 'item', name: 'HD webcam' }]), []);
  assert.deepEqual(sceneMissing(sc, []), []);
});

test('the first start: the Recorder\'s and Live Broadcast\'s kept stacks become setups with their scenes made ordinary, the old keys untouched', () => {
  const recorderStack = { items: [{ id: 'screen-1', kind: 'screen', ref: 'screen:0:0', name: 'Screen 1', on: true }, { id: 'camera-1', kind: 'camera', ref: '', on: true }], sounds: [{ id: 'mic-1', kind: 'mic', live: true }], scenes: [{ name: 'Mine', layers: [{ id: 'screen-1', box: { x: 0, y: 0, w: 1, h: 1 } }] }] };
  const recorderStudio = { size: 720, countdown: 5, shotTimer: 3, folder: 'D:/rec' };
  const broadcastStack = { items: [{ id: 'camera-1', kind: 'camera', ref: '', on: true }, { id: 'camera-2', kind: 'camera', ref: 'b', on: true }] };
  const before = JSON.stringify({ recorderStack, recorderStudio, broadcastStack });
  const r = migrate({ recorderStack, recorderStudio, broadcastStack, broadcastStudio: { fps: 60 }, now: 100 });
  assert.equal(JSON.stringify({ recorderStack, recorderStudio, broadcastStack }), before, 'nothing written into what it was given');
  assert.deepEqual(r.setups.map((s) => s.name), ['My recording setup', 'My live setup']);
  const [rec, live] = r.setups;
  assert.deepEqual(rec.stack.scenes.map((s) => s.key), ['screen-camera', 'screen', 'split', 'camera', 'scene-1']);
  assert.equal(rec.settings.size, 720);
  assert.equal(rec.settings.folder, 'D:/rec');
  assert.deepEqual(live.stack.scenes.map((s) => s.key), ['cam:camera-1', 'cam:camera-2', 'wide']);
  assert.equal(live.settings.fps, 60);
  assert.deepEqual(r.global, { countdown: 5, minimise: false, shot: 'program', shotTimer: 3 });
  assert.equal(openingSetup(r.index, 'recorder'), rec.id);
  assert.equal(openingSetup(r.index, 'broadcast'), live.id);
});

test('the first start: Live Broadcast that was the same as the Recorder shares its setup; nothing kept makes My setup', () => {
  const stack = { items: [{ id: 'camera-1', kind: 'camera', ref: '', on: true }] };
  const same = migrate({ recorderStack: stack, broadcastStack: stack });
  assert.equal(same.setups.length, 1);
  assert.equal(same.index.current.broadcast, same.setups[0].id);
  const none = migrate({ screen: screen1 });
  assert.deepEqual(none.setups.map((s) => s.name), ['My setup']);
  assert.equal(none.setups[0].stack.scenes.length, 4);
  assert.equal(none.index.current.recorder, none.setups[0].id);
  assert.equal(none.index.current.broadcast, none.setups[0].id);
  const onlySettings = migrate({ recorderStudio: { scene: 'camera', camera: true, title: 'Hi', titleOn: true } });
  assert.equal(onlySettings.setups[0].name, 'My recording setup');
  assert.ok(onlySettings.setups[0].stack.items.some((it) => it.kind === 'camera'));
});

test('materialised scenes carry the keys Auto named them by', () => {
  const st = materialiseScenes(normaliseStack({ items: [{ id: 'camera-1', kind: 'camera', on: true }, { id: 'camera-2', kind: 'camera' }], follow: { 'cam:camera-2': 'mic-1' }, rotate: { keys: ['wide'] } }));
  assert.ok(st.scenes.some((sc) => sc.key === 'cam:camera-2'));
  assert.equal(st.follow['cam:camera-2'], 'mic-1');
  assert.deepEqual(st.rotate.keys, ['wide']);
});

const withMedia = () => {
  let s = setupWith({ id: 'm', name: 'With pictures' });
  let st = s.stack;
  ({ stack: st } = addItem(st, { kind: 'picture', path: 'C:\\Users\\me\\logo.png' }, { aspect: 1 }));
  st = { ...st, placeholder: 'C:\\Users\\me\\back.jpg', music: 'C:\\Users\\me\\tune.mp3' };
  return { ...s, stack: st, settings: { ...s.settings, folder: 'D:\\Videos' }, liveOn: ['yt'] };
};
const FILES = { 'C:\\Users\\me\\logo.png': new Uint8Array([1, 2, 3]), 'C:\\Users\\me\\back.jpg': new Uint8Array([4, 5]), 'C:\\Users\\me\\tune.mp3': new Uint8Array([6]) };

test('the file holds setup.json and every picture and music file; a stream key is never written', async () => {
  const s = withMedia();
  assert.deepEqual(mediaPaths(s).length, 3);
  const { bytes, absent } = await writeSetupFile(s, {
    read: async (p) => FILES[p] || null,
    streams: [{ id: 'yt', name: 'Rutba', kind: 'youtube', url: 'rtmps://a.rtmp.youtube.com/live2', key: 'SECRET-STREAM-KEY', hasKey: true, keyTail: 'KEY1' }],
  });
  assert.deepEqual(absent, []);
  const parts = await unzip(bytes);
  assert.equal(new TextDecoder().decode(parts.mimetype), SETUP_MIME);
  assert.equal(Object.keys(parts)[0], 'mimetype', 'first, so the file says what it is from its first bytes');
  const text = new TextDecoder().decode(bytes);
  assert.ok(!text.includes('SECRET-STREAM-KEY'), 'no key in the file, anywhere');
  assert.ok(!text.includes('KEY1') && !/hasKey/.test(text));
  const json = JSON.parse(new TextDecoder().decode(parts['setup.json']));
  assert.ok(!JSON.stringify(json).includes('C:\\\\Users'), 'this computer\'s paths stay out of it');
  assert.equal(json.streams[0].name, 'Rutba');
  assert.deepEqual(Object.keys(parts).filter((k) => k.startsWith('media/')).length, 3);
});

test('a file comes back whole on another computer, its media placed where the caller keeps them', async () => {
  const s = withMedia();
  const { bytes } = await writeSetupFile(s, { read: async (p) => FILES[p] });
  const r = await readSetupFile(bytes);
  assert.equal(r.setup.name, 'With pictures');
  assert.equal(r.setup.stack.scenes.length, 4);
  assert.equal(r.media.length, 3);
  assert.deepEqual([...r.media.find((m) => m.name.endsWith('logo.png')).bytes], [1, 2, 3]);
  const placed = Object.fromEntries(r.media.map((m) => [m.entry, `E:/here/${m.name}`]));
  const done = placeMedia(r.setup, placed);
  const pic = done.stack.items.find((it) => it.kind === 'picture');
  assert.equal(pic.path, placed[r.media.find((m) => m.name.endsWith('logo.png')).entry]);
  assert.match(done.stack.placeholder, /^E:\/here\//);
  assert.match(done.stack.music, /^E:\/here\//);
  assert.deepEqual(mediaPaths(done).length, 3);
});

test('a media file that is gone is said, and the setup still saves without it', async () => {
  const s = withMedia();
  const { bytes, absent } = await writeSetupFile(s, { read: async (p) => (p.endsWith('tune.mp3') ? null : FILES[p]) });
  assert.deepEqual(absent, ['C:\\Users\\me\\tune.mp3']);
  const r = await readSetupFile(bytes);
  assert.equal(r.setup.stack.music, '');
  assert.equal(r.media.length, 2);
});

test('a file cannot reach outside itself: paths outside media/ come out empty, entry names are made plain', async () => {
  const { zipStored } = await import('../packages/studio/src/rvid.js');
  const enc = new TextEncoder();
  const setup = normaliseSetup({ name: 'Bad', stack: { items: [{ id: 'p', kind: 'picture', path: 'C:\\Windows\\win.ini', on: true }, { id: 'q', kind: 'picture', path: 'media/..\\..\\evil.png', on: true }], placeholder: '../../secret.jpg' } });
  const bytes = zipStored([
    { name: 'mimetype', bytes: enc.encode(SETUP_MIME) },
    { name: 'setup.json', bytes: enc.encode(JSON.stringify(setup)) },
    { name: 'media/..\\..\\evil.png', bytes: new Uint8Array([9]) },
    { name: 'media/ok.png', bytes: new Uint8Array([1]) },
  ]);
  const r = await readSetupFile(bytes);
  assert.deepEqual(r.setup.stack.items.map((it) => it.path), ['', '']);
  assert.equal(r.setup.stack.placeholder, '');
  assert.deepEqual(r.media.map((m) => m.entry), ['media/ok.png']);
  assert.ok(!/[\\/]/.test(r.media[0].name));
});

test('what is not a setup is refused, in words', async () => {
  await assert.rejects(() => readSetupFile(new Uint8Array([1, 2, 3, 4])), /not a Rutba Studio setup/);
  const { zipStored } = await import('../packages/studio/src/rvid.js');
  await assert.rejects(() => readSetupFile(zipStored([{ name: 'hello.txt', bytes: new Uint8Array([1]) }])), /not a Rutba Studio setup/);
  const enc = new TextEncoder();
  await assert.rejects(() => readSetupFile(zipStored([{ name: 'mimetype', bytes: enc.encode(SETUP_MIME) }, { name: 'setup.json', bytes: enc.encode('{nope') }])), /damaged/);
});

test('a stored setup, or anything, normalises; the tools\' own says so', () => {
  const s = normaliseSetup({ id: 'x', tools: true, app: 'nonsense', stack: { scenes: [{ layers: [] }] }, liveOn: ['a', 1] }, { app: 'broadcast' });
  assert.equal(s.app, 'broadcast');
  assert.equal(s.tools, true);
  assert.deepEqual(s.liveOn, ['a']);
  assert.equal(s.stack.scenes.length, 1);
  assert.ok(normaliseSetup(null).stack.items);
  assert.deepEqual(quickScenes(emptyStack()), []);
  assert.deepEqual(newSetupStack().scenes.map((sc) => sc.key), ['screen', 'window', 'screen-camera', 'camera']);
  assert.ok(updateItem(setupWith().stack, 'screen-1', { was: 'x' }));
});
