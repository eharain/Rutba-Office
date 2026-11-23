// Video's audio library (D-2026-11-22-10): a list of files added from the
// computer and of addresses that play from the web, and nothing more. An entry
// goes onto the timeline as sound.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  normaliseLibrary, addFileEntries, addRemoteEntry, removeEntry, rememberFile, searchEntries, playableKind, refuseAddress, addressName, PLAYABLE_ENDINGS,
  entrySound, SOUND_PROBLEMS,
} from '../packages/studio/src/audio-library.js';
import { emptyProject, addMedia, addSoundOnly, clipById } from '../packages/studio/src/project.js';

test('the list takes sound files and video files from the computer, once each', () => {
  const first = addFileEntries([], ['D:\\Music\\Theme.mp3', 'D:\\Clips\\talk.MP4', 'D:\\Notes\\read me.txt', 'D:\\Pictures\\a.png']);
  assert.deepEqual(first.added.map((e) => [e.kind, e.name]), [['file', 'Theme.mp3'], ['file', 'talk.MP4']]);
  assert.deepEqual(first.skipped, ['D:\\Notes\\read me.txt', 'D:\\Pictures\\a.png'], 'what is neither sound nor video is left out');
  const again = addFileEntries(first.list, ['d:/music/theme.mp3', 'D:\\Music\\Other.wav']);
  assert.deepEqual(again.added.map((e) => e.name), ['Other.wav'], 'a file already listed is not listed twice');
  assert.equal(again.list.length, 3);
  assert.equal(new Set(again.list.map((e) => e.id)).size, 3, 'every entry has an id of its own');
});

test('the list takes an https address, and refuses anything else in a way the window can say', () => {
  const ok = addRemoteEntry([], '  https://example.com/sounds/Rain%20at%20night.mp3?x=1  ');
  assert.equal(ok.entry.kind, 'remote');
  assert.equal(ok.entry.url, 'https://example.com/sounds/Rain%20at%20night.mp3?x=1');
  assert.equal(ok.entry.name, 'Rain at night.mp3');
  assert.equal(addRemoteEntry(ok.list, 'https://example.com/sounds/Rain%20at%20night.mp3?x=1').error, 'listed');
  assert.equal(addRemoteEntry([], 'http://example.com/a.mp3').error, 'not-https');
  assert.equal(addRemoteEntry([], 'rain.mp3').error, 'not-an-address');
  assert.equal(addRemoteEntry([], '').error, 'not-an-address');
  assert.equal(addRemoteEntry([], 'file:///C:/a.mp3').error, 'not-https');
  assert.equal(refuseAddress('http://127.0.0.1:8080/a.mp3'), 'not-https');
  assert.equal(refuseAddress('http://127.0.0.1:8080/a.mp3', { allowLoopback: true }), null, 'a check run reaches its own local server');
  assert.equal(addressName('https://example.com/'), 'example.com', 'an address with no file name is listed by its host');
});

test('an entry taken off the list is gone from the list alone', () => {
  let list = addFileEntries([], ['D:\\a.mp3', 'D:\\b.mp3']).list;
  list = addRemoteEntry(list, 'https://example.com/c.ogg').list;
  const without = removeEntry(list, list[0].id);
  assert.deepEqual(without.map((e) => e.name), ['b.mp3', 'c.ogg']);
  const more = addFileEntries(without, ['D:\\d.mp3']);
  assert.equal(new Set(more.list.map((e) => e.id)).size, 3, 'a new entry never takes an id still in use');
});

test('what was kept is read back, with anything that is not an entry left out', () => {
  const kept = [
    { id: 'a1', kind: 'file', path: 'D:\\a.mp3', name: 'a.mp3' },
    { id: 'a2', kind: 'remote', url: 'https://example.com/b.mp3', file: 'D:\\made\\b.mp3' },
    { id: 'a2', kind: 'file', path: 'D:\\twice.mp3' },
    { id: 'a3', kind: 'remote', url: 'javascript:alert(1)' },
    { id: 'a4', kind: 'bed', path: 'D:\\bed.mp3' },
    null, 'words', { kind: 'file', path: 'D:\\no-id.mp3' },
  ];
  assert.deepEqual(normaliseLibrary(kept), [
    { id: 'a1', kind: 'file', path: 'D:\\a.mp3', name: 'a.mp3' },
    { id: 'a2', kind: 'remote', url: 'https://example.com/b.mp3', name: 'b.mp3', file: 'D:\\made\\b.mp3' },
  ]);
  assert.deepEqual(normaliseLibrary(null), []);
  assert.deepEqual(normaliseLibrary({ a: 1 }), []);
});

test('a fetched address remembers its file, and the list is searched by name', () => {
  let list = addRemoteEntry([], 'https://example.com/rain.mp3').list;
  list = addFileEntries(list, ['D:\\Music\\Thunder.wav']).list;
  const kept = rememberFile(list, list[0].id, 'D:\\made\\rain.mp3');
  assert.equal(kept[0].file, 'D:\\made\\rain.mp3');
  assert.equal(rememberFile(list, list[1].id, 'D:\\x.wav')[1].file, undefined, 'a file from the computer is its own file');
  assert.deepEqual(searchEntries(kept, 'THUN').map((e) => e.name), ['Thunder.wav']);
  assert.equal(searchEntries(kept, '  ').length, 2);
});

test('a sound file plays as sound and a video file for its sound', () => {
  assert.equal(playableKind('a.MP3'), 'audio');
  assert.equal(playableKind('https://example.com/a.webm?download=1'), 'video');
  assert.equal(playableKind('a.png'), null);
  assert.equal(playableKind('mp3'), null);
  assert.ok(PLAYABLE_ENDINGS.includes('wav') && PLAYABLE_ENDINGS.includes('mp4'));
});

const A = { url: 'D:/clips/a.mp4', name: 'a.mp4', kind: 'video', duration: 10, hasAudio: true, width: 320, height: 240 };

test('an entry goes onto the timeline as sound alone, on a sound track free for it', () => {
  // A video on the timeline with its own sound on the first sound track.
  let doc = addMedia(emptyProject(), A, { at: 0 }).doc;
  const r = addSoundOnly(doc, { url: 'D:/music/theme.mp3', name: 'theme.mp3', duration: 4, size: 99 }, { at: 2 });
  assert.equal(r.ids.length, 1);
  const clip = clipById(r.doc, r.ids[0]);
  assert.equal(clip.type, 'sound');
  assert.deepEqual(clip.timing, { start: 2, end: 6 });
  assert.equal(clip.lane, 1, 'under the video\'s own sound, never over it');
  assert.equal(clip.group, undefined, 'a sound of its own');
  assert.equal(r.doc.sources['D:/music/theme.mp3'].kind, 'audio');
  assert.equal(r.doc.patches.filter((p) => p.type === 'video').length, 1, 'no picture is added');
  doc = r.doc;
  // A second one at the same time takes the next free track.
  const again = addSoundOnly(doc, { url: 'D:/music/theme.mp3', name: 'theme.mp3', duration: 4 }, { at: 3 });
  assert.equal(clipById(again.doc, again.ids[0]).lane, 2);
});

test('a video file from the list gives its sound and no picture, and the project keeps what it knew of the file', () => {
  const doc = addMedia(emptyProject(), A, { at: 0 }).doc;
  const r = addSoundOnly(doc, { url: A.url, name: A.name, duration: 10 }, { at: 12 });
  const clip = clipById(r.doc, r.ids[0]);
  assert.equal(clip.type, 'sound');
  assert.equal(clip.url, A.url);
  assert.deepEqual(clip.timing, { start: 12, end: 22 });
  assert.equal(r.doc.sources[A.url].kind, 'video', 'still a video to the media list');
  assert.equal(r.doc.patches.filter((p) => p.type === 'video').length, 1);
  // A video file the project has not met is recorded as the sound it is used as.
  const fresh = addSoundOnly(emptyProject(), { url: 'D:/clips/b.mp4', name: 'b.mp4', duration: 3 }, { at: 0 });
  assert.equal(fresh.doc.sources['D:/clips/b.mp4'].kind, 'audio');
  assert.equal(fresh.doc.patches.length, 1);
});

// Why an entry's sound could not be had is one of three plain reasons, a
// load that failed is not kept, and a fetch the person stopped is not an error.
const FILE = { id: 'a1', kind: 'file', path: 'D:\\Music\\Theme.mp3', name: 'Theme.mp3' };
const REMOTE = { id: 'a2', kind: 'remote', url: 'https://example.com/rain.mp3', name: 'rain.mp3' };
const SOUND = { duration: 3 };
const reasonOf = async (p) => { try { await p; } catch (err) { return err.reason; } return null; };

test('a file from the computer is read where it is, and one that is gone says so', async () => {
  const asked = [];
  const got = await entrySound(FILE, { exists: async (p) => { asked.push(p); return true; }, decode: async () => SOUND, fetchRemote: async () => { throw new Error('never'); } });
  assert.deepEqual(got, { file: FILE.path, buffer: SOUND });
  assert.deepEqual(asked, [FILE.path]);
  let decoded = false;
  assert.equal(await reasonOf(entrySound(FILE, { exists: async () => false, decode: async () => { decoded = true; return SOUND; } })), 'gone');
  assert.equal(decoded, false, 'a file that is not there is not read');
  assert.deepEqual(SOUND_PROBLEMS, ['gone', 'fetch', 'play']);
});

test('a file that cannot be played says so, by a refusal or by no sound, and the failed load is not kept', async () => {
  const forgot = [];
  const opts = { exists: async () => true, forget: (p) => forgot.push(p) };
  assert.equal(await reasonOf(entrySound(FILE, { ...opts, decode: async () => null })), 'play');
  assert.equal(await reasonOf(entrySound(FILE, { ...opts, decode: async () => { throw new Error('Unable to decode audio data'); } })), 'play');
  assert.deepEqual(forgot, [FILE.path, FILE.path], 'each failed load is forgotten, so the next asking reads the file again');
  // A load that worked is kept.
  await entrySound(FILE, { ...opts, decode: async () => SOUND });
  assert.equal(forgot.length, 2);
});

test('an address is fetched once into a file that is remembered, and a fetch that fails says so', async () => {
  const kept = [];
  let fetched = 0;
  const opts = { exists: async () => false, decode: async () => SOUND, remember: (p) => kept.push(p), fetchRemote: async (url) => { fetched += 1; assert.equal(url, REMOTE.url); return { path: 'D:\\Made\\rain.mp3' }; } };
  assert.deepEqual(await entrySound(REMOTE, opts), { file: 'D:\\Made\\rain.mp3', buffer: SOUND });
  assert.deepEqual(kept, ['D:\\Made\\rain.mp3']);
  // The file it was fetched into is still there: it is not fetched again.
  await entrySound({ ...REMOTE, file: 'D:\\Made\\rain.mp3' }, { ...opts, exists: async () => true });
  assert.equal(fetched, 1);
  // That file was deleted since: it is fetched again.
  await entrySound({ ...REMOTE, file: 'D:\\Made\\rain.mp3' }, opts);
  assert.equal(fetched, 2);
  assert.equal(await reasonOf(entrySound(REMOTE, { ...opts, fetchRemote: async () => { throw new Error('The address answered 404.'); } })), 'fetch');
  assert.equal(await reasonOf(entrySound(REMOTE, { ...opts, fetchRemote: async () => null })), 'fetch');
  assert.equal(await reasonOf(entrySound(REMOTE, { ...opts, decode: async () => null })), 'play', 'fetched, and what came is not a sound that plays');
});

test('a fetch the person stopped is neither a failure nor a sound: nothing is remembered, read or said', async () => {
  const kept = [];
  let decoded = false;
  const opts = { exists: async () => false, decode: async () => { decoded = true; return SOUND; }, remember: (p) => kept.push(p), stopped: () => true };
  assert.equal(await reasonOf(entrySound(REMOTE, { ...opts, fetchRemote: async () => { throw new Error('The fetch was stopped.'); } })), 'cancelled');
  // The stop came as the last of the file arrived: it is still a stop.
  assert.equal(await reasonOf(entrySound(REMOTE, { ...opts, fetchRemote: async () => ({ path: 'D:\\Made\\rain.mp3' }) })), 'cancelled');
  assert.deepEqual(kept, []);
  assert.equal(decoded, false);
  assert.ok(!SOUND_PROBLEMS.includes('cancelled'), 'it is the person\'s own doing, and has no message');
});
