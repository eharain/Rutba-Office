// Video starters and Replace Media (the 2.7 plan's Video starters): each
// starter laid out on its stand-in pictures with its titles, frame and
// transitions; a stand-in, or any picture or clip, replaced by another file
// of its kind in the same place and for the same time.
import test from 'node:test';
import assert from 'node:assert/strict';
import { STARTERS, buildStarter, markStandIns, isStandIn } from '../packages/studio/src/starters.js';
import { emptyProject, addMedia, clipById, laneClips, projectDuration, replaceMedia, planPatches } from '../packages/studio/src/project.js';

const pics = (n) => Array.from({ length: n }, (_, i) => ({ url: `D:/made/stand-in ${i + 1}.png`, name: `stand-in ${i + 1}.png`, kind: 'image', width: 1280, height: 720 }));

test('each starter lays its stand-ins one after another on the main track, with its titles and its frame', () => {
  for (const s of STARTERS) {
    const doc = buildStarter(s.key, pics(s.pictures), { title: 'Our summer', end: 'Thanks' });
    const stills = laneClips(doc, 'v', 0);
    assert.equal(stills.length, s.pictures, s.key);
    assert.equal(doc.options.aspect, s.aspect, s.key);
    const texts = doc.patches.filter((p) => p.type === 'text');
    assert.ok(texts.some((p) => p.text === 'Our summer') || doc.patches.some((p) => p.type === 'wordart'), `${s.key} has its title`);
    assert.ok(texts.some((p) => p.text === 'Thanks'), `${s.key} has its end card`);
    assert.ok(projectDuration(doc) > 5, s.key);
    assert.ok(planPatches(doc).length > s.pictures, s.key);
  }
});

test('the slideshow crossfades between its pictures and the social starter slides them in', () => {
  const slides = buildStarter('slideshow', pics(6));
  assert.equal(laneClips(slides, 'v', 0).filter((p) => p.transition?.kind === 'fade').length, 5);
  const social = buildStarter('social', pics(4), { tag: 'New' });
  assert.equal(laneClips(social, 'v', 0).filter((p) => p.transition?.kind === 'slide').length, 3);
  assert.ok(social.patches.some((p) => p.type === 'wordart' && p.text === 'New'));
  assert.equal(buildStarter('nothing', pics(2)), null);
});

test('a stand-in replaced by a picture keeps its place and is a stand-in no more', () => {
  let doc = markStandIns(buildStarter('slideshow', pics(6)), pics(6).map((p) => p.url));
  const first = laneClips(doc, 'v', 0)[1];
  assert.ok(isStandIn(doc, first.id));
  doc = replaceMedia(doc, first.id, { url: 'D:/photos/beach.jpg', name: 'beach.jpg', kind: 'image', width: 4000, height: 3000 });
  const after = clipById(doc, first.id);
  assert.equal(after.url, 'D:/photos/beach.jpg');
  assert.deepEqual(after.timing, first.timing);
  assert.deepEqual(after.transition, first.transition);
  assert.equal(isStandIn(doc, first.id), false);
  assert.equal(doc.sources['D:/photos/beach.jpg'].width, 4000);
});

test('a clip replaced by a shorter one is as long as the new file, its sound going with it or going', () => {
  const A = { url: 'D:/clips/a.webm', name: 'a.webm', kind: 'video', duration: 6, hasAudio: true, width: 320, height: 240 };
  const { doc, ids: [video, sound] } = addMedia(emptyProject(), A, { at: 0 });
  const withSound = replaceMedia(doc, video, { url: 'D:/clips/b.webm', name: 'b.webm', kind: 'video', duration: 4, hasAudio: true });
  assert.deepEqual(clipById(withSound, video).timing, { start: 0, end: 4 });
  assert.equal(clipById(withSound, sound).url, 'D:/clips/b.webm');
  assert.deepEqual(clipById(withSound, sound).timing, { start: 0, end: 4 });
  const silent = replaceMedia(doc, video, { url: 'D:/clips/c.webm', name: 'c.webm', kind: 'video', duration: 10, hasAudio: false });
  assert.equal(clipById(silent, sound), null);
  assert.deepEqual(clipById(silent, video).timing, { start: 0, end: 6 });
  // Only a file of the same kind.
  assert.equal(replaceMedia(doc, video, { url: 'D:/p.png', kind: 'image' }), doc);
});
