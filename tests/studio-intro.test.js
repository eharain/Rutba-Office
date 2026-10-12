// Introductory videos (docs/STUDIO.md section 9, Phase 3): the template of
// scenes (X1), layouts switched after recording (X2), the teleprompter's
// arithmetic (X3), recording scene by scene with retakes and reordering
// (X4), the brand kit on every scene (X6), and the timeline's own edits
// (trims, words, zooms, the pointer's look) kept through every change.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  newIntro, isIntro, sceneSpans, sceneLength, sceneAt, layoutGeometry, composeIntro, relayoutIntro, addScene, removeScene,
  moveScene, setLayout, updateScene, recordScene, useTake, applyBrand, normaliseBrand, setPrompterSpeed, takeFiles,
  prompterProgress, scriptSeconds, wordCount, INTRO_TEMPLATE, LANES, BRAND_DEFAULTS, sceneById,
} from '../packages/studio/src/intro.js';
import { clipsOf, trimEnd, planPatches, laneClips } from '../packages/studio/src/project.js';
import { placeCaptions, isCaption } from '../packages/studio/src/captions.js';
import { presetDoc } from '../packages/studio/src/presets.js';
import { updateZoom } from '../packages/studio/src/pointer.js';

const WORDS = { title: 'Your title', subtitle: 'A line about it', endTitle: 'Thanks for watching', endSubtitle: 'example.org', placeholder: 'Record this scene' };
const near = (a, b, msg, eps = 0.002) => assert.ok(Math.abs(a - b) < eps, `${msg}: ${a} is not ${b}`);
const of = (doc, sceneId, role) => doc.patches.find((p) => p.scene === sceneId && p.role === role) || null;

const screenFile = (n = 1, extra = {}) => ({ role: 'screen', url: `C:/rec/Screen ${n}.webm`, name: `Screen ${n}`, duration: 6, width: 1920, height: 1080, hasAudio: false, ...extra });
const cameraFile = (n = 1, extra = {}) => ({ role: 'camera', url: `C:/rec/Camera ${n}.webm`, name: `Camera ${n}`, duration: 6, width: 1280, height: 720, hasAudio: true, mirror: true, ...extra });

test('a new introduction: the template\'s scenes one after another, cards in words, recorded scenes waiting', () => {
  const doc = newIntro({ words: WORDS, brand: { name: 'Ayesha Khan', role: 'Designer' } });
  assert.ok(isIntro(doc));
  assert.deepEqual(doc.intro.scenes.map((s) => s.layout), INTRO_TEMPLATE);
  const spans = sceneSpans(doc);
  assert.deepEqual(spans.map((s) => [s.start, s.end]), [[0, 4], [4, 10], [10, 18], [18, 23]]);
  const [title, talking] = doc.intro.scenes;
  assert.equal(of(doc, title.id, 'title').text, 'Ayesha Khan', 'the title card says who it is');
  assert.equal(of(doc, title.id, 'subtitle').text, 'Designer');
  assert.equal(of(doc, title.id, 'ground').color, BRAND_DEFAULTS.colors.background);
  assert.equal(of(doc, talking.id, 'placeholder').text, 'Record this scene');
  assert.equal(of(doc, doc.intro.scenes[3].id, 'title').text, 'Thanks for watching');
  assert.equal(sceneAt(doc, 11).layout, 'screenCamera');
  // Every scene's layers on their own lanes, picture lanes in role order.
  assert.equal(of(doc, title.id, 'ground').lane, LANES.ground);
  assert.equal(of(doc, title.id, 'title').lane, LANES.title);
  // The fades between scenes.
  assert.equal(of(doc, title.id, 'ground').exit.kind, 'fade');
});

test('recording scene by scene: a take lands in its scene, the scenes after move on, a retake keeps the first, an earlier take comes back', () => {
  let doc = newIntro({ words: WORDS, brand: { name: 'Ayesha Khan', role: 'Designer' } });
  const [, talking, screen] = doc.intro.scenes;
  let r = recordScene(doc, talking.id, [screenFile(1, { duration: 3 }), cameraFile(1, { duration: 3 })]);
  doc = r.doc;
  assert.equal(sceneLength(sceneById(doc, talking.id)), 3);
  assert.deepEqual(sceneSpans(doc).map((s) => s.start), [0, 4, 7, 15]);
  const cam = of(doc, talking.id, 'camera');
  assert.equal(cam.fw, undefined, 'a talking head fills the frame');
  assert.equal(cam.flipX, true, 'mirrored as recorded');
  assert.equal(of(doc, talking.id, 'screen').visible, false, 'its screen kept but not shown');
  assert.equal(of(doc, talking.id, 'voice').lane, 0, 'the voice on the first sound lane');
  assert.equal(of(doc, talking.id, 'camera').group, of(doc, talking.id, 'voice').group);
  // The lower third: the brand's name and role over the first scene with a camera.
  assert.equal(of(doc, talking.id, 'ltName').text, 'Ayesha Khan');
  assert.equal(of(doc, talking.id, 'ltRole').text, 'Designer');
  assert.equal(of(doc, talking.id, 'band').color, BRAND_DEFAULTS.colors.accent);
  assert.equal(of(doc, talking.id, 'placeholder'), null);
  // A retake: the new take on the timeline, the first kept.
  r = recordScene(doc, talking.id, [screenFile(2, { duration: 4 }), cameraFile(2, { duration: 4 })]);
  doc = r.doc;
  const s = sceneById(doc, talking.id);
  assert.equal(s.takes.length, 2);
  assert.equal(s.take, r.take.id);
  assert.equal(of(doc, talking.id, 'camera').url, 'C:/rec/Camera 2.webm');
  assert.equal(sceneLength(s), 4);
  doc = useTake(doc, talking.id, s.takes[0].id);
  assert.equal(of(doc, talking.id, 'camera').url, 'C:/rec/Camera 1.webm');
  assert.equal(takeFiles(doc).filter((f) => !f.used).length, 2, 'the other take\'s two files kept, off the timeline');
  // The screen scene, recorded with clicks, zooms on them.
  const clicks = { samples: [[0, 0.5, 0.5], [6, 0.5, 0.5]], clicks: [[1.5, 0.7, 0.3], [2.4, 0.71, 0.31]] };
  doc = recordScene(doc, screen.id, [screenFile(3, { pointer: clicks }), cameraFile(3)]).doc;
  const scr = of(doc, screen.id, 'screen');
  assert.equal(scr.fit, 'contain');
  assert.equal(scr.autoZoom.length, 1);
  assert.ok(scr.keys.zoom.some((k) => k.v > 1));
  assert.equal(of(doc, screen.id, 'pointer').screenUrl, scr.url);
  assert.equal(of(doc, screen.id, 'camera').insetShape, 'circle');
  assert.ok(of(doc, screen.id, 'camera').fx > 0.5, 'the camera in the bottom right corner');
  assert.equal(of(doc, screen.id, 'computer'), null, 'a screen without sound brings none');
  assert.equal(of(doc, screen.id, 'ltName'), null, 'the lower third once, on the first');
});

test('a layout switched after recording (X2): the same take, laid out the new way, and back', () => {
  let doc = newIntro({ words: WORDS });
  const screen = doc.intro.scenes[2];
  doc = recordScene(doc, screen.id, [screenFile(), cameraFile()]).doc;
  doc = setLayout(doc, screen.id, 'sideBySide');
  const cam = of(doc, screen.id, 'camera');
  const scr = of(doc, screen.id, 'screen');
  assert.equal(sceneById(doc, screen.id).layout, 'sideBySide');
  assert.ok(cam.fx < 0.1 && cam.fw > 0.3, 'the camera on the left');
  assert.equal(cam.insetShape, 'rounded');
  assert.ok(scr.fx > 0.4 && scr.fit === 'contain', 'the screen beside it, whole');
  assert.equal(scr.url, 'C:/rec/Screen 1.webm', 'the same take');
  doc = setLayout(doc, screen.id, 'talking');
  assert.equal(of(doc, screen.id, 'screen').visible, false);
  assert.equal(of(doc, screen.id, 'camera').fw, undefined);
  doc = setLayout(doc, screen.id, 'screenCamera');
  assert.ok(of(doc, screen.id, 'camera').fw < 0.3);
  // A camera-less take shows its screen whole whatever the layout asks.
  let solo = newIntro({ words: WORDS });
  solo = recordScene(solo, solo.intro.scenes[1].id, [screenFile()]).doc;
  assert.equal(of(solo, solo.intro.scenes[1].id, 'screen').fw, 1);
});

test('scenes reordered, added and removed, each change keeping what the timeline says', () => {
  let doc = newIntro({ words: WORDS });
  const [title, talking, screen, end] = doc.intro.scenes;
  doc = recordScene(doc, talking.id, [screenFile(1, { duration: 5 }), cameraFile(1, { duration: 5 })]).doc;
  // The timeline's edits: the take trimmed to 3 s, the title's words typed.
  doc = trimEnd(doc, of(doc, talking.id, 'camera').id, 4 + 3, { ripple: true });
  doc = { ...doc, patches: doc.patches.map((p) => (p.scene === title.id && p.role === 'title' ? { ...p, text: 'Hello from Lahore' } : p)) };
  doc = moveScene(doc, screen.id, -1);
  assert.deepEqual(doc.intro.scenes.map((s) => s.id), [title.id, screen.id, talking.id, end.id]);
  assert.equal(sceneLength(sceneById(doc, talking.id)), 3, 'the trim kept');
  assert.equal(of(doc, title.id, 'title').text, 'Hello from Lahore', 'the words kept');
  assert.deepEqual(sceneSpans(doc).map((s) => s.start), [0, 4, 12, 15]);
  assert.equal(of(doc, talking.id, 'camera').timing.start, 12);
  assert.equal(moveScene(doc, title.id, -1), doc, 'the first cannot go earlier');
  const added = addScene(doc, 'sideBySide', { after: talking.id });
  assert.equal(added.doc.intro.scenes[3].id, added.id);
  doc = removeScene(added.doc, end.id);
  assert.equal(doc.intro.scenes.length, 4);
  assert.equal(doc.patches.some((p) => p.scene === end.id), false);
  // A card's length and a script are the scene's.
  doc = updateScene(doc, title.id, { seconds: 2.5, script: 'Hi there, I am Ayesha.' });
  assert.equal(sceneSpans(doc)[1].start, 2.5);
  assert.equal(sceneById(doc, title.id).script, 'Hi there, I am Ayesha.');
  // Captions and anything else not a scene's are left alone.
  doc = placeCaptions(doc, [{ start: 0, end: 2, text: 'hello' }]).doc;
  doc = moveScene(doc, title.id, 1);
  assert.equal(clipsOf(doc).filter(isCaption).length, 1);
  assert.ok(clipsOf(doc).find(isCaption).lane > LANES.subtitle, 'captions above every scene\'s lane');
});

test('the brand kit on every scene (X6): colours, font, logo and the lower third, kept as a brand', () => {
  assert.deepEqual(normaliseBrand({ colors: { background: 'red', accent: '#ABCDEF' }, font: '  ', logo: { url: '' } }).colors, { background: BRAND_DEFAULTS.colors.background, text: '#ffffff', accent: '#abcdef' });
  let doc = newIntro({ words: WORDS });
  const [title, talking, screen] = doc.intro.scenes;
  doc = recordScene(doc, talking.id, [cameraFile()]).doc;
  doc = recordScene(doc, screen.id, [screenFile(), cameraFile(2)]).doc;
  doc = applyBrand(doc, { name: 'عائشہ خان', role: 'ڈیزائنر', colors: { background: '#0a2540', text: '#f5f5f5', accent: '#00d4aa' }, font: 'Georgia', logo: { url: 'C:/brand/logo.png', name: 'logo' } });
  for (const s of doc.intro.scenes) assert.equal(of(doc, s.id, 'ground').color, '#0a2540');
  assert.equal(of(doc, title.id, 'subtitle').color, '#00d4aa');
  assert.equal(of(doc, title.id, 'title').text, 'عائشہ خان', "the title card still in the template's words says the brand's name");
  assert.equal(of(doc, title.id, 'subtitle').text, 'ڈیزائنر');
  assert.equal(of(doc, title.id, 'title').fontFamily, 'Georgia');
  assert.equal(of(doc, title.id, 'logo').url, 'C:/brand/logo.png');
  assert.ok(of(doc, screen.id, 'logo').fw < 0.1, 'a small logo in the corner of a recorded scene');
  assert.equal(of(doc, screen.id, 'logo').fx > 0.5, true, 'away from a camera in the bottom right');
  // An Urdu name: the lower third on the right, right to left.
  const name = of(doc, talking.id, 'ltName');
  assert.equal(name.text, 'عائشہ خان');
  assert.equal(name.align, 'right');
  assert.ok(of(doc, talking.id, 'band').fx > 0.5);
  // Without the lower third or the corner logo.
  const quiet = applyBrand(doc, { ...doc.intro.brand, lowerThird: false, watermark: false });
  assert.equal(of(quiet, talking.id, 'ltName'), null);
  assert.equal(of(quiet, screen.id, 'logo'), null);
  assert.ok(of(quiet, title.id, 'logo'), 'the card keeps its logo');
});

test('laid out for another shape: an upright export stacks the screen and the camera, and the project itself is left as it was', () => {
  let doc = newIntro({ words: WORDS });
  const screen = doc.intro.scenes[2];
  doc = recordScene(doc, screen.id, [screenFile(), cameraFile()]).doc;
  doc = placeCaptions(doc, [{ start: 0, end: 2, text: 'hello' }]).doc;
  const tall = presetDoc(doc, 'vertical', { relayout: (d, aspect) => relayoutIntro(d, aspect) });
  assert.equal(tall.options.aspect, 'vertical');
  const cam = of(tall, screen.id, 'camera');
  const scr = of(tall, screen.id, 'screen');
  assert.ok(scr.fy < 0.2 && scr.fh <= 0.5, 'the screen at the top');
  assert.ok(cam.fy > 0.5, 'the camera under it');
  near(cam.fw * 1080, cam.fh * 1920, 'a circle that is round', 1);
  assert.ok(of(tall, doc.intro.scenes[0].id, 'title').sizeFrac > of(doc, doc.intro.scenes[0].id, 'title').sizeFrac, 'bigger words in a narrow frame');
  assert.equal(tall.patches.find(isCaption).fy, 0.76);
  assert.equal(of(doc, screen.id, 'screen').fh, 1, 'the project untouched');
  assert.deepEqual(layoutGeometry('screenCamera', 'square').screen, { fx: 0, fy: 0, fw: 1, fh: 1, fit: 'contain' });
  // The plan draws every scene layer it is handed.
  assert.ok(planPatches(tall).length >= tall.patches.length - 1);
  assert.ok(laneClips(tall, 'v', LANES.ground).length === tall.intro.scenes.length);
});

test('the timeline\'s zooms and the pointer\'s look are kept through a layout switch', () => {
  let doc = newIntro({ words: WORDS });
  const screen = doc.intro.scenes[2];
  doc = recordScene(doc, screen.id, [screenFile(1, { pointer: { samples: [[0, 0.5, 0.5]], clicks: [[1, 0.2, 0.2], [5, 0.9, 0.9]] } }), cameraFile()]).doc;
  const id = of(doc, screen.id, 'screen').id;
  assert.equal(of(doc, screen.id, 'screen').autoZoom.length, 2);
  doc = updateZoom(doc, id, 1, null);
  doc = { ...doc, patches: doc.patches.map((p) => (p.scene === screen.id && p.role === 'pointer' ? { ...p, rings: false, color: '#ff5c8a' } : p)) };
  doc = setLayout(doc, screen.id, 'sideBySide');
  assert.equal(of(doc, screen.id, 'screen').autoZoom.length, 1, 'the zoom taken away stays away');
  assert.equal(of(doc, screen.id, 'pointer').rings, false);
  assert.equal(of(doc, screen.id, 'pointer').color, '#ff5c8a');
  // Hidden by a talking head and shown again, the zoom is still there.
  doc = setLayout(setLayout(doc, screen.id, 'talking'), screen.id, 'screenCamera');
  assert.equal(of(doc, screen.id, 'screen').autoZoom.length, 1);
  // A retake without clicks has no zoom; the first take, put back, zooms on its own clicks again.
  doc = recordScene(doc, screen.id, [screenFile(2), cameraFile(2)]).doc;
  assert.equal(of(doc, screen.id, 'screen').autoZoom, undefined);
  doc = useTake(doc, screen.id, sceneById(doc, screen.id).takes[0].id);
  assert.equal(of(doc, screen.id, 'screen').autoZoom.length, 2);
});

test('the teleprompter: words counted, a script\'s reading time, and how far it has scrolled', () => {
  const script = 'Hello, I am Ayesha and this is how the report is made in three steps.';
  assert.equal(wordCount(script), 15);
  assert.equal(scriptSeconds(script, 150), 6);
  assert.equal(prompterProgress(script, 0, 150), 0);
  assert.equal(prompterProgress(script, 1, 150), 0, 'a second to begin');
  near(prompterProgress(script, 4, 150), 0.5, 'half way');
  assert.equal(prompterProgress(script, 60, 150), 1);
  assert.equal(prompterProgress('', 10), 0);
  const doc = setPrompterSpeed(newIntro({ words: WORDS }), 400);
  assert.equal(doc.intro.prompter.speed, 260, 'held to a speed one can read at');
  assert.deepEqual(composeIntro(composeIntro(doc)).patches, composeIntro(doc).patches, 'composing twice is composing once');
});
