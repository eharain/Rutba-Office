// Transitions → Sound — the engine's half, and the suite's own sounds.
//
// A sound goes in as PowerPoint writes one: the WAV in ppt/media, related
// from the slide as audio, and a p:sndAc after the effect in every
// p:transition the slide has (a PowerPoint 2010 effect's and its
// fallback's). Stop Previous Sound is an endSnd; Loop Until Next Sound
// marks the one there; No Sound takes it off and leaves the effect. A
// transition another edit rewrites keeps its sound. The suite's sounds are
// real WAVs, the same bytes every time.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { SOUNDS, soundWav, soundFile, wavOf } from '../apps/desktop/renderer/apps/slides/sounds.js';

const deck = () => Deck.open(buildPptx({ title: 'Sounds', slides: [{ layout: 'title', title: 'One' }, { layout: 'obj', title: 'Two' }] }));

test('the suite\'s sounds are WAVs, named as PowerPoint names its own, the same every time', () => {
  assert.deepEqual(SOUNDS, ['Breeze', 'Camera', 'Chime', 'Click', 'Coin', 'Drum Roll', 'Laser', 'Whoosh']);
  for (const name of SOUNDS) {
    const wav = soundWav(name);
    assert.equal(String.fromCharCode(...wav.slice(0, 4)), 'RIFF', name);
    assert.equal(String.fromCharCode(...wav.slice(8, 12)), 'WAVE', name);
    assert.ok(wav.length > 1000, name);
  }
  assert.deepEqual([...soundWav('Chime')], [...soundWav('Chime')]);
  assert.equal(soundFile('Drum Roll'), 'drumroll.wav');
  assert.equal(soundWav('Nope'), null);
  assert.equal(wavOf([0, 1, -1]).length, 44 + 6);
});

test('a sound goes in after the effect, related as audio, and comes back', () => {
  const d = deck();
  d.setTransition(1, { type: 'fade' });
  assert.equal(d.setTransitionSound(1, { data: Buffer.from(soundWav('Chime')), name: 'chime.wav' }), true);
  const xml = d.pkg.text('ppt/slides/slide2.xml');
  assert.match(xml, /<p:fade\/><p:sndAc><p:stSnd><p:snd r:embed="(rId\d+)" name="chime\.wav"\/><\/p:stSnd><\/p:sndAc><\/p:transition>/);
  const rel = [...d.pkg.rels('ppt/slides/slide2.xml')].find((r) => /\/audio$/.test(r.Type));
  assert.equal(rel.Target, '../media/audio1.wav');
  assert.match(d.pkg.text('[Content_Types].xml'), /<Default Extension="wav" ContentType="audio\/wav"\/>/);
  assert.deepEqual(d.transitionSound(1), { name: 'chime.wav', part: 'ppt/media/audio1.wav', loop: false });
  // Another edit to the transition keeps it.
  d.setTransition(1, { duration: 1.5 });
  assert.equal(d.transitionSound(1)?.name, 'chime.wav');
  // Loop Until Next Sound, then Stop Previous Sound, then No Sound.
  d.setTransitionSound(1, { loop: true });
  assert.equal(d.transitionSound(1).loop, true);
  d.setTransitionSound(1, { stop: true });
  assert.deepEqual(d.transitionSound(1), { stop: true });
  d.setTransitionSound(1, null);
  assert.equal(d.transitionSound(1), null);
  assert.equal(d.slide(1).transition.type, 'fade', 'the effect stays');
  assert.throws(() => d.setTransitionSound(1, { data: Buffer.from('ID3 not a wav'), name: 'x.mp3' }), /WAV/);
});

test('a slide with no transition gets one that only plays the sound, and an effect written twice (a stated duration) gets it in both', () => {
  const d = deck();
  d.setTransitionSound(0, { data: Buffer.from(soundWav('Click')), name: 'click.wav' });
  assert.match(d.pkg.text('ppt/slides/slide1.xml'), /<p:transition[^>]*><p:sndAc><p:stSnd><p:snd r:embed="rId\d+" name="click\.wav"\/>/);
  d.setTransition(1, { type: 'push', duration: 1.25 });
  assert.match(d.pkg.text('ppt/slides/slide2.xml'), /<mc:AlternateContent/, 'written twice, for 2010 and for 2007');
  d.setTransitionSound(1, { data: Buffer.from(soundWav('Whoosh')), name: 'whoosh.wav' });
  const xml = d.pkg.text('ppt/slides/slide2.xml');
  assert.equal((xml.match(/<p:sndAc>/g) || []).length, (xml.match(/<p:transition\b/g) || []).length, 'one in each transition element');
  const back = Deck.open(d.save());
  assert.equal(back.transitionSound(1)?.name, 'whoosh.wav');
});
