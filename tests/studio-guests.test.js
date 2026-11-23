// Guests from Talk as sources of the studio (packages/studio/src/guests.js,
// Office 2.27.0): three seats, the scenes ready made for one, two and three
// guests, who sits where as people come and go, and that nobody is put on
// the frame by arriving.

import test from 'node:test';
import assert from 'node:assert/strict';
import { emptyStack, normaliseStack, addItem, applyScene, removeItem, duplicateItem, PICTURE_KINDS } from '../packages/studio/src/stack.js';
import { GUEST, SEATS, GUEST_SCENES, isGuest, seatNumber, seatItem, tiles, guestScenes, withGuests, seatPeople, seatsShown, guestsShown, hideSeats } from '../packages/studio/src/guests.js';

const inside = (b) => b.x >= 0 && b.y >= 0 && b.x + b.w <= 1.0001 && b.y + b.h <= 1.0001;
const apart = (a, b) => a.x + a.w <= b.x + 0.0001 || b.x + b.w <= a.x + 0.0001 || a.y + a.h <= b.y + 0.0001 || b.y + b.h <= a.y + 0.0001;
const BEN = { id: 'u-ben', name: 'Ben Ortiz' };
const CARA = { id: 'u-cara', name: 'Cara Lund' };
const DAN = { id: 'u-dan', name: 'Dan Moyo' };
const EVE = { id: 'u-eve', name: 'Eve Park' };

function studio({ camera = true } = {}) {
  let s = emptyStack();
  ({ stack: s } = addItem(s, { kind: 'screen', ref: 'screen:1', name: 'Screen 1' }));
  if (camera) ({ stack: s } = addItem(s, { kind: 'camera', ref: 'cam-a', name: 'USB camera' }));
  return s;
}

test('a guest is a kind of picture the stack keeps, with no person written into it', () => {
  assert.ok(PICTURE_KINDS.includes(GUEST));
  const s = normaliseStack(withGuests(studio()));
  const seats = s.items.filter((it) => it.kind === GUEST);
  assert.deepEqual(seats.map((it) => it.ref).sort(), [...SEATS]);
  for (const it of seats) {
    assert.equal(it.on, false, 'a seat is not shown until someone is put into a scene');
    assert.equal(it.name, '');
    assert.equal(it.mirror, false);
    assert.ok(isGuest(it));
  }
  assert.deepEqual(SEATS.map((ref) => seatNumber(seatItem(s, ref))), [1, 2, 3]);
  assert.equal(seatNumber({ kind: 'camera', ref: 'seat-1' }), 0);
  assert.equal(isGuest({ kind: GUEST, ref: 'elsewhere' }), false);
  // Kept and read back, it is the same studio.
  assert.deepEqual(normaliseStack(JSON.parse(JSON.stringify(s))), s);
  assert.ok(!JSON.stringify(s).includes('u-ben'));
});

test('pictures sit side by side without overlapping, each inside the frame and as wide to high as the frame', () => {
  for (let n = 1; n <= 4; n++) {
    const boxes = tiles(n);
    assert.equal(boxes.length, n);
    for (const b of boxes) { assert.ok(inside(b), `${n}: inside`); assert.ok(Math.abs(b.w - b.h) < 0.0001, 'the frame\'s own shape'); }
    for (let i = 0; i < n; i++) for (let j = i + 1; j < n; j++) assert.ok(apart(boxes[i], boxes[j]), `${n}: ${i} and ${j} do not overlap`);
  }
  assert.deepEqual(tiles(1), [{ x: 0, y: 0, w: 1, h: 1 }]);
});

test('scenes for one, two and three guests are ready made, the camera beside them when there is one', () => {
  const s = withGuests(studio(), { 'guests-2': 'Zwei Gäste' });
  const made = s.scenes.filter((sc) => GUEST_SCENES.includes(sc.key));
  assert.deepEqual(made.map((sc) => sc.key), [...GUEST_SCENES]);
  assert.deepEqual(made.map((sc) => sc.name), ['One guest', 'Zwei Gäste', 'Three guests']);
  const camera = s.items.find((it) => it.kind === 'camera');
  made.forEach((sc, i) => {
    assert.equal(sc.layers.length, i + 2, 'the camera and the guests');
    assert.equal(sc.layers[0].id, camera.id);
    assert.deepEqual(sc.layers.slice(1).map((l) => s.items.find((it) => it.id === l.id).ref), SEATS.slice(0, i + 1));
  });
  // A scene put on the frame shows its guests and the camera, and nothing else.
  const two = applyScene(s, made[1]);
  assert.deepEqual(seatsShown(two).sort(), ['seat-1', 'seat-2']);
  assert.deepEqual(two.items.filter((it) => it.on).map((it) => it.kind).sort(), ['camera', 'guest', 'guest']);

  // With no camera the guests have the frame: one fills it.
  const bare = withGuests(studio({ camera: false }));
  const alone = bare.scenes.find((sc) => sc.key === 'guests-1');
  assert.equal(alone.layers.length, 1);
  assert.deepEqual(alone.layers[0].box, { x: 0, y: 0, w: 1, h: 1 });
  assert.equal(bare.scenes.find((sc) => sc.key === 'guests-3').layers.length, 3);
});

test('asked for again, the seats and scenes are left as the person has them', () => {
  const once = withGuests(studio());
  assert.equal(withGuests(once), once, 'nothing to add: the same stack');
  // A scene the person removed is made again only with a seat that was removed too.
  const renamed = { ...once, scenes: once.scenes.map((sc) => (sc.key === 'guests-1' ? { ...sc, name: 'Interview' } : sc)) };
  assert.equal(withGuests(renamed).scenes.find((sc) => sc.key === 'guests-1').name, 'Interview');
  const fewer = removeItem(once, seatItem(once, 'seat-3').id);
  const again = withGuests(fewer);
  assert.ok(seatItem(again, 'seat-3'), 'a seat taken out comes back when guests are asked in again');
  assert.equal(again.scenes.filter((sc) => GUEST_SCENES.includes(sc.key)).length, 3);
  assert.deepEqual(guestScenes(fewer).map((sc) => sc.key), ['guests-1', 'guests-2'], 'no scene for a seat there is not');
});

test('people take seats in turn, keep them while others come and go, and a fourth waits unseated', () => {
  let r = seatPeople({}, [BEN]);
  assert.deepEqual(r.seats, { 'seat-1': BEN });
  assert.deepEqual(r.came, [{ seat: 'seat-1', ...BEN }]);
  r = seatPeople(r.seats, [CARA, BEN, DAN]);
  assert.deepEqual(r.seats, { 'seat-1': BEN, 'seat-2': CARA, 'seat-3': DAN });
  assert.deepEqual(r.came.map((c) => c.id), ['u-cara', 'u-dan']);
  // Cara leaves and Eve comes: Eve has Cara's seat, the others stay where they are.
  r = seatPeople(r.seats, [BEN, DAN, EVE]);
  assert.deepEqual(r.seats, { 'seat-1': BEN, 'seat-2': EVE, 'seat-3': DAN });
  assert.deepEqual(r.went, [{ seat: 'seat-2', ...CARA }]);
  // Every seat taken: a fourth is left out, and nobody is moved.
  r = seatPeople(r.seats, [BEN, DAN, EVE, CARA]);
  assert.deepEqual(Object.values(r.seats).map((w) => w.id), ['u-ben', 'u-eve', 'u-dan']);
  assert.deepEqual(r.came, []);
  // A name that changed is the name shown.
  r = seatPeople(r.seats, [{ id: 'u-ben', name: 'Benjamin Ortiz' }]);
  assert.deepEqual(r.seats, { 'seat-1': { id: 'u-ben', name: 'Benjamin Ortiz' } });
  assert.equal(r.went.length, 2);
});

test('someone who arrives is not put on the frame: a seat that is on it is taken off before they sit', () => {
  const s = withGuests(studio());
  // The scene for three guests is up with nobody in it yet: the first to come takes the first seat.
  const all = applyScene(s, s.scenes.find((sc) => sc.key === 'guests-3'));
  const full = seatPeople({}, [BEN]);
  assert.deepEqual(full.seats, { 'seat-1': BEN });
  assert.deepEqual(full.came.map((c) => c.seat).filter((ref) => seatsShown(all).includes(ref)), ['seat-1'], 'the seat is on the frame');
  const hidden = hideSeats(all, full.came.map((c) => c.seat));
  assert.deepEqual(seatsShown(hidden).sort(), ['seat-2', 'seat-3']);
  assert.deepEqual(guestsShown(hidden, full.seats), []);
  assert.equal(hidden.scenes.find((sc) => sc.key === 'guests-3').layers.length, 4, 'the scene still has the seat');
  // Put back by the scene, the guest is in the picture.
  assert.deepEqual(guestsShown(applyScene(hidden, hidden.scenes.find((sc) => sc.key === 'guests-3')), full.seats), ['u-ben']);
});

test('who is in the picture is who sits in a seat that is shown, a copy of a seat counted once', () => {
  let s = withGuests(studio());
  s = applyScene(s, s.scenes.find((sc) => sc.key === 'guests-2'));
  const seats = { 'seat-1': BEN, 'seat-3': DAN };
  assert.deepEqual(guestsShown(s, seats), ['u-ben'], 'seat 2 is shown and empty, seat 3 is taken and not shown');
  const copy = duplicateItem(s, seatItem(s, 'seat-1').id).stack;
  assert.deepEqual(guestsShown(copy, seats), ['u-ben']);
  assert.deepEqual(seatsShown(hideSeats(copy, ['seat-1'])), ['seat-2'], 'every copy of the seat is taken off');
  assert.deepEqual(guestsShown(s, {}), []);
});
