// Guests from Talk in the studio (Office 2.27.0, the Talk, Meet and Live
// plan's phase 3): up to three people, each a source like any other, so a
// scene can place their picture and the mix carries their sound.
//
// A guest is an item of kind `guest` whose `ref` is a seat: seat-1, seat-2
// or seat-3. Who sits in a seat is known only while the broadcast or the
// recording is open and is never written into the stack, so a setup that is
// kept holds seats and scenes, never people. A seat nobody sits in draws
// nothing.
//
// Three scenes are ready made, for one, two and three guests, each with the
// first camera beside them when there is one. They are ordinary scenes from
// then on: renamed, changed or removed as any other.
//
// Someone who comes in takes the first empty seat, so the scene for one
// guest shows the first to come. A seat that is on the frame is taken off it
// before anyone sits in it, so nobody is put on air by arriving: they wait
// until a scene or the list shows them.
//
// Nothing here touches a canvas, a stream or Talk, so it is tested in Node.

import { addItem, clampBox, showItem } from './stack.js';

export const GUEST = 'guest';
/** The seats, in the order they are taken. */
export const SEATS = Object.freeze(['seat-1', 'seat-2', 'seat-3']);
/** The keys of the ready-made scenes, by how many guests each shows. */
export const GUEST_SCENES = Object.freeze(['guests-1', 'guests-2', 'guests-3']);

export const isGuest = (item) => item?.kind === GUEST && SEATS.includes(item.ref);
/** A seat's number, 1 to 3; 0 for anything that is not a seat. */
export const seatNumber = (item) => (item?.kind === GUEST ? SEATS.indexOf(item.ref) + 1 : 0);
/** The first item of a seat, or null. */
export const seatItem = (stack, ref) => stack.items.find((it) => it.kind === GUEST && it.ref === ref) || null;

const GAP = 0.02;
const TILE = 0.47;
const NEAR = GAP;
const FAR = GAP * 2 + TILE;
const MID = (1 - TILE) / 2;

/**
 * Where `n` pictures sit on the frame, each as wide to high as the frame
 * itself, so nobody is cut: one fills it, two side by side, three as two
 * over one, four in a square.
 */
export function tiles(n) {
    const at = (x, y) => clampBox({ x, y, w: TILE, h: TILE });
    if (n <= 1) return [clampBox({ x: 0, y: 0, w: 1, h: 1 })];
    if (n === 2) return [at(NEAR, MID), at(FAR, MID)];
    if (n === 3) return [at(NEAR, NEAR), at(FAR, NEAR), at(MID, FAR)];
    return [at(NEAR, NEAR), at(FAR, NEAR), at(NEAR, FAR), at(FAR, FAR)];
}

/**
 * The ready-made scenes for the seats the stack has: one guest, two and
 * three, the first camera first in each when there is a camera. A scene
 * that needs a seat the stack has not is left out.
 */
export function guestScenes(stack, names = {}) {
    const camera = stack.items.filter((it) => it.kind === 'camera').sort((x, y) => x.id.localeCompare(y.id, undefined, { numeric: true }))[0] || null;
    const seats = SEATS.map((ref) => seatItem(stack, ref));
    const fallback = ['One guest', 'Two guests', 'Three guests'];
    const out = [];
    for (let n = 1; n <= SEATS.length; n++) {
        const guests = seats.slice(0, n);
        if (guests.some((it) => !it)) break;
        const shown = camera ? [camera, ...guests] : guests;
        const boxes = tiles(shown.length);
        const shape = shown.length === 1 ? 'rect' : 'rounded';
        out.push({ key: GUEST_SCENES[n - 1], name: names[GUEST_SCENES[n - 1]] || fallback[n - 1], layers: shown.map((it, i) => ({ id: it.id, box: boxes[i], crop: null, shape })) });
    }
    return out;
}

/**
 * The stack with its three seats, none shown, and the ready-made scenes
 * after the scenes it has. What is already there is left as it is, so this
 * can be asked for again and again.
 */
export function withGuests(stack, names = {}) {
    let s = stack;
    for (const ref of SEATS) {
        if (seatItem(s, ref)) continue;
        ({ stack: s } = addItem(s, { kind: GUEST, ref, name: '', on: false, shape: 'rounded', mirror: false, box: tiles(1)[0] }));
    }
    const have = new Set(s.scenes.map((sc) => sc.key));
    const made = guestScenes(s, names).filter((sc) => !have.has(sc.key));
    return made.length ? { ...s, scenes: [...s.scenes, ...made] } : s;
}

/**
 * Who sits where, after the people in the meeting changed. `seats` is
 * { [seat]: { id, name } }, `people` the guests there now ([{ id, name }],
 * this computer's own seat left out). Someone already seated keeps their
 * seat; someone new takes the first empty one; with every seat taken they
 * are left unseated. Answers { seats, came: [{ seat, id, name }], went }.
 * A seat in `came` that is on a frame is taken off it (hideSeats) before
 * the new seats are used.
 */
export function seatPeople(seats = {}, people = []) {
    const here = new Map(people.filter((p) => p?.id).map((p) => [p.id, p]));
    const next = {};
    const went = [];
    for (const ref of SEATS) {
        const who = seats[ref];
        if (!who) continue;
        if (here.has(who.id)) next[ref] = { id: who.id, name: here.get(who.id).name || who.name || '' };
        else went.push({ seat: ref, id: who.id, name: who.name || '' });
    }
    const seated = new Set(Object.values(next).map((w) => w.id));
    const came = [];
    for (const p of here.values()) {
        if (seated.has(p.id)) continue;
        const ref = SEATS.find((r) => !next[r]);
        if (!ref) break;
        next[ref] = { id: p.id, name: p.name || '' };
        seated.add(p.id);
        came.push({ seat: ref, id: p.id, name: p.name || '' });
    }
    return { seats: next, came, went };
}

/** The seats on this stack's frame. */
export const seatsShown = (stack) => [...new Set(stack.items.filter((it) => it.kind === GUEST && it.on && SEATS.includes(it.ref)).map((it) => it.ref))];

/** The people on this stack's frame, by their ids: who is told they are in the picture. */
export const guestsShown = (stack, seats = {}) => seatsShown(stack).map((ref) => seats[ref]?.id).filter(Boolean);

/** These seats taken off the frame, every copy of each, their places in the scenes kept. */
export function hideSeats(stack, refs = []) {
    let s = stack;
    for (const it of stack.items) if (it.kind === GUEST && it.on && refs.includes(it.ref)) s = showItem(s, it.id, false);
    return s;
}
