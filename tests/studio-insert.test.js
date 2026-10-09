/**
 * The insert palette's catalogue, and the shape vocabulary under it.
 *
 * The palette is DERIVED from the renderer — `SHAPE_GROUPS` decides the
 * headings, `GEOMETRIES` decides what exists — and that derivation is the whole
 * point: a geometry added to shapes.js should appear in the panel with no
 * change to the app. These are the assertions that keep the derivation honest,
 * because the failure mode is silent. A geometry missing from a group is not an
 * error anywhere; it is simply a shape nobody can ever find.
 *
 *   node --test tests/insert.test.mjs
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';

import { GEOMETRIES, SHAPE_GROUPS, GEOMETRY_LABELS, ICONS, ICON_NAMES } from '../packages/studio/src/toolset.js';
import { readFileSync, existsSync } from 'node:fs';
import { insertCatalogFor, searchCatalog, timelineInsertsFor } from '../packages/studio/src/insert-catalog.js';
import { emptyDoc, addLayer } from '../packages/studio/src/doc.js';

// ── the vocabulary ──────────────────────────────────────────────────────────

test('every geometry belongs to exactly one palette group', () => {
    const grouped = SHAPE_GROUPS.flatMap((g) => g.geometries);
    const seen = new Set();
    for (const g of grouped) {
        assert.ok(!seen.has(g), `${g} appears in two groups`);
        seen.add(g);
    }
    for (const g of GEOMETRIES) {
        assert.ok(seen.has(g), `${g} exists but is in no group — it would be unreachable`);
    }
    assert.equal(grouped.length, GEOMETRIES.length, 'a group names a geometry that does not exist');
});

test('every geometry has a human name', () => {
    for (const g of GEOMETRIES) {
        assert.ok(GEOMETRY_LABELS[g], `${g} has no label, so its tile would read "${g}"`);
    }
});

test('the geometry list still starts with rect', () => {
    // `compileShape` falls back to 'rect' for an unknown geometry, so a recipe
    // written by a newer editor degrades to a box rather than to nothing. That
    // only reads as a deliberate fallback while rect is in the list at all.
    assert.ok(GEOMETRIES.includes('rect'));
});

// ── the catalogue ───────────────────────────────────────────────────────────

const allItems = (kind) => insertCatalogFor(kind)
    .flatMap((s) => s.groups.flatMap((g) => g.items.map((it) => ({ ...it, section: s.key }))));

/**
 * Items that insert a LAYER. Picker items ('Picture', 'Video clip', 'Sound')
 * have no patch — they carry a document transform and ask what to insert first
 * — so the assertions about patches are about these.
 */
const itemsOf = (kind) => allItems(kind).filter((it) => it.patch);

test('every shape in the catalogue names a real geometry', () => {
    for (const it of itemsOf('image')) {
        if (it.patch.type !== 'shape') continue;
        assert.ok(GEOMETRIES.includes(it.patch.geometry), `${it.patch.geometry} is not a geometry`);
        assert.equal(it.tile.geometry, it.patch.geometry, 'the tile must draw what the click inserts');
    }
});

test('the palette offers every geometry, not a curated subset of them', () => {
    // The regression this guards: a hand-listed palette that quietly falls a
    // few shapes behind the renderer every time one is added.
    const offered = new Set(itemsOf('image').filter((i) => i.patch.type === 'shape').map((i) => i.patch.geometry));
    for (const g of GEOMETRIES) assert.ok(offered.has(g), `${g} is not offered anywhere`);
});

test('every catalogue item can actually be added to a document', () => {
    let doc = emptyDoc('video');
    for (const it of itemsOf('video')) doc = addLayer(doc, it.patch);
    assert.equal(doc.patches.length, itemsOf('video').length);
    for (const p of doc.patches) {
        assert.ok(p.id, 'addLayer minted an id');
        assert.ok(p.type, 'and the patch carries a type');
    }
});

test('every item has a tile the panel knows how to draw', () => {
    // Enumerated rather than sniffed for a field, because InsertPanel branches
    // on exactly these three and anything else falls through to a blank square.
    for (const it of itemsOf('video')) {
        assert.ok(it.tile, `${it.label} has no tile`);
        if (it.tile.type === 'shape') assert.ok(it.tile.geometry, `${it.label}: shape tile with no geometry`);
        else if (it.tile.type === 'path') assert.ok(it.tile.d, `${it.label}: path tile with no path`);
        else assert.ok(it.tile.icon, `${it.label}: tile is neither a shape, a path, nor a named icon`);
    }
});

test('every icon in the set is offered, and its tile carries the real path', () => {
    const offered = itemsOf('image').filter((i) => i.patch.type === 'icon');
    for (const name of ICON_NAMES) {
        const hit = offered.find((i) => i.patch.name === name);
        assert.ok(hit, `${name} exists but is not in the palette`);
        assert.equal(hit.tile.d, ICONS[name], 'the tile must draw the glyph the click inserts');
    }
});

test('ids are unique across the whole catalogue', () => {
    // They are React keys, and a duplicate key silently drops a tile.
    const ids = itemsOf('video').map((i) => i.id);
    assert.equal(new Set(ids).size, ids.length);
});

test('sound is offered on video and withheld from stills', () => {
    assert.ok(itemsOf('video').some((i) => i.patch.type === 'sound'), 'a video can carry sound');
    assert.ok(!itemsOf('image').some((i) => i.patch.type === 'sound'), 'a still has no timeline for a clip');
});

test('the voice-over ducks the music and the sound effect does not', () => {
    // The default IS the feature: if these two ever carry the same mix, one of
    // the entries has stopped earning its place.
    const sounds = itemsOf('video').filter((i) => i.patch?.type === 'sound');
    assert.equal(sounds.find((i) => i.label === 'Voice-over').patch.mix, 'duck');
    assert.equal(sounds.find((i) => i.label === 'Sound effect').patch.mix, 'mix');
});

// ── the timeline's add row ────────────────────────────────────────────────────

test('the timeline opens the Insert tab on media and sound, not a menu of its own', () => {
    const doors = timelineInsertsFor('video');
    assert.deepEqual(doors.map((d) => d.section), ['media', 'sound']);
    const sections = new Set(insertCatalogFor('video').map((s) => s.key));
    for (const d of doors) {
        assert.ok(sections.has(d.section), d.section + ' is a section the tab shows');
        assert.ok(d.label && d.icon, 'and the button has a name and a glyph');
    }
    // A still has no timeline and no sound section; a door to nowhere is never offered.
    assert.deepEqual(timelineInsertsFor('image').map((d) => d.section), ['media']);
});

test('the add row carries no menu of its own', () => {
    const timeline = readFileSync(new URL('../packages/studio/port/editor/Timeline.js', import.meta.url), 'utf8');
    assert.ok(!timeline.includes('add-menu'), 'the timeline no longer imports a second list');
    assert.ok(timeline.includes('askInsert(door.section)'), 'each button opens the Insert tab');
    assert.ok(!existsSync(new URL('../packages/studio/src/add-menu.js', import.meta.url)), 'and the list itself is retired');
});

test('the adjustment section does not offer the no-op preset', () => {
    const adjusts = itemsOf('image').filter((i) => i.patch.type === 'adjust');
    assert.ok(adjusts.length > 0);
    assert.ok(!adjusts.some((i) => i.patch.preset === 'none'), "'none' is what a new adjustment already is");
});

// ── insert defaults ─────────────────────────────────────────────────────────

test('a pie is inserted as a slice rather than as a whole circle', () => {
    // The reason defaultsFor exists: several geometries at their RENDERER
    // defaults are indistinguishable from a simpler shape, and the first click
    // should produce the thing the tile promised.
    const pie = itemsOf('image').find((i) => i.patch.geometry === 'pie');
    assert.ok(pie.patch.sweep < 1, 'a full turn is an ellipse');
});

test('a ring is inserted with a hole in it', () => {
    const ring = itemsOf('image').find((i) => i.patch.geometry === 'ring');
    assert.ok(ring.patch.innerRatio > 0.05 && ring.patch.innerRatio < 0.95);
});

test('a ring takes its height from the compiler, so it is round on any aspect', () => {
    // compileShape derives a circular height from the width when `fh` is
    // absent. Pinning an fh here would make the ring an oval on 9:16.
    const ring = itemsOf('image').find((i) => i.patch.geometry === 'ring');
    assert.equal(ring.patch.fh, undefined);
});

test('a rule is inserted thin', () => {
    const line = itemsOf('image').find((i) => i.patch.geometry === 'line');
    assert.ok(line.patch.fh < 0.1, 'a rule with a square box is a rectangle');
});

// ── search ──────────────────────────────────────────────────────────────────

test('an empty query means "no search", not "no results"', () => {
    // The panel shows its sections when search returns null; returning [] here
    // would blank the palette the moment someone focused the box.
    assert.equal(searchCatalog('image', ''), null);
    assert.equal(searchCatalog('image', '   '), null);
});

test('search crosses sections and reports which one it found', () => {
    const hits = searchCatalog('image', 'star');
    assert.ok(hits.length > 0);
    assert.ok(hits.every((h) => h.section), 'a result says where it lives, so the answer teaches too');
});

test('search matches the group name, not only the item name', () => {
    // Someone looking for "a round one" should find the round ones.
    const hits = searchCatalog('image', 'rounds');
    assert.ok(hits.some((h) => h.patch.geometry === 'ellipse'));
});

test('search is case-insensitive and finds nothing gracefully', () => {
    assert.ok(searchCatalog('image', 'SHIELD').length > 0);
    assert.deepEqual(searchCatalog('image', 'zzzznotathing'), []);
});

// ── picker items ────────────────────────────────────────────────────────────

/**
 * The second shape of catalogue item.
 *
 * Three insert entries do genuinely different things to the document — a
 * picture joins `doc.images` and is not a layer at all, a clip becomes TWO
 * layers, a sound becomes one — so they carry a transform rather than a patch.
 * These pin that the transform is real and that the two routes to a clip (the
 * library and the recorder) produce the same thing.
 */
const pickersOf = (kind) => allItems(kind).filter((it) => it.pick);

test('every item is either a patch or a picker, never both and never neither', () => {
    for (const it of allItems('video')) {
        const isPatch = Boolean(it.patch);
        const isPick = Boolean(it.pick);
        assert.ok(isPatch !== isPick, `${it.label} is ${isPatch && isPick ? 'both' : 'neither'}`);
        if (isPick) assert.equal(typeof it.apply, 'function', `${it.label} has no transform`);
    }
});

test('a still offers pictures but not clips or sound', () => {
    // The same rule the renderer keeps: a still has no timeline for a clip.
    assert.deepEqual(pickersOf('image').map((p) => p.pick), ['image']);
    assert.deepEqual(pickersOf('video').map((p) => p.pick).sort(), ['audio', 'image', 'video']);
});

test('a picture joins the picture list rather than becoming a layer', () => {
    const item = pickersOf('image').find((p) => p.pick === 'image');
    const before = emptyDoc('image');
    const after = item.apply(before, ['https://m.test/a.jpg']);
    assert.deepEqual(after.images, ['https://m.test/a.jpg']);
    assert.equal(after.patches.length, 0, 'a picture is not a layer');
});

test('a clip from the library becomes the same two layers a recorded one does', () => {
    // Picture and sound on ONE url, so the two routes into the timeline cannot
    // diverge — trim what you see and what you hear independently either way.
    const item = pickersOf('video').find((p) => p.pick === 'video');
    const after = item.apply(emptyDoc('video'), ['https://m.test/clip.webm']);
    const types = after.patches.map((p) => p.type).sort();
    assert.deepEqual(types, ['sound', 'video']);
    assert.equal(new Set(after.patches.map((p) => p.url)).size, 1, 'both name the same file');
});

test('a clip arrives with no timing, which the renderer reads as the whole video', () => {
    // A recorded take knows its own length because the recorder measured it; a
    // library file does not without decoding. The full span is the honest
    // default, and the lane is dragged from there.
    const item = pickersOf('video').find((p) => p.pick === 'video');
    const after = item.apply(emptyDoc('video'), ['https://m.test/clip.webm']);
    for (const p of after.patches) assert.ok(!p.timing, `${p.type} pinned a timing it cannot know`);
});

test('inserted media is named from its file, not left as "Clip"', () => {
    const item = pickersOf('video').find((p) => p.pick === 'video');
    const after = item.apply(emptyDoc('video'), ['https://m.test/holiday%20cut.webm']);
    assert.ok(after.patches[0].name.includes('holiday cut'), `got ${after.patches[0].name}`);
});

test('several files at once each become their own layers', () => {
    const item = pickersOf('video').find((p) => p.pick === 'audio');
    const after = item.apply(emptyDoc('video'), ['https://m.test/a.wav', 'https://m.test/b.wav']);
    assert.equal(after.patches.length, 2);
    assert.equal(new Set(after.patches.map((p) => p.id)).size, 2, 'ids are distinct');
});
