/**
 * The creative document — what a Studio project's `doc` column holds.
 *
 * ONE FORMAT FOR BOTH EDITORS. An image is a one-frame video (see
 * editor-core/image.js), so an image project and a video project are the same
 * document with the same layer semantics. That is not tidiness for its own
 * sake: it means one template can serve both, a caption that reads well in a
 * reel reads the same in the still, and "this creative at four sizes" is a loop
 * rather than a feature.
 *
 * NOTHING HERE IS PIXELS. Layer geometry is fractional throughout — `fx`/`fy`
 * of the frame, `sizeFrac` of its width — because the frame size is a RENDER
 * decision, not a document one. The compiler resolves fractions to pixels while
 * it builds the plan, which is why a size has to be chosen before building and
 * cannot be applied afterwards.
 *
 * THE DOC IS THE ONLY THING THAT PERSISTS. A plan is derived, a canvas is
 * derived, loaded bitmaps are derived. Everything the editor lets someone
 * change writes here and nowhere else, so undo is a snapshot of this object and
 * autosave is a PATCH of this object — neither needs to know what a layer is.
 */

export const DOC_VERSION = 1;

/** A blank creative. `kind` decides which editor opens it, nothing else. */
export function emptyDoc(kind = 'image') {
    return {
        v: DOC_VERSION,
        kind,
        // Images name a size from IMAGE_SIZES; videos name an aspect. Both are
        // stored, so switching kind keeps the other's last choice.
        size: 'square',
        options: {
            aspect: 'vertical',
            theme: 'dark',
            fit: 'blur',
        },
        title: '',
        body: '',
        // Source pictures, as media URLs. The editor resolves and loads these;
        // the doc only ever holds the reference.
        images: [],
        // Everything added or changed on the canvas, in the renderer's own
        // patch format — so the doc IS the recipe, with no translation step.
        patches: [],
        // Which instant a still freezes. Meaningless for video.
        at: 0,
    };
}

/** Coerce anything we were handed into a usable doc, without throwing. */
export function normaliseDoc(raw, kind) {
    const base = emptyDoc(kind || raw?.kind || 'image');
    if (!raw || typeof raw !== 'object') return base;
    return {
        ...base,
        ...raw,
        v: DOC_VERSION,
        kind: raw.kind || base.kind,
        options: { ...base.options, ...(raw.options || {}) },
        images: Array.isArray(raw.images) ? raw.images : [],
        patches: Array.isArray(raw.patches) ? raw.patches : [],
    };
}

// ── patches ──────────────────────────────────────────────────────────────────

/**
 * A fresh layer id.
 *
 * Prefixed by type and suffixed by a counter over the EXISTING ids rather than
 * by a clock or a random value. Two reasons, and the second is the real one:
 * ids show up in the layer list, so `wordart-2` is worth more than `l-1h8k2x`;
 * and a render's dedupe key is a hash of the doc, so an id that varied between
 * two otherwise identical docs would defeat it and re-render for nothing.
 */
export function nextId(doc, type) {
    let n = 1;
    const taken = new Set((doc.patches || []).map((p) => p.id));
    while (taken.has(`${type}-${n}`)) n += 1;
    return `${type}-${n}`;
}

/** Append a layer. Returns a NEW doc — nothing here mutates. */
export function addLayer(doc, patch) {
    const id = patch.id || nextId(doc, patch.type || 'layer');
    return { ...doc, patches: [...doc.patches, { ...patch, id }] };
}

/**
 * Merge a change into one layer's patch.
 *
 * Deep-merges the two objects a panel edits piecemeal — `style` for word art,
 * `adjust` for an adjustment, `mask` for either — so a slider can send
 * `{ style: { curve: 0.4 } }` without carrying the rest of the style with it,
 * which is the shape that quietly drops the other fields.
 */
export function updateLayer(doc, id, change) {
    return {
        ...doc,
        patches: doc.patches.map((p) => {
            if (p.id !== id) return p;
            const next = { ...p, ...change };
            for (const key of ['style', 'adjust', 'mask', 'fill', 'stroke', 'shadow', 'timing']) {
                if (change[key] && typeof change[key] === 'object' && p[key] && typeof p[key] === 'object') {
                    next[key] = { ...p[key], ...change[key] };
                }
            }
            return next;
        }),
    };
}

export function removeLayer(doc, id) {
    return { ...doc, patches: doc.patches.filter((p) => p.id !== id) };
}

/**
 * Move a layer up or down the stack.
 *
 * Reorders the PATCH LIST and lets compilation assign z from it, rather than
 * writing explicit z values. A stored z has to be kept consistent with every
 * other layer's z forever; a list position is consistent by construction.
 */
export function reorderLayer(doc, id, delta) {
    const patches = [...doc.patches];
    const i = patches.findIndex((p) => p.id === id);
    if (i < 0) return doc;
    const j = Math.max(0, Math.min(patches.length - 1, i + delta));
    if (i === j) return doc;
    const [moved] = patches.splice(i, 1);
    patches.splice(j, 0, moved);
    // Any z the layers were carrying is now a lie — drop it so the compiler's
    // "appended layers land on top, in order" rule decides again.
    return { ...doc, patches: patches.map(({ z, ...rest }) => rest) };
}

/**
 * Move a layer so it lands at a given z — the timeline's way of reordering.
 *
 * TWO EDITORS, ONE ORDER. The layer stack reorders by moving a row up or down
 * the list; the timeline reorders by dragging a lane, and it reports the result
 * as an absolute `z` midway between the lanes it was dropped between. Storing
 * that z would give the document two competing notions of order — and
 * `reorderLayer` above strips z precisely so there is only one — so a stored z
 * would survive exactly until the next drag in the other editor, then vanish
 * with no way to tell why.
 *
 * So the z is translated into a list position and thrown away. It needs the
 * PLAN to do that, because the z values the timeline compared against are the
 * compiled ones: patch layers get theirs from `nextZ()` at compile time, and
 * the built-in chrome has z values this list never sees.
 */
export function moveLayerToZ(doc, plan, id, z) {
    if (!plan) return doc;
    const patchIds = new Set((doc.patches || []).map((p) => p.id));
    // The patch-backed layers as the renderer sees them, in paint order.
    const ordered = plan.layers
        .filter((l) => patchIds.has(l.id) && l.id !== id)
        .sort((a, b) => (a.z || 0) - (b.z || 0))
        .map((l) => l.id);

    // Where the new z falls among them. Everything with a lower z paints first,
    // so the moved layer goes after the last of those.
    const zOf = new Map(plan.layers.map((l) => [l.id, l.z || 0]));
    let at = ordered.length;
    for (let i = 0; i < ordered.length; i++) {
        if (zOf.get(ordered[i]) > z) { at = i; break; }
    }

    const moved = (doc.patches || []).find((p) => p.id === id);
    if (!moved) return doc;
    const rest = (doc.patches || []).filter((p) => p.id !== id);
    // Rebuilt in the order the plan reports rather than the order the list
    // happened to be in: a layer whose z came from somewhere else would
    // otherwise be re-sorted by a list that never agreed with it.
    const byId = new Map(rest.map((p) => [p.id, p]));
    const sorted = ordered.map((lid) => byId.get(lid)).filter(Boolean);
    for (const p of rest) if (!ordered.includes(p.id)) sorted.push(p);
    sorted.splice(Math.min(at, sorted.length), 0, moved);

    return { ...doc, patches: sorted.map(({ z: _z, ...rest2 }) => rest2) };
}

export function duplicateLayer(doc, id) {
    const src = doc.patches.find((p) => p.id === id);
    if (!src) return doc;
    const copy = {
        ...src,
        id: nextId(doc, src.type || 'layer'),
        // Offset so the copy is visible rather than exactly behind the original.
        fx: Math.min(0.95, (src.fx ?? 0.5) + 0.03),
        fy: Math.min(0.95, (src.fy ?? 0.5) + 0.03),
    };
    delete copy.z;
    const i = doc.patches.findIndex((p) => p.id === id);
    const patches = [...doc.patches];
    patches.splice(i + 1, 0, copy);
    return { ...doc, patches };
}

/** The patch a layer id came from, or null. */
export const patchOf = (doc, id) => doc.patches.find((p) => p.id === id) || null;

// ── titles ───────────────────────────────────────────────────────────────────

/**
 * What to call a project nobody has named.
 *
 * Reads the doc rather than defaulting to "Untitled": the first headline is
 * almost always what the person thinks of the creative as, and a projects list
 * of nine "Untitled" rows is a list you cannot use.
 */
export function suggestedTitle(doc) {
    const headline = (doc.patches || []).find((p) => (p.type === 'wordart' || p.type === 'text') && p.text);
    const from = headline?.text || doc.title || doc.body;
    if (!from) return doc.kind === 'video' ? 'Untitled video' : 'Untitled image';
    const line = String(from).split('\n')[0].trim();
    return line.length > 60 ? `${line.slice(0, 57)}…` : line;
}

/**
 * Change the document's own settings — the ones that belong to the creative
 * rather than to any layer in it.
 *
 * A SHALLOW MERGE, where `updateLayer` deep-merges a named handful of
 * sub-objects. `options` is flat by construction: it is the twenty-odd scalars
 * the renderer reads straight off the plan, and there is no nesting to
 * preserve. A deep merge here would not merely be pointless, it would be
 * wrong — `audioTrackId: null` is how "no track chosen" is written, and the
 * usual "skip nullish so a partial patch cannot blank a field" rule would make
 * that value the one thing this function could never set.
 */
export function setOption(doc, change) {
    return { ...doc, options: { ...(doc.options || {}), ...change } };
}
