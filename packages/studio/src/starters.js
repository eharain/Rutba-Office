// Video starters (the 2.7 plan's templates: "Video starters (slideshow,
// tutorial, vertical social, montage)"): a project laid out with its titles,
// its timing and its transitions, on stand-in pictures the window draws, so
// each one plays as soon as it is opened. Replace Media puts a person's own
// picture or clip where a stand-in is, keeping its place and length.

import { emptyProject, addMedia, addTitle, addGraphic, setTransition, clipById, laneClips } from './project.js';

/** The starters, in the order offered: key, the frame's shape, how many stand-in pictures, and how long each shows. */
export const STARTERS = [
    { key: 'slideshow', aspect: 'landscape', pictures: 6, seconds: 4 },
    { key: 'tutorial', aspect: 'landscape', pictures: 3, seconds: 6 },
    { key: 'social', aspect: 'vertical', pictures: 4, seconds: 3 },
    { key: 'montage', aspect: 'landscape', pictures: 10, seconds: 1.2 },
];

export const starterByKey = (key) => STARTERS.find((s) => s.key === key) || null;

/**
 * A starter's project. `pictures` are the stand-ins as addMedia takes them
 * (kind 'image', with their width and height), `words` the titles in the
 * window's language: { title, subtitle, step1, step2, tip, end, tag }.
 * Answers the project, every picture a still of the starter's length on the
 * main track, titles above.
 */
export function buildStarter(key, pictures, words = {}) {
    const def = starterByKey(key);
    if (!def) return null;
    const w = { title: 'Title', subtitle: '', step1: 'Step one', step2: 'Step two', tip: 'A tip', end: 'Thank you for watching', tag: 'New', ...words };
    let doc = emptyProject();
    doc = { ...doc, options: { ...doc.options, aspect: def.aspect }, settings: { ...doc.settings, stillSeconds: def.seconds } };
    for (const pic of pictures.slice(0, def.pictures)) doc = addMedia(doc, { ...pic, kind: 'image' }, { lane: 0 }).doc;
    const stills = laneClips(doc, 'v', 0);
    const end = stills.length ? stills[stills.length - 1].timing.end : 0;
    const title = (d, preset, at, text) => addTitle(d, preset, { at, text }).doc;
    if (key === 'slideshow') {
        // A crossfade between every picture; the title over the first, an end card over the last.
        for (const s of stills.slice(1)) doc = setTransition(doc, s.id, { kind: 'fade', seconds: 1 });
        doc = title(doc, 'title', 0, w.title);
        const last = laneClips(doc, 'v', 0).at(-1);
        doc = title(doc, 'endCard', Math.max(0, last.timing.end - 3), w.end);
    } else if (key === 'tutorial') {
        // A title, then each step named in a lower third over its picture, a callout for a tip.
        doc = title(doc, 'title', 0, w.title);
        const [, two, three] = laneClips(doc, 'v', 0);
        if (two) doc = title(doc, 'lowerThird', two.timing.start + 0.5, w.step1);
        if (three) doc = title(doc, 'lowerThird', three.timing.start + 0.5, w.step2);
        if (three) doc = title(doc, 'callout', three.timing.start + 2.5, w.tip);
        doc = title(doc, 'endCard', Math.max(0, end - 3), w.end);
    } else if (key === 'social') {
        // Upright, quick slides in, a sticker over the first picture.
        for (const s of stills.slice(1)) doc = setTransition(doc, s.id, { kind: 'slide', seconds: 0.4 });
        doc = addGraphic(doc, { type: 'wordart', text: w.tag, preset: 'sticker', fx: 0.5, fy: 0.16, sizeFrac: 0.14 }, { at: 0, seconds: def.seconds }).doc;
        doc = title(doc, 'callout', 0.6, w.title);
        const last = laneClips(doc, 'v', 0).at(-1);
        doc = title(doc, 'endCard', Math.max(0, last.timing.end - 2.5), w.end);
    } else if (key === 'montage') {
        // Quick cuts with no transitions, a title over the first two, an end card at the close.
        doc = title(doc, 'title', 0, w.title);
        doc = title(doc, 'endCard', Math.max(0, end - 2.4), w.end);
    }
    return doc;
}

/** Whether a clip is a starter's stand-in picture, which Replace Media is for. */
export const isStandIn = (doc, id) => Boolean(clipById(doc, id)?.standIn);

/** The stand-ins marked, so the window can say what to do with them. */
export function markStandIns(doc, urls) {
    const set = new Set(urls);
    return { ...doc, patches: doc.patches.map((p) => (p.type === 'image' && set.has(p.url) ? { ...p, standIn: true } : p)) };
}
