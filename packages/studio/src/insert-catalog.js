import {
    SHAPE_GROUPS, GEOMETRY_LABELS, WORDART_PRESETS, ADJUST_PRESETS,
    ICONS, ICON_GROUPS, ICON_LABELS,
    CHART_KINDS, CHART_LABELS,
} from "./host.js";
import { addLayer } from "./doc.js";

/**
 * Everything that can be put on a creative, arranged for a panel.
 *
 * WHY THIS EXISTS SEPARATELY FROM `ADD_MENU`. That list is six entries in a row
 * of buttons, and it was the right shape while there were six things to add.
 * There are now thirty geometries alone, with icons, charts and text blocks
 * behind them — a row of buttons cannot hold that, and the first casualty of
 * trying is discovery: a `blob` nobody can find is a `blob` nobody uses.
 *
 * THE STRUCTURE IS UNIFORM ON PURPOSE — every section has `groups`, every group
 * has `items`, and a group with no label renders no heading. The alternative
 * (sections that are sometimes a flat list and sometimes grouped) puts a branch
 * in the panel for every category added, and the panel should not have opinions
 * about what a category is.
 *
 * A TILE IS EITHER DRAWN OR NAMED. Shapes carry `{ type: 'shape', geometry }`
 * and the panel paints them with the real painter, so the tile is the shape
 * rather than a picture of something like it. Everything else carries an icon
 * name, because "text", "adjustment" and "sound" are concepts and there is no
 * geometry to draw.
 */

const group = (label, items) => ({ label, items });

/** A stable id per item — the panel keys on it and search matches on label. */
let seq = 0;
const item = (label, patch, tile, hint) => ({
    id: `${patch.type}-${label.toLowerCase().replace(/\W+/g, "-")}-${seq++}`,
    label, patch, tile, hint,
});

const icon = (name) => ({ type: "icon", icon: name });
const shapeTile = (geometry) => ({ type: "shape", geometry });

/**
 * An item that opens the media picker instead of inserting straight away.
 *
 * TWO SHAPES OF ITEM, and the second exists because three of these do genuinely
 * different things to the document. A picture is appended to `doc.images` and
 * is not a layer at all; a clip becomes TWO layers; a sound becomes one. A
 * single `patch` cannot express that, so a picker item carries `apply` — a
 * document transform — and the panel hands it whatever was chosen.
 */
const pickItem = (label, glyph, hint, pick, apply) => ({
    id: `pick-${pick}-${seq++}`, label, tile: icon(glyph), hint, pick, apply,
});

/** A readable name from a media url, for the layer it becomes. */
function stemOf(url) {
    try {
        const path = new URL(url, "https://x.invalid").pathname;
        return decodeURIComponent(path.split("/").pop() || "").replace(/\.[^.]+$/, "") || "Clip";
    } catch { return "Clip"; }
}

// ── text ─────────────────────────────────────────────────────────────────────

/**
 * The word-art presets ARE the text catalogue.
 *
 * A preset is the difference between "some text" and a headline someone would
 * actually ship, and until now they were reachable only from the inspector —
 * which meant adding word art, then discovering that a dropdown you had not
 * looked at yet held the whole point of the feature. Offering them at insert
 * time puts the decision where it is being made.
 */
const TEXT_ITEMS = Object.entries(WORDART_PRESETS).map(([key, preset]) =>
    item(
        preset.label || key,
        { type: "wordart", text: "HEADLINE", preset: key, fx: 0.5, fy: 0.32, sizeFrac: 0.11 },
        icon("fa-wand-magic-sparkles"),
        "word art",
    ));

/**
 * Paragraphs and lists — all `textblock`, differing only in `list`.
 *
 * They are one layer type because a list IS a paragraph whose lines carry a
 * marker and a hanging indent; the plan had them as two items at two
 * difficulties and building the block once turned out to be less code than
 * building a list. Offering the modes as separate tiles is a UI decision, not a
 * model one — nobody looking for a bulleted list thinks of it as a setting on
 * something else.
 */
const BODY_ITEMS = [
    item("Paragraph",
        { type: "textblock", text: "Say a little more here.", fx: 0.5, fy: 0.62, fw: 0.72, sizeFrac: 0.035 },
        icon("fa-align-left"), "text that wraps in a box"),
    item("Bulleted",
        { type: "textblock", list: "bullet", text: "First point\nSecond point\nThird point", fx: 0.5, fy: 0.55, fw: 0.66, sizeFrac: 0.035 },
        icon("fa-list-ul"), "a list"),
    item("Numbered",
        { type: "textblock", list: "number", text: "Step one\nStep two\nStep three", fx: 0.5, fy: 0.55, fw: 0.66, sizeFrac: 0.035 },
        icon("fa-list-ol"), "a numbered list"),
    item("Checklist",
        { type: "textblock", list: "check", text: "Pack it\nLabel it\nSend it", fx: 0.5, fy: 0.55, fw: 0.66, sizeFrac: 0.035 },
        icon("fa-list-check"), "a checklist"),
    item("Caption chip",
        { type: "text", text: "A short label", fx: 0.5, fy: 0.86, sizeFrac: 0.03 },
        icon("fa-tag"), "one line, on a pill"),
];

// ── shapes ───────────────────────────────────────────────────────────────────

/**
 * Every geometry the renderer knows, grouped the way it groups itself.
 *
 * Derived from `SHAPE_GROUPS` rather than listed again here: a geometry added
 * to `shapes.js` appears in this panel with no change to the app, which is the
 * property that stops the palette from silently falling behind the renderer.
 */
const SHAPE_SECTIONS = SHAPE_GROUPS.map((g) => group(
    g.label,
    g.geometries.map((geometry) => item(
        GEOMETRY_LABELS[geometry] || geometry,
        { type: "shape", geometry, fx: 0.5, fy: 0.5, fw: 0.36, ...defaultsFor(geometry) },
        shapeTile(geometry),
    )),
));

/**
 * Per-geometry insert defaults.
 *
 * A shape inserted at its parameter defaults is not always a shape anyone
 * wants: a `pie` at a full turn is a circle, a `frame` with a hairline band is
 * a rectangle. These are the values that make the first click produce the thing
 * the tile promised — the RENDERER's defaults stay as they are, because a
 * stored recipe must not change meaning.
 */
function defaultsFor(geometry) {
    switch (geometry) {
        case "rect": return { radius: 0.08 };
        case "line": return { fh: 0.02, thickness: 0.5 };
        case "wave": return { fh: 0.14, points: 2, thickness: 0.3 };
        // No `fh`: compileShape derives a circular height from the width when
        // it is absent, which is what a ring wants on any aspect.
        case "ring": return { innerRatio: 0.62 };
        case "pie": return { sweep: 0.72 };
        case "frame": return { thickness: 0.1 };
        case "gear": return { points: 8, thickness: 0.24, innerRatio: 0.34 };
        case "burst": return { innerRatio: 0.82 };
        case "star": return { points: 5, innerRatio: 0.45 };
        case "polygon": return { points: 6 };
        case "bubble": return { fh: 0.28, radius: 0.14 };
        case "banner": case "chevron": return { fh: 0.16, notch: 0.14 };
        case "tag": return { fh: 0.22, notch: 0.16, innerRatio: 0.12 };
        case "bookmark": return { fh: 0.34, notch: 0.24 };
        case "arrow": case "doubleArrow": return { fh: 0.2 };
        case "plus": return { thickness: 0.34 };
        default: return {};
    }
}

// ── icons ────────────────────────────────────────────────────────────────────

/**
 * The icon set, grouped the way it groups itself.
 *
 * The tile carries the PATH rather than a geometry, so the panel renders it as
 * inline SVG — crisp at any size, no canvas, and the same path data the
 * renderer will paint. Same principle as a shape tile: the tile is the thing.
 */
const ICON_SECTIONS = ICON_GROUPS.map((g) => group(
    g.label,
    g.names.map((name) => item(
        ICON_LABELS[name] || name,
        { type: "icon", name, fx: 0.5, fy: 0.5, fw: 0.16 },
        { type: "path", d: ICONS[name] },
        "icon",
    )),
));

// ── charts ───────────────────────────────────────────────────────────────────

/**
 * One tile per chart kind, each carrying sample data.
 *
 * A chart inserted EMPTY draws nothing, and a tile that produces an invisible
 * layer is indistinguishable from one that is broken. Sample numbers mean the
 * first click puts a real chart on the canvas, which is also the fastest way to
 * see which kind you actually wanted.
 */
const CHART_TILE_ICON = {
    bar: "fa-chart-bar", column: "fa-chart-column", line: "fa-chart-line",
    area: "fa-chart-area", pie: "fa-chart-pie", donut: "fa-circle-notch",
};

const CHART_ITEMS = CHART_KINDS.map((chart) => item(
    CHART_LABELS[chart] || chart,
    {
        type: "chart", chart,
        labels: ["Mon", "Tue", "Wed", "Thu"],
        values: [12, 19, 9, 24],
        fx: 0.5, fy: 0.5, fw: 0.72, fh: 0.34,
        ...(chart === "pie" || chart === "donut"
            // A four-slice pie of near-equal values says nothing; a share is
            // what a pie is for.
            ? { labels: ["Ours", "Theirs", "Other"], values: [58, 27, 15], fh: 0.28 }
            : {}),
    },
    icon(CHART_TILE_ICON[chart] || "fa-chart-simple"),
    "chart",
));

// ── effects and media ────────────────────────────────────────────────────────

const ADJUST_ITEMS = Object.entries(ADJUST_PRESETS)
    // 'none' is what an adjustment layer already is before you touch it, so
    // offering it as something to insert is offering a no-op.
    .filter(([key]) => key !== "none")
    .map(([key, preset]) => item(
        preset.label || key,
        { type: "adjust", preset: key },
        icon("fa-sliders"),
        "adjustment",
    ));

/**
 * Media somebody ALREADY has.
 *
 * The recorder could place a take but there was no way to place a file that was
 * already in the library — half a feature, and the half people reach for more
 * often. These are the other half.
 *
 * A clip becomes the SAME two layers a recorded one does (picture and sound on
 * one url), so the two routes into the timeline cannot diverge. It arrives with
 * no `timing`, which the renderer reads as "the whole video": a recorded take
 * knows its own length because the recorder measured it, and a library file
 * does not without decoding it first — so the honest default is the full span,
 * and the lane is dragged from there.
 */
const MEDIA_ITEMS = [
    pickItem("Picture", "fa-image", "from the library", "image",
        // Pictures are not layers — they are the creative's picture list, which
        // the caption timing and the Ken Burns pass are both derived from.
        (doc, urls) => ({ ...doc, images: [...(doc.images || []), ...urls] })),

    pickItem("Video clip", "fa-film", "picture and sound", "video",
        (doc, urls) => urls.reduce((d, url) => {
            const stem = stemOf(url);
            const withPicture = addLayer(d, { type: "video", url, name: stem });
            return addLayer(withPicture, {
                type: "sound", url, name: `${stem} (sound)`, mix: "duck", volume: 1,
            });
        }, doc)),

    pickItem("Sound", "fa-file-audio", "from the library", "audio",
        (doc, urls) => urls.reduce((d, url) => addLayer(d, {
            type: "sound", url, name: stemOf(url), mix: "mix", volume: 0.9,
        }), doc)),
];

const SOUND_ITEMS = [
    // The default IS the feature — `duck` is what pulls a music bed down while
    // someone is talking, and it is the field nobody would go looking for.
    item("Voice-over", { type: "sound", name: "Voice-over", mix: "duck", volume: 1 },
        icon("fa-microphone"), "ducks the music under it"),
    item("Sound effect", { type: "sound", name: "Sound", mix: "mix", volume: 0.9 },
        icon("fa-volume-high"), "plays over everything"),
];

/**
 * The catalogue for a document of this kind.
 *
 * Sound is video-only — a still has no timeline for a clip to sit on — and that
 * filter lives here rather than in the renderer, which has no notion of a
 * document kind and should not acquire one.
 */
export function insertCatalogFor(kind) {
    const sections = [
        {
            key: "media",
            label: "Media",
            icon: "fa-photo-film",
            // A still can carry pictures but not clips or sound — the same rule
            // the renderer keeps, stated once here rather than per item.
            groups: [group(null, kind === "video"
                ? MEDIA_ITEMS
                : MEDIA_ITEMS.filter((m) => m.pick === "image"))],
        },
        {
            key: "text",
            label: "Text",
            icon: "fa-font",
            groups: [group("Headlines", TEXT_ITEMS), group("Paragraphs & lists", BODY_ITEMS)],
        },
        {
            key: "shapes",
            label: "Shapes",
            icon: "fa-shapes",
            groups: SHAPE_SECTIONS,
        },
        {
            key: "icons",
            label: "Icons",
            icon: "fa-icons",
            groups: ICON_SECTIONS,
        },
        {
            key: "charts",
            label: "Charts",
            icon: "fa-chart-simple",
            groups: [group(null, CHART_ITEMS)],
        },
        {
            key: "effects",
            label: "Effects",
            icon: "fa-sliders",
            groups: [group(null, ADJUST_ITEMS)],
        },
    ];

    if (kind === "video") {
        sections.push({ key: "sound", label: "Sound", icon: "fa-music", groups: [group(null, SOUND_ITEMS)] });
    }

    return sections;
}

/**
 * The timeline's add row: doors into the Insert tab, not a second menu.
 *
 * The row used to carry its own buttons (the toolset's six plus two sounds),
 * which meant two lists of what can be added, and the shorter one had no clips,
 * no library sound and no charts. Each entry now opens the Insert tab on one
 * section — the two a timeline is for, media and sound — so there is one
 * catalogue and the row can never offer less than the tab does.
 */
const TIMELINE_DOORS = [
    { section: "media", label: "Clip or picture", icon: "fa-photo-film" },
    { section: "sound", label: "Sound", icon: "fa-music" },
];

/** The timeline's doors for a document of this kind — only sections that exist for it. */
export function timelineInsertsFor(kind) {
    const keys = new Set(insertCatalogFor(kind).map((s) => s.key));
    return TIMELINE_DOORS.filter((d) => keys.has(d.section));
}

/**
 * Everything, flattened, for search.
 *
 * Search crosses sections deliberately: someone typing "star" should not first
 * have to know that a star is a shape rather than an icon. The section is shown
 * on the result so the answer also teaches where it lives.
 */
export function searchCatalog(kind, query) {
    const q = String(query || "").trim().toLowerCase();
    if (!q) return null;
    const out = [];
    for (const section of insertCatalogFor(kind)) {
        for (const g of section.groups) {
            for (const it of g.items) {
                const haystack = `${it.label} ${it.hint || ""} ${g.label || ""} ${section.label}`.toLowerCase();
                if (haystack.includes(q)) out.push({ ...it, section: section.label });
            }
        }
    }
    return out;
}
