/**
 * The same recipe, checked and rendered at every shape.
 *
 * Fractional geometry means a recipe already COMPILES at any aspect — that
 * was the point of storing fractions. What it does not guarantee is that a
 * layout which reads well in 9:16 still reads in 1:1: the frame changes shape
 * around fixed fractions, so a chip that sat clear of the caption in one can
 * land under it in another, and platform UI covers different parts of each.
 *
 * So the useful thing to build first is not three renders — it is knowing
 * whether the three would be worth having. This checks.
 */

// Lifted from the ERP unchanged except this line: upstream it imported the
// host's `./video-maker` wrapper (which installed a media transport and
// re-exported the renderer). A package has no host, so it imports the renderer
// directly; installing the transport is the app's job, not this module's.
import { ASPECTS, buildPlan, layerBounds } from "../index.js";

// The same bands the studio's guides draw. Fractions of the frame.
export const UI_BANDS = { bottom: 0.18, right: 0.12, titleSafe: 0.05 };

// Layers that ARE the frame, or that the renderer places itself — flagging a
// full-bleed photo for touching the edges would be noise, every time.
const IGNORED = new Set(["gradient", "photo", "video", "outro", "progress", "edges", "title"]);

/**
 * Compile `args` at each aspect and report what falls somewhere it should not.
 *
 * @param {object} args    the same object buildPlan takes, minus `canvas`
 * @param {HTMLCanvasElement} canvas  scratch canvas — buildPlan resizes it
 * @param {string[]} [keys]
 * @returns {{aspect: string, label: string, issues: {id: string, name: string, what: string}[]}[]}
 */
export function checkAspects(args, canvas, keys = Object.keys(ASPECTS)) {
    const out = [];
    for (const key of keys) {
        const plan = buildPlan({ ...args, canvas, options: { ...args.options, aspect: key } });
        const ctx = canvas.getContext("2d");
        const issues = [];
        for (const layer of plan.layers) {
            if (layer.visible === false || IGNORED.has(layer.type)) continue;
            const b = layerBounds(ctx, plan, layer);
            if (!b) continue;
            // A full-frame bounds means "not really placed" (the caption, a
            // cover photo) — the same reason IGNORED exists, caught by size.
            if (b.w >= plan.W * 0.98 && b.h >= plan.H * 0.98) continue;
            const name = layer.name || layer.text || layer.id;
            const safe = UI_BANDS.titleSafe;
            if (b.x < plan.W * safe || b.x + b.w > plan.W * (1 - safe)
                || b.y < plan.H * safe || b.y + b.h > plan.H * (1 - safe)) {
                issues.push({ id: layer.id, name, what: "outside title-safe" });
            }
            if (b.y + b.h > plan.H * (1 - UI_BANDS.bottom)) {
                issues.push({ id: layer.id, name, what: "under the platform's bottom bar" });
            }
            if (b.x + b.w > plan.W * (1 - UI_BANDS.right) && b.y < plan.H * (1 - UI_BANDS.bottom)) {
                issues.push({ id: layer.id, name, what: "under the platform's right rail" });
            }
        }
        out.push({ aspect: key, label: ASPECTS[key]?.label || key, issues });
    }
    return out;
}

/** A one-line summary of a checkAspects result, for a toast. */
export function summarise(reports) {
    const bad = reports.filter((r) => r.issues.length);
    if (!bad.length) return "every aspect looks clear";
    return bad.map((r) => `${r.label}: ${[...new Set(r.issues.map((i) => i.name))].join(", ")}`).join(" · ");
}
