import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import History from "@rutba/editing/history";
import {
    ASPECTS, THEMES, DEFAULTS,
    buildPlan, paintFrame, loadImages, releaseImages, loadVideo, releaseVideos,
    layerBounds, hitTestLayers, layerHandles, hitTestHandles, scaleFromDrag, resizePatch,
    wordArtResize, buildImagePlan, IMAGE_SIZES,
} from "../../lib/renderer";
import { normaliseDoc, updateLayer, removeLayer, patchOf } from "../../lib/doc";

/**
 * The angle of a point about a centre, in degrees, with 0 pointing UP.
 *
 * Up rather than right because the rotate stalk sits above the layer: grabbing
 * it should read as "0° so far", and an atan2 that answers -90 there makes
 * every rotation off by a quarter turn.
 */
function angleAt(center, x, y) {
    return (Math.atan2(y - center.y, x - center.x) * 180) / Math.PI + 90;
}

/**
 * useLayoutEffect on the client, useEffect on the server.
 *
 * These pages prerender, and React warns that useLayoutEffect does nothing
 * during SSR — correctly, since there is no layout to read. The canvas paint
 * genuinely needs the layout phase (it must land before the browser composites
 * the commit, or it trails the overlay by a frame), so the hook is swapped
 * rather than downgraded.
 */
const useIsomorphicLayoutEffect = typeof window !== "undefined" ? useLayoutEffect : useEffect;

/**
 * The editor engine: a document goes in, a painted canvas and a set of gestures
 * come out.
 *
 * WHY ONE HOOK FOR BOTH EDITORS. The image editor and the video studio differ
 * in exactly two places — which builder makes the plan, and whether there is a
 * clock — and in nothing else. Selection, dragging, resizing, undo, the layer
 * stack and the inspector are identical, because an image IS a one-frame video
 * in this renderer. Two hooks would have meant two subtly different drag
 * implementations within a month.
 *
 * THE DERIVATION CHAIN, and why each step is where it is:
 *
 *   doc  ──(load)──▶  bitmaps  ──(compile)──▶  plan  ──(paint)──▶  canvas
 *
 * Loading is async and expensive, so it is keyed on the image list alone and
 * survives every other edit. Compiling is synchronous and cheap but NOT free —
 * it resolves every fraction to pixels — so it reruns when the doc or the frame
 * size changes, and never on a mere repaint. Painting is pure and runs whenever
 * anything at all changes, including the playhead.
 *
 * The plan is REBUILT rather than mutated on every edit. That is the whole
 * reason the geometry stays honest: `compileLayers` is the only thing that
 * turns a fraction into a pixel, so a patch that changed a pixel directly would
 * be correct until the next size change and wrong after it.
 */
export function useCreative({ doc: initialDoc, kind, onDirty }) {
    const [doc, setDocState] = useState(() => normaliseDoc(initialDoc, kind));
    const [selectedId, setSelectedId] = useState(null);
    const [time, setTime] = useState(0);
    /**
     * A request, from outside the left rail, to open its Insert tab on one
     * section — the timeline's add row makes it. A counter rides with the
     * section so asking twice for the same one still reaches the rail.
     */
    const [insertAsk, setInsertAsk] = useState(null);
    const askInsert = useCallback((section) => setInsertAsk((prev) => ({ section, n: (prev?.n || 0) + 1 })), []);
    const [images, setImages] = useState([]);
    const [loading, setLoading] = useState(false);
    const [loadError, setLoadError] = useState(null);
    const [, forceHistoryTick] = useState(0);

    const canvasRef = useRef(null);
    const historyRef = useRef(null);
    if (!historyRef.current) historyRef.current = new History({ limit: 120 });

    /**
     * A scratch canvas that compilation owns, so the VISIBLE one is never
     * cleared out from under a frame.
     *
     * THIS IS THE FIX FOR THE FLICKER, and the cause is worth writing down
     * because it is not obvious. `buildPlan` assigns `canvas.width` and
     * `canvas.height` unconditionally (index.js), and assigning to either
     * RESETS the bitmap to transparent — that is what the width setter does,
     * even when the value is unchanged. Compilation runs inside the `plan`
     * memo, which is React's RENDER phase, while the paint used to run in a
     * passive effect, which is after the browser has had its chance to
     * composite. So every drag tick cleared the canvas, let a frame through
     * blank, and then drew. Hence: flicker, worst while dragging, because that
     * is when recompiles happen fastest.
     *
     * Compiling against a canvas nobody is looking at removes the window
     * entirely. It is sound because a plan is not bound to the canvas it was
     * compiled on — `paintFrame(ctx, plan, t)` takes its target explicitly, the
     * deck presenter already paints one plan across two stacked canvases, and
     * `renderImageSet` is built on the same property. The scratch canvas is
     * still a REAL canvas rather than an OffscreenCanvas because compilation
     * measures text through its 2D context, and `measureText` must see the same
     * font stack the visible one will paint with.
     *
     * It also means the display canvas is only ever written by `paintFrame`,
     * which opens by filling the whole frame with the theme background — so it
     * never needs clearing and never shows a gap.
     */
    const scratchRef = useRef(null);
    const scratch = useCallback(() => {
        if (!scratchRef.current && typeof document !== "undefined") {
            scratchRef.current = document.createElement("canvas");
        }
        return scratchRef.current;
    }, []);
    /** The measuring context — the one compilation and hit-testing share. */
    const measureCtx = useCallback(() => scratch()?.getContext("2d") || null, [scratch]);

    /**
     * A callback ref, not a plain one, because the plan is DERIVED FROM the
     * canvas (buildPlan sizes it and compiles against its dimensions) and a ref
     * assignment does not re-render. With a plain ref the first compile runs
     * before the canvas exists and returns null, and the only thing that would
     * ever retry is the image load — so a creative with no pictures in it, which
     * is the common starting point, would sit blank forever.
     */
    const [canvasReady, setCanvasReady] = useState(false);
    const attachCanvas = useCallback((node) => {
        canvasRef.current = node;
        setCanvasReady(Boolean(node));
    }, []);

    // ── editing, with undo ───────────────────────────────────────────────────

    /**
     * Apply a change to the doc and record the state BEFORE it.
     *
     * `group` is what makes dragging one undo step instead of ninety: every
     * pointermove during a drag records under the same group, and History
     * coalesces them. Releasing the pointer breaks the group, so the next drag
     * is its own step.
     */
    const setDoc = useCallback((next, { label = "edit", group = null } = {}) => {
        setDocState((current) => {
            const value = typeof next === "function" ? next(current) : next;
            if (value === current) return current;
            historyRef.current.record({ state: current, label, group, meta: selectedId });
            forceHistoryTick((n) => n + 1);
            if (onDirty) onDirty(value);
            return value;
        });
    }, [onDirty, selectedId]);

    const undo = useCallback(() => {
        setDocState((current) => {
            const entry = historyRef.current.undo(current, selectedId);
            if (!entry) return current;
            // The selection is restored with the state. An undo that leaves you
            // selecting a layer that no longer exists is disorienting, and an
            // undo that clears the selection loses your place.
            setSelectedId(entry.meta ?? null);
            forceHistoryTick((n) => n + 1);
            if (onDirty) onDirty(entry.state);
            return entry.state;
        });
    }, [onDirty, selectedId]);

    const redo = useCallback(() => {
        setDocState((current) => {
            const entry = historyRef.current.redo(current, selectedId);
            if (!entry) return current;
            setSelectedId(entry.meta ?? null);
            forceHistoryTick((n) => n + 1);
            if (onDirty) onDirty(entry.state);
            return entry.state;
        });
    }, [onDirty, selectedId]);

    const endGesture = useCallback(() => { historyRef.current.break(); }, []);
    const history = historyRef.current.describe();

    // ── bitmaps ──────────────────────────────────────────────────────────────

    // Keyed on the URL list, not on the doc: everything else about a creative
    // can change without the pictures needing to be fetched and decoded again.
    const imageKey = useMemo(() => JSON.stringify(doc.images || []), [doc.images]);

    useEffect(() => {
        const urls = JSON.parse(imageKey);
        if (!urls.length) { setImages([]); return undefined; }
        let cancelled = false;
        const controller = new AbortController();
        setLoading(true);
        setLoadError(null);
        loadImages(urls, { signal: controller.signal })
            .then((entries) => {
                if (cancelled) { releaseImages(entries); return; }
                setImages(entries);
            })
            .catch((err) => {
                if (!cancelled) setLoadError(err?.message || "Could not load the pictures.");
            })
            .finally(() => { if (!cancelled) setLoading(false); });
        return () => { cancelled = true; controller.abort(); };
    }, [imageKey]);

    // Release decoded bitmaps when the editor goes away. An ImageBitmap is not
    // garbage — it holds GPU memory until it is closed, and a session that
    // opens twenty projects would hold all twenty.
    useEffect(() => () => releaseImages(images), [images]);

    /**
     * Video clips, decoded into `plan.videos`.
     *
     * WITHOUT THIS A CLIP LAYER IS INVISIBLE AND SAYS NOTHING. `compileVideo`
     * sets `visible: patch.visible !== false && !!entry` — no decoded entry, no
     * layer — so a recorded take used to land on the timeline as a lane that
     * drew nothing, with no error to explain it. That is the failure mode the
     * recorder would otherwise have shipped with.
     *
     * KEYED ON THE URL LIST, exactly like the images above, and for the same
     * reason: decoding a clip is expensive and everything else about a creative
     * can change without needing it again. Videos live in the PATCHES rather
     * than in `doc.images`, so the key is gathered from there.
     *
     * A `<video>` element per clip is heavier than a bitmap — it holds a decode
     * pipeline open — so the release is not optional housekeeping, and it runs
     * on every change of the set rather than only on unmount.
     */
    const videoKey = useMemo(() => JSON.stringify(
        [...new Set((doc.patches || [])
            .filter((p) => p.type === 'video' && p.url)
            .map((p) => p.url))].sort(),
    ), [doc.patches]);

    const [videos, setVideos] = useState({});

    useEffect(() => {
        const urls = JSON.parse(videoKey);
        if (!urls.length) { setVideos({}); return undefined; }
        let cancelled = false;
        const controller = new AbortController();
        setLoading(true);
        Promise.all(urls.map((url) => loadVideo(url, { signal: controller.signal })
            // One clip that will not decode loses that clip, not the editor —
            // the same posture loadImages keeps for a missing picture.
            .then((entry) => [url, entry])
            .catch(() => [url, null])))
            .then((pairs) => {
                const map = Object.fromEntries(pairs.filter(([, v]) => v));
                if (cancelled) { releaseVideos(map); return; }
                setVideos(map);
            })
            .finally(() => { if (!cancelled) setLoading(false); });
        return () => { cancelled = true; controller.abort(); };
    }, [videoKey]);

    useEffect(() => () => releaseVideos(videos), [videos]);

    // ── the plan ─────────────────────────────────────────────────────────────

    const plan = useMemo(() => {
        // The SCRATCH canvas, not the visible one — see the note on scratchRef.
        // Compilation sizes whatever it is handed and blanks it in the process,
        // which must not happen to something on screen.
        const canvas = scratch();
        if (!canvas) return null;
        const shared = {
            canvas,
            images,
            // Decoded clips, keyed by url — `compileVideo` looks each layer up
            // here and marks it invisible when there is no entry, which is why
            // this must be threaded through rather than left to default.
            videos,
            title: doc.title || "",
            body: doc.body || "",
            options: doc.options || {},
            layerPatches: doc.patches || [],
            context: doc.context || {},
            assets: doc.assets || {},
        };
        try {
            return doc.kind === "video"
                ? buildPlan(shared)
                : buildImagePlan({ ...shared, size: doc.size || "square" });
        } catch (err) {
            // A malformed recipe must not take the whole editor down with it —
            // the layer list is how someone would FIX it.
            console.error("[studio] could not compile the creative", err);
            return null;
        }
        // No longer waits on the visible canvas: compilation owns a scratch one
        // that exists from the first call, so a creative with no pictures — the
        // common starting point — compiles immediately rather than on mount.
    }, [doc, images, videos, scratch]);

    // ── painting ─────────────────────────────────────────────────────────────

    const repaint = useCallback(() => {
        const canvas = canvasRef.current;
        if (!canvas || !plan) return;
        const ctx = canvas.getContext("2d");
        paintFrame(ctx, plan, doc.kind === "video" ? time : (doc.at || 0));
    }, [plan, time, doc.kind, doc.at]);

    /**
     * Paint BEFORE the browser composites the commit, not after it.
     *
     * A passive `useEffect` runs once the browser has already had the chance to
     * paint, so the canvas trailed the DOM by a frame: the SVG selection box and
     * its handles moved with the commit while the picture under them arrived
     * later. During a drag that reads as the handles sliding off the thing they
     * are supposed to be gripping.
     *
     * `useLayoutEffect` runs synchronously after the DOM is mutated and before
     * the paint, so the canvas and the overlay are always the same frame of the
     * same gesture. It is affordable because the gesture is throttled to one
     * commit per animation frame (see onPointerMove) — without that, this would
     * be a full repaint per pointer event, and a 120Hz trackpad delivers those
     * faster than the screen can show them.
     *
     * `canvasReady` is in the deps because `canvasRef.current` is not reactive:
     * it is what turns "the canvas has mounted" into the first paint.
     */
    useIsomorphicLayoutEffect(() => { repaint(); }, [repaint, canvasReady]);

    /**
     * Repaint when a video clip finishes seeking.
     *
     * `paintVideo` follows the playhead by setting `currentTime` on a paused
     * element — but a seek COMPLETES ASYNCHRONOUSLY, so the paint that asked
     * for the new frame draws the old one and nothing ever comes back to draw
     * the new. The symptom is a clip stuck a frame or two behind the scrubber,
     * which reads as the video being broken rather than as a timing detail.
     *
     * The listener goes through a ref because `repaint`'s identity changes on
     * every plan, and re-binding a listener per keystroke on every decoded clip
     * is churn for nothing.
     */
    const repaintRef = useRef(repaint);
    repaintRef.current = repaint;

    useEffect(() => {
        const entries = Object.values(videos);
        if (!entries.length) return undefined;
        const onSeeked = () => repaintRef.current?.();
        for (const v of entries) v.el.addEventListener("seeked", onSeeked);
        return () => { for (const v of entries) v.el.removeEventListener("seeked", onSeeked); };
    }, [videos]);

    // ── gestures ─────────────────────────────────────────────────────────────

    const dragRef = useRef(null);

    /** Canvas coordinates from a pointer event, through the on-screen scale. */
    const toCanvas = useCallback((event) => {
        const canvas = canvasRef.current;
        if (!canvas) return { x: 0, y: 0 };
        const rect = canvas.getBoundingClientRect();
        return {
            x: ((event.clientX - rect.left) / rect.width) * canvas.width,
            y: ((event.clientY - rect.top) / rect.height) * canvas.height,
        };
    }, []);

    const onPointerDown = useCallback((event) => {
        const canvas = canvasRef.current;
        if (!canvas || !plan) return;
        // The MEASURING context, which is the one the plan was compiled
        // against. Hit-testing lays text out to find a word-art layer's box, so
        // it has to measure the same way compilation did.
        const ctx = measureCtx();
        if (!ctx) return;
        const { x, y } = toCanvas(event);

        // Handles first. A corner handle sits ON the layer, so testing the body
        // first would make a resize impossible to start.
        const selected = selectedId ? plan.layers.find((l) => l.id === selectedId) : null;
        if (selected) {
            const handle = hitTestHandles(ctx, plan, selected, x, y);
            if (handle) {
                // The rotate stalk is one of the handles, but it is not a
                // resize: resizePatch only answers for the four corners and
                // returns null for anything else, so routing it through there
                // would silently do nothing.
                dragRef.current = handle.kind === "rotate"
                    ? {
                        mode: "rotate",
                        id: selectedId,
                        center: handle.center,
                        // Where the pointer started, relative to the layer's
                        // own current rotation — so grabbing the stalk does not
                        // snap the layer to the pointer before you have moved.
                        offset: (patchOf(doc, selectedId)?.rot || 0)
                            - angleAt(handle.center, x, y),
                    }
                    : { mode: "resize", handle, start: { x, y }, id: selectedId };
                canvas.setPointerCapture?.(event.pointerId);
                return;
            }
        }

        const hit = hitTestLayers(ctx, plan, x, y);
        if (!hit) { setSelectedId(null); endGesture(); return; }
        setSelectedId(hit.layer.id);
        const patch = patchOf(doc, hit.layer.id);
        dragRef.current = {
            mode: "move",
            id: hit.layer.id,
            start: { x, y },
            // The layer's fractions at the moment the drag began — deltas are
            // applied to THESE, not accumulated, so a drag cannot drift.
            from: { fx: patch?.fx ?? (hit.layer.x / plan.W), fy: patch?.fy ?? (hit.layer.y / plan.H) },
        };
        canvas.setPointerCapture?.(event.pointerId);
    }, [plan, selectedId, doc, toCanvas, endGesture, measureCtx]);

    /**
     * One drag step, applied at a known point. Separated from the event handler
     * so it can be driven by an animation frame rather than by the pointer.
     */
    const applyDrag = useCallback(({ x, y }, shiftKey) => {
        const drag = dragRef.current;
        if (!drag || !plan) return;
        const ctx = measureCtx();
        if (!ctx) return;

        if (drag.mode === "move") {
            const fx = Math.max(0, Math.min(1, drag.from.fx + (x - drag.start.x) / plan.W));
            const fy = Math.max(0, Math.min(1, drag.from.fy + (y - drag.start.y) / plan.H));
            setDoc((d) => updateLayer(d, drag.id, { fx: +fx.toFixed(4), fy: +fy.toFixed(4) }),
                { label: "move", group: `move:${drag.id}` });
            return;
        }

        if (drag.mode === "rotate") {
            let deg = angleAt(drag.center, x, y) + drag.offset;
            // Shift snaps to 15°, the increment people actually want when they
            // are straightening something rather than tilting it.
            if (shiftKey) deg = Math.round(deg / 15) * 15;
            setDoc((d) => updateLayer(d, drag.id, { rot: +(((deg % 360) + 360) % 360).toFixed(2) }),
                { label: "rotate", group: `rotate:${drag.id}` });
            return;
        }

        if (drag.mode === "resize") {
            const layer = plan.layers.find((l) => l.id === drag.id);
            if (!layer) return;
            const k = scaleFromDrag(drag.handle, drag.start, { x, y });
            // Word art has no stored width — its width is a consequence of the
            // text — so it resizes by font size and needs its own patch.
            const patch = layer.type === "wordart"
                ? wordArtResize(layer, k)
                : resizePatch(ctx, plan, layer, drag.handle.kind, k);
            if (!patch) return;
            const { id, ...change } = patch;
            setDoc((d) => updateLayer(d, drag.id, change), { label: "resize", group: `resize:${drag.id}` });
        }
    }, [plan, setDoc, measureCtx]);

    /**
     * Drag steps are coalesced to ONE PER ANIMATION FRAME.
     *
     * A pointer reports faster than a screen refreshes — a 120Hz trackpad
     * delivers roughly twice as many moves as there are frames to show them in,
     * and a burst can deliver several at once. Each one used to recompile the
     * entire plan (laying out every caption line, re-measuring every word-art
     * headline, resolving every fraction to pixels) and repaint the whole
     * frame, so most of that work was for a picture nobody would ever see.
     *
     * Keeping only the LATEST position is right rather than merely cheap: an
     * intermediate point during a drag has no meaning of its own, and the one
     * that matters is where the pointer is now. Undo is unaffected — the whole
     * gesture already coalesces into one history entry through its `group`.
     */
    const pendingRef = useRef({ raf: 0, point: null, shiftKey: false });

    const flushDrag = useCallback(() => {
        const pending = pendingRef.current;
        pending.raf = 0;
        if (pending.point) applyDrag(pending.point, pending.shiftKey);
    }, [applyDrag]);

    const onPointerMove = useCallback((event) => {
        if (!dragRef.current) return;
        const pending = pendingRef.current;
        // Read off the event NOW: the frame that consumes this runs later, and
        // by then the event is long gone.
        pending.point = toCanvas(event);
        pending.shiftKey = event.shiftKey;
        if (pending.raf) return;
        pending.raf = requestAnimationFrame(flushDrag);
    }, [toCanvas, flushDrag]);

    const onPointerUp = useCallback((event) => {
        if (dragRef.current) {
            const pending = pendingRef.current;
            // FLUSH rather than cancel. A drop lands between frames more often
            // than not, so discarding the queued step would lose the last few
            // pixels of every drag — and it would be the pixels someone was
            // aiming with.
            if (pending.raf) {
                cancelAnimationFrame(pending.raf);
                pending.raf = 0;
                if (pending.point) applyDrag(pending.point, pending.shiftKey);
            }
            pending.point = null;
            dragRef.current = null;
            // The gesture is over, so the next one is its own undo step.
            endGesture();
        }
        canvasRef.current?.releasePointerCapture?.(event.pointerId);
    }, [endGesture, applyDrag]);

    // A frame queued when the editor unmounts would call into a dead closure.
    useEffect(() => () => {
        if (pendingRef.current.raf) cancelAnimationFrame(pendingRef.current.raf);
    }, []);

    /** The selection outline and handles, in canvas coordinates, for an overlay. */
    const selection = useMemo(() => {
        if (!plan || !selectedId) return null;
        const layer = plan.layers.find((l) => l.id === selectedId);
        if (!layer) return null;
        // The measuring context again — the box has to be the one the painter
        // will draw into, and that is decided by how the text measures.
        const ctx = measureCtx();
        if (!ctx) return null;
        const box = layerBounds(ctx, plan, layer);
        if (!box) return null;
        // layerHandles answers { bounds, center, handles } — the overlay wants
        // the points, and the rotate gesture wants the centre.
        const h = layerHandles(ctx, plan, layer);
        return { layer, box, handles: h?.handles || [], center: h?.center || null };
    }, [plan, selectedId, time, doc, measureCtx]);

    // ── keyboard ─────────────────────────────────────────────────────────────

    /**
     * The editing keyboard — what every graphics tool trains a hand to expect.
     *
     * The usability pass added everything after undo/redo: an editor where
     * Delete does nothing and the arrows scroll the page instead of nudging the
     * selection reads as broken to anybody who has used any other editor, and
     * the inspector's trash button is no substitute — deleting is done LOOKING
     * AT THE CANVAS, not at a panel.
     *
     * Nudges step in FRACTIONS (1% of the frame, 5% with shift) because that is
     * what the geometry is stored in — a pixel step would mean a different
     * distance on every aspect. They coalesce under one undo group, so tapping
     * an arrow eight times is one ctrl-Z, the same bargain a drag makes.
     *
     * Layers without their own anchor — a sound, a cover-the-frame clip — are
     * skipped rather than given one: writing fx/fy onto a clip is exactly the
     * edit that silently turns "fills the frame" into an inset.
     */
    useEffect(() => {
        function onKey(event) {
            const typing = /^(INPUT|TEXTAREA|SELECT)$/.test(event.target?.tagName || "")
                || event.target?.isContentEditable;
            if (typing) return;
            const mod = event.ctrlKey || event.metaKey;
            if (mod && event.key.toLowerCase() === "z") {
                event.preventDefault();
                if (event.shiftKey) redo(); else undo();
                return;
            }
            if (mod && event.key.toLowerCase() === "y") { event.preventDefault(); redo(); return; }

            if (event.key === "Escape") { setSelectedId(null); return; }
            if (!selectedId) return;

            if (event.key === "Delete" || event.key === "Backspace") {
                event.preventDefault();
                setDoc((d) => removeLayer(d, selectedId), { label: "delete" });
                setSelectedId(null);
                return;
            }

            const nudge = {
                ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1],
            }[event.key];
            if (!nudge) return;
            const patch = patchOf(doc, selectedId);
            if (!patch || patch.fx == null || patch.fy == null) return;
            event.preventDefault();
            const step = event.shiftKey ? 0.05 : 0.01;
            const clamp01 = (v) => Math.max(0, Math.min(1, v));
            setDoc((d) => updateLayer(d, selectedId, {
                fx: +clamp01((patchOf(d, selectedId)?.fx ?? patch.fx) + nudge[0] * step).toFixed(4),
                fy: +clamp01((patchOf(d, selectedId)?.fy ?? patch.fy) + nudge[1] * step).toFixed(4),
            }), { label: "move", group: `nudge:${selectedId}` });
        }
        window.addEventListener("keydown", onKey);
        return () => window.removeEventListener("keydown", onKey);
    }, [undo, redo, selectedId, setDoc, doc]);

    // ── frame size ───────────────────────────────────────────────────────────

    /** The pixel size this creative renders at — what the canvas is sized to. */
    const frame = useMemo(() => {
        if (doc.kind === "video") {
            const a = ASPECTS[doc.options?.aspect || DEFAULTS.aspect] || ASPECTS.vertical;
            return { width: a.width, height: a.height, label: a.label };
        }
        const s = typeof doc.size === "string" ? IMAGE_SIZES[doc.size] : doc.size;
        const dims = s && s.width ? s : IMAGE_SIZES.square;
        return { width: dims.width, height: dims.height, label: dims.label || `${dims.width}×${dims.height}` };
    }, [doc.kind, doc.size, doc.options]);

    return {
        doc, setDoc, setDocState,
        selectedId, setSelectedId,
        selectedPatch: selectedId ? patchOf(doc, selectedId) : null,
        selectedLayer: selection?.layer || null,
        plan, canvasRef, attachCanvas, repaint, frame, selection,
        images, videos, loading, loadError,
        time, setTime,
        insertAsk, askInsert,
        undo, redo, history, endGesture,
        onPointerDown, onPointerMove, onPointerUp,
        THEMES, ASPECTS, IMAGE_SIZES,
    };
}
