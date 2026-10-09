import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

/**
 * The timeline under the video: one lane per layer, the bar sized by duration
 * and placed by start time, against a ruler that spans the whole video.
 *
 * On the shared shelf (moved from the social app, 2026-08-24) because it is a
 * pure view over @rutba/video layer envelopes with no app dependency — the
 * social Video Studio drives it today, and any other editor over the same
 * engine (Studio's web app, when it lands) drives it the same way.
 *
 * Everything here is a VIEW onto the layer envelope. The component writes
 * exactly three things, always as layer patches through onPatch: `timing`
 * (drag the bar to move, drag an end to trim), `z` (drag a lane header to
 * reorder — the top lane paints on top), and `visible` (the eye). Sorting by
 * entry or exit time is a view, so it never silently changes what paints over
 * what — the reorder handle only works in the z view.
 *
 * Entry/exit ramps draw as wedges at the bar's ends, sized by their seconds,
 * so an overlap between two photos reads as a crossfade at a glance. The
 * playhead is shared with the preview scrubber: click or drag the ruler to
 * scrub, and the canvas repaints at that instant.
 */

const LANE_H = 26;
const HEADER_W = 156;

const TYPE_ICONS = {
    photo: "fa-image", video: "fa-film", image: "fa-star", text: "fa-font", caption: "fa-keyboard",
    qr: "fa-qrcode", sound: "fa-music", gradient: "fa-adjust", title: "fa-heading",
    outro: "fa-flag-checkered", progress: "fa-bars-progress", edges: "fa-circle-half-stroke",
    // Studio's own layer types (packages/editor-core/{wordart,shapes,adjust}.js).
    // They arrive through the renderer's layer-type registry, so this view has
    // to know them or every one of them draws as a nameless grey lane.
    wordart: "fa-wand-magic-sparkles", shape: "fa-shapes", adjust: "fa-sliders", icon: "fa-icons", textblock: "fa-align-left", chart: "fa-chart-simple",
};

const TYPE_COLORS = {
    photo: "#3f6fce", video: "#b03a48", image: "#8a5fbf", text: "#b5852c", caption: "#7a4fb0",
    qr: "#2c7a6e", sound: "#2f9e6e",
    gradient: "#5a5a66", title: "#5a5a66", outro: "#5a5a66", progress: "#5a5a66", edges: "#5a5a66",
    // Warm for the two that draw, cool for the one that only changes what is
    // already beneath it — an adjustment lane reads as different at a glance,
    // which is what its position in the stack is about.
    wordart: "#b5852c", shape: "#2c7a6e", adjust: "#4a6fa5", icon: "#a8632c", textblock: "#b5852c", chart: "#2c7a6e",
};

// Layer types a lane can duplicate — the appendable set. Chrome layers
// (gradient, title, outro, progress, edges) exist once by construction.
const DUPLICABLE = new Set(["photo", "video", "image", "text", "qr", "caption", "sound",
    "wordart", "shape", "adjust", "icon", "textblock", "chart"]);

const fmtT = (s) => {
    if (s >= 60) {
        const m = Math.floor(s / 60);
        return `${m}:${String(Math.floor(s % 60)).padStart(2, "0")}`;
    }
    return s % 1 === 0 ? `${s}s` : `${s.toFixed(1)}s`;
};

export default function VideoTimeline({
    plan,
    previewTime,
    busy,
    selectedLayerId,
    appendedIds, // Set<string> — patch-backed layers, the ones delete can remove
    addRow, // host-provided "add a layer" buttons, shown in the toolbar
    // Lanes that are NOT plan layers — audio the video really carries, but
    // that lives in the options rather than the layer stack (the music bed).
    // They draw and select like any lane; they cannot be dragged, trimmed,
    // hidden, duplicated or deleted, because none of that would have anywhere
    // to be written. Display only: these never reach renderVideo.
    extraLanes,
    onSelect,
    onScrub,
    onPatch,
    onRemove,
    onDuplicate,
}) {
    const [zoom, setZoom] = useState(1); // 1 = fit, 2, 4
    const [sortBy, setSortBy] = useState("z"); // 'z' | 'enter' | 'exit'
    const scrollRef = useRef(null);
    const trackRef = useRef(null);
    const dragRef = useRef(null); // { id, mode: 'move'|'trim-start'|'trim-end'|'scrub', x0, t0, orig }
    const [laneDrag, setLaneDrag] = useState(null); // dragged layer id (header reorder)

    const duration = plan?.duration || 1;

    // Window helpers — timing null means the whole video.
    const winOf = (l) => (l.timing ? l.timing : { start: 0, end: duration });

    const lanes = useMemo(() => {
        if (!plan) return [];
        const ls = [...plan.layers, ...(extraLanes || [])];
        if (sortBy === "enter") ls.sort((a, b) => winOf(a).start - winOf(b).start);
        else if (sortBy === "exit") ls.sort((a, b) => winOf(b).end - winOf(a).end);
        else ls.sort((a, b) => (b.z || 0) - (a.z || 0)); // top lane paints on top
        return ls;
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [plan, extraLanes, sortBy, duration]);

    // Ruler ticks at a step that keeps labels readable at any zoom.
    const [trackW, setTrackW] = useState(600);
    useEffect(() => {
        const el = scrollRef.current;
        if (!el) return;
        const measure = () => setTrackW(Math.max(200, el.clientWidth - HEADER_W));
        measure();
        const ro = new ResizeObserver(measure);
        ro.observe(el);
        return () => ro.disconnect();
    }, []);
    const pxPerSec = (trackW * zoom) / duration;
    const tickStep = [0.5, 1, 2, 5, 10, 30].find((s) => s * pxPerSec >= 56) || 60;
    const ticks = [];
    for (let t = 0; t <= duration + 1e-6; t += tickStep) ticks.push(+t.toFixed(3));

    const pct = (t) => `${(t / duration) * 100}%`;

    // ── scrub + bar dragging (one shared pointer machine) ──────────────────
    const timeAtClientX = useCallback((clientX) => {
        const track = trackRef.current;
        if (!track) return 0;
        const r = track.getBoundingClientRect();
        return Math.max(0, Math.min(duration, ((clientX - r.left) / r.width) * duration));
    }, [duration]);

    const beginDrag = (e, layer, mode) => {
        if (busy) return;
        e.preventDefault();
        e.stopPropagation();
        const w = winOf(layer);
        dragRef.current = { id: layer.id, mode, x0: e.clientX, orig: { ...w }, moved: false };
        e.currentTarget.setPointerCapture?.(e.pointerId);
    };

    const onTrackPointerDown = (e) => {
        // Empty track / ruler → scrub.
        dragRef.current = { mode: "scrub" };
        onScrub?.(timeAtClientX(e.clientX));
        e.currentTarget.setPointerCapture?.(e.pointerId);
    };

    // ── keyframe diamonds (drag to retime, double-click to delete) ─────────
    const beginKeyDrag = (e, layer, kt) => {
        if (busy) return;
        e.preventDefault();
        e.stopPropagation();
        const w = winOf(layer);
        dragRef.current = { id: layer.id, mode: "key", kt, barLen: w.end - w.start, keys: layer.keys, x0: e.clientX };
        e.currentTarget.setPointerCapture?.(e.pointerId);
    };
    const retimeKeys = (keys, from, to) => {
        const out = {};
        for (const [prop, list] of Object.entries(keys || {})) {
            out[prop] = list
                .map((k) => (Math.abs(k.t - from) < 0.02 ? { ...k, t: +to.toFixed(3) } : k))
                .sort((a, b) => a.t - b.t);
        }
        return out;
    };
    const deleteKeysAt = (layer, kt) => {
        if (busy) return;
        const out = {};
        let any = false;
        for (const [prop, list] of Object.entries(layer.keys || {})) {
            const rest = list.filter((k) => Math.abs(k.t - kt) > 0.02);
            if (rest.length) { out[prop] = rest; any = true; }
        }
        onPatch?.({ id: layer.id, keys: any ? out : null });
    };

    const onPointerMove = (e) => {
        const d = dragRef.current;
        if (!d) return;
        if (d.mode === "scrub") { onScrub?.(timeAtClientX(e.clientX)); return; }
        const track = trackRef.current;
        if (!track) return;
        const dt = ((e.clientX - d.x0) / track.getBoundingClientRect().width) * duration;
        if (d.mode === "key") {
            const to = Math.max(0, Math.min(d.barLen, d.kt + dt));
            onPatch?.({ id: d.id, keys: retimeKeys(d.keys, d.kt, to) });
            return;
        }
        if (Math.abs(e.clientX - d.x0) > 2) d.moved = true;
        const MIN = 0.2;
        let { start, end } = d.orig;
        if (d.mode === "move") {
            const len = end - start;
            start = Math.max(0, Math.min(duration - len, start + dt));
            end = start + len;
        } else if (d.mode === "trim-start") {
            start = Math.max(0, Math.min(end - MIN, start + dt));
        } else if (d.mode === "trim-end") {
            end = Math.min(duration, Math.max(start + MIN, end + dt));
        }
        onPatch?.({ id: d.id, timing: { start: +start.toFixed(3), end: +end.toFixed(3) } });
    };

    const onPointerUp = (e) => {
        const d = dragRef.current;
        dragRef.current = null;
        if (d && d.mode !== "scrub" && d.mode !== "key" && !d.moved) onSelect?.(d.id);
    };

    // ── lane reorder = z (only meaningful in the z view) ───────────────────
    // Dropping a lane on another takes that lane's place: the dragged layer's
    // new z is midway between the target and whatever sits above it.
    const dropOnLane = (targetId) => {
        if (!laneDrag || laneDrag === targetId || sortBy !== "z") { setLaneDrag(null); return; }
        // Extra lanes have no z to trade places with — dropping on one is a
        // no-op rather than a z computed against something that isn't stacked.
        const ordered = lanes.filter((l) => l.id !== laneDrag && !l.readOnly); // z-descending, without the dragged one
        const to = ordered.findIndex((l) => l.id === targetId);
        if (to < 0) { setLaneDrag(null); return; }
        const above = ordered[to - 1] || null;
        const below = ordered[to];
        const hi = above ? (above.z || 0) : (below.z || 0) + 20;
        const lo = below.z || 0;
        onPatch?.({ id: laneDrag, z: +((hi + lo) / 2).toFixed(3) });
        setLaneDrag(null);
    };

    if (!plan) return null;

    return (
        <div className="w-100 mt-2 border rounded" style={{ userSelect: "none", fontSize: 12 }}>
            {/* toolbar — the layer list's top line: sort views, ADD buttons, zoom */}
            <div className="d-flex align-items-center gap-2 px-2 py-1 border-bottom flex-wrap" style={{ background: "rgba(0,0,0,0.03)" }}>
                <span className="text-muted">Timeline</span>
                <div className="btn-group btn-group-sm ms-2" role="group" aria-label="sort">
                    {["z", "enter", "exit"].map((s) => (
                        <button key={s} type="button"
                            className={`btn py-0 px-2 ${sortBy === s ? "btn-secondary" : "btn-outline-secondary"}`}
                            title={s === "z" ? "Stack order — drag lane names to change what paints on top" : `Sort lanes by ${s} time (a view — the stack is untouched)`}
                            onClick={() => setSortBy(s)}>
                            {s === "z" ? "stack" : s}
                        </button>
                    ))}
                </div>
                {addRow && (
                    <div className="d-flex align-items-center gap-1 flex-wrap ms-2 border-start ps-2">
                        <span className="text-muted" style={{ fontSize: 11 }}>Add</span>
                        {addRow}
                    </div>
                )}
                <div className="btn-group btn-group-sm ms-auto" role="group" aria-label="zoom">
                    {[1, 2, 4].map((zf) => (
                        <button key={zf} type="button"
                            className={`btn py-0 px-2 ${zoom === zf ? "btn-secondary" : "btn-outline-secondary"}`}
                            onClick={() => setZoom(zf)}>
                            {zf === 1 ? "fit" : `${zf}×`}
                        </button>
                    ))}
                </div>
            </div>

            <div ref={scrollRef} style={{ overflowX: zoom > 1 ? "auto" : "hidden", overflowY: "hidden" }}>
                <div style={{ width: HEADER_W + trackW * zoom, minWidth: "100%" }}>
                    {/* ruler */}
                    <div className="d-flex" style={{ height: 22 }}>
                        <div className="flex-shrink-0 border-end text-muted px-2" style={{ width: HEADER_W, lineHeight: "22px" }}>
                            {fmtT(0)} → {fmtT(duration)}
                        </div>
                        <div ref={trackRef} className="position-relative flex-grow-1 border-bottom"
                            style={{ cursor: "col-resize", touchAction: "none" }}
                            onPointerDown={onTrackPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerUp}>
                            {ticks.map((t) => (
                                <div key={t} className="position-absolute text-muted" style={{ left: pct(t), top: 0, bottom: 0 }}>
                                    <div style={{ width: 1, height: 6, background: "currentColor", opacity: 0.5 }} />
                                    <div style={{ fontSize: 10, transform: "translateX(2px)", lineHeight: "14px" }}>{fmtT(t)}</div>
                                </div>
                            ))}
                            {/* start / finish markers */}
                            <div className="position-absolute" style={{ left: 0, top: 0, bottom: 0, width: 2, background: "#888" }} />
                            <div className="position-absolute" style={{ left: "100%", top: 0, bottom: 0, width: 2, background: "#888", transform: "translateX(-2px)" }} />
                        </div>
                    </div>

                    {/* lanes */}
                    <div className="position-relative">
                        {lanes.map((l) => {
                            const w = winOf(l);
                            const sel = selectedLayerId === l.id;
                            const hidden = l.visible === false;
                            const color = TYPE_COLORS[l.type] || "#5a5a66";
                            const barLen = Math.max(0.0001, w.end - w.start);
                            const enterFrac = l.enter?.seconds > 0 && l.enter.kind !== "none" && l.enter.kind !== "type-on"
                                ? Math.min(1, l.enter.seconds / barLen) : 0;
                            const exitFrac = l.exit?.seconds > 0 && l.exit.kind !== "none" && l.exit.kind !== "type-on"
                                ? Math.min(1, l.exit.seconds / barLen) : 0;
                            const name = l.name || l.text || l.id;
                            const appended = appendedIds?.has(l.id) || false;
                            const fixed = !!l.readOnly; // an extra lane: selectable, not editable here
                            return (
                                <div key={l.id} data-lane-id={l.id} className="d-flex border-bottom" style={{ height: LANE_H, opacity: hidden ? 0.45 : 1, background: sel ? "rgba(13,110,253,0.08)" : "transparent" }}>
                                    {/* header */}
                                    <div className={`flex-shrink-0 border-end d-flex align-items-center gap-1 px-1 ${sel ? "fw-bold" : ""}`}
                                        style={{ width: HEADER_W, cursor: "pointer" }}
                                        draggable={sortBy === "z" && !busy && !fixed}
                                        onDragStart={() => setLaneDrag(l.id)}
                                        onDragOver={(e) => { if (laneDrag) e.preventDefault(); }}
                                        onDrop={() => dropOnLane(l.id)}
                                        onClick={() => onSelect?.(sel ? null : l.id)}>
                                        {sortBy === "z" && !fixed && (
                                            <i className="fas fa-grip-vertical text-muted" style={{ fontSize: 9, cursor: "grab" }} title="Drag to reorder — the top lane paints on top" />
                                        )}
                                        {!fixed && (
                                            <button className="btn btn-sm p-0 border-0" disabled={busy}
                                                title={hidden ? "Show" : "Hide"}
                                                onClick={(e) => { e.stopPropagation(); onPatch?.({ id: l.id, visible: hidden }); }}>
                                                <i className={`fas ${hidden ? "fa-eye-slash text-muted" : "fa-eye"}`} style={{ fontSize: 10 }} />
                                            </button>
                                        )}
                                        <i className={`fas ${TYPE_ICONS[l.type] || "fa-square"} text-muted`} style={{ fontSize: 9 }} />
                                        <span className="text-truncate flex-grow-1" title={name}>{name}</span>
                                        {!fixed && DUPLICABLE.has(l.type) && (
                                            <button className="btn btn-sm p-0 border-0 text-muted" disabled={busy}
                                                title="Duplicate — same layer again, nudged later"
                                                onClick={(e) => { e.stopPropagation(); onDuplicate?.(l); }}>
                                                <i className="fas fa-clone" style={{ fontSize: 10 }} />
                                            </button>
                                        )}
                                        {appended && (
                                            <button className="btn btn-sm p-0 border-0 text-danger" disabled={busy}
                                                title="Delete this layer"
                                                onClick={(e) => { e.stopPropagation(); onRemove?.(l.id); }}>
                                                <i className="fas fa-trash" style={{ fontSize: 10 }} />
                                            </button>
                                        )}
                                    </div>
                                    {/* track */}
                                    <div className="position-relative flex-grow-1" style={{ touchAction: "none" }}
                                        onPointerDown={onTrackPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerUp}>
                                        <div className="position-absolute rounded"
                                            style={{
                                                left: pct(w.start), width: pct(barLen),
                                                top: 4, bottom: 4,
                                                background: color, opacity: hidden ? 0.4 : 0.9,
                                                outline: sel ? "2px solid #0d6efd" : "1px solid rgba(0,0,0,0.25)",
                                                cursor: fixed ? "pointer" : "grab", minWidth: 6,
                                                ...(fixed ? { backgroundImage: "repeating-linear-gradient(45deg, rgba(255,255,255,0.16) 0 6px, transparent 6px 12px)" } : {}),
                                            }}
                                            title={fixed
                                                ? `${name} — under the whole video; its properties are in the rail`
                                                : `${name}: ${fmtT(w.start)} → ${fmtT(w.end)}${l.timing ? "" : " (whole video)"}`}
                                            onPointerDown={(e) => (fixed ? onSelect?.(l.id) : beginDrag(e, l, "move"))}
                                            onPointerMove={onPointerMove} onPointerUp={onPointerUp}>
                                            {/* entry / exit wedges */}
                                            {enterFrac > 0 && (
                                                <div className="position-absolute" style={{
                                                    left: 0, top: 0, bottom: 0, width: `${enterFrac * 100}%`,
                                                    background: "rgba(255,255,255,0.45)",
                                                    clipPath: "polygon(0 100%, 100% 0, 100% 100%)",
                                                    pointerEvents: "none",
                                                }} />
                                            )}
                                            {exitFrac > 0 && (
                                                <div className="position-absolute" style={{
                                                    right: 0, top: 0, bottom: 0, width: `${exitFrac * 100}%`,
                                                    background: "rgba(255,255,255,0.45)",
                                                    clipPath: "polygon(0 0, 0 100%, 100% 100%)",
                                                    pointerEvents: "none",
                                                }} />
                                            )}
                                            {/* trim handles */}
                                            {!fixed && (
                                                <>
                                                    <div data-trim="start" className="position-absolute" style={{ left: -3, top: 0, bottom: 0, width: 8, cursor: "ew-resize" }}
                                                        onPointerDown={(e) => beginDrag(e, l, "trim-start")} onPointerMove={onPointerMove} onPointerUp={onPointerUp} />
                                                    <div data-trim="end" className="position-absolute" style={{ right: -3, top: 0, bottom: 0, width: 8, cursor: "ew-resize" }}
                                                        onPointerDown={(e) => beginDrag(e, l, "trim-end")} onPointerMove={onPointerMove} onPointerUp={onPointerUp} />
                                                </>
                                            )}
                                            {/* keyframe diamonds — the union of all keyed properties' times */}
                                            {sel && l.keys && [...new Set(
                                                Object.values(l.keys).flat().map((k) => +k.t.toFixed(3)),
                                            )].sort((a, b) => a - b).map((kt) => (
                                                <div key={kt} className="position-absolute"
                                                    title={`Key at ${fmtT(kt)} into the layer — drag to retime, double-click to delete`}
                                                    style={{
                                                        left: `${Math.min(100, (kt / barLen) * 100)}%`, top: "50%",
                                                        width: 9, height: 9, marginLeft: -4.5, marginTop: -4.5,
                                                        background: "#fff", outline: "1px solid rgba(0,0,0,0.6)",
                                                        transform: "rotate(45deg)", cursor: "ew-resize", zIndex: 2,
                                                    }}
                                                    onPointerDown={(e) => beginKeyDrag(e, l, kt)}
                                                    onPointerMove={onPointerMove} onPointerUp={onPointerUp}
                                                    onDoubleClick={(e) => { e.stopPropagation(); deleteKeysAt(l, kt); }} />
                                            ))}
                                        </div>
                                    </div>
                                </div>
                            );
                        })}

                        {/* playhead — spans every lane */}
                        <div className="position-absolute" style={{
                            left: `calc(${HEADER_W}px + ${(Math.min(previewTime, duration) / duration) * 100} * (100% - ${HEADER_W}px) / 100)`,
                            top: 0, bottom: 0, width: 2, background: "#dc3545", pointerEvents: "none",
                        }} />
                    </div>
                </div>
            </div>
        </div>
    );
}
