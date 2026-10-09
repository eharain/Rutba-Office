import { useCallback, useEffect, useLayoutEffect, useRef, useState } from "react";
import { clampZoom, fitScale, stepZoom, zoomLabel, scrollToKeep, zoomKey } from "../../lib/zoom";

/** The space around the frame inside the stage, in screen pixels. */
const PAD = 16;

/**
 * The canvas, and the selection drawn over it.
 *
 * THE OVERLAY IS NOT PAINTED ONTO THE CANVAS. It would be easier — paintFrame
 * already has the context — and it would be wrong twice over: the selection
 * would be recorded into the video, and it would be rasterised at frame
 * resolution, so a handle would be four pixels wide on screen for a 1080-wide
 * creative. It is an SVG sibling at the canvas' on-screen size instead, which
 * also means handles stay a constant, clickable size at any zoom.
 *
 * THE CANVAS IS ALWAYS AT FULL RENDER SIZE. It is never sized to what fits on
 * screen, because the plan compiles its geometry to the canvas' pixels: a
 * preview canvas at half size would need a second compile to look right, and
 * the two would drift. CSS scales it down for display, so what you see is the
 * real frame, just smaller.
 *
 * ZOOM AND PAN are that same CSS size, larger or smaller. The stage opens on
 * Fit; the − and + buttons, Ctrl/Cmd with − + 0, and Ctrl/Cmd with the wheel
 * (about the pointer) change it, and past Fit the stage scrolls. Hold Space
 * and drag, or drag with the middle button, to pan. Pointer gestures on the
 * canvas need nothing from here: useCreative maps them through the canvas'
 * on-screen box, which already is the zoomed one. The arithmetic is
 * `lib/zoom.js`.
 */
export default function Stage({
    attachCanvas, frame, selection, loading, loadError,
    onPointerDown, onPointerMove, onPointerUp,
    fit = "contain", children,
}) {
    const wrapRef = useRef(null);
    const scrollRef = useRef(null);
    const [fitted, setFitted] = useState(1);
    // Null is Fit: it follows the stage's size. A number is a chosen zoom.
    const [zoom, setZoom] = useState(null);
    const scale = zoom ?? fitted;
    const scaleRef = useRef(scale);
    scaleRef.current = scale;
    const fittedRef = useRef(fitted);
    fittedRef.current = fitted;
    // Where to scroll once a zoom has been drawn, so the point under the
    // pointer stays under it.
    const pendingScroll = useRef(null);
    const [spaceHeld, setSpaceHeld] = useState(false);
    const pan = useRef(null);
    // Space pans only while the pointer is over the stage, so it still
    // presses a focused button everywhere else.
    const over = useRef(false);

    // The Fit scale, remeasured whenever the box or the frame changes.
    // ResizeObserver rather than a window listener: the stage shrinks when the
    // inspector opens, and the window never resized.
    useEffect(() => {
        const el = wrapRef.current;
        if (!el) return undefined;
        const measure = () => {
            const box = el.getBoundingClientRect();
            if (!box.width || !box.height) return;
            const room = { width: box.width - 2 * PAD, height: box.height - 2 * PAD };
            setFitted(fit === "width" ? room.width / frame.width : fitScale(room, frame));
        };
        measure();
        const ro = new ResizeObserver(measure);
        ro.observe(el);
        return () => ro.disconnect();
    }, [frame.width, frame.height, fit]);

    /** Zoom to `next` (null for Fit), keeping `pointer` — scroller coordinates — still. */
    const zoomTo = useCallback((next, pointer) => {
        const el = scrollRef.current;
        const to = next == null ? fittedRef.current : clampZoom(next);
        if (el) {
            const at = pointer ?? { x: el.clientWidth / 2, y: el.clientHeight / 2 };
            pendingScroll.current = scrollToKeep(el, at, scaleRef.current, to, PAD);
        }
        setZoom(next == null ? null : to);
    }, []);

    useLayoutEffect(() => {
        const el = scrollRef.current;
        const want = pendingScroll.current;
        pendingScroll.current = null;
        if (!el || !want) return;
        el.scrollLeft = want.left;
        el.scrollTop = want.top;
    }, [scale]);

    // Ctrl/Cmd + wheel zooms about the pointer. Not a React prop: React's wheel
    // listener is passive, and the browser's own page zoom must be stopped.
    useEffect(() => {
        const el = scrollRef.current;
        if (!el) return undefined;
        const onWheel = (e) => {
            if (!(e.ctrlKey || e.metaKey)) return;
            e.preventDefault();
            const box = el.getBoundingClientRect();
            zoomTo(stepZoom(scaleRef.current, e.deltaY < 0 ? 1 : -1), { x: e.clientX - box.left, y: e.clientY - box.top });
        };
        el.addEventListener("wheel", onWheel, { passive: false });
        return () => el.removeEventListener("wheel", onWheel);
    }, [zoomTo]);

    // Ctrl/Cmd with + − 0, and Space held for panning — never while typing.
    useEffect(() => {
        const typing = (t) => /^(INPUT|TEXTAREA|SELECT)$/.test(t?.tagName || "") || t?.isContentEditable;
        const onDown = (e) => {
            if (typing(e.target)) return;
            const ask = zoomKey(e);
            if (ask) {
                e.preventDefault();
                if (ask === "fit") zoomTo(null);
                else zoomTo(stepZoom(scaleRef.current, ask === "in" ? 1 : -1));
                return;
            }
            if (e.key === " " && !e.repeat && over.current) { e.preventDefault(); setSpaceHeld(true); }
        };
        const onUp = (e) => { if (e.key === " ") setSpaceHeld(false); };
        const onBlur = () => setSpaceHeld(false);
        window.addEventListener("keydown", onDown);
        window.addEventListener("keyup", onUp);
        window.addEventListener("blur", onBlur);
        return () => {
            window.removeEventListener("keydown", onDown);
            window.removeEventListener("keyup", onUp);
            window.removeEventListener("blur", onBlur);
        };
    }, [zoomTo]);

    // Panning: Space + drag, or the middle button. Caught before the canvas
    // sees the press, so a pan never starts a move or a selection.
    const onPanDown = (e) => {
        if (!(spaceHeld || e.button === 1)) return;
        const el = scrollRef.current;
        if (!el) return;
        e.preventDefault();
        e.stopPropagation();
        pan.current = { x: e.clientX, y: e.clientY, left: el.scrollLeft, top: el.scrollTop, id: e.pointerId };
        el.setPointerCapture?.(e.pointerId);
    };
    const onPanMove = (e) => {
        const p = pan.current;
        const el = scrollRef.current;
        if (!p || !el) return;
        e.stopPropagation();
        el.scrollLeft = p.left - (e.clientX - p.x);
        el.scrollTop = p.top - (e.clientY - p.y);
    };
    const onPanUp = (e) => {
        if (!pan.current) return;
        e.stopPropagation();
        scrollRef.current?.releasePointerCapture?.(pan.current.id);
        pan.current = null;
    };

    const shownW = frame.width * scale;
    const shownH = frame.height * scale;

    return (
        <div ref={wrapRef} className="studio-stage" style={{ position: "relative", padding: 0, overflow: "hidden", minHeight: 240 }}
            onPointerEnter={() => { over.current = true; }}
            onPointerLeave={() => { over.current = false; }}>
            <div
                ref={scrollRef}
                className="studio-stage-scroll"
                style={{
                    position: "absolute", inset: 0, overflow: "auto", display: "flex", padding: PAD,
                    cursor: spaceHeld ? (pan.current ? "grabbing" : "grab") : undefined,
                }}
                onPointerDownCapture={onPanDown}
                onPointerMoveCapture={onPanMove}
                onPointerUpCapture={onPanUp}
                onPointerCancelCapture={onPanUp}
            >
            <div className="studio-stage-inner" style={{ width: shownW, height: shownH, margin: "auto", flex: "none" }}>
                <canvas
                    ref={attachCanvas}
                    width={frame.width}
                    height={frame.height}
                    className="studio-canvas"
                    style={{ width: shownW, height: shownH, cursor: spaceHeld ? "inherit" : undefined }}
                    onPointerDown={onPointerDown}
                    onPointerMove={onPointerMove}
                    onPointerUp={onPointerUp}
                    onPointerCancel={onPointerUp}
                />

                {selection && (
                    <svg
                        className="studio-overlay"
                        width={shownW}
                        height={shownH}
                        viewBox={`0 0 ${frame.width} ${frame.height}`}
                        // The overlay must never eat the pointer — every gesture
                        // belongs to the canvas underneath, which is the only
                        // thing that can hit-test against the compiled plan.
                        style={{ pointerEvents: "none" }}
                    >
                        <g transform={selection.layer.rot
                            ? `rotate(${selection.layer.rot} ${selection.box.x + selection.box.w / 2} ${selection.box.y + selection.box.h / 2})`
                            : undefined}
                        >
                            <rect
                                x={selection.box.x} y={selection.box.y}
                                width={selection.box.w} height={selection.box.h}
                                fill="none" stroke="#22d3ee"
                                // Divided by the scale so the outline is one
                                // screen pixel at any zoom, not one frame pixel.
                                strokeWidth={1.5 / scale}
                                strokeDasharray={`${6 / scale} ${4 / scale}`}
                            />
                        </g>
                        {selection.handles.map((h) => (
                            h.kind === "rotate" ? (
                                <circle key={h.kind} cx={h.x} cy={h.y} r={7 / scale}
                                    fill="#22d3ee" stroke="#0b1220" strokeWidth={1.5 / scale} />
                            ) : (
                                <rect key={h.kind}
                                    x={h.x - 6 / scale} y={h.y - 6 / scale}
                                    width={12 / scale} height={12 / scale}
                                    fill="#ffffff" stroke="#22d3ee" strokeWidth={1.5 / scale} />
                            )
                        ))}
                    </svg>
                )}

                {loading && (
                    <div className="studio-stage-veil">
                        <div className="spinner-border spinner-border-sm text-info me-2" role="status" />
                        Loading pictures…
                    </div>
                )}
                {loadError && !loading && (
                    <div className="studio-stage-veil text-danger">
                        <i className="fa-solid fa-triangle-exclamation me-2" />
                        {loadError}
                    </div>
                )}
                {children}
            </div>
            </div>

            {/* The zoom bar sits over the stage, outside the scroller, so it
                stays put while the frame scrolls under it. */}
            <div className="btn-group btn-group-sm" role="group" aria-label="Zoom"
                style={{ position: "absolute", right: 10, bottom: 10, zIndex: 2, opacity: 0.92 }}>
                <button type="button" className="btn btn-dark" title="Zoom out (Ctrl −)"
                    onClick={() => zoomTo(stepZoom(scale, -1))}>
                    <i className="fa-solid fa-magnifying-glass-minus" />
                </button>
                <button type="button" className="btn btn-dark" style={{ minWidth: 56 }}
                    title={zoom == null ? "Fitted to the stage" : "Fit the whole frame (Ctrl 0)"}
                    onClick={() => zoomTo(null)}>
                    {zoom == null ? `Fit ${zoomLabel(scale)}` : zoomLabel(scale)}
                </button>
                <button type="button" className="btn btn-dark" title="Zoom in (Ctrl +)"
                    onClick={() => zoomTo(stepZoom(scale, 1))}>
                    <i className="fa-solid fa-magnifying-glass-plus" />
                </button>
            </div>
        </div>
    );
}
