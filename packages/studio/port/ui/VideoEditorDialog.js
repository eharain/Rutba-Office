/**
 * VideoEditorDialog - trim a library video, optionally re-frame it, and hand the
 * result back as a File. The host owns the upload.
 *
 * There is no way to cut a video in a browser without re-encoding it: nothing
 * ships a muxer that could copy the compressed frames between the in and out
 * points. So the clip is PLAYED and re-recorded - MediaRecorder over the
 * element's own stream - which means an export takes as long as the clip does
 * and the tab has to stay in front. That is the same bargain the Video Studio
 * makes, and the progress bar says so.
 *
 * Two shapes of export, one skeleton:
 *   - a plain trim records the element's video track directly, so the picture is
 *     passed through at source resolution;
 *   - a re-frame (crop or resize) draws each frame into a canvas first and
 *     records that instead.
 * Audio comes from the element's captured stream either way, which is why the
 * element is never muted while a take with sound is running - a muted element
 * has historically captured silence.
 *
 * The bytes are fetched through the app's media transport (`fetchMedia`), not
 * assigned as a remote src, because a canvas that has drawn a cross-origin frame
 * is tainted and cannot be recorded at all.
 */
import React, { useCallback, useEffect, useRef, useState } from "react";
import {
    encodeBlocker, extForMime, fetchMediaBytes, fmtBytes, fmtPrecise, frameFit, safeStem, verifiedRecordMime,
} from "../lib/media-encode";

/**
 * Can this element's frames be read back? A video drawn from a cross-origin
 * source without CORS taints the canvas, and a tainted canvas - or a stream
 * captured from that element - cannot be recorded at all. Drawing two pixels
 * and reading one back is the only answer the browser gives honestly.
 */
/**
 * How long a clip really is, when its own header will not say.
 *
 * A webm written by MediaRecorder carries NO duration - and that is every clip
 * this app produces, from the recorder and from the Video Studio alike - so the
 * element reports `Infinity`. Seeking far past the end makes the browser go and
 * find the real end; the answer can only be read once the seek has LANDED,
 * because until then `currentTime` reads back the time that was asked for.
 *
 * Returns 0 when even that fails, which the caller treats as "not trimmable"
 * rather than as a zero-length clip - the difference between saying so and
 * silently exporting an empty file.
 */
function discoverDuration(el) {
    return new Promise((resolve) => {
        let settled = false;
        const finish = () => {
            if (settled) return;
            settled = true;
            el.removeEventListener("seeked", finish);
            el.removeEventListener("durationchange", onDuration);
            clearTimeout(timer);
            const found = Number.isFinite(el.duration) && el.duration > 0 ? el.duration : el.currentTime;
            const ok = Number.isFinite(found) && found > 0 && found < 1e5;
            try { el.currentTime = 0; } catch { /* the rewind is a courtesy */ }
            resolve(ok ? found : 0);
        };
        const onDuration = () => { if (Number.isFinite(el.duration) && el.duration > 0) finish(); };
        const timer = setTimeout(finish, 5000);
        el.addEventListener("seeked", finish);
        el.addEventListener("durationchange", onDuration);
        try { el.currentTime = 1e6; } catch { finish(); }
    });
}

function framesAreReadable(el) {
    try {
        const c = document.createElement("canvas");
        c.width = 2;
        c.height = 2;
        const ctx = c.getContext("2d");
        ctx.drawImage(el, 0, 0, 2, 2);
        ctx.getImageData(0, 0, 1, 1);
        return true;
    } catch {
        return false;
    }
}

const CROPS = [
    { k: "original", label: "As-is", ratio: null },
    { k: "9:16", label: "9:16 vertical", ratio: 9 / 16 },
    { k: "1:1", label: "1:1 square", ratio: 1 },
    { k: "16:9", label: "16:9 wide", ratio: 16 / 9 },
];

const SIZES = [
    { k: "original", label: "Source size", short: 0 },
    { k: "1080", label: "1080 (short edge)", short: 1080 },
    { k: "720", label: "720 (short edge)", short: 720 },
    { k: "480", label: "480 (short edge)", short: 480 },
];

/** The chosen framing, resolved against the source's real dimensions. */
const framing = (sw, sh, cropKey, sizeKey) => frameFit(
    sw, sh,
    (CROPS.find((c) => c.k === cropKey) || CROPS[0]).ratio,
    (SIZES.find((s) => s.k === sizeKey) || SIZES[0]).short,
);

export default function VideoEditorDialog({
    show = false,
    source,                    // { url, name, duration?, width?, height? }
    fetchMedia,                // (url) => Promise<Blob> - the app's media transport
    onClose,
    onSave,                    // (file, meta) => void | Promise<void>
    zIndex = 10600,
}) {
    const [blocked, setBlocked] = useState(null);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState(null);
    const [ready, setReady] = useState(false);
    const [duration, setDuration] = useState(0);
    const [dims, setDims] = useState({ w: 0, h: 0 });

    const [inAt, setInAt] = useState(0);
    const [outAt, setOutAt] = useState(0);
    const [head, setHead] = useState(0);
    const [playing, setPlaying] = useState(false);

    const [keepAudio, setKeepAudio] = useState(true);
    const [crop, setCrop] = useState("original");
    const [size, setSize] = useState("original");

    const [exporting, setExporting] = useState(false);
    const [progress, setProgress] = useState(0);
    const [result, setResult] = useState(null); // { blob, url, mimeType, seconds }
    const [outName, setOutName] = useState("");
    const [saving, setSaving] = useState(false);

    const videoRef = useRef(null);
    const canvasRef = useRef(null);
    const objectUrlRef = useRef(null);
    const resultUrlRef = useRef(null);
    const recRef = useRef(null);
    const rafRef = useRef(0);
    const stopAtRef = useRef(0);
    const capturedRef = useRef(null);  // the element's captured stream, held open between exports
    const canvasStreamRef = useRef(null);

    useEffect(() => { setBlocked(encodeBlocker()); }, []);

    /**
     * Get the clip playing in a form we are allowed to re-record.
     *
     * The media file server answers with `Access-Control-Allow-Origin: *` and
     * honours ranges, so the video is STREAMED straight from it whenever the
     * browser will let us read the frames back. That is the path that matters:
     * the app's proxy buffers a whole file in memory and refuses anything over
     * its cap, which most real library videos exceed. The proxy is the fallback
     * for the sources that need it - anything serving no CORS headers at all.
     */
    useEffect(() => {
        if (!show || !source?.url) return undefined;
        let dead = false;
        setLoading(true);
        setError(null);
        setReady(false);
        setResult(null);

        const direct = () => new Promise((resolve) => {
            const el = videoRef.current;
            if (!el) { resolve(false); return; }
            const settle = (ok) => {
                el.removeEventListener("loadeddata", onData);
                el.removeEventListener("error", onErr);
                clearTimeout(timer);
                resolve(ok);
            };
            const onData = () => settle(framesAreReadable(el));
            const onErr = () => settle(false);
            const timer = setTimeout(() => settle(false), 15000);
            el.addEventListener("loadeddata", onData);
            el.addEventListener("error", onErr);
            el.crossOrigin = "anonymous";
            el.src = source.url;
            el.load();
        });

        (async () => {
            try {
                if (await direct()) {
                    if (!dead) setLoading(false);
                    return;
                }
                if (dead) return;
                const blob = await fetchMediaBytes(source.url, fetchMedia);
                if (dead) return;
                const url = URL.createObjectURL(blob);
                objectUrlRef.current = url;
                const el = videoRef.current;
                if (el) { el.removeAttribute("crossorigin"); el.src = url; el.load(); }
            } catch (err) {
                if (dead) return;
                console.error("Failed to fetch the video for editing", err);
                setError(
                    "Could not load that video for editing. It is either too large for the media proxy "
                    + "(which is the fallback when a host sends no CORS headers) or the browser cannot decode it.",
                );
            } finally {
                if (!dead) setLoading(false);
            }
        })();
        return () => { dead = true; };
    }, [show, source?.url, fetchMedia]);

    const cleanup = useCallback(() => {
        if (rafRef.current) cancelAnimationFrame(rafRef.current);
        rafRef.current = 0;
        if (recRef.current && recRef.current.state !== "inactive") {
            try { recRef.current.stop(); } catch { /* already stopping */ }
        }
        recRef.current = null;
        // Capturing a media element keeps a live pipeline open - including, on
        // some machines, a hold on the audio device that makes the NEXT
        // getUserMedia (the recorder, in another dialog) sit and wait. Closing
        // the dialog has to let go of it.
        for (const s of [capturedRef.current, canvasStreamRef.current]) {
            if (!s) continue;
            for (const t of s.getTracks()) { try { t.stop(); } catch { /* already ended */ } }
        }
        capturedRef.current = null;
        canvasStreamRef.current = null;
        const el = videoRef.current;
        if (el) { try { el.pause(); } catch { /* gone */ } }
        if (objectUrlRef.current) { URL.revokeObjectURL(objectUrlRef.current); objectUrlRef.current = null; }
        if (resultUrlRef.current) { URL.revokeObjectURL(resultUrlRef.current); resultUrlRef.current = null; }
    }, []);

    useEffect(() => () => cleanup(), [cleanup]);
    useEffect(() => { if (!show) cleanup(); }, [show, cleanup]);

    useEffect(() => {
        if (show) setOutName(`${String(source?.name || "clip").replace(/\.[a-z0-9]+$/i, "")}-trim`);
    }, [show, source?.name]);

    const onLoaded = async () => {
        const el = videoRef.current;
        if (!el) return;
        // The media server's metadata can be missing or wrong; the decoded
        // element is the one source of truth for both length and size.
        let d = Number.isFinite(el.duration) && el.duration > 0 ? el.duration : 0;
        if (!d) d = await discoverDuration(el);
        if (!d) d = Number(source?.duration) || 0;
        setDuration(d);
        setDims({ w: el.videoWidth, h: el.videoHeight });
        setInAt(0);
        setOutAt(d);
        setHead(0);
        setReady(d > 0);
        if (!d) setError("This clip reports no length, so it cannot be trimmed here.");
    };

    const seekTo = (t) => new Promise((resolve) => {
        const el = videoRef.current;
        if (!el) { resolve(); return; }
        if (Math.abs(el.currentTime - t) < 0.02) { resolve(); return; }
        const timer = setTimeout(finish, 600); // a seek that never reports back must not hang the export
        function finish() { clearTimeout(timer); el.removeEventListener("seeked", finish); resolve(); }
        el.addEventListener("seeked", finish);
        el.currentTime = t;
    });

    const scrub = async (t) => {
        const el = videoRef.current;
        if (!el) return;
        el.pause();
        setPlaying(false);
        setHead(t);
        await seekTo(t);
    };

    const playPreview = async () => {
        const el = videoRef.current;
        if (!el) return;
        if (playing) { el.pause(); setPlaying(false); return; }
        if (el.currentTime < inAt - 0.05 || el.currentTime >= outAt - 0.02) await seekTo(inAt);
        el.muted = !keepAudio;
        stopAtRef.current = outAt;
        await el.play().catch(() => { /* a gesture is required and we have one */ });
        setPlaying(true);
        const step = () => {
            const now = el.currentTime;
            setHead(now);
            if (now >= stopAtRef.current - 0.02) { el.pause(); setPlaying(false); return; }
            rafRef.current = requestAnimationFrame(step);
        };
        rafRef.current = requestAnimationFrame(step);
    };

    // == export ==============================================
    const runExport = async () => {
        const el = videoRef.current;
        if (!el || !ready) return;
        const seconds = outAt - inAt;
        if (!(seconds > 0.05)) {
            setError("Nothing to export - the out point has to come after the in point.");
            return;
        }
        setExporting(true);
        setError(null);
        setProgress(0);
        setResult(null);
        if (resultUrlRef.current) { URL.revokeObjectURL(resultUrlRef.current); resultUrlRef.current = null; }
        try {
            el.pause();
            setPlaying(false);
            await seekTo(inAt);

            const capture = el.captureStream?.bind(el) || el.mozCaptureStream?.bind(el);
            if (!capture) throw new Error("This browser cannot capture a video element.");
            el.muted = !keepAudio;
            // Captured once and reused: the tracks are the element's own, so
            // stopping them between exports would leave a second one with
            // nothing to record. They are released when the dialog closes.
            const srcStream = capturedRef.current || capture();
            capturedRef.current = srcStream;

            const f = framing(dims.w, dims.h, crop, size);
            const reframe = f.w !== dims.w || f.h !== dims.h;
            let canvasStream = null;
            let videoTrack;
            if (reframe) {
                const c = canvasRef.current;
                c.width = f.w;
                c.height = f.h;
                // captureStream(0) means "a frame only when I ask", which is the
                // accurate mode - but an engine without requestFrame would then
                // be asked for frames it can never produce and write an empty
                // file. Fall back to rate-based capture there.
                canvasStream = c.captureStream(0);
                videoTrack = canvasStream.getVideoTracks()[0];
                if (!videoTrack || typeof videoTrack.requestFrame !== "function") {
                    canvasStream.getTracks().forEach((t) => t.stop());
                    canvasStream = c.captureStream(30);
                    videoTrack = canvasStream.getVideoTracks()[0];
                }
                canvasStreamRef.current = canvasStream;
            } else {
                videoTrack = srcStream.getVideoTracks()[0];
            }
            if (!videoTrack) throw new Error("No picture came back from that clip.");
            const audioTracks = keepAudio ? srcStream.getAudioTracks() : [];

            // Asked by encoding, not by asking - see verifiedRecordMime.
            const mimeType = await verifiedRecordMime("transcode", audioTracks.length > 0);
            const rec = new MediaRecorder(new MediaStream([videoTrack, ...audioTracks]), {
                ...(mimeType ? { mimeType } : {}),
                videoBitsPerSecond: Math.round(Math.min(12e6, Math.max(1.5e6, f.w * f.h * 30 * 0.07))),
                ...(audioTracks.length ? { audioBitsPerSecond: 128000 } : {}),
            });
            const chunks = [];
            rec.ondataavailable = (e) => { if (e.data?.size) chunks.push(e.data); };
            let pushed = 0;   // frames handed to the encoder
            let ticks = 0;    // turns of the capture loop
            const stopped = new Promise((resolve, reject) => {
                rec.onstop = resolve;
                rec.onerror = (e) => reject(e?.error || new Error("The recorder failed mid-export."));
            });
            recRef.current = rec;

            const ctx = reframe ? canvasRef.current.getContext("2d") : null;
            const manual = reframe && typeof videoTrack.requestFrame === "function";
            const drawOne = () => {
                if (!ctx) return;
                ctx.drawImage(el, f.sx, f.sy, f.sw, f.sh, 0, 0, f.w, f.h);
                if (manual) { try { videoTrack.requestFrame(); pushed++; } catch { /* track ended */ } }
            };

            drawOne();          // a first frame, so the file never opens on black
            rec.start(1000);
            await el.play();

            await new Promise((resolve) => {
                const step = () => {
                    const now = el.currentTime;
                    ticks++;
                    setHead(now);
                    setProgress(Math.min(1, (now - inAt) / seconds));
                    drawOne();
                    if (now >= outAt - 0.02 || el.ended) { resolve(); return; }
                    rafRef.current = requestAnimationFrame(step);
                };
                rafRef.current = requestAnimationFrame(step);
            });

            el.pause();
            rec.stop();
            await stopped;
            recRef.current = null;
            if (canvasStream) canvasStream.getTracks().forEach((t) => t.stop());

            const type = (rec.mimeType || mimeType || "video/webm").split(";")[0];
            const blob = new Blob(chunks, { type });
            // The encoder can fail without ever raising: no error, no data, just
            // a file with nothing in it. Uploading that would be worse than
            // saying so.
            if (!blob.size) {
                throw new Error(
                    `The browser produced an empty ${type} file - it accepted that format and then encoded nothing `
                    + `(${ticks} loop turns, ${reframe ? `${pushed} frames pushed` : "direct track"}, `
                    + `played to ${el.currentTime.toFixed(2)}s). Try again, or use Chrome/Edge.`,
                );
            }
            const url = URL.createObjectURL(blob);
            resultUrlRef.current = url;
            setResult({ blob, url, mimeType: type, seconds, width: f.w, height: f.h });
            setProgress(1);
        } catch (err) {
            console.error("Export failed", err);
            setError(err?.message || "The export failed.");
        } finally {
            setExporting(false);
        }
    };

    const save = async () => {
        if (!result) return;
        setSaving(true);
        setError(null);
        try {
            const file = new File([result.blob], `${safeStem(outName, "clip-trim")}${extForMime(result.mimeType)}`, { type: result.mimeType });
            await onSave?.(file, { seconds: result.seconds, width: result.width, height: result.height });
            onClose?.();
        } catch (err) {
            console.error("Failed to save the clip", err);
            setError(err?.response?.data?.error?.message || err?.message || "Could not save that clip.");
        } finally {
            setSaving(false);
        }
    };

    if (!show) return null;

    const busy = exporting || saving;
    const f = dims.w && dims.h ? framing(dims.w, dims.h, crop, size) : null;
    const len = Math.max(0, outAt - inAt);
    const pct = (t) => (duration > 0 ? (t / duration) * 100 : 0);

    // Deliberately no click-the-backdrop-to-close: an export that has been
    // running in real time for a minute must not end on a stray click.
    return (
        <div className="modal d-block" tabIndex={-1} style={{ background: "rgba(0,0,0,0.55)", zIndex }}>
            <div className="modal-dialog modal-lg modal-dialog-centered modal-dialog-scrollable">
                <div className="modal-content">
                    <div className="modal-header py-2">
                        <h6 className="modal-title mb-0 text-truncate">
                            <i className="fas fa-scissors me-2" />Edit “{source?.name}”
                        </h6>
                        <button type="button" className="btn-close" onClick={() => onClose?.()} disabled={busy} />
                    </div>

                    <div className="modal-body py-2">
                        {blocked && <div className="alert alert-danger py-2 small">{blocked}</div>}
                        {error && <div className="alert alert-danger py-2 small">{error}</div>}
                        {loading && (
                            <div className="text-center py-3">
                                <span className="spinner-border spinner-border-sm me-2" />Fetching the video…
                            </div>
                        )}

                        <div className={result ? "d-none" : "bg-dark rounded mb-2"}>
                            {/* Driven by the transport below rather than the browser's own
                                controls, so the in/out points and the playhead agree. */}
                            <video ref={videoRef} playsInline onLoadedMetadata={onLoaded}
                                className="w-100" style={{ maxHeight: 300, display: "block" }} />
                        </div>
                        <canvas ref={canvasRef} className="d-none" />

                        {result && (
                            <>
                                <video src={result.url} controls playsInline className="w-100 bg-dark rounded mb-2" style={{ maxHeight: 300 }} />
                                <div className="alert alert-success py-2 small mb-2">
                                    <i className="fas fa-check me-1" />
                                    {fmtPrecise(result.seconds)} · {result.width}×{result.height} · {fmtBytes(result.blob.size)}
                                    <button className="btn btn-sm btn-link p-0 ms-2" onClick={() => setResult(null)} disabled={busy}>
                                        back to the editor
                                    </button>
                                </div>
                            </>
                        )}

                        {ready && !result && (
                            <>
                                {/* == trim strip == */}
                                <div className="position-relative mb-2" style={{ height: 26, background: "#e9ecef", borderRadius: 4 }}>
                                    <div className="position-absolute" style={{
                                        left: `${pct(inAt)}%`, width: `${Math.max(0, pct(outAt) - pct(inAt))}%`,
                                        top: 0, bottom: 0, background: "rgba(13,110,253,0.35)", borderRadius: 4,
                                    }} />
                                    <div className="position-absolute bg-dark" style={{ left: `${pct(head)}%`, width: 2, top: 0, bottom: 0 }} />
                                    <input type="range" className="form-range position-absolute w-100" min={0} max={duration || 0} step={0.05}
                                        value={head} disabled={busy}
                                        onChange={(e) => scrub(Number(e.target.value))}
                                        style={{ top: 2, opacity: 0.35 }} />
                                </div>

                                <div className="d-flex flex-wrap align-items-center gap-2 mb-2">
                                    <button className="btn btn-sm btn-outline-primary" onClick={playPreview} disabled={busy}>
                                        <i className={`fas ${playing ? "fa-pause" : "fa-play"} me-1`} />Preview
                                    </button>
                                    <button className="btn btn-sm btn-outline-success" disabled={busy}
                                        onClick={() => setInAt(Math.min(head, outAt - 0.2))} title="Start the clip here">
                                        <i className="fas fa-right-to-bracket me-1" />Set in
                                    </button>
                                    <button className="btn btn-sm btn-outline-danger" disabled={busy}
                                        onClick={() => setOutAt(Math.max(head, inAt + 0.2))} title="End the clip here">
                                        <i className="fas fa-right-from-bracket me-1" />Set out
                                    </button>
                                    <button className="btn btn-sm btn-outline-secondary" disabled={busy}
                                        onClick={() => { setInAt(0); setOutAt(duration); }}>Whole clip</button>
                                    <span className="badge bg-light text-dark border ms-auto">
                                        in {fmtPrecise(inAt)} · out {fmtPrecise(outAt)} · keeps {fmtPrecise(len)}
                                    </span>
                                </div>

                                <div className="row g-2">
                                    <div className="col-6 col-md-3">
                                        <label className="form-label small mb-1">In (seconds)</label>
                                        <input type="number" className="form-control form-control-sm" value={Number(inAt.toFixed(2))}
                                            step={0.1} min={0} max={duration} disabled={busy}
                                            onChange={(e) => setInAt(Math.max(0, Math.min(Number(e.target.value) || 0, outAt - 0.2)))} />
                                    </div>
                                    <div className="col-6 col-md-3">
                                        <label className="form-label small mb-1">Out (seconds)</label>
                                        <input type="number" className="form-control form-control-sm" value={Number(outAt.toFixed(2))}
                                            step={0.1} min={0} max={duration} disabled={busy}
                                            onChange={(e) => setOutAt(Math.min(duration, Math.max(Number(e.target.value) || 0, inAt + 0.2)))} />
                                    </div>
                                    <div className="col-6 col-md-3">
                                        <label className="form-label small mb-1">Framing</label>
                                        <select className="form-select form-select-sm" value={crop} disabled={busy}
                                            onChange={(e) => setCrop(e.target.value)}>
                                            {CROPS.map((c) => <option key={c.k} value={c.k}>{c.label}</option>)}
                                        </select>
                                    </div>
                                    <div className="col-6 col-md-3">
                                        <label className="form-label small mb-1">Size</label>
                                        <select className="form-select form-select-sm" value={size} disabled={busy}
                                            onChange={(e) => setSize(e.target.value)}>
                                            {SIZES.map((s) => <option key={s.k} value={s.k}>{s.label}</option>)}
                                        </select>
                                    </div>
                                </div>

                                <div className="d-flex flex-wrap align-items-center gap-3 mt-2">
                                    <div className="form-check form-switch mb-0">
                                        <input className="form-check-input" type="checkbox" id="ve-audio" checked={keepAudio}
                                            disabled={busy} onChange={(e) => setKeepAudio(e.target.checked)} />
                                        <label className="form-check-label small" htmlFor="ve-audio">Keep the sound</label>
                                    </div>
                                    <span className="text-muted small">
                                        {dims.w}×{dims.h}{f && (f.w !== dims.w || f.h !== dims.h) ? ` → ${f.w}×${f.h}` : ""}
                                        {crop !== "original" ? " · centre crop" : ""}
                                    </span>
                                </div>
                            </>
                        )}

                        {exporting && (
                            <div className="mt-2">
                                <div className="progress" style={{ height: 8 }}>
                                    <div className="progress-bar progress-bar-striped progress-bar-animated"
                                        style={{ width: `${Math.round(progress * 100)}%` }} />
                                </div>
                                <small className="text-muted">
                                    Re-encoding in real time - {fmtPrecise(len)} to go through. Leave this tab in front.
                                </small>
                            </div>
                        )}
                    </div>

                    <div className="modal-footer py-2">
                        {result ? (
                            <>
                                <input className="form-control form-control-sm me-auto" style={{ maxWidth: 280 }} value={outName}
                                    disabled={saving} onChange={(e) => setOutName(e.target.value)} />
                                <button className="btn btn-sm btn-secondary" onClick={() => onClose?.()} disabled={saving}>Cancel</button>
                                <button className="btn btn-sm btn-primary" onClick={save} disabled={saving}>
                                    {saving ? <span className="spinner-border spinner-border-sm me-1" /> : <i className="fas fa-cloud-upload-alt me-1" />}
                                    Upload as a new video
                                </button>
                            </>
                        ) : (
                            <>
                                <span className="text-muted small me-auto">
                                    The original is never touched - an edit is always a new file.
                                </span>
                                <button className="btn btn-sm btn-secondary" onClick={() => onClose?.()} disabled={busy}>Cancel</button>
                                <button className="btn btn-sm btn-primary" onClick={runExport} disabled={!ready || busy || !!blocked}>
                                    {exporting ? <span className="spinner-border spinner-border-sm me-1" /> : <i className="fas fa-scissors me-1" />}
                                    {exporting ? "Exporting…" : "Export the clip"}
                                </button>
                            </>
                        )}
                    </div>
                </div>
            </div>
        </div>
    );
}
