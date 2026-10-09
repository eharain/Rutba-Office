/**
 * RecorderDialog - capture a take from this machine and hand it back as a File.
 *
 * One dialog covers both jobs because they are the same job: pick sources, watch
 * a level meter, press record, listen back, keep it. `mode="audio"` records a
 * voice-over or a bed; `mode="video"` records the webcam (or the screen). The
 * host owns what happens next - this component never uploads, so the audio
 * library, the video library and the Video Studio each keep their own idea of
 * where a recording belongs.
 *
 * THE ONE CAPTURE SURFACE, and deliberately so. Three products draw on it: the
 * social app (audio/video libraries, Video Studio), the comms shell (a meeting
 * take, uploaded to Drive and attached to the session) and Studio's web app
 * when it lands. New capture work belongs HERE rather than beside it - a second
 * getUserMedia path is a second set of the three hard-won rules below to get
 * wrong.
 *
 * WHAT IS NOT HERE, on purpose: this records LOCALLY. Publishing these same
 * sources as live tracks - microphone, camera and screen SHARING between
 * meeting participants - is comms PLAN M2's open item, blocked on LiveKit
 * token minting (consumer/comms/infra/livekit/README.md). When it lands it
 * takes the streams this file already knows how to open and hands them to a
 * room instead of to a MediaRecorder; it does not open its own.
 *
 * Three things here are hard-won rather than obvious:
 *
 *  - SYSTEM AUDIO comes from getDisplayMedia, not getUserMedia; the browser has
 *    no other way to hear what the machine is playing. It must be requested
 *    while the click that asked for it is still "recent" - so it is always the
 *    FIRST await in arm(), before the microphone is opened, or Chrome rejects it
 *    as untrusted. Its video track is kept alive but never recorded: dropping it
 *    ends the capture session in some Chrome builds, taking the audio with it.
 *
 *  - EVERY source is routed through one AudioContext, even a lone microphone.
 *    That is what makes "mic + system" a single track the recorder can take, and
 *    it gives the level meter one place to tap.
 *
 *  - AUDIO IS HANDED OVER AS WAV. MediaRecorder produces webm/opus, and the
 *    media file server types a file by its EXTENSION - a `.webm` full of audio
 *    comes back as `video/webm` and turns up in the video library. Decoding to
 *    WAV costs a second and makes the file honest everywhere. Safari records
 *    audio/mp4 and is left alone: `.m4a` is already unambiguous.
 */
import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { blobToWav } from "../lib/audio-edit";
import {
    baseMime, extForMime, fmtBytes, fmtClock, pickRecordMime, safeStem, stampNow,
} from "../lib/media-encode";

export default function RecorderDialog({
    show = false,
    mode = "audio",              // "audio" | "video"
    onClose,
    onRecorded,                  // (file, meta) => void | Promise<void> - the host uploads
    namePrefix,                  // default file-name stem
    useLabel,                    // button text, e.g. "Add to the library"
    maxSeconds = 900,            // hard stop, so a forgotten tab cannot eat all the memory
    zIndex = 10600,              // above the media-library modal (9999)
}) {
    const isVideo = mode === "video";

    const [support, setSupport] = useState(null); // null = not checked yet
    const [error, setError] = useState(null);
    const [notice, setNotice] = useState(null);

    // audio: mic | system | mix   ·   video: camera | screen
    const [source, setSource] = useState(isVideo ? "camera" : "mic");
    const [micId, setMicId] = useState("");
    const [camId, setCamId] = useState("");
    const [devices, setDevices] = useState({ mics: [], cams: [] });

    const [armed, setArmed] = useState(false);
    const [arming, setArming] = useState(false);
    const [recording, setRecording] = useState(false);
    const [paused, setPaused] = useState(false);
    const [elapsed, setElapsed] = useState(0);
    const [take, setTake] = useState(null);       // { blob, url, mimeType, seconds }
    const [name, setName] = useState("");
    const [saving, setSaving] = useState(false);

    const streamsRef = useRef([]);        // raw device streams, stopped on teardown
    const recordStreamRef = useRef(null); // what MediaRecorder is handed
    const audioCtxRef = useRef(null);
    const analyserRef = useRef(null);
    const recorderRef = useRef(null);
    const chunksRef = useRef([]);
    const liveRef = useRef(null);         // <video> showing the camera/screen
    const meterRef = useRef(null);        // level bar, written to directly
    const rafRef = useRef(0);
    const tickRef = useRef(0);
    const runRef = useRef({ at: 0, acc: 0 }); // ms: segment start + banked time
    const takeRef = useRef(null);         // for revoking the object url on teardown

    const wantsMic = isVideo ? micId !== "none" : (source === "mic" || source === "mix");

    useEffect(() => {
        setSupport(
            typeof window !== "undefined"
            && !!navigator?.mediaDevices?.getUserMedia
            && typeof MediaRecorder !== "undefined",
        );
    }, []);

    // == stream lifecycle ====================================
    const stopMeter = useCallback(() => {
        if (rafRef.current) cancelAnimationFrame(rafRef.current);
        rafRef.current = 0;
        if (meterRef.current) meterRef.current.style.width = "0%";
    }, []);

    const releaseStreams = useCallback(() => {
        stopMeter();
        for (const s of streamsRef.current) { for (const t of s.getTracks()) { try { t.stop(); } catch { /* gone */ } } }
        streamsRef.current = [];
        recordStreamRef.current = null;
        analyserRef.current = null;
        if (audioCtxRef.current) { try { audioCtxRef.current.close(); } catch { /* already closed */ } }
        audioCtxRef.current = null;
        if (liveRef.current) liveRef.current.srcObject = null;
        setArmed(false);
    }, [stopMeter]);

    const runMeter = useCallback(() => {
        const bins = new Uint8Array(1024);
        const step = () => {
            const an = analyserRef.current;
            if (an && meterRef.current) {
                an.getByteTimeDomainData(bins);
                let sum = 0;
                for (let i = 0; i < bins.length; i++) { const v = (bins[i] - 128) / 128; sum += v * v; }
                // A little headroom: speech rarely passes 0.3 RMS, so scale it up
                // rather than showing a permanently flat bar.
                const pct = Math.min(100, Math.round(Math.sqrt(sum / bins.length) * 260));
                meterRef.current.style.width = `${pct}%`;
                meterRef.current.className = `h-100 ${pct > 88 ? "bg-danger" : pct > 4 ? "bg-success" : "bg-secondary"}`;
            }
            rafRef.current = requestAnimationFrame(step);
        };
        rafRef.current = requestAnimationFrame(step);
    }, []);

    /**
     * Open the chosen sources and build the one stream that will be recorded.
     * `cfg` is passed explicitly rather than read from state so a device change
     * can arm the NEW choice inside the same click.
     */
    const arm = useCallback(async (cfg) => {
        const src = cfg.source;
        const wantDisplay = isVideo ? src === "screen" : (src === "system" || src === "mix");
        const wantMic = isVideo ? cfg.micId !== "none" : (src === "mic" || src === "mix");
        const wantCam = isVideo && src === "camera";

        releaseStreams();
        setError(null);
        setNotice(null);
        setArming(true);
        try {
            // Display capture FIRST - see the header note about transient activation.
            let display = null;
            if (wantDisplay) {
                display = await navigator.mediaDevices.getDisplayMedia({
                    video: isVideo ? { frameRate: { ideal: 30 } } : true,
                    audio: true,
                });
                streamsRef.current.push(display);
                if (!display.getAudioTracks().length) {
                    setNotice(isVideo
                        ? "That share carries no sound - re-pick it and tick “Also share tab/system audio” if you want the machine's audio too."
                        : "That share carries no sound. Re-pick it and tick “Share tab audio” (a tab) or “Share system audio” (the whole screen).");
                }
            }

            let user = null;
            if (wantMic || wantCam) {
                user = await navigator.mediaDevices.getUserMedia({
                    audio: wantMic ? (cfg.micId ? { deviceId: { exact: cfg.micId } } : true) : false,
                    video: wantCam
                        ? {
                            ...(cfg.camId ? { deviceId: { exact: cfg.camId } } : {}),
                            width: { ideal: 1280 }, height: { ideal: 720 },
                        }
                        : false,
                });
                streamsRef.current.push(user);
            }

            const inputs = [
                ...(user ? user.getAudioTracks() : []),
                ...(display ? display.getAudioTracks() : []),
            ];
            let audioTracks = [];
            if (inputs.length) {
                // One graph for everything: it mixes when there are two sources,
                // and gives the meter somewhere to listen when there is one.
                const Ctx = window.AudioContext || window.webkitAudioContext;
                const ctx = new Ctx();
                const dest = ctx.createMediaStreamDestination();
                const analyser = ctx.createAnalyser();
                analyser.fftSize = 2048;
                for (const t of inputs) {
                    const node = ctx.createMediaStreamSource(new MediaStream([t]));
                    node.connect(dest);
                    node.connect(analyser);
                }
                audioCtxRef.current = ctx;
                analyserRef.current = analyser;
                audioTracks = dest.stream.getAudioTracks();
            }

            const videoTrack = isVideo
                ? (wantCam ? user?.getVideoTracks()[0] : display?.getVideoTracks()[0])
                : null;
            if (isVideo && !videoTrack) throw new Error("No picture came back from that source.");

            const out = new MediaStream([...(videoTrack ? [videoTrack] : []), ...audioTracks]);
            if (!out.getTracks().length) throw new Error("Nothing to record - pick a microphone or a source that carries sound.");
            recordStreamRef.current = out;

            // A screen share has its own Stop button, and unplugging a webcam
            // ends its track. Either way the stream goes silent while the UI
            // still says "recording" - so end the take and say what happened.
            for (const t of [...(user ? user.getTracks() : []), ...(display ? display.getTracks() : [])]) {
                t.addEventListener("ended", () => {
                    if (recorderRef.current && recorderRef.current.state !== "inactive") {
                        try { recorderRef.current.stop(); } catch { /* already stopping */ }
                    }
                    if (tickRef.current) { clearInterval(tickRef.current); tickRef.current = 0; }
                    setArmed(false);
                    setNotice("The source stopped - sharing was ended, or the device went away. Anything already recorded is kept.");
                });
            }

            if (isVideo && liveRef.current) {
                // Muted, or the room hears itself through the speakers.
                liveRef.current.srcObject = new MediaStream(videoTrack ? [videoTrack] : []);
                liveRef.current.play().catch(() => { /* autoplay is best-effort */ });
            }
            if (analyserRef.current) runMeter();
            setArmed(true);

            // Labels only exist once a permission has been granted, so this is
            // the moment the device lists become useful. The same fact now
            // lives in hooks/useMediaDevices.js, which the meeting room uses;
            // folding this into it touches four products and wants its own
            // change rather than riding along with a meeting feature.
            try {
                const all = await navigator.mediaDevices.enumerateDevices();
                setDevices({
                    mics: all.filter((d) => d.kind === "audioinput"),
                    cams: all.filter((d) => d.kind === "videoinput"),
                });
            } catch { /* the dropdowns just stay on "Default" */ }
        } catch (err) {
            releaseStreams();
            const n = err?.name;
            setError(
                n === "NotAllowedError" ? "Permission was refused - allow the microphone/camera for this site and try again."
                    : n === "NotFoundError" ? "No such device - pick another one."
                        : n === "NotReadableError" ? "That device is busy - another app is probably using it."
                            : String(err?.message || err),
            );
        } finally {
            setArming(false);
        }
    }, [isVideo, releaseStreams, runMeter]);

    const clearTake = useCallback(() => {
        if (takeRef.current?.url) URL.revokeObjectURL(takeRef.current.url);
        takeRef.current = null;
        setTake(null);
    }, []);

    const teardown = useCallback(() => {
        if (recorderRef.current && recorderRef.current.state !== "inactive") {
            try { recorderRef.current.stop(); } catch { /* already stopping */ }
        }
        recorderRef.current = null;
        if (tickRef.current) clearInterval(tickRef.current);
        tickRef.current = 0;
        releaseStreams();
        clearTake();
        setRecording(false);
        setPaused(false);
        setElapsed(0);
        setNotice(null);
        setError(null);
    }, [releaseStreams, clearTake]);

    // Open/close: arm on the way in, let go of every device on the way out.
    useEffect(() => {
        if (!show) { teardown(); return; }
        if (support === false) return;
        if (support) {
            setName(`${namePrefix || (isVideo ? "clip" : "voice")}-${stampNow()}`);
            arm({ source, micId, camId });
        }
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [show, support]);

    useEffect(() => () => teardown(), [teardown]);

    // The live view unmounts while a take is being auditioned, so its element is
    // a NEW one when "Record again" puts it back - re-attach on the render that
    // creates it rather than from arm(), which may run before React gets there.
    useEffect(() => {
        if (!isVideo || take || !armed) return;
        const el = liveRef.current;
        const stream = recordStreamRef.current;
        if (!el || !stream) return;
        el.srcObject = new MediaStream(stream.getVideoTracks());
        el.play().catch(() => { /* autoplay is best-effort */ });
    }, [isVideo, take, armed]);

    // == recording ===========================================
    const startTick = () => {
        if (tickRef.current) clearInterval(tickRef.current);
        tickRef.current = setInterval(() => {
            const r = runRef.current;
            const ms = r.acc + (r.at ? Date.now() - r.at : 0);
            const secs = ms / 1000;
            setElapsed(secs);
            if (secs >= maxSeconds) stop();
        }, 200);
    };

    const start = () => {
        const stream = recordStreamRef.current;
        if (!stream) return;
        clearTake();
        chunksRef.current = [];
        const mimeType = pickRecordMime(isVideo ? "video" : "audio");
        let rec;
        try {
            rec = new MediaRecorder(stream, {
                ...(mimeType ? { mimeType } : {}),
                audioBitsPerSecond: 128000,
                ...(isVideo ? { videoBitsPerSecond: 2500000 } : {}),
            });
        } catch (err) {
            setError(`This browser refused to record: ${String(err?.message || err)}`);
            return;
        }
        rec.ondataavailable = (e) => { if (e.data?.size) chunksRef.current.push(e.data); };
        rec.onerror = (e) => setError(`Recording stopped: ${String(e?.error?.message || e?.error || "unknown error")}`);
        rec.onstop = () => {
            const type = baseMime(rec.mimeType) || (isVideo ? "video/webm" : "audio/webm");
            const blob = new Blob(chunksRef.current, { type });
            chunksRef.current = [];
            // An encoder can accept a format and then write nothing at all -
            // no error, no data. Better to say so than to offer an empty take.
            if (!blob.size) {
                setError(`The browser recorded nothing in ${type}. Try again, or use Chrome/Edge.`);
                setRecording(false);
                setPaused(false);
                runRef.current = { at: 0, acc: 0 };
                return;
            }
            const r = runRef.current;
            const seconds = Math.max(0.1, (r.acc + (r.at ? Date.now() - r.at : 0)) / 1000);
            runRef.current = { at: 0, acc: 0 };
            const t = { blob, url: URL.createObjectURL(blob), mimeType: type, seconds };
            takeRef.current = t;
            setTake(t);
            setRecording(false);
            setPaused(false);
        };
        recorderRef.current = rec;
        runRef.current = { at: Date.now(), acc: 0 };
        setElapsed(0);
        // A timeslice means a crash or an abrupt stop still leaves usable chunks.
        rec.start(1000);
        setRecording(true);
        setPaused(false);
        startTick();
    };

    const pause = () => {
        const rec = recorderRef.current;
        if (!rec || rec.state !== "recording") return;
        rec.pause();
        runRef.current = { at: 0, acc: runRef.current.acc + (Date.now() - runRef.current.at) };
        setPaused(true);
    };

    const resume = () => {
        const rec = recorderRef.current;
        if (!rec || rec.state !== "paused") return;
        rec.resume();
        runRef.current = { ...runRef.current, at: Date.now() };
        setPaused(false);
    };

    const stop = () => {
        const rec = recorderRef.current;
        if (tickRef.current) { clearInterval(tickRef.current); tickRef.current = 0; }
        if (!rec || rec.state === "inactive") return;
        try { rec.stop(); } catch { /* already stopped */ }
    };

    const changeSource = (next) => {
        if (recording) return;
        setSource(next);
        arm({ source: next, micId, camId });
    };
    const changeMic = (next) => {
        if (recording) return;
        setMicId(next);
        arm({ source, micId: next, camId });
    };
    const changeCam = (next) => {
        if (recording) return;
        setCamId(next);
        arm({ source, micId, camId: next });
    };

    // == hand over ===========================================
    const use = async () => {
        if (!take) return;
        setSaving(true);
        setError(null);
        try {
            let blob = take.blob;
            let type = take.mimeType;
            if (!isVideo && baseMime(type) !== "audio/mp4" && baseMime(type) !== "audio/wav") {
                try {
                    blob = await blobToWav(take.blob);
                    type = "audio/wav";
                } catch (err) {
                    // Keeping the original beats losing the take; the file is
                    // still playable, just named after its real container.
                    console.warn("Could not re-encode the recording as WAV - keeping the original", err);
                }
            }
            const file = new File([blob], `${safeStem(name, `${isVideo ? "clip" : "voice"}-${stampNow()}`)}${extForMime(type)}`, { type: baseMime(type) });
            await onRecorded?.(file, { seconds: take.seconds, source, mimeType: baseMime(type), mode });
            teardown();
            onClose?.();
        } catch (err) {
            console.error("Failed to keep the recording", err);
            setError(err?.response?.data?.error?.message || err?.message || "Could not save that recording.");
        } finally {
            setSaving(false);
        }
    };

    const close = () => { if (!saving) { teardown(); onClose?.(); } };

    const sources = useMemo(() => (isVideo
        ? [
            { k: "camera", icon: "fa-video", label: "Camera" },
            { k: "screen", icon: "fa-desktop", label: "Screen" },
        ]
        : [
            { k: "mic", icon: "fa-microphone", label: "Microphone" },
            { k: "system", icon: "fa-volume-high", label: "System audio" },
            { k: "mix", icon: "fa-sliders", label: "Mic + system" },
        ]), [isVideo]);

    if (!show) return null;

    const busy = recording || arming || saving;

    return (
        <div className="modal d-block" tabIndex={-1} style={{ background: "rgba(0,0,0,0.55)", zIndex }}
            onClick={close}>
            <div className={`modal-dialog ${isVideo ? "modal-lg" : ""} modal-dialog-centered modal-dialog-scrollable`}
                onClick={(e) => e.stopPropagation()}>
                <div className="modal-content">
                    <div className="modal-header py-2">
                        <h6 className="modal-title mb-0">
                            <i className={`fas ${isVideo ? "fa-video" : "fa-microphone"} me-2`} />
                            {isVideo ? "Record video" : "Record audio"}
                        </h6>
                        <button type="button" className="btn-close" onClick={close} disabled={saving} />
                    </div>

                    <div className="modal-body py-2">
                        {support === false && (
                            <div className="alert alert-danger py-2 small mb-0">
                                This browser cannot record here. Recording needs a modern Chrome, Edge or Firefox
                                <strong> on a secure origin</strong> - https, or localhost. Over plain http on a LAN address the
                                browser hides the microphone and camera entirely.
                            </div>
                        )}

                        {support && (
                            <>
                                {/* == sources == */}
                                <div className="d-flex flex-wrap align-items-center gap-2 mb-2">
                                    <div className="btn-group btn-group-sm" role="group">
                                        {sources.map((s) => (
                                            <button key={s.k} type="button"
                                                className={`btn ${source === s.k ? "btn-primary" : "btn-outline-primary"}`}
                                                disabled={busy} onClick={() => changeSource(s.k)}>
                                                <i className={`fas ${s.icon} me-1`} />{s.label}
                                            </button>
                                        ))}
                                    </div>
                                    {arming && <span className="spinner-border spinner-border-sm text-muted" />}
                                </div>

                                <div className="row g-2 mb-2">
                                    {isVideo && source === "camera" && (
                                        <div className="col-md-6">
                                            <label className="form-label small mb-1">Camera</label>
                                            <select className="form-select form-select-sm" value={camId} disabled={busy}
                                                onChange={(e) => changeCam(e.target.value)}>
                                                <option value="">Default camera</option>
                                                {devices.cams.map((d, i) => (
                                                    <option key={d.deviceId || i} value={d.deviceId}>{d.label || `Camera ${i + 1}`}</option>
                                                ))}
                                            </select>
                                        </div>
                                    )}
                                    {(wantsMic || isVideo) && (
                                        <div className={isVideo ? "col-md-6" : "col-12"}>
                                            <label className="form-label small mb-1">Microphone</label>
                                            <select className="form-select form-select-sm" value={micId} disabled={busy}
                                                onChange={(e) => changeMic(e.target.value)}>
                                                {isVideo && <option value="none">No microphone</option>}
                                                <option value="">Default microphone</option>
                                                {devices.mics.map((d, i) => (
                                                    <option key={d.deviceId || i} value={d.deviceId}>{d.label || `Microphone ${i + 1}`}</option>
                                                ))}
                                            </select>
                                        </div>
                                    )}
                                </div>

                                {!isVideo && source !== "mic" && (
                                    <p className="text-muted small mb-2">
                                        System audio is captured through the screen-share picker: choose a tab or the whole screen
                                        and tick the <strong>share audio</strong> box. The picture is thrown away - only the sound is recorded.
                                    </p>
                                )}

                                {error && <div className="alert alert-danger py-2 small">{error}</div>}
                                {notice && <div className="alert alert-warning py-2 small">{notice}</div>}

                                {/* == the take, or the live view == */}
                                {take ? (
                                    isVideo
                                        ? <video src={take.url} controls playsInline className="w-100 bg-dark rounded mb-2" style={{ maxHeight: 320 }} />
                                        : <audio src={take.url} controls className="w-100 mb-2" />
                                ) : (
                                    isVideo && (
                                        <div className="bg-dark rounded mb-2 d-flex align-items-center justify-content-center" style={{ minHeight: 220 }}>
                                            <video ref={liveRef} muted playsInline autoPlay
                                                className="w-100" style={{ maxHeight: 320, transform: source === "camera" ? "scaleX(-1)" : "none" }} />
                                        </div>
                                    )
                                )}

                                {/* == level meter == */}
                                <div className="d-flex align-items-center gap-2 mb-2">
                                    <i className="fas fa-wave-square text-muted" />
                                    <div className="progress flex-grow-1" style={{ height: 8 }}>
                                        <div ref={meterRef} className="h-100 bg-secondary" style={{ width: "0%", transition: "width .06s linear" }} />
                                    </div>
                                    <span className="badge bg-light text-dark border" style={{ minWidth: 62 }}>
                                        {fmtClock(take ? take.seconds : elapsed)}
                                    </span>
                                </div>

                                {/* == transport == */}
                                <div className="d-flex flex-wrap align-items-center gap-2">
                                    {!recording && !take && (
                                        <button className="btn btn-sm btn-danger" disabled={!armed || arming} onClick={start}>
                                            <i className="fas fa-circle me-1" />Record
                                        </button>
                                    )}
                                    {recording && (
                                        <>
                                            <button className="btn btn-sm btn-outline-secondary" onClick={paused ? resume : pause}>
                                                <i className={`fas ${paused ? "fa-play" : "fa-pause"} me-1`} />{paused ? "Resume" : "Pause"}
                                            </button>
                                            <button className="btn btn-sm btn-dark" onClick={stop}>
                                                <i className="fas fa-stop me-1" />Stop
                                            </button>
                                            <span className="badge bg-danger">
                                                <i className={`fas fa-circle me-1 ${paused ? "" : "fa-beat"}`} />{paused ? "paused" : "recording"}
                                            </span>
                                        </>
                                    )}
                                    {take && (
                                        <button className="btn btn-sm btn-outline-secondary" disabled={saving}
                                            onClick={() => {
                                                clearTake();
                                                setElapsed(0);
                                                // Stopping the RECORDER never stopped the devices, so the
                                                // stream is usually still live - re-arming here would put the
                                                // screen-share picker in front of someone who only wanted a
                                                // second take.
                                                const live = recordStreamRef.current?.getTracks().some((t) => t.readyState === "live");
                                                if (!live) arm({ source, micId, camId });
                                            }}>
                                            <i className="fas fa-rotate-left me-1" />Record again
                                        </button>
                                    )}
                                    {take && (
                                        <span className="text-muted small">
                                            {fmtBytes(take.blob.size)}
                                            {!isVideo && baseMime(take.mimeType) !== "audio/mp4" && " · saved as WAV"}
                                        </span>
                                    )}
                                    <span className="text-muted small ms-auto">
                                        {recording ? `stops on its own at ${fmtClock(maxSeconds)}` : armed ? "ready" : ""}
                                    </span>
                                </div>

                                {take && (
                                    <div className="mt-2">
                                        <label className="form-label small mb-1">Name</label>
                                        <input className="form-control form-control-sm" value={name} disabled={saving}
                                            onChange={(e) => setName(e.target.value)} />
                                    </div>
                                )}
                            </>
                        )}
                    </div>

                    <div className="modal-footer py-2">
                        <button className="btn btn-sm btn-secondary" onClick={close} disabled={saving}>Cancel</button>
                        <button className="btn btn-sm btn-primary" disabled={!take || saving} onClick={use}>
                            {saving ? <span className="spinner-border spinner-border-sm me-1" /> : <i className="fas fa-check me-1" />}
                            {useLabel || (isVideo ? "Use this clip" : "Use this take")}
                        </button>
                    </div>
                </div>
            </div>
        </div>
    );
}
